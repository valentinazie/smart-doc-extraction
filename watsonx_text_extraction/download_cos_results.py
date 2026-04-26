#!/usr/bin/env python3
"""
Download text extraction results from COS.

Usage:
    python download_cos_results.py                     # List all results in COS
    python download_cos_results.py --download          # Download all results
    python download_cos_results.py --download --filter "Galaxy"  # Download matching results only
    python download_cos_results.py --download --output ./my_results  # Custom output directory

    # Batch by local folder: walk the folder, take each file's stem, and
    # download the COS results that belong to those files.
    python download_cos_results.py --download --from-folder client_data/data
"""

import argparse
import os
import sys
from pathlib import Path

_SRC_ROOT = Path(__file__).resolve().parent.parent
if str(_SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(_SRC_ROOT))

from common.config import load_env, get_space_cos_client  # noqa: E402

# Shared with delete_cos_results.py — don't duplicate the paginator.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from cos_results_utils import COS_PREFIX, list_all_objects  # noqa: E402,F401

load_env()

DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "cos_downloads"


def get_cos_client():
    """Connect to COS using space credentials."""
    return get_space_cos_client()


def group_by_document(objects):
    """Group COS objects by their source document name."""
    groups = {}
    for obj in objects:
        key = obj["Key"]
        parts = key.removeprefix(COS_PREFIX).split("/")
        doc_folder = parts[0] if parts else "unknown"
        groups.setdefault(doc_folder, []).append(obj)
    return groups


def print_summary(groups):
    """Print a summary of what's in COS."""
    print(f"\n{'='*70}")
    print(f"  COS Text Extraction Results")
    print(f"{'='*70}")

    total_files = 0
    for doc_name in sorted(groups.keys()):
        files = groups[doc_name]
        total_files += len(files)
        total_size = sum(f["Size"] for f in files)

        md_count = sum(1 for f in files if f["Key"].endswith(".md"))
        json_count = sum(1 for f in files if f["Key"].endswith(".json"))
        img_count = sum(1 for f in files if any(f["Key"].endswith(ext) for ext in [".png", ".jpg", ".jpeg"]))
        other_count = len(files) - md_count - json_count - img_count

        parts = []
        if md_count:
            parts.append(f"{md_count} md")
        if json_count:
            parts.append(f"{json_count} json")
        if img_count:
            parts.append(f"{img_count} img")
        if other_count:
            parts.append(f"{other_count} other")

        size_str = f"{total_size / 1024:.1f} KB" if total_size < 1024 * 1024 else f"{total_size / (1024*1024):.1f} MB"
        print(f"\n  {doc_name}")
        print(f"    {', '.join(parts)} ({size_str})")

    print(f"\n{'='*70}")
    print(f"  Total: {len(groups)} documents, {total_files} files")
    print(f"{'='*70}\n")


def collect_stems_from_folder(folder):
    """Walk a local folder recursively and return the set of file stems.

    Mirrors the extraction scripts' filename normalization: spaces in the
    original filename are converted to underscores before upload, so the COS
    folder uses the underscore form. We do the same here so the stems match.
    """
    stems = set()
    for _root, _dirs, files in os.walk(folder):
        for fname in files:
            if fname.startswith("."):
                continue
            normalized = fname.replace(" ", "_")
            stem = os.path.splitext(normalized)[0]
            if stem:
                stems.add(stem)
    return stems


def filter_objects_by_stems(objects, stems):
    """Keep only objects whose COS doc folder is `<stem>_<timestamp>/...`.

    Match against the stem followed by `_` so that e.g. stem "report" doesn't
    accidentally pull in COS folders for "report_v2_<timestamp>".
    """
    if not stems:
        return []
    prefixes = tuple(s + "_" for s in stems)
    matched = []
    for obj in objects:
        relative = obj["Key"].removeprefix(COS_PREFIX)
        doc_folder = relative.split("/", 1)[0] if relative else ""
        if doc_folder.startswith(prefixes):
            matched.append(obj)
    return matched


def download_results(cos_client, bucket, objects, output_dir, filter_str=None):
    """Download files from COS to local directory."""
    output_dir = Path(output_dir)
    downloaded = 0
    skipped = 0

    for obj in objects:
        key = obj["Key"]
        relative_path = key.removeprefix(COS_PREFIX)

        if filter_str and filter_str.lower() not in relative_path.lower():
            continue

        local_path = output_dir / relative_path
        local_path.parent.mkdir(parents=True, exist_ok=True)

        if local_path.exists() and local_path.stat().st_size == obj["Size"]:
            skipped += 1
            continue

        print(f"  Downloading: {relative_path}")
        cos_client.download_file(bucket, key, str(local_path))
        downloaded += 1

    print(f"\n  Downloaded: {downloaded}, Skipped (already exists): {skipped}")
    print(f"  Output directory: {output_dir}")


def main():
    parser = argparse.ArgumentParser(description="Download text extraction results from COS")
    parser.add_argument("--download", "-d", action="store_true", help="Download files (default is list only)")
    parser.add_argument("--filter", "-f", type=str, default=None, help="Filter results by name (case-insensitive)")
    parser.add_argument("--from-folder", "-F", type=str, default=None,
                        help="Local folder: walk it recursively and download COS results "
                             "matching each file's stem (cannot combine with --filter).")
    parser.add_argument("--output", "-o", type=str, default=str(DEFAULT_OUTPUT_DIR), help="Output directory")
    args = parser.parse_args()

    if args.filter and args.from_folder:
        parser.error("--filter and --from-folder are mutually exclusive.")

    print("Connecting to COS...")
    cos_client, bucket = get_cos_client()

    print(f"Listing objects in {bucket}/{COS_PREFIX}...")
    objects = list_all_objects(cos_client, bucket)

    if not objects:
        print("No extraction results found in COS.")
        return

    groups = group_by_document(objects)

    if args.from_folder:
        folder = Path(args.from_folder).resolve()
        if not folder.is_dir():
            print(f"Not a directory: {folder}")
            return
        stems = collect_stems_from_folder(folder)
        print(f"Folder mode: {folder}")
        print(f"  Collected {len(stems)} file stem(s) to look up in COS")
        objects = filter_objects_by_stems(objects, stems)
        groups = group_by_document(objects)
        if not groups:
            print(f"No COS results match any file under {folder}")
            return
        unmatched = sorted(s for s in stems
                           if not any(g.startswith(s + "_") for g in groups))
        if unmatched:
            print(f"  ⚠️  {len(unmatched)} local file(s) had no matching COS results:")
            for s in unmatched[:10]:
                print(f"     - {s}")
            if len(unmatched) > 10:
                print(f"     ... and {len(unmatched) - 10} more")

    elif args.filter:
        groups = {k: v for k, v in groups.items() if args.filter.lower() in k.lower()}
        objects = [obj for objs in groups.values() for obj in objs]
        if not groups:
            print(f"No results matching '{args.filter}'")
            return

    print_summary(groups)

    if args.download:
        print("Downloading results...")
        download_results(cos_client, bucket, objects, args.output, args.filter)
    else:
        print("Run with --download to download these files.")
        print(f"  Example: python download_cos_results.py --download")
        if len(groups) > 1:
            sample_name = next(iter(groups)).rsplit("_", 1)[0]
            print(f"  Example: python download_cos_results.py --download --filter \"{sample_name}\"")


if __name__ == "__main__":
    main()
