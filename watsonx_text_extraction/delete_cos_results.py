#!/usr/bin/env python3
"""
Delete text extraction results from COS by name.

Usage:
    # List what would be deleted (dry run — default)
    python delete_cos_results.py "Galaxy"

    # Actually delete (prompts for confirmation)
    python delete_cos_results.py "Galaxy" --delete

    # Skip the confirmation prompt
    python delete_cos_results.py "Galaxy" --delete --yes

    # Match exact key (no substring/case-insensitive matching)
    python delete_cos_results.py "text_extraction_results/Galaxy_S24/assembly.md" --delete --exact

    # Search the whole bucket, not just text_extraction_results/
    python delete_cos_results.py "stale_run_2026" --delete --prefix ""

The NAME argument is a case-insensitive substring match against the COS key
by default — pass --exact to require the key match the name verbatim.
"""

import argparse
import os
import sys
from pathlib import Path

_SRC_ROOT = Path(__file__).resolve().parent.parent
if str(_SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(_SRC_ROOT))

from common.config import load_env, get_space_cos_client  # noqa: E402

# Shared with download_cos_results.py.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from cos_results_utils import (  # noqa: E402,F401
    COS_PREFIX,
    list_all_objects,
    find_matches,
    fmt_size,
)

load_env()


def get_cos_client():
    """Connect to COS using space credentials (same as download_cos_results.py)."""
    return get_space_cos_client()


def delete_objects(cos_client, bucket, objects):
    """Delete via batched delete_objects (max 1000 keys per call)."""
    deleted, errors = 0, []
    for i in range(0, len(objects), 1000):
        batch = objects[i:i + 1000]
        resp = cos_client.delete_objects(
            Bucket=bucket,
            Delete={"Objects": [{"Key": o["Key"]} for o in batch], "Quiet": False},
        )
        deleted += len(resp.get("Deleted", []))
        errors.extend(resp.get("Errors", []))
    return deleted, errors


def main():
    parser = argparse.ArgumentParser(
        description="Delete text extraction results from COS by name.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("name",
                        help="Name to match against COS keys (substring, case-insensitive by default)")
    parser.add_argument("--delete", "-d", action="store_true",
                        help="Actually delete (default: dry run / list only)")
    parser.add_argument("--yes", "-y", action="store_true",
                        help="Skip the y/N confirmation prompt (use with --delete)")
    parser.add_argument("--exact", "-e", action="store_true",
                        help="Require exact key match instead of substring")
    parser.add_argument("--prefix", "-p", default=COS_PREFIX,
                        help=f"COS key prefix to search under (default: {COS_PREFIX!r}; "
                             f"pass '' to search whole bucket)")
    args = parser.parse_args()

    print("Connecting to COS...")
    cos_client, bucket = get_cos_client()

    print(f"Listing objects in {bucket}/{args.prefix or '(whole bucket)'}...")
    objects = list_all_objects(cos_client, bucket, args.prefix)
    if not objects:
        print("No objects found under that prefix.")
        return

    matches = find_matches(objects, args.name, args.exact)
    if not matches:
        match_kind = "exact key" if args.exact else "substring"
        print(f"No objects matching {match_kind} '{args.name}' "
              f"(searched {len(objects)} keys).")
        return

    total_size = sum(o["Size"] for o in matches)
    print(f"\n{'=' * 70}")
    print(f"  MATCHES — {len(matches)} object(s), total {fmt_size(total_size)}")
    print(f"{'=' * 70}")
    for o in matches:
        print(f"  {fmt_size(o['Size']):>10}  {o['Key']}")
    print(f"{'=' * 70}\n")

    if not args.delete:
        print("Dry run only — re-run with --delete to actually remove these objects.")
        return

    if not args.yes:
        confirm = input(f"Delete {len(matches)} object(s) from {bucket}? [y/N] ").strip().lower()
        if confirm not in ("y", "yes"):
            print("Aborted.")
            sys.exit(1)

    print(f"Deleting {len(matches)} object(s)...")
    deleted, errors = delete_objects(cos_client, bucket, matches)
    print(f"  Deleted: {deleted}")
    if errors:
        print(f"  Errors:  {len(errors)}")
        for e in errors[:10]:
            print(f"    - {e.get('Key')}: {e.get('Code')} {e.get('Message')}")
        sys.exit(1)


if __name__ == "__main__":
    main()
