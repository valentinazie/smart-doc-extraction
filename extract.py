#!/usr/bin/env python3
"""
Unified entry point for the src/ extraction pipelines.

Given any supported document (or a folder of them), this script picks the
right extraction method based on the file extension and delegates to the
corresponding pipeline in src/:

    .pdf                → pdf_extraction.comprehensive_pdf_analyzer
    .pptx               → ppt_extraction.comprehensive_presentation_analyzer
    .xlsx / .xlsm /.xls → excel_extraction.excel_to_jsonl_pipeline
    .docx               → watsonx_text_extraction.text_extraction

.csv and other formats are not supported and will be skipped.

Usage
    # Single file (type auto-detected)
    python src/extract.py path/to/file.pdf

    # Folder (walks recursively, dispatches per file)
    python src/extract.py --folder path/to/docs

    # Custom output root
    python src/extract.py path/to/file.pptx --output ./my_output

Outputs
    Each underlying pipeline writes into its own subfolder of --output:
        <output>/<stem>_analysis/     (PDF / PPTX: comprehensive_analysis_complete.json
                                       + comprehensive_labeling_with_vlm.md +
                                       comprehensive_summary.json + visual_captures/)
        <output>/<stem>/              (Excel: assembly.md + sheets/ + excel_extracted.jsonl)
        <output>/<stem>/              (DOCX: assembly.md + assets, downloaded from COS)
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path
from typing import Callable, Iterable, List

# Make every subpackage under src/ importable regardless of invocation cwd.
_SRC_DIR = Path(__file__).resolve().parent
for _p in (_SRC_DIR, _SRC_DIR / "ppt_extraction", _SRC_DIR / "pdf_extraction"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))


# ---------------------------------------------------------------------------
# File-type routing
# ---------------------------------------------------------------------------

PDF_EXTS = {".pdf"}
PPTX_EXTS = {".pptx"}
EXCEL_EXTS = {".xlsx", ".xlsm", ".xls"}
DOCX_EXTS = {".docx"}
UNSUPPORTED_EXTS = {".csv", ".ppt", ".doc"}

SUPPORTED_EXTS = PDF_EXTS | PPTX_EXTS | EXCEL_EXTS | DOCX_EXTS


def _extension(path: Path) -> str:
    return path.suffix.lower()


def _kind(path: Path) -> str:
    ext = _extension(path)
    if ext in PDF_EXTS:
        return "pdf"
    if ext in PPTX_EXTS:
        return "pptx"
    if ext in EXCEL_EXTS:
        return "excel"
    if ext in DOCX_EXTS:
        return "docx"
    return "unsupported"


# ---------------------------------------------------------------------------
# Adapters — call each pipeline's `main()` with a crafted sys.argv
# ---------------------------------------------------------------------------

def _run_with_argv(entry: Callable[[], object], argv: List[str]) -> int:
    """Run `entry()` as if `argv` had been passed on the command line.

    Returns the exit code (int). Most pipelines use `sys.exit(code)`, some
    just `return` — we normalize both to 0/1 here.
    """
    original = sys.argv
    sys.argv = argv
    try:
        rc = entry()
        if isinstance(rc, int):
            return rc
        return 0
    except SystemExit as e:
        code = e.code
        return 0 if code in (None, 0) else int(code) if isinstance(code, int) else 1
    finally:
        sys.argv = original


def _run_pdf(path: Path, output: Path, dpi: int) -> int:
    from pdf_extraction import comprehensive_pdf_analyzer as m  # type: ignore
    argv = ["comprehensive_pdf_analyzer.py",
            str(path),
            "--output", str(output),
            "--dpi", str(dpi)]
    return _run_with_argv(m.main, argv)


def _run_pptx(path: Path, output: Path) -> int:
    from ppt_extraction import comprehensive_presentation_analyzer as m  # type: ignore
    argv = ["comprehensive_presentation_analyzer.py",
            str(path),
            "--output", str(output)]
    return _run_with_argv(m.main, argv)


def _run_excel(path: Path, output: Path) -> int:
    from excel_extraction import excel_to_jsonl_pipeline as m  # type: ignore
    # Place the JSONL inside the per-doc folder so it never leaks next to
    # outputs from other pipelines (e.g. PDF/PPTX `<stem>_analysis/` siblings).
    per_doc_dir = output / path.stem
    per_doc_dir.mkdir(parents=True, exist_ok=True)
    jsonl_path = per_doc_dir / "excel_extracted.jsonl"
    argv = ["excel_to_jsonl_pipeline.py",
            str(path),
            "--output-dir", str(output),
            "--jsonl", str(jsonl_path)]
    return _run_with_argv(m.main, argv)


def _run_docx(path: Path, output: Path) -> int:
    # DOCX pipeline pushes results to COS and (when --output is given)
    # downloads them locally under `<output>/<stem>/`.
    from watsonx_text_extraction import text_extraction as m  # type: ignore
    argv = ["text_extraction.py", str(path), "--output", str(output)]
    return _run_with_argv(m.main, argv)


_DISPATCH = {
    "pdf": _run_pdf,
    "pptx": _run_pptx,
    "excel": _run_excel,
    "docx": _run_docx,
}


# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

def _discover(path: Path, recursive: bool = True) -> Iterable[Path]:
    if path.is_file():
        yield path
        return
    if not path.is_dir():
        raise FileNotFoundError(f"Path not found: {path}")
    pattern = "**/*" if recursive else "*"
    for p in sorted(path.glob(pattern)):
        if p.is_file() and _extension(p) in SUPPORTED_EXTS | UNSUPPORTED_EXTS:
            yield p


# ---------------------------------------------------------------------------
# Public API — callable from other Python code
# ---------------------------------------------------------------------------

def extract(path: str | Path,
            output: str | Path = "output",
            *,
            dpi: int = 200) -> int:
    """Extract a single document to `output/`. Returns the exit code."""
    path = Path(path)
    output = Path(output)
    kind = _kind(path)
    if kind == "unsupported":
        print(f"⚠️  Unsupported file type: {path.name}")
        return 2
    runner = _DISPATCH[kind]
    if kind == "pdf":
        return runner(path, output, dpi)
    return runner(path, output)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _build_argparser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Unified dispatcher for src/ extraction pipelines "
                    "(PDF / PPTX / Excel / DOCX).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("path", nargs="?", type=Path,
                   help="Document file, or folder (with --folder)")
    p.add_argument("--folder", "-f", type=Path, default=None,
                   help="Process all supported documents under this folder")
    p.add_argument("--output", "-o", type=Path, default=Path("output"),
                   help="Output root directory (default: ./output)")
    p.add_argument("--dpi", type=int, default=200,
                   help="DPI for PDF page rendering (PDF only, default: 200)")
    p.add_argument("--continue-on-error", action="store_true",
                   help="In folder mode, keep going after a failure")
    return p


def main() -> int:
    args = _build_argparser().parse_args()

    if not args.path and not args.folder:
        print("❌ Provide either a file path or --folder <dir>", file=sys.stderr)
        return 1

    target = args.folder or args.path
    target = Path(target).resolve()
    output = Path(args.output).resolve()
    output.mkdir(parents=True, exist_ok=True)

    files = list(_discover(target, recursive=True))
    if not files:
        print(f"❌ No supported documents found under {target}")
        return 1

    stats = {"success": [], "skipped": [], "failed": []}
    for i, fp in enumerate(files, 1):
        kind = _kind(fp)
        tag = f"[{i}/{len(files)}]"
        print(f"\n{'=' * 70}")
        print(f"{tag} {fp.name}  (→ {kind})")
        print("=" * 70)

        if kind == "unsupported":
            print(f"⏭️  Skipped: unsupported extension {_extension(fp)}")
            stats["skipped"].append(fp.name)
            continue

        t0 = time.time()
        try:
            rc = extract(fp, output=output, dpi=args.dpi)
        except Exception as e:  # noqa: BLE001 — we want to survive any crash
            print(f"❌ Unhandled error: {e}")
            import traceback
            traceback.print_exc()
            rc = 1

        dt = time.time() - t0
        if rc == 0:
            print(f"✅ {tag} done in {dt:.1f}s")
            stats["success"].append(fp.name)
        else:
            print(f"❌ {tag} failed (rc={rc}) in {dt:.1f}s")
            stats["failed"].append(fp.name)
            if not args.continue_on_error:
                break

    print(f"\n{'=' * 70}\n📊 BATCH SUMMARY")
    print(f"   ✅ Success: {len(stats['success'])}/{len(files)}")
    print(f"   ⏭️  Skipped: {len(stats['skipped'])}/{len(files)}")
    print(f"   ❌ Failed:  {len(stats['failed'])}/{len(files)}")
    if stats["failed"]:
        print("   Failed files:")
        for n in stats["failed"]:
            print(f"     - {n}")
    print(f"   📂 Output:  {output}")
    return 0 if not stats["failed"] else 1


if __name__ == "__main__":
    sys.exit(main())
