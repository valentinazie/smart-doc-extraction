"""Unified LibreOffice / soffice wrapper.

Before: PPT mixin, excel pipeline, and watsonx_text_extraction pipeline all
had their own copy of "find libreoffice, run --headless --convert-to pdf".
Now every caller goes through `convert_to_pdf()`.

Typical usage:

    from common.libreoffice import convert_to_pdf
    pdf_path = convert_to_pdf(src=Path("report.pptx"), out_dir=Path("./pdfs"))

Returns a `Path` to the produced `<stem>.pdf`. Raises `RuntimeError` on
failure. A PDFExport filter string can be passed via `pdf_filter` to tune
output (image resolution, tagged PDF, etc.) — the PPT pipeline uses this
to render high-quality slides.
"""

from __future__ import annotations

import shutil
import subprocess
from functools import lru_cache
from pathlib import Path
from typing import List, Optional


_BIN_CANDIDATES = (
    "/opt/homebrew/bin/soffice",  # macOS (brew)
    "/usr/local/bin/soffice",     # macOS (Intel brew) / manual install
    "soffice",                    # PATH
    "libreoffice",                # Linux PATH
)


@lru_cache(maxsize=1)
def find_libreoffice() -> Optional[str]:
    """Return the path/name of a usable LibreOffice executable, or None."""
    for candidate in _BIN_CANDIDATES:
        if candidate.startswith("/"):
            if Path(candidate).exists():
                return candidate
        elif shutil.which(candidate):
            return candidate
    return None


def convert_to_pdf(
    src: Path,
    out_dir: Path,
    *,
    timeout: int = 180,
    pdf_filter: Optional[str] = None,
    soffice_cmd: Optional[str] = None,
) -> Path:
    """Convert `src` (pptx / xlsx / xlsm / xls / docx / …) to PDF via
    LibreOffice and return the output path.

    Parameters
    ----------
    src : Path
        Source document.
    out_dir : Path
        Directory to write the PDF into. Will be created if missing.
    timeout : int, default 180
        Seconds before the conversion subprocess is killed.
    pdf_filter : str, optional
        Advanced PDFExport filter (e.g. `"writer_pdf_Export:{...}"`). Mainly
        the PPTX pipeline uses this to pin ImageResolution for visual capture.
        If None, a plain `--convert-to pdf` is used.
    soffice_cmd : str, optional
        Override the executable (otherwise auto-detected).

    Raises
    ------
    RuntimeError
        On missing LibreOffice, non-zero exit, timeout, or missing output.
    """
    src = Path(src)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    cmd_bin = soffice_cmd or find_libreoffice()
    if not cmd_bin:
        raise RuntimeError(
            "LibreOffice not found. Install it (brew install --cask libreoffice) "
            "or pass soffice_cmd=... explicitly."
        )

    convert_to = f"pdf:{pdf_filter}" if pdf_filter else "pdf"
    cmd: List[str] = [
        cmd_bin,
        "--headless",
        "--convert-to", convert_to,
        "--outdir", str(out_dir),
        str(src),
    ]

    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            f"LibreOffice conversion timed out (>{timeout}s) on {src.name}"
        )

    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice failed (rc={result.returncode}) on {src.name}: "
            f"{result.stderr.strip()}"
        )

    pdf_path = out_dir / f"{src.stem}.pdf"
    if not pdf_path.exists():
        raise RuntimeError(
            f"LibreOffice reported success but no PDF produced at {pdf_path}"
        )
    return pdf_path
