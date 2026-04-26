"""PPTX → PDF (→ page images) conversion for the PPT analyzer.

Thin wrappers around `common.libreoffice.convert_to_pdf`. All the previous
duplication (three separate copies of "find soffice, run subprocess, check
returncode") has been replaced by a single call.

Public methods the rest of the mixin stack consumes:

    convert_pptx_to_pdf_images(pptx_path) -> bool
        Convert PPTX → PDF (if needed) and rasterize pages via pdf2image.
        Populates `self.slide_images[i] = PIL.Image` and writes page_*.png.

    convert_pptx_to_pdf_for_watsonx(pptx_path) -> Optional[Path]
        Ensure a PDF exists for watsonx text extraction. Reuses existing
        `<stem>.pdf` if present.
"""

from __future__ import annotations

import shutil
from pathlib import Path

# Add src/ to path for common.* imports when run standalone.
import sys
_SRC = Path(__file__).resolve().parents[2]
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from common.libreoffice import convert_to_pdf  # noqa: E402

try:
    from pdf2image import convert_from_path  # noqa: F401
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False


# PDF-export filter tuned for Korean text + high-res images (previously
# duplicated in four places inside this file).
_HQ_PDF_FILTER = (
    'writer_pdf_Export:{"UseTaggedPDF":false,'
    '"ExportNotesPages":false,'
    '"SelectPdfVersion":1,'
    '"CompressMode":2,'
    '"ImageResolution":600}'
)


class ConversionMixin:
    """PPTX → PDF → page images conversion helpers."""

    # ------------------------------------------------------------------ #
    # Public: PPTX → PDF + rasterize pages for visual capture
    # ------------------------------------------------------------------ #
    def convert_pptx_to_pdf_images(self, pptx_path) -> bool:
        if not PDF2IMAGE_AVAILABLE:
            print("⚠️  pdf2image not available - visual PDF capture disabled")
            self.slide_images = {}
            return False

        print("📄 Converting PowerPoint to PDF images for visual capture...")
        pptx_path = Path(pptx_path)
        pdf_path = pptx_path.with_suffix(".pdf")

        if pdf_path.exists():
            print(f"✅ Found existing PDF file: {pdf_path.name} — reusing")
        else:
            print(f"🔄 Converting {pptx_path.name} → PDF via LibreOffice (high-res)")
            try:
                convert_to_pdf(
                    pptx_path,
                    pptx_path.parent,
                    pdf_filter=_HQ_PDF_FILTER,
                    timeout=60,
                )
            except RuntimeError as e:
                print(f"❌ Failed to convert PPTX to PDF: {e}")
                print("   → Visual captures will be skipped.")
                self.slide_images = {}
                return False

        try:
            from pdf2image import convert_from_path
            pages = convert_from_path(pdf_path, dpi=300, fmt="PNG")
        except Exception as e:
            print(f"❌ Error rendering PDF pages to images: {e}")
            self.slide_images = {}
            return False

        print(f"🖼️  Converted {len(pages)} PDF pages to images")
        self.slide_images = {}
        for i, page_image in enumerate(pages, 1):
            page_filepath = self.pdf_pages_dir / f"page_{i:02d}_full.png"
            page_image.save(page_filepath, "PNG")
            self.slide_images[i] = page_image
            print(f"   💾 Saved page {i}: {page_filepath.name}")

        print("✅ PDF images ready for visual capture")
        return True

    # ------------------------------------------------------------------ #
    # Public: PPTX → PDF for watsonx text extraction
    # ------------------------------------------------------------------ #
    def convert_pptx_to_pdf_for_watsonx(self, pptx_path):
        pptx_path = Path(pptx_path)

        existing_pdf = pptx_path.with_suffix(".pdf")
        if existing_pdf.exists():
            print(f"   ✅ Reusing existing PDF: {existing_pdf.name}")
            return existing_pdf

        cached = pptx_path.parent / f"{pptx_path.stem}_for_watsonx.pdf"
        if cached.exists():
            print(f"   ✅ Reusing watsonx PDF cache: {cached.name}")
            return cached

        print(f"   🔄 Converting {pptx_path.name} → PDF for watsonx")
        try:
            produced = convert_to_pdf(
                pptx_path,
                pptx_path.parent,
                pdf_filter=_HQ_PDF_FILTER,
                timeout=120,
            )
        except RuntimeError as e:
            print(f"   ❌ Conversion failed: {e}")
            return None

        if produced != cached:
            shutil.move(str(produced), str(cached))
        print(f"   ✅ PPTX converted to PDF: {cached}")
        return cached
