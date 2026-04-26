"""
PDF page rendering. Equivalent of ConversionMixin in ppt_extraction, except
the input is already a PDF — we just rasterize each page to PNG so the
visual-capture and VLM stages have something to crop.
"""

from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image


class PDFConversionMixin:
    """Render every page of a PDF to a PNG that downstream stages can crop."""

    def render_pdf_pages(self, pdf_path: Path) -> bool:
        pdf_path = Path(pdf_path)
        doc = fitz.open(pdf_path)

        self.slide_images = {}  # 1-indexed page → PIL.Image, mirrors PPTX naming
        self.page_size_pt = {}  # 1-indexed page → (width_pt, height_pt)

        for page_idx, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=self.render_dpi)
            png_path = self.pdf_pages_dir / f"page_{page_idx:02d}_full.png"
            pix.save(png_path)
            self.slide_images[page_idx] = Image.open(png_path).convert("RGB")
            self.page_size_pt[page_idx] = (page.rect.width, page.rect.height)
            print(f"   📄 Rendered page {page_idx}/{len(doc)} → {png_path.name} "
                  f"({pix.width}x{pix.height} px)")

        doc.close()
        print(f"✅ Rendered {len(self.slide_images)} pages at {self.render_dpi} DPI")
        return True
