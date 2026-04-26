"""
PDF spatial extraction.

Equivalent of SpatialMixin in ppt_extraction. Walks each page's primitives
via PyMuPDF and produces a `boxes` list with the same schema the rest of the
pipeline already understands.

Coordinate system: positions are emitted in **EMUs** (English Metric Units,
914400 / inch, 12700 / pt). PyMuPDF works in PDF points, so we multiply by
EMU_PER_PT here. The reason: the shared mixins inherited from PPTX have
hardcoded absolute thresholds in EMU (e.g. row tolerance = 300000 EMU ≈
23 pt). Storing PDF positions in EMU lets us reuse those mixins unchanged.
The visual-capture mixin reverses the conversion when cropping.
"""

import json
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image


EMU_PER_PT = 12700  # PDF points → English Metric Units


class PDFSpatialMixin:
    """Produce a boxes-per-page spatial map from PyMuPDF primitives."""

    # Shape types we want VLM captures + group-level treatment for.
    PDF_CAPTURE_TARGET_TYPES = {"Picture", "Table"}

    def extract_spatial_analysis(self, pdf_path):
        print(f"\n📦 EXTRACTING PDF SPATIAL ANALYSIS")
        print("=" * 50)

        pdf_path = Path(pdf_path)
        doc = fitz.open(pdf_path)

        spatial_analysis = {"slides": [], "captured_shapes": []}

        # Metadata mirrors what the PPTX pipeline records.
        page_w_pt, page_h_pt = doc[0].rect.width, doc[0].rect.height
        self.comprehensive_data["metadata"] = {
            "file_path": str(pdf_path),
            "file_type": "pdf",
            "total_slides": len(doc),
            # Reuse the PPTX field names so shared mixins read them transparently.
            "slide_width": page_w_pt * EMU_PER_PT,
            "slide_height": page_h_pt * EMU_PER_PT,
            # Keep raw points around in case anything needs them.
            "slide_width_pt": page_w_pt,
            "slide_height_pt": page_h_pt,
            "title": doc.metadata.get("title", "") or pdf_path.stem,
            "author": doc.metadata.get("author", "") or "",
        }

        for page_idx, page in enumerate(doc, start=1):
            print(f"   📦 Processing page {page_idx}/{len(doc)}...")
            slide_data = self._extract_page_boxes(page, page_idx)
            spatial_analysis["slides"].append(slide_data)

        self.comprehensive_data["spatial_analysis"] = spatial_analysis

        # Save mirror of PPTX spatial_analysis.json
        spatial_path = self.spatial_analysis_dir / "spatial_analysis.json"
        with open(spatial_path, "w", encoding="utf-8") as f:
            json.dump(spatial_analysis, f, ensure_ascii=False, indent=2, default=str)

        doc.close()
        print(f"✅ Spatial analysis extracted: {len(spatial_analysis['slides'])} pages")
        return True

    # ---------- per-page extraction ----------

    def _extract_page_boxes(self, page, page_number):
        """Walk one PDF page and produce its boxes list."""
        boxes = []
        next_idx = 0

        # 1. Text blocks → TextBox boxes (skip empty / sliver blocks).
        for block in page.get_text("dict")["blocks"]:
            if block["type"] != 0:
                continue
            text = self._block_text(block)
            if not text.strip():
                continue
            x0, y0, x1, y1 = block["bbox"]
            if (x1 - x0) <= 0 or (y1 - y0) <= 0:
                continue
            font_size = self._dominant_font_size(block)
            boxes.append({
                "box_id": f"S{next_idx}",
                "shape_index": next_idx,
                "shape_type": "TextBox",
                "position": self._bbox_to_position((x0, y0, x1, y1)),
                "text": text.strip(),
                "has_text": True,
                "is_target_for_capture": False,
                "dominant_font_size": font_size,
            })
            next_idx += 1

        # 2. Images with real placement bboxes → Picture boxes.
        for img in page.get_images(full=True):
            try:
                bbox = page.get_image_bbox(img)
            except Exception:
                continue
            x0, y0, x1, y1 = bbox.x0, bbox.y0, bbox.x1, bbox.y1
            if x1 < x0:
                x0, x1 = x1, x0
            if y1 < y0:
                y0, y1 = y1, y0
            if (x1 - x0) < 5 or (y1 - y0) < 5:
                # PyMuPDF returns degenerate bboxes for XObjects that are
                # referenced by the page resource dict but not actually placed
                # in its content stream (slide masters, unused templates, etc).
                # The classic shape is (1,1,-1,-1) → 2pt × 2pt. Anything under
                # ~5pt is also too small to be a real on-page picture.
                continue
            boxes.append({
                "box_id": f"S{next_idx}",
                "shape_index": next_idx,
                "shape_type": "Picture",
                "position": self._bbox_to_position((x0, y0, x1, y1)),
                "text": "",
                "has_text": False,
                "is_target_for_capture": True,
                "xref": img[0],
                "source_pixels": [img[2], img[3]],
            })
            next_idx += 1

        slide_data = {
            "slide_number": page_number,
            "boxes": boxes,
            "total_shapes": len(boxes),
            "spatial_map": {},
            "local_sections": [],
            "line_dividers": [],
        }
        return slide_data

    # ---------- helpers ----------

    @staticmethod
    def _block_text(block):
        return "\n".join(
            "".join(span["text"] for span in line["spans"])
            for line in block["lines"]
        )

    @staticmethod
    def _dominant_font_size(block):
        sizes = []
        for line in block["lines"]:
            for span in line["spans"]:
                sizes.append(span["size"])
        return max(sizes) if sizes else 0.0

    @staticmethod
    def _bbox_to_position(bbox):
        """Convert a PDF-points bbox to an EMU position dict."""
        x0, y0, x1, y1 = bbox
        return {
            "left": float(x0) * EMU_PER_PT,
            "top": float(y0) * EMU_PER_PT,
            "width": float(x1 - x0) * EMU_PER_PT,
            "height": float(y1 - y0) * EMU_PER_PT,
        }
