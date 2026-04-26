"""
PDF text structure.

Equivalent of TextExtractionMixin in ppt_extraction. Produces the
`text_structure` dict the reading-order mixin reads, using PyMuPDF text
blocks sorted in standard reading order (top→bottom rows, left→right
within each row).

This is the fallback path. A future revision can call watsonx Text
Extraction V2 against the same PDF for higher-quality OCR + table
structure, in the same way the PPTX side already does.
"""

from pathlib import Path

import fitz  # PyMuPDF


# Two text blocks belong to the same row when their vertical centers are
# within ROW_TOLERANCE * line_height of each other.
ROW_TOLERANCE = 0.6


class PDFTextExtractionMixin:

    def extract_text_structure(self, pdf_path):
        print(f"\n📝 EXTRACTING PDF TEXT STRUCTURE")
        print("=" * 50)

        pdf_path = Path(pdf_path)
        doc = fitz.open(pdf_path)

        text_structure = {
            "file_info": {
                "file_path": str(pdf_path),
                "file_type": "pdf",
                "extraction_method": "native_pymupdf",
                "total_slides": len(doc),
            },
            "slides": [],
        }

        for page_idx, page in enumerate(doc, start=1):
            content = self._extract_page_reading_order(page)
            text_structure["slides"].append({
                "slide_number": page_idx,
                "content": content,
                "title": content[0]["text"] if content else f"Page {page_idx}",
                "notes": "",
                "layout_name": "PDF Page",
            })
            print(f"   📝 Page {page_idx}: {len(content)} text items")

        self.comprehensive_data["text_structure"] = text_structure
        doc.close()
        print("✅ Text structure extracted")
        return True

    @staticmethod
    def _extract_page_reading_order(page):
        """Sort text blocks into row-then-column reading order."""
        blocks = []
        for b in page.get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            text = "\n".join(
                "".join(span["text"] for span in line["spans"])
                for line in b["lines"]
            ).strip()
            if not text:
                continue
            x0, y0, x1, y1 = b["bbox"]
            blocks.append({
                "bbox": (x0, y0, x1, y1),
                "text": text,
                "height": y1 - y0,
            })

        if not blocks:
            return []

        # Group by row, then sort each row left-to-right.
        median_h = sorted(b["height"] for b in blocks)[len(blocks) // 2]
        tol = max(ROW_TOLERANCE * median_h, 4.0)

        blocks.sort(key=lambda b: (b["bbox"][1], b["bbox"][0]))
        rows = []
        for b in blocks:
            placed = False
            cy = (b["bbox"][1] + b["bbox"][3]) / 2
            for row in rows:
                ry = (row[0]["bbox"][1] + row[0]["bbox"][3]) / 2
                if abs(cy - ry) <= tol:
                    row.append(b)
                    placed = True
                    break
            if not placed:
                rows.append([b])

        ordered = []
        ro = 1
        for row in rows:
            row.sort(key=lambda b: b["bbox"][0])
            for b in row:
                x0, y0, x1, y1 = b["bbox"]
                ordered.append({
                    "reading_order": ro,
                    "text": b["text"],
                    "is_title": (ro == 1),
                    "position": {
                        "left": float(x0),
                        "top": float(y0),
                        "width": float(x1 - x0),
                        "height": float(y1 - y0),
                    },
                })
                ro += 1
        return ordered
