"""
PDF table extraction.

Uses pdfplumber.find_tables() to detect tables (ruled or text-aligned), then:
  1. Removes the text/picture boxes consumed by each table from the page boxes list.
  2. Builds a Table box with cell_contents in the same schema the PPTX
     pipeline produces, so the labeling/reading-order mixins read it identically.

For each cell:
  - text comes from pdfplumber's extracted cell text;
  - any boxes that overlap the cell ≥50% (Picture) or ≥30% (TextBox) get
    listed in `overlapping_shapes` (mirrors the PPTX behaviour after we
    fixed the dryer-motor case).
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pdfplumber


EMU_PER_PT = 12700  # match PDFSpatialMixin's coordinate convention


class PDFTablesMixin:
    """Add tables to the spatial analysis and consume their cell contents."""

    PICTURE_OVERLAP_THRESHOLD = 50.0
    TEXTBOX_OVERLAP_THRESHOLD = 30.0

    # pdfplumber's default strategy infers tables from text alignment, which
    # routinely "finds" a giant phantom NxN table that swallows the whole page.
    # Restrict to actually-ruled tables so we only consume real ones.
    TABLE_SETTINGS = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
    }

    # Defensive caps on what counts as a real table after the line-strategy
    # filter. A ruled table that covers > MAX_PAGE_AREA_FRAC of the page is
    # almost certainly the page border being read as a table.
    MAX_PAGE_AREA_FRAC = 0.85
    MIN_NON_EMPTY_CELL_FRAC = 0.20

    def extract_tables(self, pdf_path):
        print(f"\n📋 EXTRACTING PDF TABLES")
        print("=" * 50)

        spatial = self.comprehensive_data.get("spatial_analysis", {})
        if not spatial.get("slides"):
            print("❌ No spatial analysis available — run extract_spatial_analysis first")
            return False

        with pdfplumber.open(pdf_path) as pp:
            for slide_data in spatial["slides"]:
                page_num = slide_data["slide_number"]
                pp_page = pp.pages[page_num - 1]
                self._process_page_tables(slide_data, pp_page)

        print(f"✅ Tables processed across {len(spatial['slides'])} pages")
        return True

    # ---------- per-page work ----------

    def _process_page_tables(self, slide_data, pp_page):
        page_num = slide_data["slide_number"]
        boxes = slide_data["boxes"]

        tables = pp_page.find_tables(table_settings=self.TABLE_SETTINGS)
        page_area = pp_page.width * pp_page.height
        tables = [t for t in tables if self._table_passes_sanity(t, page_area, page_num)]

        if not tables:
            print(f"   📋 Page {page_num}: no tables")
            return

        # Process tables in reverse-area order so smaller tables nested inside
        # larger ones don't have their boxes consumed twice.
        tables_sorted = sorted(
            tables,
            key=lambda t: (t.bbox[2] - t.bbox[0]) * (t.bbox[3] - t.bbox[1]),
            reverse=True,
        )

        consumed_box_ids = set()

        for t_idx, table in enumerate(tables_sorted):
            table_box = self._build_table_box(table, t_idx, boxes, consumed_box_ids)
            if table_box is None:
                continue

            # Mark which existing boxes the table swallowed.
            for cell in table_box["cell_contents"]:
                for shape in cell.get("overlapping_shapes", []):
                    consumed_box_ids.add(shape["box_id"])

            print(
                f"   📋 Page {page_num} T{t_idx}: "
                f"{table_box['table_rows']}x{table_box['table_cols']} table, "
                f"consumed {len(consumed_box_ids)} boxes"
            )

        if not consumed_box_ids:
            return

        # Drop consumed boxes (their content now lives inside table cells) and
        # append the new Table boxes. Re-number box_ids contiguously so the
        # downstream "S0, S1, S2..." invariant is preserved.
        kept = [b for b in boxes if b["box_id"] not in consumed_box_ids]
        # Table boxes were collected on the side — fetch them via slide_data.
        for tb in slide_data.get("_pending_table_boxes", []):
            kept.append(tb)

        # Renumber so each page is still S0..S(n-1).
        for new_idx, box in enumerate(kept):
            box["shape_index"] = new_idx
            box["box_id"] = f"S{new_idx}"
            # Cell-contents references are keyed by the OLD ids, but those
            # old shapes are gone now anyway — the cell now owns the text.

        slide_data["boxes"] = kept
        slide_data["_pending_table_boxes"] = []
        slide_data["total_shapes"] = len(kept)

    @classmethod
    def _table_passes_sanity(cls, table, page_area, page_num):
        x0, y0, x1, y1 = table.bbox
        area = (x1 - x0) * (y1 - y0)
        if area / page_area > cls.MAX_PAGE_AREA_FRAC:
            print(f"   📋 Page {page_num}: skipping table covering "
                  f"{100*area/page_area:.0f}% of page (likely page border)")
            return False
        rows = table.extract()
        if not rows:
            return False
        total = sum(len(r) for r in rows)
        non_empty = sum(1 for r in rows for c in r if c and c.strip())
        if total and non_empty / total < cls.MIN_NON_EMPTY_CELL_FRAC:
            print(f"   📋 Page {page_num}: skipping {len(rows)}x{max(len(r) for r in rows)} "
                  f"table — only {non_empty}/{total} cells filled")
            return False
        return True

    # ---------- per-table work ----------

    def _build_table_box(self, table, t_idx, boxes, already_consumed):
        rows = table.extract()
        if not rows:
            return None
        n_rows = len(rows)
        n_cols = max(len(r) for r in rows)

        x0, y0, x1, y1 = table.bbox
        table_position = {
            "left": float(x0) * EMU_PER_PT,
            "top": float(y0) * EMU_PER_PT,
            "width": float(x1 - x0) * EMU_PER_PT,
            "height": float(y1 - y0) * EMU_PER_PT,
        }

        cell_contents = []
        cell_rects = self._extract_cell_rects(table, n_rows, n_cols)

        for r in range(n_rows):
            for c in range(n_cols):
                if c >= len(rows[r]):
                    continue
                cell_text = (rows[r][c] or "").strip()
                cell_rect = cell_rects.get((r, c))
                cell_pos = (
                    self._rect_to_position(cell_rect) if cell_rect else None
                )

                overlapping = []
                if cell_pos:
                    overlapping = self._find_overlapping_boxes(
                        cell_pos, boxes, already_consumed
                    )

                has_visual = bool(overlapping)
                cell_contents.append({
                    "row": r,
                    "col": c,
                    "text": cell_text,
                    "display_text": cell_text,
                    "position": cell_pos,
                    "shapes": [
                        {
                            "box_id": s["box_id"],
                            "type": s["shape_type"],
                            "text": s.get("text", ""),
                            "overlap_percentage": s["overlap_percentage"] / 100.0,
                            "visual_capture": "",
                        }
                        for s in overlapping
                    ],
                    "overlapping_shapes": overlapping,
                    "has_visual_content": has_visual,
                    "has_content": bool(cell_text) or has_visual,
                })

        joined_text = " | ".join(
            " | ".join((cell or "").strip() for cell in row) for row in rows
        )

        # Find a slide_data we belong to and stash the table box on it. The
        # caller (process_page_tables) extracts these after walking all tables.
        # We rely on the position to attach later — see _process_page_tables.
        table_box = {
            "box_id": "PENDING",  # filled in by _process_page_tables renumber
            "shape_index": -1,
            "shape_type": "Table",
            "box_type": "table",
            "position": table_position,
            "text": joined_text,
            "has_text": bool(joined_text.strip()),
            "is_target_for_capture": True,
            "table_rows": n_rows,
            "table_cols": n_cols,
            "table_dimensions": f"{n_rows}x{n_cols}",
            "cell_contents": cell_contents,
        }

        # Stash on the spatial slide via the boxes list's parent. The caller
        # owns the slide_data and reads back via _pending_table_boxes.
        # (We piggy-back here since boxes is a list reference — we attach to
        #  whichever slide_data it belongs to via attribute on the list.)
        # Simpler: the caller pre-set slide_data['_pending_table_boxes'] = [];
        # find that list by walking spatial_analysis on the analyzer.
        for sd in self.comprehensive_data["spatial_analysis"]["slides"]:
            if sd["boxes"] is boxes:
                sd.setdefault("_pending_table_boxes", []).append(table_box)
                break

        return table_box

    # ---------- geometry helpers ----------

    @staticmethod
    def _extract_cell_rects(table, n_rows, n_cols):
        """Return {(row, col): (x0, y0, x1, y1)} from pdfplumber's cell list."""
        rects = {}
        cells = table.cells  # list[(x0, y0, x1, y1) | None] in row-major order
        for idx, cell in enumerate(cells):
            if not cell:
                continue
            r = idx // n_cols
            c = idx % n_cols
            rects[(r, c)] = cell
        return rects

    @staticmethod
    def _rect_to_position(rect):
        x0, y0, x1, y1 = rect
        return {
            "left": float(x0) * EMU_PER_PT,
            "top": float(y0) * EMU_PER_PT,
            "width": float(x1 - x0) * EMU_PER_PT,
            "height": float(y1 - y0) * EMU_PER_PT,
        }

    @classmethod
    def _find_overlapping_boxes(cls, cell_pos, boxes: Iterable[dict], already_consumed):
        """Return boxes that overlap the given cell ≥ type-specific threshold."""
        overlaps = []
        for box in boxes:
            if box["box_id"] in already_consumed:
                continue
            shape_type = box.get("shape_type")
            threshold = (
                cls.PICTURE_OVERLAP_THRESHOLD
                if shape_type == "Picture"
                else cls.TEXTBOX_OVERLAP_THRESHOLD
            )
            pct = cls._overlap_percentage(box["position"], cell_pos)
            if pct >= threshold:
                overlaps.append({
                    "box_id": box["box_id"],
                    "shape_type": shape_type,
                    "text": box.get("text", ""),
                    "overlap_percentage": pct,
                    "visual_capture": "",
                })
        return overlaps

    @staticmethod
    def _overlap_percentage(box_pos, cell_pos):
        """% of the box's area that lies inside the cell."""
        bx0, by0 = box_pos["left"], box_pos["top"]
        bx1, by1 = bx0 + box_pos["width"], by0 + box_pos["height"]
        cx0, cy0 = cell_pos["left"], cell_pos["top"]
        cx1, cy1 = cx0 + cell_pos["width"], cy0 + cell_pos["height"]
        ix0, iy0 = max(bx0, cx0), max(by0, cy0)
        ix1, iy1 = min(bx1, cx1), min(by1, cy1)
        if ix1 <= ix0 or iy1 <= iy0:
            return 0.0
        inter = (ix1 - ix0) * (iy1 - iy0)
        box_area = (bx1 - bx0) * (by1 - by0)
        if box_area <= 0:
            return 0.0
        return 100.0 * inter / box_area
