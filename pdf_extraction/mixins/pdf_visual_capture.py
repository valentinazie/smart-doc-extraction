"""
PDF visual captures.

For each box marked `is_target_for_capture` (Picture / Table) we crop the
corresponding region out of the rendered page PNG. Filename format mirrors
the PPTX side so the labeling mixin's `find_visual_capture_file` can be
reused unchanged:

    slide_{page:02d}_{shape_type_lower}_{shape_index:02d}_visual.png

Plus the "G{n}" / "S{n}" capture filename pattern for unified groups, which
the smart-grouping mixin emits via capture_unified_group_visual().
"""

from datetime import datetime

from PIL import Image


EMU_PER_PT = 12700  # match PDFSpatialMixin's coordinate convention


class PDFVisualCaptureMixin:

    PADDING_PT = 4.0  # padding in PDF points around each crop

    def capture_all_targets(self):
        print(f"\n📸 CAPTURING TARGET REGIONS")
        print("=" * 50)

        spatial = self.comprehensive_data.get("spatial_analysis", {})
        captures = self.comprehensive_data.setdefault("visual_captures", [])

        for slide_data in spatial.get("slides", []):
            page_num = slide_data["slide_number"]
            page_image = self.slide_images.get(page_num)
            if page_image is None:
                continue

            page_w_pt, page_h_pt = self.page_size_pt[page_num]

            for box in slide_data["boxes"]:
                if not box.get("is_target_for_capture"):
                    continue
                capture = self._crop_box(box, page_image, page_w_pt, page_h_pt, page_num)
                if capture:
                    captures.append(capture)

            n = sum(1 for c in captures if c["slide_number"] == page_num)
            print(f"   📸 Page {page_num}: {n} captures")

        print(f"✅ Visual captures: {len(captures)}")

    def _crop_box(self, box, page_image, page_w_pt, page_h_pt, page_num):
        pos = box["position"]
        # Positions are stored in EMU; convert back to points before scaling
        # to page-image pixels.
        left_pt = pos["left"] / EMU_PER_PT
        top_pt = pos["top"] / EMU_PER_PT
        width_pt = pos["width"] / EMU_PER_PT
        height_pt = pos["height"] / EMU_PER_PT

        sx = page_image.width / page_w_pt
        sy = page_image.height / page_h_pt
        pad = self.PADDING_PT
        left_px = max(0, int((left_pt - pad) * sx))
        top_px = max(0, int((top_pt - pad) * sy))
        right_px = min(page_image.width,
                       int((left_pt + width_pt + pad) * sx))
        bottom_px = min(page_image.height,
                        int((top_pt + height_pt + pad) * sy))

        if right_px <= left_px or bottom_px <= top_px:
            return None

        cropped = page_image.crop((left_px, top_px, right_px, bottom_px))

        shape_type_name = box["shape_type"].lower()
        filename = (f"slide_{page_num:02d}_{shape_type_name}_"
                    f"{box['shape_index']:02d}_visual.png")
        filepath = self.visual_captures_dir / filename
        cropped.save(filepath, "PNG")

        return {
            "slide_number": page_num,
            "box_id": box["box_id"],
            "shape_index": box["shape_index"],
            "shape_type": shape_type_name,
            "filename": filename,
            "filepath": str(filepath),
            "capture_type": "pdf_crop",
            "capture_timestamp": datetime.now().isoformat(),
        }
