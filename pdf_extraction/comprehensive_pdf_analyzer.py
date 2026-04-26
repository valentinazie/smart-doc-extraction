#!/usr/bin/env python3
"""
Comprehensive PDF Analyzer.

Mirror of ppt_extraction's ComprehensivePresentationAnalyzer, but the input
is a PDF. PyMuPDF + pdfplumber replace python-pptx for spatial / table
extraction; everything downstream (smart grouping, reading order, VLM,
labeling) is reused from ppt_extraction unchanged.

Usage:
    source venv/bin/activate
    # Single file
    python pdf_extraction/comprehensive_pdf_analyzer.py <file.pdf>
    # All .pdf files under a folder (recursive) — just pass the folder path
    python pdf_extraction/comprehensive_pdf_analyzer.py <dir>
    # Or use the explicit --folder flag
    python pdf_extraction/comprehensive_pdf_analyzer.py --folder <dir>
    # Custom output root
    python pdf_extraction/comprehensive_pdf_analyzer.py <dir> --output <out_root>

Output: <out_root>/<file_stem>_pdf_analysis/
    - comprehensive_analysis_complete.json
    - comprehensive_summary.json
    - comprehensive_labeling_with_vlm.md
    - visual_captures/
    - pdf_pages/
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

# Make ppt_extraction (which holds the shared mixins) and our own package
# importable when running this file directly.
_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))
sys.path.insert(0, str(_ROOT / "ppt_extraction"))

from common.config import load_env, get_watsonx_credentials
load_env()

# Reused, source-agnostic mixins from the PPTX pipeline.
from ppt_extraction.mixins import (
    SmartGroupingMixin,
    ReadingOrderMixin,
    VLMMixin,
    LabelingMixin,
    VisualizationMixin,
    VisualCaptureMixin as PPTXVisualCaptureMixin,  # for find_visual_capture_file
)

from pdf_extraction.mixins import (
    PDFConversionMixin,
    PDFTextExtractionMixin,
    PDFSpatialMixin,
    PDFTablesMixin,
    PDFVisualCaptureMixin,
)


class ComprehensivePDFAnalyzer(
    PDFConversionMixin,
    PDFTextExtractionMixin,
    PDFSpatialMixin,
    PDFTablesMixin,
    PDFVisualCaptureMixin,
    # Borrow find_visual_capture_file + a couple helpers from the PPTX
    # capture mixin without overriding its should_capture_shape (we don't
    # call that on the PDF side anyway).
    PPTXVisualCaptureMixin,
    SmartGroupingMixin,
    ReadingOrderMixin,
    VLMMixin,
    LabelingMixin,
    VisualizationMixin,
):
    """Run the full text + spatial + grouping + VLM pipeline on a PDF."""

    def __init__(self, output_dir: Path, render_dpi: int = 200):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.render_dpi = render_dpi

        self.streamlined_mode = True
        self.config = {
            "generate_spatial_visualizations": False,
            "generate_smart_group_visualizations": True,
            "generate_hierarchical_flow_visualizations": True,
            "generate_local_sectioning_visualizations": False,
            "generate_enhanced_group_summaries": False,
            "generate_visual_captures": True,
            "generate_reading_order_groups": True,
            "generate_comprehensive_json": True,
        }

        # Mirror of PPTX directory layout.
        self.text_structure_dir = self.output_dir / "text_structure"
        self.spatial_analysis_dir = self.output_dir / "spatial_analysis"
        self.visual_captures_dir = self.output_dir / "visual_captures"
        self.smart_groups_dir = self.output_dir / "smart_groups"
        self.reading_order_groups_dir = self.output_dir / "reading_order_groups"
        self.reading_order_dir = self.output_dir / "reading_order"
        self.pdf_pages_dir = self.output_dir / "pdf_pages"
        self.table_extractions_dir = self.output_dir / "table_extractions"
        self.watsonx_raw_outputs_dir = self.output_dir / "watsonx_raw_outputs"

        for d in [
            self.text_structure_dir, self.spatial_analysis_dir,
            self.visual_captures_dir, self.smart_groups_dir,
            self.reading_order_groups_dir, self.reading_order_dir,
            self.pdf_pages_dir, self.table_extractions_dir,
            self.watsonx_raw_outputs_dir,
        ]:
            d.mkdir(parents=True, exist_ok=True)

        self.comprehensive_data = {
            "metadata": {},
            "text_structure": {},
            "spatial_analysis": {},
            "visual_captures": [],
            "smart_groups": {},
            "reading_order": {},
            "slides": [],
            "extraction_method": "comprehensive_pdf_analysis",
            "timestamp": datetime.now().isoformat(),
        }

        # Shared smart-grouping state (matches PPTX analyzer)
        self.slide_images = {}
        self.page_size_pt = {}
        self.overlap_threshold = 0.3
        self.containment_threshold = 0.5
        self.unified_group_counter = 0
        self.unified_group_members = {}

        # PPTX VisualCaptureMixin's should_capture_shape touches
        # capture_target_types — we never call it for PDFs, but the attribute
        # has to exist to avoid surprise if a shared method consults it.
        self.capture_target_types = set()

        # watsonx setup is needed only for the VLM caption stage. We don't
        # need a Space or COS bucket here — text extraction is local.
        self.watsonx_available = False
        self.credentials = None
        if os.environ.get("WATSONX_APIKEY") and os.environ.get("WATSONX_URL"):
            try:
                self.credentials = get_watsonx_credentials()
                self.watsonx_available = True
                print("✅ watsonx credentials loaded for VLM captioning")
            except Exception as e:
                print(f"⚠️  watsonx credentials setup failed: {e}")
        else:
            print("⚠️  WATSONX_APIKEY / WATSONX_URL missing — VLM captions disabled")

    # ------------------------------------------------------------------
    # Helpers consumed by SmartGroupingMixin (live on the PPT analyzer
    # class itself there, so we mirror them here).
    # ------------------------------------------------------------------

    def _analyze_member_types(self, members):
        type_counts = {}
        for member in members:
            box_type = member['box'].get('box_type', 'unknown')
            type_counts[box_type] = type_counts.get(box_type, 0) + 1
        return type_counts

    def _get_table_extractions_info(self):
        """Reading-order mixin asks for this; for the PDF pipeline tables are
        already inlined into spatial boxes, so just return a minimal stub."""
        return {
            "enabled": False,
            "total_tables": 0,
            "extraction_method": "pdfplumber_inline",
            "table_files": [],
        }

    def create_comprehensive_summary(self):
        """Reading-order integration calls this at the end. Mirror of the PPT
        analyzer's comprehensive_summary.txt writer."""
        summary_path = self.reading_order_dir / "comprehensive_summary.txt"
        ro = self.comprehensive_data.get("reading_order", {})
        with open(summary_path, "w", encoding="utf-8") as f:
            f.write("COMPREHENSIVE PDF ANALYSIS SUMMARY\n")
            f.write("=" * 70 + "\n\n")
            f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            for slide_data in ro.get("slides", []):
                slide_num = slide_data["slide_number"]
                f.write(f"\n{'='*20} PAGE {slide_num}: {slide_data.get('title','')} {'='*20}\n\n")
                f.write(f"Layout: {slide_data.get('layout_name','PDF Page')}\n")
                f.write(f"Reading Order Groups: {len(slide_data.get('reading_order_groups', []))}\n")
                f.write(f"Smart Groups: {len(slide_data.get('smart_groups', {}))}\n\n")
                for ro_group in slide_data.get("reading_order_groups", []):
                    text = ro_group.get("text_content", "")
                    f.write(f"\n{ro_group['reading_order']}. TEXT: {text[:100]}{'...' if len(text) > 100 else ''}\n")
                    f.write(f"   📦 Mapped to Group: {ro_group.get('mapped_spatial_group','')}\n")
        print(f"   📋 Comprehensive summary saved: {summary_path}")

    def has_mixed_content_types(self, group):
        try:
            shape_types = set()
            has_text = False
            has_images = False
            for component in group:
                shape_type = component.get('shape_type', '')
                shape_types.add(shape_type)
                if component.get('has_text', False):
                    has_text = True
                if shape_type in ['Picture', 'Group']:
                    has_images = True
            return ((len(shape_types) >= 2)
                    or (has_text and has_images)
                    or (len(group) >= 3 and has_text))
        except Exception:
            return True

    # ------------------------------------------------------------------
    # PDF-specific apply_smart_grouping: same as the shared mixin's body
    # but inserts a hierarchy-population step before reading-order runs.
    # The PPT spatial mixin computes hierarchical_groups + local_sections
    # alongside boxes; on the PDF side those don't exist yet and the
    # downstream reading-order/VLM code expects them.
    # ------------------------------------------------------------------

    def apply_smart_grouping(self):
        if "spatial_analysis" not in self.comprehensive_data:
            print("❌ No spatial analysis available for smart grouping")
            return False

        print(f"\n🧠 APPLYING SMART GROUPING (PDF)")
        print("=" * 50)

        smart_groups_analysis = {"slides": [], "overall_statistics": {}}
        for slide_data in self.comprehensive_data["spatial_analysis"]["slides"]:
            print(f"   🧠 Processing slide {slide_data['slide_number']}...")
            analysis = self.create_smart_groups_for_slide(
                slide_data["boxes"], slide_data["slide_number"])
            if analysis:
                smart_groups_analysis["slides"].append(analysis)
                if self.config["generate_smart_group_visualizations"]:
                    self.create_enhanced_group_visualization(analysis)

        smart_groups_analysis["overall_statistics"] = {
            "total_slides": len(smart_groups_analysis["slides"]),
            "total_boxes": sum(s.get("total_boxes", 0) for s in smart_groups_analysis["slides"]),
            "total_groups": sum(len(s.get("smart_groups", {})) for s in smart_groups_analysis["slides"]),
            "overlap_threshold": self.overlap_threshold,
            "containment_threshold": self.containment_threshold,
        }
        self.comprehensive_data["smart_groups"] = smart_groups_analysis

        smart_groups_path = self.smart_groups_dir / "smart_groups_analysis.json"
        with open(smart_groups_path, "w", encoding="utf-8") as f:
            json.dump(smart_groups_analysis, f, ensure_ascii=False,
                      indent=2, default=str)

        print(f"✅ Smart grouping applied: "
              f"{smart_groups_analysis['overall_statistics']['total_groups']} groups "
              f"across {smart_groups_analysis['overall_statistics']['total_slides']} slides")

        # Synthesize per-slide hierarchical_groups + local_sections so the
        # shared reading-order/VLM mixins have something to work with.
        self._populate_hierarchy_from_smart_groups()
        self.create_reading_order_based_groups()
        return True

    def _populate_hierarchy_from_smart_groups(self):
        smart = self.comprehensive_data.get("smart_groups", {})
        smart_by_page = {s["slide_number"]: s for s in smart.get("slides", [])}

        for slide_data in self.comprehensive_data["spatial_analysis"]["slides"]:
            slide_smart = smart_by_page.get(slide_data["slide_number"])
            boxes = slide_data.get("boxes", [])

            if slide_smart:
                # Convert each smart_groups entry → the {group_id, members,
                # root_component, total_members} shape that
                # flatten_hierarchical_groups_to_components understands.
                groups = []
                for gid, g in slide_smart.get("smart_groups", {}).items():
                    root_box = g["root_component"]
                    member_boxes = [m["box"] for m in g.get("members", [])]
                    all_members = [root_box] + member_boxes
                    groups.append({
                        "group_id": gid,
                        "members": all_members,
                        "root_component": root_box,
                        "total_members": len(all_members),
                    })
            elif boxes:
                # SmartGroupingMixin returns None for slides with <2 boxes,
                # so e.g. "page is a single full-page table" produces no
                # smart_groups entry. Synthesize a one-group fallback so
                # reading-order/VLM still see the content.
                root_box = boxes[0]
                groups = [{
                    "group_id": root_box["box_id"],
                    "members": list(boxes),
                    "root_component": root_box,
                    "total_members": len(boxes),
                }]
            else:
                slide_data["hierarchical_groups"] = []
                continue

            ordered = self.sort_groups_by_simple_reading_order(groups)
            slide_data["hierarchical_groups"] = ordered

            # Annotates each box in-place with `hierarchical_info`, and
            # returns the flat reading-order list (we keep the original
            # `boxes` order too — it's still indexed S0..Sn).
            self.flatten_hierarchical_groups_to_components(ordered, [])

            # One full-page section. Intentionally **no `shapes` field** —
            # PPT's spatial mixin doesn't set one either. Setting it would
            # exercise an else-branch in create_hierarchical_reading_order_summary
            # that has latent KeyError bugs (it expects `hierarchical_info`
            # keys that the section-component dicts don't carry). Without
            # `shapes`, create_reading_order_from_local_sections returns an
            # empty reading_order_groups and the summary writer falls back to
            # spatial.hierarchical_groups + boxes[*].hierarchical_info, which
            # is what we just populated.
            slide_data["local_sections"] = [{
                "section_id": "main",
                "section_type": "pdf_full_page",
                "bounds": (self.calculate_groups_bounds(ordered)
                           if ordered else {}),
                "groups": ordered,
                "reading_order": 1,
            }]

    # ------------------------------------------------------------------
    # Pipeline orchestration
    # ------------------------------------------------------------------

    def process_pdf(self, pdf_path: Path) -> bool:
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            print(f"❌ File not found: {pdf_path}")
            return False

        print(f"🚀 COMPREHENSIVE PDF ANALYSIS: {pdf_path.name}")
        print("=" * 80)

        # 1. Render pages so visual capture has something to crop.
        print(f"\n🔄 STEP 1/7: PAGE RENDERING")
        print("=" * 40)
        self.render_pdf_pages(pdf_path)

        # 2. Text structure (reading-ordered text blocks).
        print(f"\n📝 STEP 2/7: TEXT STRUCTURE EXTRACTION")
        print("=" * 40)
        self.extract_text_structure(pdf_path)

        # 3. Spatial analysis (boxes per page).
        print(f"\n📦 STEP 3/7: SPATIAL ANALYSIS")
        print("=" * 40)
        self.extract_spatial_analysis(pdf_path)

        # 3b. Tables (consume cell shapes into Table boxes with cell_contents).
        self.extract_tables(pdf_path)

        # 3c. Visual captures (after tables so renumbered indices are stable).
        self.capture_all_targets()

        # 4. Smart grouping (shared mixin).
        print(f"\n🧠 STEP 4/7: SMART GROUPING")
        print("=" * 40)
        self.apply_smart_grouping()

        # 5. VLM captions for each capture.
        print(f"\n🏷️  STEP 5/7: VLM CAPTIONS")
        print("=" * 40)
        try:
            self.create_comprehensive_labeling_output()
        except Exception as e:
            print(f"⚠️  VLM caption stage failed: {e}")

        # 6. Reading-order integration.
        print(f"\n📖 STEP 6/7: READING ORDER INTEGRATION")
        print("=" * 40)
        try:
            self.create_reading_order_integration()
        except Exception as e:
            print(f"⚠️  Reading order integration failed: {e}")

        try:
            if not hasattr(self, "vlm_captions_cache"):
                self.vlm_captions_cache = {}
            self.create_enhanced_reading_order_summary_with_vlm()
        except Exception as e:
            print(f"⚠️  Enhanced reading order summary failed: {e}")

        # 7. Persist results.
        print(f"\n💾 STEP 7/7: SAVE RESULTS")
        print("=" * 40)
        self._save_results()

        # Clean up intermediate dirs (mirror PPTX analyzer). Keep only final
        # outputs: comprehensive_analysis_complete.json + comprehensive_labeling_with_vlm.md
        # + comprehensive_summary.json + visual_captures/.
        for d in [
            self.text_structure_dir,
            self.spatial_analysis_dir,
            self.smart_groups_dir,
            self.reading_order_dir,
            self.reading_order_groups_dir,
            self.table_extractions_dir,
            self.pdf_pages_dir,
            self.watsonx_raw_outputs_dir,
        ]:
            if d.exists():
                shutil.rmtree(d)

        print(f"\n🎉 COMPLETE → {self.output_dir}")
        return True

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------

    def _save_results(self):
        # Strip in-memory PIL images from the comprehensive_data we serialize.
        complete_path = self.output_dir / "comprehensive_analysis_complete.json"
        with open(complete_path, "w", encoding="utf-8") as f:
            json.dump(self.comprehensive_data, f, ensure_ascii=False,
                      indent=2, default=str)

        summary = {
            "file": self.comprehensive_data["metadata"].get("file_path"),
            "pages": self.comprehensive_data["metadata"].get("total_slides"),
            "boxes_per_page": [
                {"page": s["slide_number"], "boxes": s["total_shapes"]}
                for s in self.comprehensive_data["spatial_analysis"]["slides"]
            ],
            "visual_captures": len(self.comprehensive_data.get("visual_captures", [])),
            "smart_groups_total": sum(
                len(s.get("smart_groups", {}))
                for s in self.comprehensive_data.get("smart_groups", {}).get("slides", [])
            ),
        }
        with open(self.output_dir / "comprehensive_summary.json", "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)


def _run_one(pdf_path: Path, output_root: Path, dpi: int) -> bool:
    out_dir = output_root / f"{pdf_path.stem}_pdf_analysis"
    analyzer = ComprehensivePDFAnalyzer(out_dir, render_dpi=dpi)
    try:
        return analyzer.process_pdf(pdf_path)
    finally:
        import gc
        del analyzer
        gc.collect()


def main():
    ap = argparse.ArgumentParser(
        description="Comprehensive PDF Analyzer",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  Single file:
    python pdf_extraction/comprehensive_pdf_analyzer.py file.pdf
  Folder (all .pdf inside, recursively) — just pass the folder path:
    python pdf_extraction/comprehensive_pdf_analyzer.py /path/to/folder
  Or use the explicit --folder flag:
    python pdf_extraction/comprehensive_pdf_analyzer.py --folder /path/to/folder
  With custom output root:
    python pdf_extraction/comprehensive_pdf_analyzer.py /path/to/folder --output ./results
""",
    )
    ap.add_argument("path", nargs="?", type=Path,
                    help="A .pdf file OR a folder containing .pdf files (recursive)")
    ap.add_argument("--folder", "-f", type=Path,
                    help="Folder path to process all .pdf files recursively "
                         "(equivalent to passing a folder as the positional arg)")
    ap.add_argument("--output", "-o", type=Path, default=Path("output"),
                    help="Output root directory (default: ./output)")
    ap.add_argument("--dpi", type=int, default=200)
    args = ap.parse_args()

    # Resolve input: --folder takes priority, otherwise auto-detect from the
    # positional arg (file vs directory).
    folder: Path | None = None
    single_file: Path | None = None

    if args.folder:
        folder = args.folder
    elif args.path:
        if args.path.is_dir():
            folder = args.path
        elif args.path.is_file() or args.path.suffix.lower() == ".pdf":
            single_file = args.path
        else:
            print(f"❌ Path not found: {args.path}")
            sys.exit(1)
    else:
        ap.print_help()
        sys.exit(1)

    if folder is not None:
        folder = folder.resolve()
        if not folder.is_dir():
            print(f"❌ Folder not found: {folder}")
            sys.exit(1)

        pdf_files = sorted(folder.rglob("*.pdf"))
        if not pdf_files:
            print(f"❌ No .pdf files found in {folder}")
            sys.exit(1)

        print(f"📁 BATCH FOLDER MODE")
        print(f"   📂 Source: {folder}")
        print(f"   📄 Found {len(pdf_files)} .pdf files")
        print("=" * 70)

        results = {"success": [], "failed": []}
        for idx, pdf_path in enumerate(pdf_files, 1):
            print(f"\n{'='*70}")
            print(f"📄 [{idx}/{len(pdf_files)}] {pdf_path.name}")
            print(f"{'='*70}")
            try:
                ok = _run_one(pdf_path, args.output, args.dpi)
                if ok:
                    print(f"✅ [{idx}/{len(pdf_files)}] {pdf_path.name} - Done")
                    results["success"].append(pdf_path.name)
                else:
                    print(f"❌ [{idx}/{len(pdf_files)}] {pdf_path.name} - Failed")
                    results["failed"].append(pdf_path.name)
            except Exception as e:
                print(f"❌ [{idx}/{len(pdf_files)}] {pdf_path.name} - Error: {e}")
                import traceback
                traceback.print_exc()
                results["failed"].append(pdf_path.name)

        print(f"\n{'='*70}")
        print(f"📊 BATCH SUMMARY")
        print(f"   ✅ Success: {len(results['success'])}/{len(pdf_files)}")
        print(f"   ❌ Failed:  {len(results['failed'])}/{len(pdf_files)}")
        if results["failed"]:
            print(f"   Failed files:")
            for f in results["failed"]:
                print(f"     - {f}")
        print(f"   📂 Output: {args.output.resolve()}")
        sys.exit(0 if not results["failed"] else 1)

    # Single-file mode
    if not single_file.exists():
        print(f"❌ File not found: {single_file}")
        sys.exit(1)
    ok = _run_one(single_file, args.output, args.dpi)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
