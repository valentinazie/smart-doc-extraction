#!/usr/bin/env python3
"""
Comprehensive Presentation Analyzer.
Combines: Text structure + Spatial analysis + Visual captures + Smart grouping +
Reading order + VLM image captioning (Mistral Small 3.1 24B via watsonx).

Usage:
    python ppt_extraction/comprehensive_presentation_analyzer.py <file.pptx>
    python ppt_extraction/comprehensive_presentation_analyzer.py --folder <dir>
    python ppt_extraction/comprehensive_presentation_analyzer.py --folder <dir> --output <out_root>

Examples (run from repo root, with the venv active):
    source venv/bin/activate

    # Single PPTX
    python ppt_extraction/comprehensive_presentation_analyzer.py \\
        "client_data/data/F250904-5269_240828_VR8500배터리_연소_pptx.pptx"

    # All .pptx files under a folder (recursive). Default output root is ./output
    python ppt_extraction/comprehensive_presentation_analyzer.py \\
        --folder client_data/data

    # Batch with a custom output root
    python ppt_extraction/comprehensive_presentation_analyzer.py \\
        --folder client_data/data --output ppt_extraction/output

Output (per file): <output_root>/<file_stem>_analysis/
    - comprehensive_analysis_complete.json     (full structured analysis)
    - comprehensive_summary.json               (high-level summary)
    - comprehensive_labeling_with_vlm.md       (human-readable + VLM captions)
    - visual_captures/                         (rendered shape/group images)
Intermediate dirs (text_structure/, spatial_analysis/, smart_groups/, reading_order*/,
table_extractions/, watsonx_raw_outputs/, pdf_pages/) are deleted at the end of a
successful run. Comment out the shutil.rmtree loop in process_presentation to keep them.

Requirements:
    - LibreOffice on PATH (used to convert .pptx -> .pdf). If <stem>.pdf already
      exists alongside the .pptx, it is reused and LibreOffice is not invoked.
    - Credentials in .env (loaded from ./.env then ../text_extraction/.env if present):
        WATSONX_APIKEY, WATSONX_URL, SPACE_ID, COS_BUCKET_NAME
    - For VLM captions to work, also set WATSONX_PROJECT_ID (project-scoped).
      Without it, every image renders "❌ VLM Caption Failed" in the .md output —
      the rest of the pipeline (text, spatial, grouping, reading order) still runs.
"""

import os
import sys
import json
import time
from pathlib import Path
import base64

# Bootstrap sys.path so `from mixins import ...` and `from utils.geometry import ...`
# resolve whether this file is run directly, from src/, or from the repo root.
_THIS_DIR = Path(__file__).resolve().parent          # .../src/ppt_extraction
_SRC_DIR = _THIS_DIR.parent                           # .../src
for _p in (_SRC_DIR, _THIS_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from datetime import datetime
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageOps
import shutil
import subprocess

# Safe imports with error handling
try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    PPTX_AVAILABLE = True
    print("✅ python-pptx available")
except ImportError:
    PPTX_AVAILABLE = False
    print("⚠️  python-pptx not available. Install with: pip install python-pptx")

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
    print("✅ pdf2image available")
except ImportError:
    PDF2IMAGE_AVAILABLE = False
    print("⚠️  pdf2image not available. Install with: pip install pdf2image pillow")

from common.config import load_env, get_watsonx_credentials
load_env()

# watsonx imports
try:
    from ibm_watsonx_ai import APIClient, Credentials
    from ibm_watsonx_ai.foundation_models.extractions import TextExtractionsV2, TextExtractionsV2ResultFormats
    from ibm_watsonx_ai.helpers import DataConnection, S3Location
    from ibm_watsonx_ai.metanames import TextExtractionsV2ParametersMetaNames
    WATSONX_AVAILABLE = True
    print("✅ watsonx libraries available")
except ImportError as e:
    WATSONX_AVAILABLE = False  
    print(f"⚠️  watsonx libraries not available: {e}")
    print("Install with: pip install ibm-watsonx-ai ibm-cloud-sdk-core ibm-cos-sdk")


def get_credentials():
    return get_watsonx_credentials()


from ppt_extraction.mixins import (
    ConversionMixin,
    TextExtractionMixin,
    SpatialMixin,
    TablesMixin,
    ReadingOrderMixin,
    SmartGroupingMixin,
    VisualCaptureMixin,
    VLMMixin,
    VisualizationMixin,
    LabelingMixin,
)

class ComprehensivePresentationAnalyzer(
    ConversionMixin,
    TextExtractionMixin,
    SpatialMixin,
    TablesMixin,
    ReadingOrderMixin,
    SmartGroupingMixin,
    VisualCaptureMixin,
    VLMMixin,
    VisualizationMixin,
    LabelingMixin,
):
    """Complete presentation analysis: Text + Spatial + Visual + Grouping + Reading Order"""

    def __init__(self, output_dir="./output/HKMG_test/comprehensive_analysis_watsonx", streamlined_mode=True):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Streamlined mode configuration - ONLY generate essential outputs + essential maps
        self.streamlined_mode = streamlined_mode
        self.config = {
            'generate_spatial_visualizations': not streamlined_mode,
            'generate_smart_group_visualizations': True,  # ALWAYS - Essential maps for user
            'generate_hierarchical_flow_visualizations': True,  # ALWAYS - Essential maps for user
            'generate_local_sectioning_visualizations': not streamlined_mode,
            'generate_enhanced_group_summaries': not streamlined_mode,
            'generate_visual_captures': True,  # ALWAYS - User wants captured images
            'generate_reading_order_groups': True,  # ALWAYS needed
            'generate_comprehensive_json': True,    # ALWAYS needed
        }
        
        if streamlined_mode:
            print("🚀 STREAMLINED MODE: Essential outputs (reading_order_groups + comprehensive_analysis_complete.json + essential maps)")
        
        # Create comprehensive subdirectories (only create needed directories)
        self.text_structure_dir = self.output_dir / "text_structure"
        self.spatial_analysis_dir = self.output_dir / "spatial_analysis"
        self.visual_captures_dir = self.output_dir / "visual_captures"
        self.smart_groups_dir = self.output_dir / "smart_groups"
        self.reading_order_groups_dir = self.output_dir / "reading_order_groups"  # ALWAYS needed
        self.reading_order_dir = self.output_dir / "reading_order"
        self.pdf_pages_dir = self.output_dir / "pdf_pages"
        self.table_extractions_dir = self.output_dir / "table_extractions"  # NEW: watsonx table extractions
        self.watsonx_raw_outputs_dir = self.output_dir / "watsonx_raw_outputs"  # NEW: raw watsonx markdown & assembly
        
        required_dirs = [
            self.reading_order_groups_dir,
            self.text_structure_dir,
            self.spatial_analysis_dir,
            self.smart_groups_dir,
            self.reading_order_dir,
            self.table_extractions_dir,
            self.watsonx_raw_outputs_dir,
        ]
        optional_dirs = []
        
        if self.config['generate_visual_captures']:
            optional_dirs.append(self.visual_captures_dir)
            optional_dirs.append(self.pdf_pages_dir)
        
        for dir_path in required_dirs + optional_dirs:
            dir_path.mkdir(parents=True, exist_ok=True)
        
        # Comprehensive data structure
        self.comprehensive_data = {
            'metadata': {},
            'text_structure': {},
            'spatial_analysis': {},
            'visual_captures': [],
            'smart_groups': {},
            'reading_order': {},
            'slides': [],
            'extraction_method': 'comprehensive_unified_analysis',
            'timestamp': datetime.now().isoformat()
        }
        
        # Store slide images for cropping
        self.slide_images = {}
        
        # Target shape types for visual capture
        self.capture_target_types = [
            MSO_SHAPE_TYPE.PICTURE,
            MSO_SHAPE_TYPE.TABLE,                    # Tables (requested back for captures + VLM)
            MSO_SHAPE_TYPE.GROUP,
            MSO_SHAPE_TYPE.CHART,                    # Charts (requested)
            MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT       # OLE Objects (requested)
            # REMOVED: MSO_SHAPE_TYPE.AUTO_SHAPE     # no AutoShape captures
            # REMOVED: MSO_SHAPE_TYPE.PLACEHOLDER    # no Placeholder captures
        ]
        
        # Smart grouping parameters
        self.overlap_threshold = 0.3  # 30% overlap = components interfere
        self.containment_threshold = 0.5  # 50% contained = assign to group
        
        # Incremental counters for unified groups
        # G = Unified Groups (overlapping images + 80% overlapping text → 1 entity)
        self.unified_group_counter = 0  # For G1, G2, G3...
        
        # Setup watsonx if available
        self.watsonx_available = False
        self.watsonx_client = None
        self.credentials = None
        # SPACE_ID / COS_BUCKET_NAME must be configured in .env when watsonx
        # is wired up. If missing, watsonx setup is skipped and the analyzer
        # falls back to native python-pptx text extraction.
        self.space_id = os.environ.get("SPACE_ID")
        self.cos_bucket_name = os.environ.get("COS_BUCKET_NAME")

        if WATSONX_AVAILABLE:
            try:
                self.credentials = get_credentials()
                self.watsonx_client = APIClient(credentials=self.credentials, space_id=self.space_id)
                self.setup_watsonx_cos()
                self.watsonx_available = True
                print("✅ watsonx Text Extraction V2 client initialized")
            except Exception as e:
                print(f"⚠️  watsonx setup failed: {e}")
                self.watsonx_available = False
        else:
            print("⚠️  watsonx not available - falling back to native python-pptx text extraction")
    
    def reset_consolidation_counters(self):
        """Reset consolidation counters for new presentation processing"""
        self.unified_group_counter = 0
        self.unified_group_members = {}  # Reset image-to-group tracking
        print("🔄 Reset unified group counter for new presentation")

    def setup_watsonx_cos(self):
        """Setup COS connection for watsonx text extraction."""
        try:
            print("☁️  Setting up COS connection...")
            from common.config import get_space_cos_client
            self.cos_client, _ = get_space_cos_client(self.watsonx_client)

            cos_credentials = self.watsonx_client.spaces.get_details(
                space_id=self.space_id)["entity"]["storage"]["properties"]
            connection_details = self.watsonx_client.connections.create({
                "datasource_type": self.watsonx_client.connections.get_datasource_type_uid_by_name("bluemixcloudobjectstorage"),
                "name": "Comprehensive Analyzer COS Connection",
                "properties": {
                    "bucket": self.cos_bucket_name,
                    "access_key": cos_credentials["credentials"]["editor"]["access_key_id"],
                    "secret_key": cos_credentials["credentials"]["editor"]["secret_access_key"],
                    "iam_url": self.watsonx_client.service_instance._href_definitions.get_iam_token_url(),
                    "url": cos_credentials["endpoint_url"],
                },
            })
            self.cos_connection_id = self.watsonx_client.connections.get_id(connection_details)
            print(f"✅ COS Connection ID: {self.cos_connection_id}")
            print("🔧 watsonx COS connection verified")
        except Exception as e:
            print(f"⚠️  COS connection issue: {e}")
            raise


    def _analyze_member_types(self, members):
        """Analyze member types"""
        type_counts = {}
        for member in members:
            box_type = member['box'].get('box_type', 'unknown')
            type_counts[box_type] = type_counts.get(box_type, 0) + 1
        return type_counts
    

    def has_mixed_content_types(self, group):
        """Check if a group has mixed content types (text, images, etc.) to be considered table-like"""
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
            
            # Consider it table-like if it has:
            # 1. Mixed content types (at least 2 different types)
            # 2. Both text and images, OR
            # 3. Multiple text components in rectangular arrangement
            return (len(shape_types) >= 2) or (has_text and has_images) or (len(group) >= 3 and has_text)
            
        except Exception as e:
            print(f"⚠️  Error checking mixed content types: {e}")
            return True  # Default to allowing consolidation


    def create_comprehensive_summary(self):
        """Create comprehensive summary with reading order"""
        print(f"   📋 Creating comprehensive summary...")
        
        summary_path = self.reading_order_dir / "comprehensive_summary.txt"
        
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write("COMPREHENSIVE PRESENTATION ANALYSIS SUMMARY\n")
            f.write("=" * 70 + "\n\n")
            f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Analysis Type: Comprehensive (Text + Spatial + Visual + Grouping + Reading Order)\n\n")
            
            # Overall statistics
            f.write("ANALYSIS COMPONENTS:\n")
            f.write("• 📝 Text Structure: Hierarchical content extraction in reading order\n")
            f.write("• 📦 Spatial Analysis: Precise box positions and shapes\n")
            f.write("• 📸 Visual Captures: Both PDF visual crops and matplotlib shape captures\n")
            f.write("• 🧠 Smart Grouping: Non-overlapping groups with best-fit assignment\n")
            f.write("• 📖 Reading Order: Text-spatial integration for content flow\n\n")
            
            # Process each slide
            for slide_data in self.comprehensive_data['reading_order']['slides']:
                slide_num = slide_data['slide_number']
                f.write(f"\n{'='*20} SLIDE {slide_num}: {slide_data['title']} {'='*20}\n\n")
                f.write(f"Layout: {slide_data['layout_name']}\n")
                f.write(f"Reading Order Groups: {len(slide_data['reading_order_groups'])}\n")
                f.write(f"Smart Groups: {len(slide_data['smart_groups'])}\n\n")
                # Write reading order groups
                f.write("CONTENT IN READING ORDER:\n")
                f.write("-" * 40 + "\n")
                for ro_group in slide_data['reading_order_groups']:
                    f.write(f"\n{ro_group['reading_order']}. TEXT: {ro_group['text_content'][:100]}{'...' if len(ro_group['text_content']) > 100 else ''}\n")
                    f.write(f"   📦 Mapped to Group: {ro_group['mapped_spatial_group']}\n")
                    
                    if ro_group['spatial_components']:
                        f.write(f"   🔍 Spatial Components:\n")
                        for comp in ro_group['spatial_components']:
                            if comp['component_type'] == 'root':
                                f.write(f"      🏠 ROOT {comp['box_id']} ({comp['box_type']})")
                            else:
                                f.write(f"      🔗 MEMBER {comp['box_id']} ({comp['box_type']}, {comp['containment_percentage']:.0%})")
                            
                            if comp['has_text'] and comp['text']:
                                preview = comp['text'][:50] + '...' if len(comp['text']) > 50 else comp['text']
                                f.write(f": {preview}")
                            f.write("\n")
                # Notes
                if slide_data['notes']:
                    f.write(f"\n📝 NOTES:\n{slide_data['notes']}\n")
        
        print(f"   📋 Comprehensive summary saved: {summary_path}")
    

    def convert_position_format(self, obj):
        """Convert position coordinates from (top, left, width, height) to (top, left, bottom, right)"""
        if isinstance(obj, dict):
            if 'position' in obj and isinstance(obj['position'], dict):
                pos = obj['position']
                if 'top' in pos and 'left' in pos and 'width' in pos and 'height' in pos:
                    # Convert to (top, left, bottom, right) format
                    obj['position'] = {
                        'top': pos['top'],
                        'left': pos['left'],
                        'right': pos['left'] + pos['width'],
                        'bottom': pos['top'] + pos['height']
                    }
            
            # Recursively process nested dictionaries and lists
            for key, value in obj.items():
                obj[key] = self.convert_position_format(value)
        elif isinstance(obj, list):
            # Process each item in the list
            for i, item in enumerate(obj):
                obj[i] = self.convert_position_format(item)
        
        return obj


    def save_comprehensive_results(self):
        """Save all comprehensive analysis results"""
        print(f"\n💾 SAVING COMPREHENSIVE RESULTS")
        print("=" * 50)
        
        # Create a deep copy of comprehensive_data and convert position format
        print(f"   🔄 Converting position coordinates from (top, left, width, height) to (top, left, bottom, right)...")
        comprehensive_data_converted = self.convert_position_format(json.loads(json.dumps(self.comprehensive_data, default=str)))
        
        # Save main comprehensive data with converted coordinates
        main_results_path = self.output_dir / "comprehensive_analysis_complete.json"
        with open(main_results_path, 'w', encoding='utf-8') as f:
            json.dump(comprehensive_data_converted, f, ensure_ascii=False, indent=2, default=str)
        
        # Save summary
        summary_path = self.output_dir / "comprehensive_summary.json"
        
        # Calculate local sectioning statistics
        total_local_sections = 0
        total_line_dividers = 0
        slides_with_lines = 0
        
        for slide in self.comprehensive_data.get('spatial_analysis', {}).get('slides', []):
            local_sections = slide.get('local_sections', [])
            line_dividers = slide.get('line_dividers', [])
            total_local_sections += len(local_sections)
            total_line_dividers += len(line_dividers)
            if line_dividers:
                slides_with_lines += 1
        
        summary = {
            'analysis_method': 'comprehensive_unified_with_local_line_sectioning',
            'timestamp': self.comprehensive_data['timestamp'],
            'components': {
                'text_structure': bool(self.comprehensive_data.get('text_structure')),
                'spatial_analysis': bool(self.comprehensive_data.get('spatial_analysis')),
                'local_line_sectioning': total_line_dividers > 0,
                'visual_captures': len(self.comprehensive_data.get('visual_captures', [])),
                'smart_groups': bool(self.comprehensive_data.get('smart_groups')),
                'reading_order': bool(self.comprehensive_data.get('reading_order'))
            },
            'statistics': {
                'total_slides': len(self.comprehensive_data.get('text_structure', {}).get('slides', [])),
                'total_visual_captures': len(self.comprehensive_data.get('visual_captures', [])),
                'total_smart_groups': sum(len(s.get('smart_groups', {})) for s in self.comprehensive_data.get('smart_groups', {}).get('slides', [])),
                'local_sectioning': {
                    'total_line_dividers': total_line_dividers,
                    'total_local_sections': total_local_sections,
                    'slides_with_line_dividers': slides_with_lines,
                    'average_sections_per_slide': total_local_sections / len(self.comprehensive_data.get('spatial_analysis', {}).get('slides', [])) if self.comprehensive_data.get('spatial_analysis', {}).get('slides') else 0
                }
            },
            'output_directory': str(self.output_dir)
        }
        
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        print(f"✅ Comprehensive results saved:")
        print(f"   📄 Main results: {main_results_path}")
        print(f"   📊 Summary: {summary_path}")
        print(f"   📁 All components in: {self.output_dir}")
    

    def process_presentation(self, pptx_path):
        """Main method to process presentation with comprehensive analysis"""
        pptx_path = Path(pptx_path)
        
        if not pptx_path.exists():
            print(f"❌ File not found: {pptx_path}")
            return False
        
        # Reset consolidation counters for new presentation
        self.reset_consolidation_counters()
        
        print(f"🚀 COMPREHENSIVE PRESENTATION ANALYSIS: {pptx_path.name}")
        print("=" * 80)
        print("🔍 Unified approach combining:")
        print("   📝 Text structure (reading order)")
        print("   📦 Spatial analysis (precise positions)")
        print("   📏 Local line sectioning (lines divide only their vicinity)")
        print("   📸 Visual captures (PDF + matplotlib)")
        print("   🧠 Smart grouping (containment logic)")
        print("   📖 Reading order integration")
        print("=" * 80)
        
        success_count = 0
        total_steps = 7  # Note: Smart grouping now includes reading order reordering + VLM processing
        
        # Step 1: Convert to PDF images (optional for visual capture) - only if needed
        if self.config['generate_visual_captures']:
            print(f"\n🔄 STEP 1/6: PDF IMAGE CONVERSION")
            print("=" * 40)
            pdf_success = self.convert_pptx_to_pdf_images(pptx_path)
            if pdf_success:
                success_count += 1
                print("✅ PDF images ready for visual capture")
            else:
                print("⚠️  PDF visual capture will be skipped")
        else:
            print(f"\n🚀 STEP 1/6: PDF IMAGE CONVERSION - SKIPPED (Streamlined Mode)")
            print("=" * 40)
            print("⚡ Skipping PDF conversion for faster processing")
            success_count += 1  # Consider it successful since we intentionally skipped it
        
        # Step 2: Extract text structure
        print(f"\n📝 STEP 2/6: TEXT STRUCTURE EXTRACTION")
        print("=" * 40)
        if self.extract_text_structure(pptx_path):
            success_count += 1
            print("✅ Text structure extracted successfully")
        else:
            print("❌ Text structure extraction failed")
            return False
        
        # Step 3: Extract spatial analysis
        print(f"\n📦 STEP 3/6: SPATIAL ANALYSIS")
        print("=" * 40)
        # Get watsonx content for spatial analysis
        watsonx_content = None
        if 'text_structure' in self.comprehensive_data:
            text_slides = self.comprehensive_data['text_structure'].get('slides', [])
            if text_slides:
                first_slide = text_slides[0]
                watsonx_content = first_slide.get('content', [])
                print(f"   📝 Found {len(watsonx_content)} watsonx content items for slide 1 spatial matching")
        
        if self.extract_spatial_analysis(pptx_path, watsonx_content):
            success_count += 1
            print("✅ Spatial analysis completed successfully")
        else:
            print("❌ Spatial analysis failed")
            return False
        
        # Step 4: Apply smart grouping
        print(f"\n🧠 STEP 4/6: SMART GROUPING")
        print("=" * 40)
        if self.apply_smart_grouping():
            success_count += 1
            print("✅ Smart grouping applied successfully")
        else:
            print("❌ Smart grouping failed")
            return False
        
        # Step 5: Create comprehensive labeling output with VLM captioning (before reading order)
        try:
            print(f"\n🏷️  STEP 5/7: CREATE VLM CAPTIONS")
            print("=" * 60)
            self.create_comprehensive_labeling_output()
            success_count += 1
            print("✅ VLM captions created successfully")
        except Exception as e:
            print(f"⚠️  Warning: Could not create VLM captions: {e}")
            # Continue without VLM captions
            success_count += 1
        
        # Step 6: Create reading order integration (now with VLM captions available)
        print(f"\n📖 STEP 6/7: READING ORDER INTEGRATION")
        print("=" * 40)
        if self.create_reading_order_integration():
            success_count += 1
            print("✅ Reading order integration completed")
        else:
            print("❌ Reading order integration failed")
            return False
        
        # Step 6.5: Create enhanced reading order summary with VLM captions (separate file)
        try:
            print(f"\n🏷️  CREATING ENHANCED READING ORDER WITH VLM CAPTIONS")
            print("=" * 60)
            # Always create the VLM enhanced file, even if no captions are available
            if not hasattr(self, 'vlm_captions_cache'):
                self.vlm_captions_cache = {}
            vlm_count = len(self.vlm_captions_cache) if self.vlm_captions_cache else 0
            print(f"   📊 Available VLM captions: {vlm_count}")
            self.create_enhanced_reading_order_summary_with_vlm()
            print("✅ Enhanced reading order summary with VLM captions created")
        except Exception as e:
            print(f"⚠️ Warning: Could not create enhanced reading order summary: {e}")
            import traceback
            traceback.print_exc()
        
        # Step 7: Save comprehensive results
        print(f"\n💾 STEP 7/7: SAVE RESULTS")
        print("=" * 40)
        self.save_comprehensive_results()
        success_count += 1
        print("✅ All results saved successfully")
        
        # Final summary
        print(f"\n🎉 COMPREHENSIVE ANALYSIS COMPLETE!")
        print("=" * 80)
        print(f"✅ Completed {success_count}/{total_steps} steps successfully")
        
        # Component summary
        total_visual = len(self.comprehensive_data.get('visual_captures', []))
        total_groups = sum(len(s.get('smart_groups', {})) for s in self.comprehensive_data.get('smart_groups', {}).get('slides', []))
        total_slides = len(self.comprehensive_data.get('text_structure', {}).get('slides', []))
        
        print(f"📊 ANALYSIS RESULTS:")
        if self.streamlined_mode:
            print(f"   🚀 MODE: Streamlined (essential outputs only)")
            print(f"   📁 OUTPUT: reading_order_groups/ + comprehensive_analysis_complete.json")
        else:
            print(f"   📊 MODE: Full analysis (all outputs)")
        
        # Show extraction method used
        extraction_method = self.comprehensive_data.get('text_structure', {}).get('file_info', {}).get('extraction_method', 'unknown')
        if extraction_method == 'watsonx_with_spatial_mapping':
            print(f"   🤖 TEXT EXTRACTION: watsonx Text Extraction V2 API (high quality)")
        elif extraction_method == 'native_python_pptx':
            print(f"   📝 TEXT EXTRACTION: Native python-pptx (fallback)")
        else:
            print(f"   📝 TEXT EXTRACTION: {extraction_method}")
            
        print(f"   📄 Slides processed: {total_slides}")
        print(f"   📝 Text structure: Standard reading order (LEFT→RIGHT within rows, TOP→BOTTOM between rows)") 
        print(f"   📦 Spatial analysis: Precise box positions and shapes")
        
        if not self.streamlined_mode:
            print(f"   📏 Local line sectioning: Lines divide only their vicinity")
            print(f"   📸 Visual captures: {total_visual} (PDF visual crops)")
            
        print(f"   🧠 Smart groups: {total_groups} spatial containment groups (S2, S17...)")
        print(f"   📖 Reading order groups: Same groups reordered by content flow (G1, G2, G3...)")
        
        if not self.streamlined_mode:
            print(f"   🔄 Dual approach: Compare spatial vs reading order grouping")
        
        # Add local sectioning summary
        total_local_sections = sum(len(slide.get('local_sections', [])) for slide in self.comprehensive_data.get('spatial_analysis', {}).get('slides', []))
        total_line_dividers = sum(len(slide.get('line_dividers', [])) for slide in self.comprehensive_data.get('spatial_analysis', {}).get('slides', []))
        
        if total_line_dividers > 0:
            # Analyze line types (original simple approach)
            total_horizontal_lines = 0
            total_vertical_lines = 0
            total_diagonal_lines = 0
            
            for slide in self.comprehensive_data.get('spatial_analysis', {}).get('slides', []):
                for line in slide.get('line_dividers', []):
                    pos = line['position']
                    width = max(pos['width'], 1)
                    height = max(pos['height'], 1)
                    aspect_ratio = width / height
                    
                    if aspect_ratio > 2.0:
                        total_horizontal_lines += 1
                    elif aspect_ratio < 0.5:
                        total_vertical_lines += 1
                    else:
                        total_diagonal_lines += 1
            
            print(f"\n📏 LOCAL LINE SECTIONING:")
            print(f"   📏 Line dividers found: {total_line_dividers}")
            print(f"      • Horizontal lines: {total_horizontal_lines} (RED)")
            print(f"      • Vertical lines: {total_vertical_lines} (BLUE)")
            if total_diagonal_lines > 0:
                print(f"      • Diagonal lines: {total_diagonal_lines} (ORANGE)")
            print(f"   🏗️  Local sections created: {total_local_sections}")
            print(f"   🎯 Sectioning approach: Lines divide only their vicinity")
            print(f"   📖 Within sections: Standard reading (LEFT→RIGHT within rows, TOP→BOTTOM between rows)")
        else:
            print(f"\n📖 READING ORDER:")
            print(f"   📖 No line dividers found, using standard reading (LEFT→RIGHT within rows, TOP→BOTTOM between rows)")
        
        # Free slide images from memory before cleanup
        self.slide_images = {}
        self.comprehensive_data = {}
        
        # Clean up intermediate folders, keep only final outputs
        # (comprehensive_analysis_complete.json + comprehensive_labeling_with_vlm.md
        #  + comprehensive_summary.json + visual_captures/).
        intermediate_dirs = [
            self.text_structure_dir,
            self.spatial_analysis_dir,
            self.smart_groups_dir,
            self.reading_order_dir,
            self.reading_order_groups_dir,
            self.table_extractions_dir,
            self.watsonx_raw_outputs_dir,
            self.pdf_pages_dir,
        ]
        for d in intermediate_dirs:
            if d.exists():
                shutil.rmtree(d)
        
        print(f"\n📁 FINAL OUTPUT:")
        print(f"   📂 {self.output_dir}")
        print(f"   ├── 📄 comprehensive_analysis_complete.json")
        print(f"   ├── 📄 comprehensive_summary.json")
        print(f"   ├── 📝 comprehensive_labeling_with_vlm.md")
        print(f"   └── 📸 visual_captures/")
        
        return True
    

    def _get_table_extractions_info(self):
        """Get table extraction information to include in reading order groups"""
        table_info = {
            'enabled': False,
            'total_tables': 0,
            'extraction_method': 'none',
            'table_files': []
        }
        
        try:
            # Check if we have watsonx text structure with table extractions
            if 'text_structure' in self.comprehensive_data:
                text_structure = self.comprehensive_data['text_structure']
                if 'table_extractions' in text_structure:
                    table_extractions = text_structure['table_extractions']
                    table_info.update({
                        'enabled': table_extractions.get('enabled', False),
                        'total_tables': table_extractions.get('total_tables', 0),
                        'extraction_method': 'watsonx_assembly_analysis',
                        'table_files': table_extractions.get('table_files', [])
                    })
            
            # Also check if table extraction directory has files
            if self.table_extractions_dir.exists():
                table_files = list(self.table_extractions_dir.glob('*.json'))
                if table_files and not table_info['enabled']:
                    table_info.update({
                        'enabled': True,
                        'total_tables': len([f for f in table_files if 'summary' not in f.name]),
                        'extraction_method': 'file_system_detection',
                        'table_files': [f.name for f in table_files]
                    })
            
        except Exception as e:
            print(f"   ⚠️  Error getting table extraction info: {e}")
        
        return table_info
    

    def process_multiple_presentations(self, pptx_files, base_output_dir="./output"):
        """Process multiple PPTX files, each in its own output directory"""
        print(f"\n🎯 BATCH PROCESSING: {len(pptx_files)} files")
        print("=" * 70)
        
        results = []
        
        for idx, pptx_file in enumerate(pptx_files, 1):
            pptx_path = Path(pptx_file)
            
            if not pptx_path.exists():
                print(f"❌ File {idx}/{len(pptx_files)}: {pptx_path.name} - NOT FOUND")
                results.append({'file': str(pptx_path), 'status': 'file_not_found', 'success': False})
                continue
            
            # Create file-specific output directory
            file_stem = pptx_path.stem
            file_output_dir = Path(base_output_dir) / f"{file_stem}_analysis"
            
            print(f"\n📁 Processing {idx}/{len(pptx_files)}: {pptx_path.name}")
            print(f"   📂 Output directory: {file_output_dir}")
            
            try:
                # Create new analyzer instance for this file
                file_analyzer = ComprehensivePresentationAnalyzer(
                    output_dir=str(file_output_dir),
                    streamlined_mode=self.streamlined_mode
                )
                # Process the file
                success = file_analyzer.process_presentation(pptx_file)
                if success:
                    print(f"   ✅ {pptx_path.name} - Analysis completed successfully!")
                    results.append({
                        'file': str(pptx_path), 
                        'output_dir': str(file_output_dir),
                        'status': 'completed', 
                        'success': True
                    })
                else:
                    print(f"   ❌ {pptx_path.name} - Analysis failed")
                    results.append({
                        'file': str(pptx_path), 
                        'output_dir': str(file_output_dir),
                        'status': 'analysis_failed', 
                        'success': False
                    })
                    
            except Exception as e:
                print(f"   ❌ {pptx_path.name} - Error: {e}")
                results.append({
                    'file': str(pptx_path), 
                    'output_dir': str(file_output_dir),
                    'status': 'error', 
                    'error': str(e),
                    'success': False
                })
        
        # Print batch summary
        print(f"\n📊 BATCH PROCESSING SUMMARY")
        print("=" * 50)
        successful = sum(1 for r in results if r['success'])
        failed = len(results) - successful
        
        print(f"   ✅ Successful: {successful}/{len(results)}")
        print(f"   ❌ Failed: {failed}/{len(results)}")
        
        if successful > 0:
            print(f"\n📁 OUTPUT DIRECTORIES:")
            for result in results:
                if result['success']:
                    print(f"   📂 {Path(result['file']).name} → {result['output_dir']}")
        
        if failed > 0:
            print(f"\n❌ FAILED FILES:")
            for result in results:
                if not result['success']:
                    status = result.get('error', result['status'])
                    print(f"   ❌ {Path(result['file']).name} - {status}")
        
        return results



def main():
    """Main function with support for single file and batch folder processing"""
    import sys
    import argparse
    from pathlib import Path

    parser = argparse.ArgumentParser(
        description="Comprehensive Presentation Analyzer",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  Single file:
    python comprehensive_presentation_analyzer.py presentation.pptx
  Folder (all .pptx inside, recursively):
    python comprehensive_presentation_analyzer.py --folder /path/to/folder
  Folder with custom output root:
    python comprehensive_presentation_analyzer.py --folder /path/to/folder --output ./results
""",
    )
    parser.add_argument("file", nargs="?", help="Single .pptx file to process")
    parser.add_argument("--folder", "-f", help="Folder path to process all .pptx files recursively")
    parser.add_argument("--output", "-o", default="./output", help="Output root directory (default: ./output)")

    args = parser.parse_args()

    if not args.file and not args.folder:
        parser.print_help()
        return

    print("🚀 COMPREHENSIVE PRESENTATION ANALYZER")
    print("=" * 70)
    print()

    if args.folder:
        folder = Path(args.folder).resolve()
        if not folder.is_dir():
            print(f"❌ Folder not found: {folder}")
            return

        pptx_files = sorted(folder.rglob("*.pptx"))
        if not pptx_files:
            print(f"❌ No .pptx files found in {folder}")
            return

        print(f"📁 BATCH FOLDER MODE")
        print(f"   📂 Source: {folder}")
        print(f"   📄 Found {len(pptx_files)} .pptx files")
        print("=" * 70)
        print()

        results = {"success": [], "failed": []}
        for idx, pptx_path in enumerate(pptx_files, 1):
            print(f"\n{'='*70}")
            print(f"📄 [{idx}/{len(pptx_files)}] {pptx_path.name}")
            print(f"{'='*70}")

            file_stem = pptx_path.stem
            output_dir = str(Path(args.output) / f"{file_stem}_analysis")
            analyzer = ComprehensivePresentationAnalyzer(output_dir=output_dir)

            try:
                success = analyzer.process_presentation(str(pptx_path))
                if success:
                    print(f"✅ [{idx}/{len(pptx_files)}] {pptx_path.name} - Done")
                    results["success"].append(pptx_path.name)
                else:
                    print(f"❌ [{idx}/{len(pptx_files)}] {pptx_path.name} - Failed")
                    results["failed"].append(pptx_path.name)
            except Exception as e:
                print(f"❌ [{idx}/{len(pptx_files)}] {pptx_path.name} - Error: {e}")
                results["failed"].append(pptx_path.name)
            finally:
                import gc
                del analyzer
                gc.collect()

        print(f"\n{'='*70}")
        print(f"📊 BATCH SUMMARY")
        print(f"   ✅ Success: {len(results['success'])}/{len(pptx_files)}")
        print(f"   ❌ Failed:  {len(results['failed'])}/{len(pptx_files)}")
        if results["failed"]:
            print(f"   Failed files:")
            for f in results["failed"]:
                print(f"     - {f}")
        print(f"   📂 Output: {Path(args.output).resolve()}")

    else:
        test_file = str(Path(args.file).resolve())

        if not Path(test_file).exists():
            print(f"❌ File not found: {test_file}")
            return

        print(f"📄 SINGLE FILE MODE: {os.path.basename(test_file)}")
        print("=" * 70)

        file_stem = Path(test_file).stem
        output_dir = str(Path(args.output) / f"{file_stem}_analysis")
        analyzer = ComprehensivePresentationAnalyzer(output_dir=output_dir)

        if analyzer.watsonx_available:
            print("✅ watsonx Text Extraction V2 ready")
        else:
            print("⚠️  watsonx not available - will use native python-pptx extraction")
        print()

        try:
            success = analyzer.process_presentation(test_file)
            if success:
                print("✅ Analysis completed successfully!")
                print(f"📂 Output directory: {output_dir}")
            else:
                print("❌ Analysis failed")
        except Exception as e:
            print(f"❌ Error during analysis: {e}")
            import traceback
            traceback.print_exc()
    


if __name__ == "__main__":
    main() 