"""
Labeling output, Korean filename generation, and summary text creation.
"""

import os
import json
import time
from pathlib import Path
from datetime import datetime


class LabelingMixin:
    """Methods for creating labeled output and summaries."""

    def create_comprehensive_labeling_output(self):
        """Create comprehensive labeling output with VLM image captioning from reading_order_groups_summary.txt"""
        try:
            print("\n📋 CREATING COMPREHENSIVE LABELING OUTPUT WITH VLM CAPTIONING")
            
            # Import VLM functionality
            try:
                from ibm_watsonx_ai.foundation_models import ModelInference
                import base64
                import os
                self.ModelInference = ModelInference
                vlm_available = True
                print("✅ VLM functionality loaded successfully")
            except Exception as e:
                vlm_available = False
                print(f"❌ VLM not available: {e}")
            
            # Load the reading order summary text file (not JSON)
            summary_file = self.output_dir / "reading_order_groups" / "reading_order_groups_summary.txt"
            if not summary_file.exists():
                print(f"❌ Reading order summary file not found: {summary_file}")
                return
            
            print(f"📖 Processing summary file: {summary_file}")
            
            # Parse the summary file to extract image references
            image_references = self.extract_image_references_from_summary(summary_file)
            if not image_references:
                print("❌ No image references found in summary file")
                return
            
            print(f"📸 Found {len(image_references)} image references")
            
            # Create enhanced output with VLM captions
            if vlm_available and getattr(self, 'watsonx_available', False) and getattr(self, 'credentials', None):
                enhanced_summary = self.create_vlm_enhanced_summary(summary_file, image_references)
            else:
                if not getattr(self, 'watsonx_available', False):
                    print("⚠️ watsonx credentials not configured — skipping VLM captions")
                else:
                    print("⚠️ VLM not available, creating summary without captions")
                enhanced_summary = self.create_basic_enhanced_summary(summary_file, image_references)
            
            # Save the enhanced summary
            output_file = self.output_dir / "comprehensive_labeling_with_vlm.md"
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(enhanced_summary)
            
            print(f"✅ Comprehensive labeling output created: {output_file}")
            
        except Exception as e:
            print(f"❌ Error creating comprehensive labeling output: {e}")
            import traceback
            traceback.print_exc()



    def extract_image_references_from_summary(self, summary_file):
        """Extract all PNG image references from the reading order summary file"""
        try:
            import re
            image_references = []
            
            with open(summary_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Look for patterns like "→ slide_01_autoshape_23_visual.png"
            png_pattern = r'→\s*([^→\n]*\.png)'
            matches = re.findall(png_pattern, content)
            
            for match in matches:
                png_filename = match.strip()
                # Construct full path relative to visual_captures directory
                full_path = self.output_dir / "visual_captures" / png_filename
                if full_path.exists():
                    image_references.append({
                        'filename': png_filename,
                        'full_path': str(full_path),
                        'exists': True
                    })
                else:
                    image_references.append({
                        'filename': png_filename,
                        'full_path': str(full_path),
                        'exists': False
                    })
                    print(f"⚠️ Image not found: {full_path}")
            
            return image_references
            
        except Exception as e:
            print(f"❌ Error extracting image references: {e}")
            return []



    def create_vlm_enhanced_summary(self, summary_file, image_references):
        """Create enhanced summary with VLM captions"""
        try:
            # Initialize VLM captions cache if not exists
            if not hasattr(self, 'vlm_captions_cache'):
                self.vlm_captions_cache = {}
            
            # Read original summary
            with open(summary_file, 'r', encoding='utf-8') as f:
                original_content = f.read()
            
            # Create enhanced version
            enhanced_content = f"""# 🏷️ COMPREHENSIVE LABELING OUTPUT WITH VLM CAPTIONING

**Analysis Date:** {datetime.now().isoformat()}
**Source:** {summary_file.name}
**VLM Model:** mistralai/mistral-medium-2505
**Total Images:** {len(image_references)}

---

## 📸 VLM IMAGE CAPTIONS

"""
            
            # Process each image with VLM
            caption_count = 0
            for i, img_ref in enumerate(image_references, 1):
                print(f"   🤖 Processing image {i}/{len(image_references)}: {img_ref['filename']}")
                enhanced_content += f"### 📷 Image {i}: `{img_ref['filename']}`\n\n"
                if img_ref['exists']:
                    # Generate VLM caption
                    caption_info = self.generate_vlm_caption(img_ref['full_path'])
                    if caption_info:
                        enhanced_content += f"**📝 VLM Caption:**\n"
                        enhanced_content += f"- **Title:** {caption_info['title']}\n"
                        enhanced_content += f"- **Description:** {caption_info['description']}\n\n"
                        
                        # Generate Korean filename
                        korean_filename = self.generate_korean_filename_for_image(img_ref['filename'], caption_info)
                        enhanced_content += f"**🏷️ Korean Filename:** `{korean_filename}`\n\n"
                        
                        # Cache the VLM caption for later use in reading order summary
                        self.vlm_captions_cache[img_ref['filename']] = {
                            'title': caption_info['title'],
                            'description': caption_info['description'],
                            'korean_filename': korean_filename,
                            'vlm_raw_response': caption_info.get('vlm_raw_response', '')
                        }
                        
                        caption_count += 1
                    else:
                        enhanced_content += f"❌ **VLM Caption Failed**\n\n"
                else:
                    enhanced_content += f"❌ **Image Not Found:** {img_ref['full_path']}\n\n"
                enhanced_content += "---\n\n"
            
            # Add original summary at the end
            enhanced_content += f"""## 📋 ORIGINAL READING ORDER SUMMARY

{original_content}

---

**Summary:** Successfully generated {caption_count} VLM captions out of {len(image_references)} images.
"""
            
            return enhanced_content
            
        except Exception as e:
            print(f"❌ Error creating VLM enhanced summary: {e}")
            return f"Error creating enhanced summary: {e}"



    def create_basic_enhanced_summary(self, summary_file, image_references):
        """Create basic enhanced summary without VLM (fallback)"""
        try:
            with open(summary_file, 'r', encoding='utf-8') as f:
                original_content = f.read()
            
            enhanced_content = f"""# 🏷️ COMPREHENSIVE LABELING OUTPUT (No VLM)

**Analysis Date:** {datetime.now().isoformat()}
**Source:** {summary_file.name}
**VLM Status:** Not Available
**Total Images:** {len(image_references)}

---

## 📸 IMAGE INVENTORY

"""
            
            for i, img_ref in enumerate(image_references, 1):
                status = "✅ Found" if img_ref['exists'] else "❌ Missing"
                enhanced_content += f"{i}. `{img_ref['filename']}` - {status}\n"
            
            enhanced_content += f"""

---

## 📋 ORIGINAL READING ORDER SUMMARY

{original_content}
"""
            
            return enhanced_content
            
        except Exception as e:
            print(f"❌ Error creating basic enhanced summary: {e}")
            return f"Error creating basic summary: {e}"



    def generate_korean_filename_for_image(self, original_filename, caption_info):
        """Generate Korean filename for a single image"""
        try:
            # Extract slide number and component info from filename
            # e.g., "slide_01_autoshape_23_visual.png" 
            import re
            match = re.match(r'slide_(\d+)_(\w+)_(\d+)_visual\.png', original_filename)
            
            if match:
                slide_num = match.group(1)
                comp_type = match.group(2)
                comp_num = match.group(3)
            else:
                slide_num = "01"
                comp_type = "content"
                comp_num = "001"
            
            # Get presentation name
            source_file = Path(self.pptx_path).stem if hasattr(self, 'pptx_path') else "presentation"
            
            # Determine if it's a table image
            is_table = 'table' in comp_type.lower() or 'table' in caption_info.get('title', '').lower()
            
            # Generate type code
            type_code = f"{int(slide_num):03d}"
            
            # Clean title for filename
            title = caption_info.get('title', 'image')
            clean_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).strip()
            clean_title = clean_title.replace(' ', '_')[:50]  # Limit length
            
            if is_table:
                # Table image: tableimg_033_{filename}_content_001
                korean_filename = f"tableimg_{type_code}_{{{source_file}}}_content_{comp_num:>03s}.png"
            else:
                # Regular image: img_032_{filename}_content_001_caption
                korean_filename = f"img_{type_code}_{{{source_file}}}_content_{comp_num:>03s}_{clean_title}.png"
            
            return korean_filename
            
        except Exception as e:
            print(f"   ⚠️ Error generating Korean filename for {original_filename}: {e}")
            return f"img_001_{{unknown}}_content_001_image.png"



    def generate_korean_image_filename(self, slide_number, component, caption_info):
        """Generate filename according to Korean labeling convention"""
        try:
            # Extract base info
            source_file = Path(self.pptx_path).stem if hasattr(self, 'pptx_path') else "presentation"
            comp_type = component.get('type', 'unknown')
            comp_id = component.get('component_id', '001')
            
            # Determine if it's a table image
            is_table = 'table' in comp_type.lower() or 'table' in caption_info.get('title', '').lower()
            
            # Generate code (arbitrary as mentioned in requirements)
            type_code = f"{int(slide_number):03d}"
            
            # Clean title for filename
            title = caption_info.get('title', 'image')
            # Remove special characters and limit length
            clean_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).strip()
            clean_title = clean_title.replace(' ', '_')[:50]  # Limit length
            
            # Get sequential number
            seq_num = comp_id.split('_')[-1] if '_' in str(comp_id) else "001"
            
            if is_table:
                # Table image: tableimg_033_{filename}_key_001
                new_filename = f"tableimg_{type_code}_{{{source_file}}}_content_{seq_num}"
            else:
                # Regular image: img_032_{filename}_key_001_caption
                new_filename = f"img_{type_code}_{{{source_file}}}_content_{seq_num}_{clean_title}"
            
            return new_filename
            
        except Exception as e:
            print(f"   ⚠️ Error generating Korean filename: {e}")
            return f"img_001_{{unknown}}_content_001_image"



    def rename_captured_image(self, old_path, new_filename):
        """Rename captured image according to new convention"""
        try:
            old_path = Path(old_path)
            if not old_path.exists():
                return False
            
            extension = old_path.suffix
            new_path = old_path.parent / f"{new_filename}{extension}"
            
            old_path.rename(new_path)
            print(f"   📁 Renamed: {old_path.name} → {new_path.name}")
            return str(new_path)
            
        except Exception as e:
            print(f"   ⚠️ Error renaming {old_path}: {e}")
            return False



    def create_labeling_summary_text(self, enhanced_output):
        """Create human-readable summary with VLM captions"""
        try:
            summary_file = self.output_dir / "comprehensive_labeling_summary.md"
            
            with open(summary_file, 'w', encoding='utf-8') as f:
                f.write("# 📋 COMPREHENSIVE LABELING OUTPUT WITH VLM CAPTIONING\n\n")
                f.write(f"**Analysis Date:** {enhanced_output['metadata']['creation_time']}\n")
                f.write(f"**Source File:** {enhanced_output['metadata']['source_file']}\n")
                f.write(f"**VLM Captioning:** {'✅ Enabled' if enhanced_output['metadata']['vlm_captioning_enabled'] else '❌ Disabled'}\n\n")
                # Process each slide
                for slide_key, slide_data in enhanced_output['slides'].items():
                    slide_num = slide_key.replace('slide_', '')
                    f.write(f"## 📄 SLIDE {slide_num}\n\n")
                    
                    # Reading order groups with position info
                    f.write("### 📖 READING ORDER GROUPS WITH SPATIAL INFO\n\n")
                    
                    for group in slide_data['reading_order_groups']:
                        f.write(f"#### 🏷️ GROUP {group['group_id']} (Order: {group['reading_order']})\n")
                        f.write(f"**Type:** {group['group_type']}\n")
                        f.write(f"**Summary:** {group['content_summary']}\n")
                        
                        # Spatial information
                        spatial = group['spatial_info']
                        if spatial.get('position'):
                            pos = spatial['position']
                            f.write(f"**Position:** ({pos.get('left', 0):.1f}, {pos.get('top', 0):.1f}) - Size: {pos.get('width', 0):.1f} × {pos.get('height', 0):.1f}\n")
                        
                        f.write("\n**Components:**\n")
                        
                        for comp in group['components']:
                            f.write(f"- **{comp['type']}** (ID: {comp['component_id']})\n")
                            if comp.get('content'):
                                content = str(comp['content'])[:200] + "..." if len(str(comp['content'])) > 200 else str(comp['content'])
                                f.write(f"  Content: {content}\n")
                            
                            # Position info
                            if comp.get('position'):
                                pos = comp['position']
                                f.write(f"  Position: ({pos.get('left', 0):.1f}, {pos.get('top', 0):.1f})\n")
                            
                            # VLM Caption info
                            if comp.get('vlm_caption'):
                                caption = comp['vlm_caption']
                                f.write(f"\n  📸 **IMAGE CAPTION (VLM Generated):**\n")
                                f.write(f"  Title: {caption['title']}\n")
                                f.write(f"  Description: {caption['description']}\n")
                                
                                if comp.get('renamed_capture'):
                                    f.write(f"  New Filename: {Path(comp['renamed_capture']).name}\n")
                            
                            f.write("\n")
                        
                        f.write("\n---\n\n")
                    
                    # Image rename summary
                    if slide_data.get('renamed_images'):
                        f.write("### 📁 IMAGE RENAMING SUMMARY\n\n")
                        for old_name, new_name in slide_data['renamed_images'].items():
                            f.write(f"- `{Path(old_name).name}` → `{Path(new_name).name}`\n")
                        f.write("\n")
            
            print(f"✅ Human-readable summary created: {summary_file}")
            
        except Exception as e:
            print(f"❌ Error creating labeling summary: {e}")


