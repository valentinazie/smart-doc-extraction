"""
VLM (Vision Language Model) caption generation and enhanced summaries.
"""

import os
import time
import base64
from pathlib import Path
from datetime import datetime

from utils.geometry import (
    calculate_group_bounds,
)



class VLMMixin:
    """Methods for VLM captioning and enhanced summary generation."""

    def create_enhanced_reading_order_summary_with_vlm(self):
        """Create enhanced reading order summary with VLM captions integrated"""
        summary_path = self.reading_order_groups_dir / "reading_order_groups_with_vlm_captions.md"
        
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write("# 🏷️ ENHANCED READING ORDER GROUPS WITH VLM CAPTIONS\n")
            f.write("=" * 80 + "\n\n")
            f.write(f"**Analysis Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("**Approach:** Hierarchical Group-based Reading Order with VLM Image Captioning\n")
            f.write("**Reading Pattern:** TOP→BOTTOM between groups, LEFT→RIGHT within groups\n")
            f.write("**VLM Model:** mistralai/mistral-medium-2505\n\n")
            
            # Process each slide
            total_groups = 0
            total_components = 0
            
            for slide_data in self.comprehensive_data['spatial_analysis']['slides']:
                slide_num = slide_data['slide_number']
                hierarchical_groups = slide_data.get('hierarchical_groups', [])
                if not hierarchical_groups:
                    continue
                f.write(f"{'='*25} SLIDE {slide_num} {'='*25}\n\n")
                # Get all components in reading order (excluding layout containers)
                all_components = []
                for box in slide_data.get('boxes', []):
                    if 'hierarchical_info' in box:
                        # Skip ONLY layout containers (1x1 empty tables used as containers)
                        is_layout_container = (
                            box.get('shape_type') == 'LayoutContainer' and 
                            box.get('layout_container_type') == 'empty_table'
                        )
                        if is_layout_container:
                            continue
                        all_components.append(box)
                # Sort by hierarchical order
                all_components.sort(key=lambda x: (
                    x['hierarchical_info']['group_order'],
                    x['hierarchical_info']['component_order_in_group']
                ))
                # Statistics
                f.write(f"## 📊 SLIDE SUMMARY\n")
                f.write(f"   • **Groups:** {len(hierarchical_groups)}\n")
                f.write(f"   • **Content Components:** {len(all_components)} (layout containers excluded)\n\n")
                # Group by group with sequential numbering
                current_group = None
                group_components = []
                group_counter = 0
                for component in all_components:
                    group_id = component['hierarchical_info']['group_id']
                    
                    if current_group != group_id:
                        # Write previous group
                        if current_group and group_components:
                            self._write_enhanced_group_summary_with_vlm(f, group_counter, current_group, group_components, slide_num)
                        
                        # Start new group
                        current_group = group_id
                        group_components = []
                        group_counter += 1
                    
                    group_components.append(component)
                # Write last group
                if current_group and group_components:
                    self._write_enhanced_group_summary_with_vlm(f, group_counter, current_group, group_components, slide_num)
                f.write("\n")
                total_groups += len(hierarchical_groups)
                total_components += len(all_components)
            
            # Overall statistics
            f.write(f"## 📈 OVERALL STATISTICS\n")
            f.write(f"**Total Groups:** {total_groups}\n")
            f.write(f"**Total Components:** {total_components}\n")
            f.write(f"**VLM Captions:** {len(self.vlm_captions_cache) if hasattr(self, 'vlm_captions_cache') else 0}\n")
            
        print(f"✅ Enhanced reading order summary with VLM captions saved: {summary_path}")
        return summary_path



    def _write_enhanced_group_summary_with_vlm(self, f, group_number, original_group_id, components, slide_number):
        """Write enhanced summary for a single group with VLM captions"""
        f.write(f"### 📦 GROUP {group_number} ({len(components)} components)\n")
        
        # Add group location first - clearly about the group
        if components:
            group_bounds = calculate_group_bounds(components)
            if group_bounds:
                f.write(f"**📍 Group Location:** Top={group_bounds['top']:.0f}, Bottom={group_bounds['bottom']:.0f}, Left={group_bounds['left']:.0f}, Right={group_bounds['right']:.0f} (W×H: {group_bounds['width']:.0f}×{group_bounds['height']:.0f})\n")
        f.write(f"**🔗 Original ID:** {original_group_id}\n\n")
        
        for i, component in enumerate(components, 1):
            # Check if this is a table - if so, use hierarchical display
            if component.get('shape_type') == 'Table' and 'cell_contents' in component:
                self._write_enhanced_hierarchical_table_summary_with_vlm(f, component, i, slide_number)
            # Check if this is a UnifiedGroup - if so, use hierarchical display
            elif component.get('shape_type') == 'UnifiedGroup' and ('component_images' in component or 'component_texts' in component):
                self._write_enhanced_hierarchical_unified_group_summary_with_vlm(f, component, i, slide_number)
            # Check if this is a SmartVisualGroup that contains tables
            elif component.get('shape_type') == 'SmartVisualGroup' and 'component_visuals' in component:
                self._write_enhanced_smart_visual_group_summary_with_vlm(f, component, i, slide_number)
            else:
                self._write_enhanced_standard_component_summary_with_vlm(f, component, i, slide_number)
        
        f.write("---\n\n")



    def _write_enhanced_hierarchical_unified_group_summary_with_vlm(self, f, component, index, slide_number):
        """Write hierarchical unified group display with VLM captions"""
        box_id = component['box_id']
        
        # Check for visual capture
        visual_capture = self.find_visual_capture_file(box_id, slide_number)
        visual_info = f" → `{visual_capture}`" if visual_capture else ""
        
        # Add position information
        pos = component.get('position', {})
        position_info = ""
        if pos:
            position_info = f" [T={pos.get('top', 0):.0f}, L={pos.get('left', 0):.0f}, W={pos.get('width', 0):.0f}×H={pos.get('height', 0):.0f}]"
        
        # Get text content for the unified group
        text_content = component.get('text', '').strip()
        text_display = f'"{text_content}"' if text_content else "[Combined Visual Group]"
        
        # Write unified group container
        f.write(f" {index:2d}. **{box_id}** (UnifiedGroup): {text_display}{position_info}{visual_info}\n")
        
        # Add VLM caption for the unified group
        if visual_capture:
            vlm_caption = self.find_vlm_caption_for_visual_file(visual_capture)
            if vlm_caption:
                f.write(f"   - **🤖 Image Title:** {vlm_caption['title']}\n")
                f.write(f"   - **📝 Image Description:** {vlm_caption['description']}\n")
                f.write(f"   - **🏷️ Korean Filename:** `{vlm_caption['korean_filename']}`\n")
        
        # Count total components
        component_images = component.get('component_images', [])
        component_texts = component.get('component_texts', [])
        component_lines = component.get('component_lines', [])
        total_components = len(component_images) + len(component_texts) + len(component_lines)
        
        f.write(f"\n   **📦 Contains {total_components} visual components:**\n")
        
        # Write individual images
        for img in component_images:
            img_text = img.get('text', '').strip()
            img_display = f': "{img_text}"' if img_text else ': [No text]'
            f.write(f"   - **{img['box_id']}** ({img.get('shape_type', 'Unknown')}){img_display}\n")
        
        # Write individual texts  
        for txt in component_texts:
            txt_text = txt.get('text', '').strip()
            txt_display = f': "{txt_text}"' if txt_text else ': [No text]'
            f.write(f"   - **{txt['box_id']}** ({txt.get('shape_type', 'Unknown')}){txt_display}\n")
        
        # Write individual lines (if any)
        for line in component_lines:
            f.write(f"   - **{line['box_id']}** (Line): [Structural line]\n")
        
        f.write("\n")



    def _write_enhanced_hierarchical_table_summary_with_vlm(self, f, component, index, slide_number):
        """Write hierarchical table display with VLM captions"""
        box_id = component['box_id']
        table_dimensions = component.get('table_dimensions', 'unknown')
        
        # Check for visual capture
        visual_capture = self.find_visual_capture_file(box_id, slide_number)
        visual_info = f" → `{visual_capture}`" if visual_capture else ""
        
        # Add position information
        pos = component.get('position', {})
        position_info = ""
        if pos:
            position_info = f" [T={pos.get('top', 0):.0f}, L={pos.get('left', 0):.0f}, W={pos.get('width', 0):.0f}×H={pos.get('height', 0):.0f}]"
        
        # Write table container
        f.write(f" {index:2d}. **{box_id}** (Table): [Table Container {table_dimensions}]{position_info}{visual_info}\n")
        
        # Add VLM caption for the table
        if visual_capture:
            vlm_caption = self.find_vlm_caption_for_visual_file(visual_capture)
            if vlm_caption:
                f.write(f"   - **🤖 Image Title:** {vlm_caption['title']}\n")
                f.write(f"   - **📝 Image Description:** {vlm_caption['description']}\n")
                f.write(f"   - **🏷️ Korean Filename:** `{vlm_caption['korean_filename']}`\n")
        
        # Write individual cells
        cell_contents = component.get('cell_contents', [])
        if cell_contents:
            # Group cells by row for better display
            rows = {}
            for cell in cell_contents:
                # Show cells with text content OR visual content (embedded components)
                if cell['has_content'] or cell.get('has_visual_content') or cell.get('shapes'):
                    row_idx = cell['row']
                    if row_idx not in rows:
                        rows[row_idx] = []
                    rows[row_idx].append(cell)
            
            f.write(f"\n   **📋 Table cells:**\n")
            # Display cells row by row
            for row_idx in sorted(rows.keys()):
                for cell in sorted(rows[row_idx], key=lambda c: c['col']):
                    # Show ALL components in the cell from BOTH systems
                    cell_text = cell.get('text', '').strip()
                    cell_shapes = cell.get('shapes', [])  # NEW system
                    overlapping_shapes = cell.get('overlapping_shapes', [])  # OLD system (for UnifiedGroups)
                    
                    # Build display showing ALL components
                    components_display = []
                    
                    # Add original cell text if exists
                    if cell_text:
                        components_display.append(f'"{cell_text}"')
                    
                    # Add ALL shapes/components from NEW system (shapes array)
                    for shape in cell_shapes:
                        if shape['type'] == 'Table':
                            # Table - show image path
                            table_id = shape['box_id']
                            table_num = table_id[1:] if len(table_id) > 1 else '00'
                            try:
                                table_visual = f"slide_{slide_number:02d}_table_{int(table_num):02d}_visual.png"
                            except (ValueError, TypeError):
                                table_visual = f"slide_{slide_number:02d}_table_{table_num}_visual.png"
                            components_display.append(f"`{table_visual}`")
                        elif shape['type'] == 'Picture':
                            # Picture - show image path
                            picture_id = shape['box_id']
                            picture_num = picture_id[1:] if len(picture_id) > 1 else '00'
                            picture_visual = f"slide_{slide_number:02d}_picture_{picture_num}_visual.png"
                            components_display.append(f"`{picture_visual}`")
                        elif shape['type'] == 'UnifiedGroup':
                            # UnifiedGroup - show visual capture path directly
                            if shape.get('visual_capture'):
                                components_display.append(f"`{shape['visual_capture']}`")
                        elif shape.get('text'):
                            # Text component - show the text
                            components_display.append(f'"{shape["text"]}"')
                    
                    # Add ALL components from OLD system (overlapping_shapes array - includes UnifiedGroups)
                    for shape in overlapping_shapes:
                        if shape.get('visual_capture'):
                            # Component with visual capture (Groups, Pictures, Tables)
                            components_display.append(f"`{shape['visual_capture']}`")
                        elif shape.get('text_content'):
                            # Text component
                            components_display.append(f'"{shape["text_content"]}"')
                    
                    # Join all components with " + " separator
                    if components_display:
                        display_content = " + ".join(components_display)
                        visual_emoji = " 🖼️" if cell.get('has_visual_content', False) else ""
                        f.write(f"   - **Cell({cell['row']},{cell['col']}):** {display_content}{visual_emoji}\n")
                    else:
                        f.write(f"   - **Cell({cell['row']},{cell['col']}):** [Empty]\n")
        else:
            f.write(f"\n   **📋 Table cells:** [Empty table or no cell data available]\n")
        
        # Display not-embedded components (components that overlap with table but not assigned to specific cells)
        # FILTER OUT consumed/converted components (like 1x2 tables that were split into groups)
        if 'not_embedded_components' in component and component['not_embedded_components']:
            # Get list of all consumed shape IDs to filter out
            all_consumed_ids = getattr(self, '_all_consumed_shape_ids', [])  # IDs of shapes consumed by spatial grouping
            converted_table_ids = getattr(self, '_converted_table_ids', [])  # IDs of tables converted to groups (like S2→G2,G3)
            
            # Filter out consumed/converted components
            active_not_embedded = []
            for not_embedded in component['not_embedded_components']:
                box_id = not_embedded.get('box_id')
                # Skip if this component was consumed by spatial grouping or converted to groups
                if box_id not in all_consumed_ids and box_id not in converted_table_ids:
                    active_not_embedded.append(not_embedded)
            
            # Only show section if there are active not-embedded components
            if active_not_embedded:
                f.write(f"\n   **📎 Not-embedded components:**\n")
                for not_embedded in active_not_embedded:
                    text_info = f" - \"{not_embedded['text']}\"" if not_embedded.get('text') else ""
                    visual_path = f" → `{not_embedded['visual_capture']}`" if not_embedded.get('visual_capture') else ""
                    overlap_info = f" *({not_embedded['table_overlap_percentage']:.1f}% table overlap)*" if not_embedded.get('table_overlap_percentage') else ""
                    f.write(f"   - **{not_embedded['box_id']}** ({not_embedded['type']}){overlap_info}{text_info}{visual_path}\n")
        
        f.write("\n")



    def _write_enhanced_smart_visual_group_summary_with_vlm(self, f, component, index, slide_number):
        """Write SmartVisualGroup summary with VLM captions (existing logic)"""
        # This would be similar to the existing _write_smart_visual_group_summary but with VLM captions
        # For now, fall back to standard component display
        self._write_enhanced_standard_component_summary_with_vlm(f, component, index, slide_number)



    def _write_enhanced_standard_component_summary_with_vlm(self, f, component, index, slide_number):
        """Write standard component summary with VLM captions"""
        # Get component info
        box_id = component['box_id']
        shape_type = component.get('shape_type', 'Unknown')
        text_content = component.get('text', '').strip()
        
        # Show full text content (no truncation)
        if text_content:
            text_display = f'"{text_content}"'
        else:
            text_display = "[No text]"
        
        # Add position information
        pos = component.get('position', {})
        position_info = ""
        if pos:
            position_info = f" [T={pos.get('top', 0):.0f}, L={pos.get('left', 0):.0f}, W={pos.get('width', 0):.0f}×H={pos.get('height', 0):.0f}]"
        
        # Base component line
        f.write(f" {index:2d}. **{box_id}** ({shape_type}): {text_display}{position_info}")
        
        # Check for visual capture
        visual_capture = self.find_visual_capture_file(box_id, slide_number)
        
        if visual_capture:
            f.write(f" → `{visual_capture}`\n")
            # Add VLM caption information with clear hierarchy
            vlm_caption = self.find_vlm_caption_for_visual_file(visual_capture)
            if vlm_caption:
                f.write(f"   - **🤖 Image Title:** {vlm_caption['title']}\n")
                f.write(f"   - **📝 Image Description:** {vlm_caption['description']}\n")
                f.write(f"   - **🏷️ Korean Filename:** `{vlm_caption['korean_filename']}`\n")
            else:
                f.write(f"   - **❌ VLM Caption:** Not available\n")
        else:
            f.write("\n")
        
        f.write("\n")



    def find_vlm_caption_for_visual_file(self, visual_filename):
        """Find VLM caption information for a visual file"""
        try:
            # Check if we have stored VLM captions (they should be stored during VLM processing)
            if not hasattr(self, 'vlm_captions_cache') or not self.vlm_captions_cache:
                return None
            
            # Look for the caption in our cache
            for image_filename, caption_info in self.vlm_captions_cache.items():
                if image_filename == visual_filename:
                    return caption_info
            
            return None
            
        except Exception as e:
            print(f"⚠️ Error finding VLM caption for {visual_filename}: {e}")
            return None



    def generate_vlm_caption(self, image_path, slide_number=None):
        """Generate VLM caption with boundary table exclusion"""
        # Skip VLM captioning for boundary tables
        if 'table_02_visual' in str(image_path) or 'table_04_visual' in str(image_path):
            print(f'      🔲 Skipping VLM caption for boundary table: {image_path}')
            return {
                'title': 'Boundary Table (Text Only)',
                'description': 'This is a boundary table processed as text-only content.',
                'korean_filename': 'boundary_table_text_only.txt'
            }
        
        # For all other images, call the original VLM generation logic
        return self.original_generate_vlm_caption(image_path)



    def original_generate_vlm_caption(self, image_path):
        """Generate VLM caption for an image using direct VLM functions"""
        try:
            print(f"   🤖 Generating VLM caption for: {Path(image_path).name}")
            
            # Encode image using VLM functions
            image_encoded_list = self.get_image_encode(image_path)
            image_encoded = image_encoded_list[0]  # get_image_encode returns a list
            
            # Generate description
            user_request = """[Guide]
You are seeing the image we got from automobile presentation.
So to match the contents location, we need information on what is the image is about and the title so we can identify it.

[Output format]
Output:
- Image title:
- Image description:

[Trial]
<|assistant|>
Output:
"""
            
            description = self.get_image_description(image_encoded, user_request)
            
            # Parse the response to extract title and description
            lines = description.split('\n')
            title = ""
            full_description = description
            
            for line in lines:
                if 'image title:' in line.lower():
                    title = line.split(':', 1)[1].strip()
                    break
            
            if not title:
                # Try to extract from first line or use a default
                first_line = lines[0].strip() if lines else ""
                title = first_line[:100] if first_line else f"Content from {Path(image_path).stem}"
            
            return {
                'title': title,
                'description': full_description,
                'vlm_raw_response': description
            }
            
        except Exception as e:
            print(f"   ⚠️ Error generating VLM caption for {image_path}: {e}")
            import traceback
            traceback.print_exc()
            return None



    def get_image_encode(self, images_path):
        """Encode image to base64 (from VLM functionality)"""
        encoded_images = []
        with open(images_path, 'rb') as image_file:
            encoded_images.append(base64.b64encode(image_file.read()).decode("utf-8"))
            return encoded_images



    def get_image_description(self, image, user_request):
        """Get image description using VLM (from VLM functionality)"""
        messages = self.augment_api_request_body(user_request, image)
        response = self.get_model().chat(messages=messages)
        return response['choices'][0]['message']['content']



    def augment_api_request_body(self, user_query, image):
        """Prepare API request body (from VLM functionality)"""
        messages = [
            {
                "role": "user",
                "content": [{
                    "type": "text",
                    "text": 'You are a helpful assistant. Do the following:\n ' + user_query
                },
                {
                    "type": "image_url",
                    "image_url": {
                    "url": f"data:image/jpeg;base64,{image}",
                    }
                }]
            }
        ]
        return messages



    def get_model(self):
        """Get VLM model (from VLM functionality)"""
        project_id = os.getenv("WATSONX_PROJECT_ID") or os.getenv("PROJECT_ID")
        if not project_id:
            raise RuntimeError(
                "VLM requires a watsonx project. Set WATSONX_PROJECT_ID or PROJECT_ID in .env"
            )
        return self.ModelInference(
            model_id = "mistralai/mistral-medium-2505",
            credentials=self.credentials,
            project_id=project_id,
            params={
                "max_tokens": 2000,
            }
        )


