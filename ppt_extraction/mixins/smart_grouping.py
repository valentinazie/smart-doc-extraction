"""
Smart grouping, unified groups, image consolidation, and hybrid entities.
"""

import json
import time
import numpy as np
from pathlib import Path
from datetime import datetime

from utils.geometry import (
    calculate_combined_boundary,
    calculate_merged_boundary,
    calculate_overlap_percentage,
    calculate_proximity,
    calculate_spatial_overlap,
    simple_boxes_overlap,
)


class SmartGroupingMixin:
    """Methods for smart spatial grouping and entity consolidation."""

    def apply_smart_grouping(self):
        """Apply smart containment grouping to spatial analysis with enhanced visualization"""
        print(f"\n🧠 APPLYING SMART GROUPING")
        print("=" * 50)
        
        if 'spatial_analysis' not in self.comprehensive_data:
            print("❌ No spatial analysis available for smart grouping")
            return False
        
        smart_groups_analysis = {
            'slides': [],
            'overall_statistics': {}
        }
        
        # Process each slide
        for slide_data in self.comprehensive_data['spatial_analysis']['slides']:
            print(f"   🧠 Processing slide {slide_data['slide_number']} for smart grouping...")
            slide_analysis = self.create_smart_groups_for_slide(slide_data['boxes'], slide_data['slide_number'])
            if slide_analysis:
                smart_groups_analysis['slides'].append(slide_analysis)
                # Create enhanced group visualizations (if enabled)
                if self.config['generate_smart_group_visualizations']:
                    self.create_enhanced_group_visualization(slide_analysis)
        
        # Calculate overall statistics
        total_slides = len(smart_groups_analysis['slides'])
        total_boxes = sum(s.get('total_boxes', 0) for s in smart_groups_analysis['slides'])
        total_groups = sum(len(s.get('smart_groups', {})) for s in smart_groups_analysis['slides'])
        
        smart_groups_analysis['overall_statistics'] = {
            'total_slides': total_slides,
            'total_boxes': total_boxes,
            'total_groups': total_groups,
            'overlap_threshold': self.overlap_threshold,
            'containment_threshold': self.containment_threshold
        }
        
        self.comprehensive_data['smart_groups'] = smart_groups_analysis
        
        # Save smart groups analysis
        smart_groups_path = self.smart_groups_dir / "smart_groups_analysis.json"
        with open(smart_groups_path, 'w', encoding='utf-8') as f:
            json.dump(smart_groups_analysis, f, ensure_ascii=False, indent=2, default=str)
        
        # Create comprehensive group content summary (if enabled)
        if self.config['generate_enhanced_group_summaries']:
            self.create_enhanced_group_content_summary(smart_groups_analysis['slides'])
        
        print(f"✅ Smart grouping applied: {total_groups} groups across {total_slides} slides")
        print(f"📊 Enhanced visualizations and summaries created")
        
        # Now create reading-order-based reordering
        self.create_reading_order_based_groups()
        
        return True



    def create_smart_groups_for_slide(self, slide_boxes, slide_number):
        """Create smart groups for a single slide (original spatial containment approach)"""
        if len(slide_boxes) < 2:
            return None
        
        # Step 1: Find non-overlapping independent groups
        independent_groups, overlap_matrix = self.find_non_overlapping_groups(slide_boxes)
        
        # Step 2: Assign remaining components to best-fit groups
        assignments = self.assign_components_to_groups(slide_boxes, independent_groups, overlap_matrix)
        
        # Step 3: Create final group structure (original approach)
        smart_groups = {}
        total_members = 0
        
        for group in independent_groups:
            group_name = group['box_id']
            members = group.get('members', [])
            
            smart_groups[group_name] = {
                'group_id': group_name,
                'root_component': group['box'],
                'root_index': group['index'],
                'members': members,
                'total_members': len(members) + 1,
                'is_empty_group': len(members) == 0,
                'member_types': self._analyze_member_types(members)
            }
            
            total_members += len(members) + 1
        
        return {
            'slide_number': slide_number,
            'total_boxes': len(slide_boxes),
            'independent_groups_count': len(independent_groups),
            'total_members': total_members,
            'assignments': assignments,
            'smart_groups': smart_groups,
            'analysis_timestamp': datetime.now().isoformat()
        }



    def find_non_overlapping_groups(self, slide_boxes):
        """Find components that don't overlap significantly"""
        n = len(slide_boxes)
        overlap_matrix = np.zeros((n, n))
        
        for i in range(n):
            for j in range(n):
                if i != j:
                    overlap_pct = calculate_overlap_percentage(slide_boxes[i], slide_boxes[j])
                    overlap_matrix[i][j] = overlap_pct
        
        independent_groups = []
        used_indices = set()
        
        # Sort by area (largest first)
        components_by_size = [(i, box) for i, box in enumerate(slide_boxes)]
        components_by_size.sort(key=lambda x: x[1]['position']['width'] * x[1]['position']['height'], reverse=True)
        
        for idx, box in components_by_size:
            if idx in used_indices:
                continue
            box_id = box.get('box_id', f'S{idx+1}')
            
            # Check conflicts with existing groups
            conflicts_with_existing = False
            for existing_group in independent_groups:
                existing_idx = existing_group['index']
                overlap_pct = overlap_matrix[idx][existing_idx]
                if overlap_pct > self.overlap_threshold:
                    conflicts_with_existing = True
                    break
            
            if not conflicts_with_existing:
                independent_groups.append({
                    'index': idx,
                    'box': box,
                    'box_id': box_id,
                    'members': []
                })
                used_indices.add(idx)
        
        return independent_groups, overlap_matrix



    def assign_components_to_groups(self, slide_boxes, independent_groups, overlap_matrix):
        """Assign remaining components to best-fit groups"""
        group_indices = set(group['index'] for group in independent_groups)
        unassigned_indices = [i for i in range(len(slide_boxes)) if i not in group_indices]
        
        assignments = []
        
        for idx in unassigned_indices:
            box = slide_boxes[idx]
            box_id = box.get('box_id', f'S{idx+1}')
            
            best_group = None
            best_containment = 0.0
            
            # Find best group
            for group in independent_groups:
                group_idx = group['index']
                containment_pct = overlap_matrix[idx][group_idx]
                if containment_pct > best_containment:
                    best_containment = containment_pct
                    best_group = group
            
            # Assign to best group if meaningful
            if best_group and best_containment >= self.containment_threshold:
                best_group['members'].append({
                    'index': idx,
                    'box': box,
                    'box_id': box_id,
                    'containment_percentage': best_containment,
                    'assignment_type': 'full' if best_containment >= 0.95 else 'partial'
                })
                assignments.append({
                    'component': box_id,
                    'assigned_to': best_group['box_id'],
                    'containment': best_containment
                })
            else:
                # Create standalone group
                standalone_group = {
                    'index': idx,
                    'box': box,
                    'box_id': box_id,
                    'members': []
                }
                independent_groups.append(standalone_group)
                assignments.append({
                    'component': box_id,
                    'assigned_to': 'standalone',
                    'containment': 0.0
                })
        
        return assignments



    def find_non_overlapping_groups_with_reading_order(self, reading_order_boxes):
        """Find components that don't overlap significantly, respecting reading order"""
        n = len(reading_order_boxes)
        overlap_matrix = np.zeros((n, n))
        
        for i in range(n):
            for j in range(n):
                if i != j:
                    overlap_pct = calculate_overlap_percentage(reading_order_boxes[i], reading_order_boxes[j])
                    overlap_matrix[i][j] = overlap_pct
        
        independent_groups = []
        used_indices = set()
        
        # Sort by reading order first, then by area (to respect content flow)
        components_by_reading_order = [(i, box) for i, box in enumerate(reading_order_boxes)]
        components_by_reading_order.sort(key=lambda x: (
            x[1].get('reading_order', 999),  # Primary: reading order
            -(x[1]['position']['width'] * x[1]['position']['height'])  # Secondary: size (larger first)
        ))
        
        for idx, box in components_by_reading_order:
            if idx in used_indices:
                continue
            box_id = box.get('box_id', f'S{idx+1}')
            
            # Check conflicts with existing groups
            conflicts_with_existing = False
            for existing_group in independent_groups:
                existing_idx = existing_group['index']
                overlap_pct = overlap_matrix[idx][existing_idx]
                if overlap_pct > self.overlap_threshold:
                    conflicts_with_existing = True
                    break
            
            if not conflicts_with_existing:
                independent_groups.append({
                    'index': idx,
                    'box': box,
                    'box_id': box_id,
                    'members': []
                })
                used_indices.add(idx)
        
        return independent_groups, overlap_matrix



    def assign_components_to_groups_with_reading_order(self, reading_order_boxes, independent_groups, overlap_matrix):
        """Assign remaining components to best-fit groups, preserving reading order"""
        group_indices = set(group['index'] for group in independent_groups)
        unassigned_indices = [i for i in range(len(reading_order_boxes)) if i not in group_indices]
        
        # Sort unassigned components by reading order
        unassigned_with_order = [(i, reading_order_boxes[i]) for i in unassigned_indices]
        unassigned_with_order.sort(key=lambda x: x[1].get('reading_order', 999))
        
        assignments = []
        
        for idx, box in unassigned_with_order:
            box_id = box.get('box_id', f'S{idx+1}')
            
            best_group = None
            best_containment = 0.0
            
            # Find best group (prefer groups with similar or earlier reading order)
            for group in independent_groups:
                group_idx = group['index']
                containment_pct = overlap_matrix[idx][group_idx]
                if containment_pct >= self.containment_threshold:
                    # Bonus for reading order proximity
                    group_reading_order = group['box'].get('reading_order', 999)
                    box_reading_order = box.get('reading_order', 999)
                    order_proximity_bonus = max(0, 1 - abs(group_reading_order - box_reading_order) / 10)
                    
                    adjusted_score = containment_pct + (order_proximity_bonus * 0.1)
                    
                    if adjusted_score > best_containment:
                        best_containment = containment_pct  # Keep original containment for records
                        best_group = group
            
            # Assign to best group if meaningful
            if best_group and best_containment >= self.containment_threshold:
                best_group['members'].append({
                    'index': idx,
                    'box': box,
                    'box_id': box_id,
                    'containment_percentage': best_containment,
                    'assignment_type': 'full' if best_containment >= 0.95 else 'partial'
                })
                assignments.append({
                    'component': box_id,
                    'assigned_to': best_group['box_id'],
                    'containment': best_containment,
                    'reading_order': box.get('reading_order', 999)
                })
            else:
                # Create standalone group (maintaining reading order)
                standalone_group = {
                    'index': idx,
                    'box': box,
                    'box_id': box_id,
                    'members': []
                }
                independent_groups.append(standalone_group)
                assignments.append({
                    'component': box_id,
                    'assigned_to': 'standalone',
                    'containment': 0.0,
                    'reading_order': box.get('reading_order', 999)
                })
        
        return assignments



    def find_spatial_proximity_groups(self, elements):
        """Group elements using connected components - elements connected through intermediate elements"""
        if len(elements) < 2:
            return [[element] for element in elements]
        
        # Build adjacency graph - which elements are close to each other
        adjacency = {}
        proximity_threshold = 900000  # PowerPoint units (~630pt, increased to capture vertical table structures)
        
        for i, element in enumerate(elements):
            adjacency[i] = []
            element_pos = element['position']
            element_center_x = element_pos['left'] + element_pos['width'] // 2
            element_center_y = element_pos['top'] + element_pos['height'] // 2
            
            for j, other_element in enumerate(elements):
                if i == j:
                    continue
                    
                other_pos = other_element['position']
                other_center_x = other_pos['left'] + other_pos['width'] // 2
                other_center_y = other_pos['top'] + other_pos['height'] // 2
                # Calculate distance between centers
                distance = ((element_center_x - other_center_x)**2 + (element_center_y - other_center_y)**2)**0.5
                if distance < proximity_threshold:
                    adjacency[i].append(j)
        
        # Find connected components using depth-first search
        visited = set()
        groups = []
        
        def dfs(node, current_group):
            if node in visited:
                return
            visited.add(node)
            current_group.append(elements[node])
            
            for neighbor in adjacency[node]:
                dfs(neighbor, current_group)
        
        for i in range(len(elements)):
            if i not in visited:
                current_group = []
                dfs(i, current_group)
                groups.append(current_group)
        
        return groups



    def create_smart_visual_groups(self, elements, line_elements=None):
        """SMART VISUAL CONSOLIDATION: Only create groups when visual elements (images/charts/groups) overlap. Ignore standalone text/images."""
        
        # Add line_elements to the main elements list
        if line_elements is None:
            line_elements = []
        all_elements = elements + line_elements
        
        # Separate visual elements that can form groups
        images = [elem for elem in all_elements if elem.get('shape_type') in ['Picture', 'Image'] and 'position' in elem]
        tables = [elem for elem in all_elements if elem.get('shape_type') == 'Table' and 'position' in elem]
        
        # Enhanced AutoShape categorization - focus on visual components only
        autoshapes_visual = []  # Only visual autoshapes (charts, shapes, short labels)
        
        for elem in all_elements:
            if elem.get('shape_type') == 'AutoShape' and 'position' in elem:
                text = elem.get('text', '').strip()
                if not text:
                    # No text - treat as visual component
                    autoshapes_visual.append(elem)
                elif len(text) <= 15 and ('<' in text or text.isupper() or text.isdigit() or any(char in text for char in ['@', 'Hz', 'mm', 'kg', '%'])):
                    # Short labels like <A9_1>, @38Hz, technical labels - treat as visual components
                    autoshapes_visual.append(elem)
                    print(f"      🏷️  AutoShape {elem['box_id']} with technical label '{text}' - treating as visual component")
                # Longer text autoshapes are ignored (not grouped)
        
        # Lines and other elements
        lines = line_elements
        others = [elem for elem in all_elements
                  if elem.get('shape_type') not in ['Picture', 'Image', 'Text', 'AutoShape', 'Table', 'Line'] 
                  and 'position' in elem]
        
        # VISUAL ELEMENTS that can form groups: images + visual autoshapes + tables + others
        visual_elements = images + autoshapes_visual + tables + others
        
        print(f"      🎯 SMART VISUAL CONSOLIDATION: {len(visual_elements)} visual elements ({len(images)} images, {len(tables)} tables, {len(autoshapes_visual)} autoshapes, {len(others)} others)")
        
        if len(visual_elements) < 2:
            print(f"      ⚪ Less than 2 visual elements - no groups possible")
            return elements  # Need at least 2 visual elements to form a group
        
        # Find overlapping visual elements to create groups
        consolidated_elements = []
        processed_visual = set()
        
        for i, primary_visual in enumerate(visual_elements):
            if id(primary_visual) in processed_visual:
                continue
                
            # Find all visual elements that overlap with this one
            overlapping_visuals = [primary_visual]
            processed_visual.add(id(primary_visual))
            
            for j, other_visual in enumerate(visual_elements):
                if i != j and id(other_visual) not in processed_visual:
                    # Check if they overlap
                    overlap_ratio = calculate_spatial_overlap(primary_visual['position'], other_visual['position'])
                    if overlap_ratio > 0.05:  # Even small overlap counts for visual elements
                        overlapping_visuals.append(other_visual)
                        processed_visual.add(id(other_visual))
                        print(f"      🔗 Visual overlap: {primary_visual['box_id']} ↔ {other_visual['box_id']} (overlap: {overlap_ratio:.2f})")
            
            # Only create a group if we have 2+ overlapping visual elements
            if len(overlapping_visuals) >= 2:
                # Find any texts and lines that overlap with the visual group boundary
                combined_boundary = calculate_merged_boundary([v['position'] for v in overlapping_visuals])
                
                overlapping_texts = []
                overlapping_lines = []
                
                # Check for text overlaps (for inclusion in the group, but text alone doesn't create groups)
                for text in all_elements:
                    if text.get('shape_type') == 'Text' and 'position' in text:
                        text_overlap = calculate_spatial_overlap(combined_boundary, text['position'])
                        if text_overlap > 0.8:  # High overlap threshold for text inclusion
                            overlapping_texts.append(text)
                            print(f"      📝 Text inclusion: {text['box_id']} overlaps with visual group (overlap: {text_overlap:.2f})")
                
                # Check for line overlaps
                for line in lines:
                    if 'position' in line:
                        line_overlap = calculate_spatial_overlap(combined_boundary, line['position'])
                        if line_overlap > 0.3:  # Lower threshold for lines
                            overlapping_lines.append(line)
                            print(f"      📏 Line inclusion: {line['box_id']} overlaps with visual group (overlap: {line_overlap:.2f})")
                
                # Create unified group entity
                unified_group = self.create_smart_visual_group_entity(
                    overlapping_visuals, 
                    overlapping_texts, 
                    overlapping_lines,
                    combined_boundary
                )
                consolidated_elements.append(unified_group)
                print(f"      ✅ Created visual group {unified_group['box_id']}: {len(overlapping_visuals)} visuals, {len(overlapping_texts)} texts, {len(overlapping_lines)} lines")
            
            # If only 1 visual element, it remains as standalone (not grouped)
        
        # Add remaining non-visual elements as individuals (texts, standalone images, etc.)
        for elem in all_elements:
            if id(elem) not in processed_visual:
                # Skip lines that were used in groups
                if elem in line_elements:
                    used_in_group = any(elem in group.get('component_lines', []) for group in consolidated_elements if group.get('box_type') == 'smart_visual_group')
                    if used_in_group:
                        continue
                consolidated_elements.append(elem)
        
        print(f"      🎯 Smart visual consolidation result: {len(consolidated_elements)} total elements")
        return consolidated_elements



    def create_smart_visual_group_entity(self, visual_elements, texts, lines, merged_boundary):
        """Create a smart visual group entity from overlapping visual elements"""
        self.unified_group_counter += 1
        group_id = f"G{self.unified_group_counter}"
        
        # Combine all text content
        all_text_parts = []
        all_text_parts.extend([text.get('text', '').strip() for text in texts if text.get('text', '').strip()])
        all_text_parts.extend([visual.get('text', '').strip() for visual in visual_elements if visual.get('text', '').strip()])
        combined_text = " ".join(all_text_parts)
        
        entity = {
            'box_id': group_id,
            'box_type': 'smart_visual_group',
            'shape_type': 'SmartVisualGroup',
            'position': merged_boundary,
            'text': combined_text,
            'component_visuals': visual_elements,  # Primary visual elements that define the group
            'component_texts': texts,             # Associated text elements
            'component_lines': lines,             # Associated line elements
            'total_components': len(visual_elements) + len(texts) + len(lines),
            'visual_component_count': len(visual_elements)  # Key metric for group formation
        }
        
        return entity



    def create_unified_groups(self, elements, line_elements=None):
        """UNIFIED CONSOLIDATION: Merge overlapping images + 80% overlapping text + lines into single groups"""

        # Reset per-slide membership tracking. box_ids (S0, S1, S2…) are
        # re-numbered from zero on every slide, so leaving stale entries from
        # earlier slides causes should_capture_shape() to skip Pictures on later
        # slides that happen to share an index with a previously-grouped shape.
        # Group counters (unified_group_counter, _next_group_id) stay live so
        # G/UG IDs remain globally unique within the file.
        self.unified_group_members = {}

        # Add line_elements to the main elements list
        if line_elements is None:
            line_elements = []
        all_elements = elements + line_elements
        
        # Separate by type (with position validation)
        images = [elem for elem in all_elements if elem.get('shape_type') in ['Picture', 'Image'] and 'position' in elem]
        texts = [elem for elem in all_elements if elem.get('shape_type') in ['Text', 'TextBox'] and 'position' in elem]
        lines = line_elements  # Lines are already properly identified
        
        # Enhanced AutoShape categorization
        autoshapes_no_text = []
        autoshapes_short_labels = []  # For labels like <A9_1>, <A8_-1>
        autoshapes_long_text = []     # For actual text content
        
        for elem in all_elements:
            if elem.get('shape_type') == 'AutoShape' and 'position' in elem:
                text = elem.get('text', '').strip()
                if not text:
                    # No text - treat as visual component
                    autoshapes_no_text.append(elem)
                elif len(text) <= 10 and ('<' in text or text.isupper() or text.isdigit()):
                    # Short labels like <A9_1>, <A8_-1>, numbers - treat as visual components
                    autoshapes_short_labels.append(elem)
                    print(f"      🏷️  AutoShape {elem['box_id']} with short label '{text}' - treating as visual component")
                else:
                    # Longer text - treat as text
                    autoshapes_long_text.append(elem)
        
        # Combine autoshapes for visual overlap checking
        autoshapes_visual = autoshapes_no_text + autoshapes_short_labels
        
        # Add long text autoshapes to texts
        texts.extend(autoshapes_long_text)
        
        others = [elem for elem in all_elements if elem.get('shape_type') not in ['Picture', 'Image', 'Text', 'TextBox', 'Line', 'AutoShape'] or 'position' not in elem]
        
        # Debug: Check for elements without position and line detection
        no_position = [elem for elem in all_elements if 'position' not in elem]
        if no_position:
            print(f"      ⚠️  {len(no_position)} elements without position field - moving to others")
            for elem in no_position[:3]:  # Show first 3
                print(f"         🔍 {elem.get('box_id', 'unknown')}: {elem.get('shape_type', 'unknown')} - {list(elem.keys())}")
            others.extend(no_position)
        
        # Debug: Show line detection specifically
        if lines:
            print(f"      🔍 DEBUG - Line elements received: {len(lines)}")
            for line_elem in lines:
                print(f"         📏 {line_elem.get('box_id')}: {line_elem.get('shape_type', 'Line')} - position: {'position' in line_elem}")
        else:
            print(f"      ⚠️  DEBUG - No line elements received")
            
        # Debug: Show all shape types and check S43, S44 specifically
        shape_types = {}
        s43_found = None
        s44_found = None
        for elem in all_elements:
            shape_type = elem.get('shape_type', 'unknown')
            shape_types[shape_type] = shape_types.get(shape_type, 0) + 1
            if elem.get('box_id') == 'S43':
                s43_found = elem
            elif elem.get('box_id') == 'S44':
                s44_found = elem
        
        print(f"      🔍 DEBUG - All shape types: {shape_types}")
        
        if s43_found:
            print(f"      🔍 FOUND S43: shape_type='{s43_found.get('shape_type')}', text='{s43_found.get('text', '')}', has_text={bool(s43_found.get('text', '').strip())}")
        else:
            print(f"      ❌ S43 NOT FOUND in all_elements")
            
        if s44_found:
            print(f"      🔍 FOUND S44: shape_type='{s44_found.get('shape_type')}', text='{s44_found.get('text', '')}', has_text={bool(s44_found.get('text', '').strip())}")
        else:
            print(f"      ❌ S44 NOT FOUND in all_elements")
        
        if len(images) == 0:
            return elements  # No images to consolidate
            
        print(f"      🎯 UNIFIED CONSOLIDATION: {len(images)} images, {len(texts)} texts, {len(lines)} lines, {len(autoshapes_visual)} autoshapes")
        
        # Debug: Show first few image elements to understand structure
        if images:
            print(f"      🔍 DEBUG - First image element structure:")
            first_img = images[0]
            print(f"         📋 {first_img.get('box_id')}: {list(first_img.keys())}")
            if 'position' in first_img:
                pos = first_img['position']
                print(f"         🎯 Position: {type(pos)} - {pos}")
            else:
                print(f"         ❌ No 'position' field!")
        
        # Step 1: Find overlapping image groups (including transitive connectivity through lines)
        image_groups = []
        used_images = set()
        
        for i, img1 in enumerate(images):
            if i in used_images:
                continue
            # Start a new group with this image
            current_group = [img1]
            used_images.add(i)
            print(f"      🔍 Starting new image group with {img1['box_id']}")
            
            # Find all images that overlap with any image in current group (direct or through lines)
            changed = True
            while changed:
                changed = False
                # Check direct image-to-image overlaps
                for j, img2 in enumerate(images):
                    if j in used_images:
                        continue
                    
                    # Check if img2 overlaps with any image in current group
                    for img_in_group in current_group:
                        try:
                            overlap_pct = calculate_spatial_overlap(img2['position'], img_in_group['position'])
                            if overlap_pct > 0:
                                print(f"      ✅ Direct overlap: {img2['box_id']} overlaps with {img_in_group['box_id']} ({overlap_pct:.1f}%)")
                                current_group.append(img2)
                                used_images.add(j)
                                changed = True
                                break
                        except Exception as e:
                            print(f"      ❌ Error checking overlap between {img2.get('box_id')} and {img_in_group.get('box_id')}: {e}")
                            continue
                # Check transitive connections through lines
                print(f"      🔍 Checking transitive connections for group with {[img['box_id'] for img in current_group]}")
                for j, img2 in enumerate(images):
                    if j in used_images:
                        print(f"         ⏭️  Skipping {img2['box_id']} (already used)")
                        continue
                    
                    print(f"         🔍 Checking if {img2['box_id']} can connect to group")
                    
                    # Check if img2 is connected to any image in current group through a shared line
                    for img_in_group in current_group:
                        connected_through_line = False
                        print(f"         🔍 Testing connection: {img_in_group['box_id']} <-> {img2['box_id']} through lines")
                        
                        for line in lines:
                            try:
                                # Check if line overlaps with both images
                                overlap1 = calculate_spatial_overlap(line['position'], img_in_group['position'])
                                overlap2 = calculate_spatial_overlap(line['position'], img2['position'])
                                
                                # If no overlap, check proximity (within 50000 EMUs ≈ 2mm)
                                proximity_threshold = 50000
                                close_enough1 = overlap1 > 0 or calculate_proximity(line['position'], img_in_group['position']) < proximity_threshold
                                close_enough2 = overlap2 > 0 or calculate_proximity(line['position'], img2['position']) < proximity_threshold
                                
                                print(f"         📏 Line {line['box_id']}: {img_in_group['box_id']}({overlap1:.1f}%) <-> {img2['box_id']}({overlap2:.1f}%)")
                                if overlap1 == 0 or overlap2 == 0:
                                    prox1 = calculate_proximity(line['position'], img_in_group['position'])
                                    prox2 = calculate_proximity(line['position'], img2['position'])
                                    print(f"            🔍 Proximity check: {img_in_group['box_id']}({prox1:.0f} EMUs) <-> {img2['box_id']}({prox2:.0f} EMUs)")
                                
                                if (overlap1 >= 5.0 or close_enough1) and (overlap2 >= 5.0 or close_enough2):
                                    print(f"      🌉 TRANSITIVE CONNECTION: {img_in_group['box_id']} <- Line {line['box_id']} -> {img2['box_id']}")
                                    print(f"         Line-Image1 overlap: {overlap1:.1f}%, Line-Image2 overlap: {overlap2:.1f}%")
                                    current_group.append(img2)
                                    used_images.add(j)
                                    changed = True
                                    connected_through_line = True
                                    break
                            except Exception as e:
                                print(f"      ❌ Error checking line connection: {e}")
                                continue
                        
                        if connected_through_line:
                            break
            
            image_groups.append(current_group)
            print(f"      📦 Final group: {[img['box_id'] for img in current_group]}")
        
        # Step 2: For each image group, use incremental boundary expansion
        consolidated_elements = []
        
        for group_images in image_groups:
            print(f"      🔄 INCREMENTAL CONSOLIDATION for images: {[img['box_id'] for img in group_images]}")
            
            # Start with the image group
            current_elements = group_images[:]
            overlapping_texts = []
            overlapping_lines = []  
            overlapping_autoshapes = []
            
            # Keep expanding until no more elements can be added
            expansion_iteration = 0
            while True:
                expansion_iteration += 1
                print(f"      🔄 Expansion iteration {expansion_iteration}")
                # Calculate current merged boundary from all current elements
                all_positions = []
                for elem in current_elements + overlapping_lines + overlapping_autoshapes:
                    if 'position' in elem:
                        all_positions.append(elem['position'])
                if not all_positions:
                    break
                    
                current_boundary = calculate_merged_boundary(all_positions)
                print(f"      📐 Current boundary: left={current_boundary['left']}, top={current_boundary['top']}, right={current_boundary['left']+current_boundary['width']}, bottom={current_boundary['top']+current_boundary['height']}")
                # Track what gets added in this iteration
                new_elements_added = False
                new_images = []
                new_lines = []
                new_autoshapes = []
                new_texts = []
                # Check for new images that overlap with current boundary
                for img in images:
                     if img in current_elements:
                         continue  # Already included
                     try:
                         overlap_pct = calculate_spatial_overlap(img['position'], current_boundary)
                         print(f"      🔍 Image {img['box_id']} vs current boundary: {overlap_pct:.1f}%")
                         
                         if overlap_pct >= 5.0:
                             print(f"      ✅ Adding image {img['box_id']} to consolidated group ({overlap_pct:.1f}%)")
                             new_images.append(img)
                             new_elements_added = True
                         elif overlap_pct == 0.0:
                             # Check proximity for very small gaps (≤ 20,000 EMUs ≈ 0.7mm)
                             proximity = calculate_proximity(img['position'], current_boundary)
                             print(f"      🔍 Proximity check: {img['box_id']} <-> boundary = {proximity:.0f} EMUs")
                             if proximity <= 20000:  # 20,000 EMUs ≈ 0.7mm tolerance
                                 print(f"      ✅ Adding image {img['box_id']} to consolidated group (proximity: {proximity:.0f} EMUs)")
                                 new_images.append(img)
                                 new_elements_added = True
                         
                     except Exception as e:
                         print(f"      ❌ Error checking image {img.get('box_id')}: {e}")
                # Check for new lines that overlap with current boundary  
                for line in lines:
                    if line in overlapping_lines:
                        continue  # Already included
                    try:
                        overlap_pct = calculate_spatial_overlap(line['position'], current_boundary)
                        print(f"      🔍 Line {line['box_id']} vs current boundary: {overlap_pct:.1f}%")
                        if overlap_pct >= 5.0:
                            print(f"      ✅ Adding line {line['box_id']} to consolidated group ({overlap_pct:.1f}%)")
                            new_lines.append(line)
                            new_elements_added = True
                    except Exception as e:
                        print(f"      ❌ Error checking line {line.get('box_id')}: {e}")
                # Check for new autoshapes that overlap with current boundary.
                # Reject autoshapes that engulf the boundary (slide-wide backgrounds /
                # decorative frames whose bbox is much larger than the picture cluster) —
                # otherwise they expand the group to swallow unrelated titles/headers.
                for autoshape in autoshapes_visual:
                    if autoshape in overlapping_autoshapes:
                        continue  # Already included
                    try:
                        overlap_pct = calculate_spatial_overlap(autoshape['position'], current_boundary)
                        if overlap_pct < 5.0:
                            continue
                        ash_pos = autoshape.get('position', {})
                        ash_area = ash_pos.get('width', 0) * ash_pos.get('height', 0)
                        boundary_area = current_boundary.get('width', 0) * current_boundary.get('height', 0)
                        if ash_area > 0 and boundary_area > 0:
                            # How much of the autoshape sits inside the boundary.
                            # If the autoshape is much larger than the boundary it likely
                            # represents a slide-wide background, not a cluster member.
                            l = max(ash_pos['left'], current_boundary['left'])
                            t = max(ash_pos['top'], current_boundary['top'])
                            r = min(ash_pos['left'] + ash_pos['width'], current_boundary['left'] + current_boundary['width'])
                            b = min(ash_pos['top'] + ash_pos['height'], current_boundary['top'] + current_boundary['height'])
                            inter = max(0, r - l) * max(0, b - t)
                            ash_coverage = inter / ash_area
                            if ash_coverage < 0.5:
                                print(f"      🚫 Skipping engulfing autoshape {autoshape['box_id']} (only {ash_coverage*100:.1f}% of it inside boundary; likely background)")
                                continue
                        print(f"      ✅ Adding autoshape {autoshape['box_id']} to consolidated group ({overlap_pct:.1f}%)")
                        new_autoshapes.append(autoshape)
                        new_elements_added = True
                    except Exception as e:
                        print(f"      ❌ Error checking autoshape {autoshape.get('box_id')}: {e}")
                # Check for new texts that overlap significantly with current boundary
                for text in texts:
                    if text in overlapping_texts:
                        continue  # Already included
                    try:
                        overlap_pct = calculate_spatial_overlap(text['position'], current_boundary)
                        if overlap_pct >= 30.0:  # 30% threshold for text
                            print(f"      ✅ Adding text {text['box_id']} to consolidated group ({overlap_pct:.1f}%)")
                            new_texts.append(text)
                            new_elements_added = True
                    except Exception as e:
                        print(f"      ❌ Error checking text {text.get('box_id')}: {e}")
                # Add new elements to the growing lists
                current_elements.extend(new_images)
                overlapping_lines.extend(new_lines)
                overlapping_autoshapes.extend(new_autoshapes)
                overlapping_texts.extend(new_texts)
                print(f"      📊 After iteration {expansion_iteration}: {len(current_elements)} images, {len(overlapping_lines)} lines, {len(overlapping_autoshapes)} autoshapes, {len(overlapping_texts)} texts")
                # If nothing was added, we're done expanding
                if not new_elements_added:
                    print(f"      🏁 No more elements to add - consolidation complete")
                    break
                # Safety check to prevent infinite loops
                if expansion_iteration > 10:
                    print(f"      ⚠️  Max iterations reached - stopping expansion")
                    break
            
            # After expansion is complete, create the final unified group entity
            all_overlapping_components = overlapping_texts + overlapping_lines + overlapping_autoshapes
            
            if len(current_elements) > 1 or all_overlapping_components:
                # Calculate final merged boundary from all elements
                final_positions = []
                for elem in current_elements + overlapping_lines + overlapping_autoshapes:
                    if 'position' in elem:
                        final_positions.append(elem['position'])
                if final_positions:
                    merged_boundary = calculate_merged_boundary(final_positions)
                    print(f"      🎯 Final merged boundary: left={merged_boundary['left']}, top={merged_boundary['top']}, right={merged_boundary['left']+merged_boundary['width']}, bottom={merged_boundary['top']+merged_boundary['height']}")
                    # Create unified group entity
                    unified_group = self.create_unified_group_entity(
                        current_elements, 
                        overlapping_texts, 
                        overlapping_lines, 
                        overlapping_autoshapes,
                        merged_boundary
                    )
                    consolidated_elements.append(unified_group)
                    
                    # Remove used components from pools
                    texts = [t for t in texts if t not in overlapping_texts]
                    lines = [l for l in lines if l not in overlapping_lines]
                    autoshapes_visual = [a for a in autoshapes_visual if a not in overlapping_autoshapes]
                    images = [img for img in images if img not in current_elements]
                else:
                    # Fallback - add images individually
                    consolidated_elements.extend(current_elements)
                    images = [img for img in images if img not in current_elements]
            else:
                # Single image with no overlapping components - keep as is
                consolidated_elements.extend(current_elements)
                images = [img for img in images if img not in current_elements]


        
        # Add remaining components that weren't absorbed into unified groups
        if not hasattr(self, 'unified_group_members'):
            self.unified_group_members = {}
        
        # Filter out absorbed components
        remaining_texts = [t for t in texts if t.get('box_id') not in self.unified_group_members]
        remaining_lines = [l for l in lines if l.get('box_id') not in self.unified_group_members]
        remaining_autoshapes = [a for a in autoshapes_visual if a.get('box_id') not in self.unified_group_members]
        remaining_others = [o for o in others if o.get('box_id') not in self.unified_group_members]
        
        consolidated_elements.extend(remaining_texts)
        consolidated_elements.extend(remaining_lines)
        consolidated_elements.extend(remaining_autoshapes)
        consolidated_elements.extend(remaining_others)
        
        # Debug: Show what was filtered out
        absorbed_count = len(texts) - len(remaining_texts) + len(lines) - len(remaining_lines) + len(autoshapes_visual) - len(remaining_autoshapes)
        if absorbed_count > 0:
            print(f"      🔄 Filtered out {absorbed_count} components absorbed into unified groups")
        
        unified_count = len([e for e in consolidated_elements if e.get('box_type') == 'unified_group'])
        print(f"      ✅ Created {unified_count} unified groups from {len(images)} images")
        
        return consolidated_elements



    def create_unified_group_entity(self, images, texts, lines=None, autoshapes=None, merged_boundary=None):
        """Create a unified group entity from images, texts, lines, and autoshapes"""
        self.unified_group_counter += 1
        group_id = f"G{self.unified_group_counter}"
        
        # Default empty lists if not provided
        if lines is None:
            lines = []
        if autoshapes is None:
            autoshapes = []
        
        # Use merged boundary or calculate from all components
        if merged_boundary is None:
            all_positions = [img['position'] for img in images]
            all_positions.extend([text['position'] for text in texts])
            all_positions.extend([line['position'] for line in lines])
            all_positions.extend([autoshape['position'] for autoshape in autoshapes])
            merged_boundary = calculate_merged_boundary(all_positions)
        
        # Combine all text content (from texts and autoshapes with text)
        combined_text = " ".join([text.get('text', '').strip() for text in texts if text.get('text', '').strip()])
        
        entity = {
            'box_id': group_id,
            'box_type': 'unified_group',
            'shape_type': 'UnifiedGroup',
            'position': merged_boundary,
            'text': combined_text,
            'component_images': images,
            'component_texts': texts,
            'component_lines': lines,
            'component_autoshapes': autoshapes,
            'total_components': len(images) + len(texts) + len(lines) + len(autoshapes)
        }
        
        # Track which components are part of this unified group (to avoid duplicate individual processing)
        if not hasattr(self, 'unified_group_members'):
            self.unified_group_members = {}
        
        # Track images for capture avoidance
        for image in images:
            image_id = image.get('box_id')
            if image_id:
                self.unified_group_members[image_id] = group_id
                print(f"         📝 Tracking image {image_id} as part of {group_id}")
        
        # Track texts for reading order avoidance
        for text in texts:
            text_id = text.get('box_id')
            if text_id:
                self.unified_group_members[text_id] = group_id
                print(f"         📝 Tracking text {text_id} as part of {group_id}")
        
        # Track lines and autoshapes
        for line in lines:
            line_id = line.get('box_id')
            if line_id:
                self.unified_group_members[line_id] = group_id
                print(f"         📝 Tracking line {line_id} as part of {group_id}")
        
        for autoshape in autoshapes:
            autoshape_id = autoshape.get('box_id')
            if autoshape_id:
                self.unified_group_members[autoshape_id] = group_id
                print(f"         📝 Tracking autoshape {autoshape_id} as part of {group_id}")
        
        print(f"      🔍 DEBUG - Created {group_id} entity:")
        print(f"         📐 Entity position: left={merged_boundary['left']}, top={merged_boundary['top']}, right={merged_boundary['left']+merged_boundary['width']}, bottom={merged_boundary['top']+merged_boundary['height']}")
        print(f"         📊 Components: {len(images)} images, {len(texts)} texts, {len(lines)} lines, {len(autoshapes)} autoshapes")
        
        return entity



    def capture_unified_group_visual(self, unified_group, slide_number):
        """Capture visual for unified group (re-enabled per user request)"""
        if slide_number not in self.slide_images:
            print(f"      ⚠️  No slide image available for slide {slide_number} - skipping unified group capture")
            return None
        
        box_id = unified_group['box_id']
        box_type = unified_group.get('box_type', 'unified_group')
        
        print(f"      📸 Capturing unified group {box_id} ({box_type})")
        
        try:
            slide_image = self.slide_images[slide_number]
            position = unified_group['position']
            
            print(f"      🔍 DEBUG - Visual capture using position:")
            print(f"         📐 Position: left={position['left']}, top={position['top']}, right={position['left']+position['width']}, bottom={position['top']+position['height']}")
            print(f"         📏 Width={position['width']}, Height={position['height']}")
            
            # Convert PowerPoint coordinates to image pixels
            img_width, img_height = slide_image.size
            # Get actual slide dimensions from metadata
            ppt_width = self.comprehensive_data.get("metadata", {}).get("slide_width", 9144000)
            ppt_height = self.comprehensive_data.get("metadata", {}).get("slide_height", 6858000)
            
            # Convert coordinates
            x_scale = img_width / ppt_width
            y_scale = img_height / ppt_height
            
            left = int(position['left'] * x_scale)
            top = int(position['top'] * y_scale)
            width = int(position['width'] * x_scale)
            height = int(position['height'] * y_scale)
            
            # Add padding
            padding = 10
            left = max(0, left - padding)
            top = max(0, top - padding)
            right = min(img_width, left + width + 2*padding)
            bottom = min(img_height, top + height + 2*padding)
            
            print(f"      🔍 DEBUG - Final crop coordinates:")
            print(f"         📐 Image size: {img_width}x{img_height} pixels")
            print(f"         ✂️  Crop bounds: left={left}, top={top}, right={right}, bottom={bottom}")
            print(f"         📏 Crop size: {right-left}x{bottom-top} pixels")
            
            # Ensure valid bounds
            if right <= left or bottom <= top:
                print(f"      ⚠️  Invalid crop bounds for {box_id}")
                return None
            
            # Crop the unified group region
            cropped_region = slide_image.crop((left, top, right, bottom))
            
            # Add green border to distinguish unified group captures
            from PIL import ImageOps
            bordered_region = ImageOps.expand(cropped_region, border=4, fill='green')
            
            # Save the unified group capture
            capture_filename = f"slide_{slide_number:02d}_{box_id.lower()}_visual.png"
            capture_path = self.visual_captures_dir / capture_filename
            bordered_region.save(capture_path, 'PNG', dpi=(600, 600))
            
            # Add to comprehensive data visual captures list
            capture_info = {
                'slide_number': slide_number,
                'box_id': box_id,
                'box_type': box_type,
                'shape_index': unified_group.get('shape_index', 0),
                'filename': capture_filename,
                'filepath': str(capture_path),
                'position': position,
                'total_components': unified_group.get('total_components', 0)
            }
            
            if 'visual_captures' not in self.comprehensive_data:
                self.comprehensive_data['visual_captures'] = []
            self.comprehensive_data['visual_captures'].append(capture_info)
            
            print(f"         ✅ Saved unified group capture: {capture_filename}")
            return capture_filename
            
        except Exception as e:
            print(f"      ❌ Error capturing unified group {box_id}: {e}")
            return None



    def consolidate_overlapping_images(self, elements):
        """Consolidate overlapping images into single entities with combined boundaries"""
        images = [elem for elem in elements if elem.get('shape_type') in ['Picture', 'Image']]
        non_images = [elem for elem in elements if elem.get('shape_type') not in ['Picture', 'Image']]
        
        if len(images) < 2:
            return elements  # No consolidation needed
            
        print(f"      🖼️  Found {len(images)} images - checking for overlaps")
        
        # Find overlapping image groups using connected components
        overlap_groups = self.find_overlapping_image_groups(images)
        
        consolidated_images = []
        processed_images = set()
        
        for group in overlap_groups:
            if len(group) > 1:
                # Multiple overlapping images - consolidate them
                consolidated_image = self.create_consolidated_image_entity(group)
                consolidated_images.append(consolidated_image)
                processed_images.update(img['box_id'] for img in group)
                print(f"      🖼️  Consolidated {len(group)} overlapping images into {consolidated_image['box_id']}")
            else:
                # Single image - keep as is
                image = group[0]
                if image['box_id'] not in processed_images:
                    consolidated_images.append(image)
                    processed_images.add(image['box_id'])
        
        # Add any remaining unprocessed images
        for image in images:
            if image['box_id'] not in processed_images:
                consolidated_images.append(image)
        
        # Combine consolidated images with non-image elements
        final_elements = consolidated_images + non_images
        print(f"      🖼️  Image consolidation: {len(images)} → {len(consolidated_images)} images")
        
        return final_elements



    def find_overlapping_image_groups(self, images):
        """Find groups of overlapping images using connected components clustering"""
        if len(images) < 2:
            return [[img] for img in images]
        
        # Build adjacency list for overlapping images
        adjacency = {i: [] for i in range(len(images))}
        
        for i in range(len(images)):
            for j in range(i + 1, len(images)):
                overlap_pct = calculate_spatial_overlap(
                    images[i]['position'], 
                    images[j]['position']
                )
                if overlap_pct > 30.0:  # 30% overlap threshold
                    adjacency[i].append(j)
                    adjacency[j].append(i)
                    print(f"      🖼️  {images[i]['box_id']} overlaps {images[j]['box_id']} by {overlap_pct:.1f}%")
        
        # Find connected components using DFS
        visited = [False] * len(images)
        groups = []
        
        def dfs(node, current_group):
            visited[node] = True
            current_group.append(images[node])
            for neighbor in adjacency[node]:
                if not visited[neighbor]:
                    dfs(neighbor, current_group)
        
        for i in range(len(images)):
            if not visited[i]:
                group = []
                dfs(i, group)
                groups.append(group)
        
        return groups



    def create_text_image_hybrids(self, elements):
        """Create hybrid entities: text overlapping with images → preserve text, capture image separately"""
        print("   🔗 Creating text+image hybrid entities...")
        
        # Separate elements by type
        text_elements = [elem for elem in elements if elem.get('shape_type') in ['TextBox', 'AutoShape'] and elem.get('text', '').strip()]
        image_elements = [elem for elem in elements if elem.get('shape_type') in ['Picture', 'ConsolidatedImage'] or elem.get('box_type') in ['consolidated_image', 'image_autoshape_combo']]
        other_elements = [elem for elem in elements if elem not in text_elements and elem not in image_elements]
        
        if not text_elements or not image_elements:
            print(f"      🔗 Found {len(text_elements)} text and {len(image_elements)} image elements - no hybrid creation needed")
            return elements
        
        hybrid_elements = []
        processed_text_ids = set()
        processed_image_ids = set()
        hybrid_count = 0
        
        # Check each text element for overlap with image elements
        for text_elem in text_elements:
            if text_elem['box_id'] in processed_text_ids:
                continue
            overlapping_images = []
            for image_elem in image_elements:
                if image_elem['box_id'] not in processed_image_ids:
                    overlap_pct = calculate_spatial_overlap(text_elem['position'], image_elem['position'])
                    if overlap_pct > 80.0:  # 80% overlap threshold
                        overlapping_images.append(image_elem)
            
            if overlapping_images:
                # Create hybrid entity - preserve text, reference image captures
                hybrid_entity = self.create_text_image_hybrid_entity(text_elem, overlapping_images)
                hybrid_elements.append(hybrid_entity)
                processed_text_ids.add(text_elem['box_id'])
                for img in overlapping_images:
                    processed_image_ids.add(img['box_id'])
                hybrid_count += 1
                image_ids = [img['box_id'] for img in overlapping_images]
                print(f"      🔗 Created hybrid: text {text_elem['box_id']} + images [{', '.join(image_ids)}] → {hybrid_entity['box_id']}")
            else:
                # No overlapping images, keep text as is
                hybrid_elements.append(text_elem)
        
        # Add non-processed images and other elements
        for image_elem in image_elements:
            if image_elem['box_id'] not in processed_image_ids:
                hybrid_elements.append(image_elem)
        
        hybrid_elements.extend(other_elements)
        
        print(f"      🔗 Created {hybrid_count} text+image hybrid entities")
        return hybrid_elements



    def create_text_image_hybrid_entity(self, text_element, overlapping_images):
        """Create a hybrid entity that preserves text but references image captures"""
        # Use text element as base (preserve text content and position)
        hybrid_entity = text_element.copy()
        
        hybrid_entity['box_type'] = 'text_image_hybrid'
        hybrid_entity['hybrid_images'] = [img['box_id'] for img in overlapping_images]
        hybrid_entity['has_visual_capture'] = True
        
        return hybrid_entity



    def super_consolidate_overlapping_entities(self, elements):
        """Super-consolidate: CI+CI, IA+IA, CI+IA, +text → mega entities"""
        print("   🚀 Super-consolidating overlapping consolidated entities...")
        
        # Separate elements by type
        consolidated_entities = [elem for elem in elements if elem.get('box_type') in ['consolidated_image', 'image_autoshape_combo']]
        text_elements = [elem for elem in elements if elem.get('shape_type') in ['TextBox', 'AutoShape'] and elem.get('text', '').strip()]
        other_elements = [elem for elem in elements if elem not in consolidated_entities and elem not in text_elements]
        
        if len(consolidated_entities) < 2:
            print(f"      🚀 Found {len(consolidated_entities)} consolidated entities - no super-consolidation needed")
            return elements
        
        super_entities = []
        processed_ids = set()
        super_count = 0
        
        # Find groups of overlapping consolidated entities
        for i, entity1 in enumerate(consolidated_entities):
            if entity1['box_id'] in processed_ids:
                continue
            # Start a new super group with this entity
            super_group = [entity1]
            processed_ids.add(entity1['box_id'])
            
            # Find all entities that overlap with any entity in the current group
            # (transitive overlapping)
            changed = True
            while changed:
                changed = False
                for entity2 in consolidated_entities:
                    if entity2['box_id'] not in processed_ids:
                        # Check if entity2 overlaps with any entity in current group
                        for group_entity in super_group:
                            overlap_pct = calculate_spatial_overlap(entity2['position'], group_entity['position'])
                            if overlap_pct > 30.0:  # Lower threshold for super-consolidation
                                super_group.append(entity2)
                                processed_ids.add(entity2['box_id'])
                                changed = True
                                break
            
            # If we have multiple entities in this group, create super-consolidated entity
            if len(super_group) > 1:
                # Find overlapping text elements (80% overlap)
                overlapping_texts = []
                for text_elem in text_elements:
                    for group_entity in super_group:
                        overlap_pct = calculate_spatial_overlap(text_elem['position'], group_entity['position'])
                        if overlap_pct > 80.0:
                            overlapping_texts.append(text_elem)
                            break
                # Create super-consolidated entity
                super_entity = self.create_super_consolidated_entity(super_group, overlapping_texts)
                super_entities.append(super_entity)
                super_count += 1
                group_ids = [ent['box_id'] for ent in super_group]
                text_ids = [txt['box_id'] for txt in overlapping_texts]
                print(f"      🚀 Super-consolidated: [{', '.join(group_ids)}] + texts [{', '.join(text_ids)}] → {super_entity['box_id']}")
                # Remove processed text elements
                text_elements = [txt for txt in text_elements if txt not in overlapping_texts]
            else:
                # Single entity, keep as is
                super_entities.append(entity1)
        
        # Combine results
        final_elements = super_entities + text_elements + other_elements
        print(f"      🚀 Created {super_count} super-consolidated entities")
        
        return final_elements



    def create_super_consolidated_entity(self, consolidated_entities, text_elements):
        """Create a super-consolidated entity from multiple consolidated entities + text"""
        # Calculate combined boundary from all entities
        all_positions = [ent['position'] for ent in consolidated_entities]
        if text_elements:
            all_positions.extend([txt['position'] for txt in text_elements])
            
        combined_boundary = calculate_combined_boundary(all_positions)
        
        # Generate incremental super-consolidated ID
        self.super_consolidated_counter += 1
        super_id = f"SC{self.super_consolidated_counter}"
        
        # Combine text from all elements
        all_texts = []
        for ent in consolidated_entities:
            if ent.get('text', '').strip():
                all_texts.append(ent['text'])
        for txt in text_elements:
            if txt.get('text', '').strip():
                all_texts.append(txt['text'])
        combined_text = " | ".join(all_texts)
        
        # Create super entity
        super_entity = {
            'box_id': super_id,
            'shape_type': 'SuperConsolidated',
            'box_type': 'super_consolidated',
            'position': combined_boundary,
            'text': combined_text,
            'has_text': bool(combined_text.strip()),
            'constituent_entities': [ent['box_id'] for ent in consolidated_entities],
            'constituent_texts': [txt['box_id'] for txt in text_elements],
            'entity_count': len(consolidated_entities),
            'text_count': len(text_elements),
            'is_target_for_capture': True
        }
        
        return super_entity



    def create_consolidated_image_entity(self, overlapping_images):
        """Create a consolidated image entity from multiple overlapping images"""
        # Calculate combined boundary
        combined_boundary = calculate_combined_boundary([img['position'] for img in overlapping_images])
        
        # Generate incremental consolidated ID (CI1, CI2, CI3...)
        self.consolidated_image_counter += 1
        consolidated_id = f"CI{self.consolidated_image_counter}"
        
        # Combine metadata
        combined_text = " | ".join([img.get('text', '') for img in overlapping_images if img.get('text', '').strip()])
        
        # Create consolidated entity
        consolidated_image = {
            'box_id': consolidated_id,
            'shape_type': 'ConsolidatedImage',
            'box_type': 'consolidated_image',
            'position': combined_boundary,
            'text': combined_text,
            'has_text': bool(combined_text.strip()),
            'consolidated_images_count': len(overlapping_images),
            'constituent_images': [
                {
                    'box_id': img['box_id'],
                    'original_position': img['position'],
                    'text': img.get('text', ''),
                    'file_info': img.get('file_info', {})
                }
                for img in overlapping_images
            ],
            'consolidation_type': 'overlapping_images'
        }
        
        # Add hierarchical info if present
        if overlapping_images[0].get('hierarchical_info'):
            # Use the first image's hierarchical info as base
            consolidated_image['hierarchical_info'] = overlapping_images[0]['hierarchical_info'].copy()
        
        return consolidated_image



    def create_image_autoshape_consolidated_entity(self, all_components, primary_image):
        """Create a consolidated entity from image + overlapping AutoShapes"""
        # Calculate combined boundary
        combined_boundary = calculate_combined_boundary([comp['position'] for comp in all_components])
        
        # Generate incremental IA group ID (IA1, IA2, IA3...)
        self.ia_group_counter += 1
        consolidated_id = f"IA{self.ia_group_counter}"
        
        # Combine text from all components
        all_texts = [comp.get('text', '') for comp in all_components if comp.get('text', '').strip()]
        combined_text = " | ".join(all_texts)
        
        # Create consolidated entity
        consolidated_entity = {
            'box_id': consolidated_id,
            'shape_type': 'ConsolidatedImageAutoShape',
            'box_type': 'image_autoshape_combo',
            'position': combined_boundary,
            'text': combined_text,
            'has_text': bool(combined_text.strip()),
            'primary_image': {
                'box_id': primary_image['box_id'],
                'original_position': primary_image['position'],
                'text': primary_image.get('text', ''),
                'file_info': primary_image.get('file_info', {})
            },
            'overlapping_autoshapes': [
                {
                    'box_id': comp['box_id'],
                    'original_position': comp['position'],
                    'text': comp.get('text', ''),
                    'file_info': comp.get('file_info', {})
                }
                for comp in all_components if comp['shape_type'] == 'AutoShape'
            ],
            'consolidation_type': 'image_with_autoshapes'
        }
        
        # Add hierarchical info from primary image if present
        if primary_image.get('hierarchical_info'):
            consolidated_entity['hierarchical_info'] = primary_image['hierarchical_info'].copy()
        
        return consolidated_entity



    def consolidate_image_autoshape_overlaps(self, elements):
        """Consolidate images that overlap with AutoShapes into single consolidated entities - PRESERVE TEXT COMPONENTS"""
        image_elements = [elem for elem in elements if elem.get('shape_type') in ['Picture', 'Image'] or elem.get('box_type') in ['consolidated_image', 'image_autoshape_combo']]
        autoshape_elements = [elem for elem in elements if elem.get('shape_type') == 'AutoShape']
        other_elements = [elem for elem in elements if elem.get('shape_type') not in ['Picture', 'Image', 'AutoShape'] and elem.get('box_type') not in ['consolidated_image', 'image_autoshape_combo']]
        
        if not image_elements or not autoshape_elements:
            print(f"      🔗 Found {len(image_elements)} images and {len(autoshape_elements)} AutoShapes - no image+AutoShape consolidation needed")
            return elements
        
        print(f"      🔗 Consolidating images with overlapping AutoShapes (PRESERVING TEXT): {len(image_elements)} images, {len(autoshape_elements)} AutoShapes")
        
        consolidated_elements = []
        processed_ids = set()
        consolidation_count = 0
        
        # Process each image to find overlapping AutoShapes
        for image_elem in image_elements:
            if image_elem['box_id'] in processed_ids:
                continue
            
            # Find AutoShapes that overlap with this image - ONLY NON-TEXT AutoShapes
            overlapping_autoshapes = []
            print(f"      🔍 Checking overlaps for image {image_elem['box_id']}...")
            
            for autoshape_elem in autoshape_elements:
                if autoshape_elem['box_id'] not in processed_ids:
                    overlap_pct = calculate_spatial_overlap(image_elem['position'], autoshape_elem['position'])
                    
                    if overlap_pct > 30.0:  # Higher threshold - be more selective
                        # STRICT: Only consolidate AutoShapes with NO significant text content
                        autoshape_text = autoshape_elem.get('text', '').strip()
                        
                        if not autoshape_text or len(autoshape_text) < 5:  # Only truly empty AutoShapes
                            overlapping_autoshapes.append(autoshape_elem)
                            print(f"         ✅ Will consolidate EMPTY AutoShape {autoshape_elem['box_id']} (overlap: {overlap_pct:.1f}%)")
                        else:
                            print(f"         📝 PRESERVING text-containing AutoShape {autoshape_elem['box_id']} (overlap: {overlap_pct:.1f}%, text: '{autoshape_text[:30]}...')")
                    else:
                        print(f"         ❌ AutoShape {autoshape_elem['box_id']} overlap too low: {overlap_pct:.1f}%")
            
            if overlapping_autoshapes:
                # Create consolidated image + autoshape entity - SIMPLE ID
                all_components = [image_elem] + overlapping_autoshapes
                consolidated_entity = self.create_simple_image_autoshape_entity(all_components, image_elem)
                consolidated_elements.append(consolidated_entity)
                processed_ids.add(image_elem['box_id'])
                for ash in overlapping_autoshapes:
                    processed_ids.add(ash['box_id'])
                consolidation_count += 1
                autoshape_ids = [ash['box_id'] for ash in overlapping_autoshapes]
                print(f"      🖼️🔗 Consolidated image {image_elem['box_id']} with EMPTY AutoShapes [{' + '.join(autoshape_ids)}] → {consolidated_entity['box_id']}")
            else:
                # No overlapping AutoShapes, keep image as is
                consolidated_elements.append(image_elem)
                processed_ids.add(image_elem['box_id'])
        
        # Add unprocessed AutoShapes and other elements (PRESERVE ALL TEXT)
        for elem in autoshape_elements + other_elements:
            if elem['box_id'] not in processed_ids:
                consolidated_elements.append(elem)
        
        print(f"      ✅ Created {consolidation_count} image+AutoShape consolidations (TEXT PRESERVED)")
        return consolidated_elements



    def create_simple_image_autoshape_entity(self, all_components, primary_image):
        """Create a simple consolidated entity with incremental naming"""
        # Calculate combined boundary
        combined_boundary = calculate_combined_boundary([comp['position'] for comp in all_components])
        
        # Generate incremental IA group ID (IA1, IA2, IA3...)
        self.ia_group_counter += 1
        simple_id = f"IA{self.ia_group_counter}"
        
        # Only include text from primary image (preserve separate text components)
        primary_text = primary_image.get('text', '').strip()
        
        # Create simple consolidated entity
        consolidated_entity = {
            'box_id': simple_id,
            'shape_index': primary_image.get('shape_index', 0),
            'position': combined_boundary,
            'shape_type': 'Picture',  # Treat as Picture for capture
            'text': primary_text,
            'has_text': bool(primary_text),
            'is_target_for_capture': True,
            'box_type': 'image_autoshape_combo',
            'image_component': primary_image['box_id'],
            'autoshape_components': [comp['box_id'] for comp in all_components if comp['shape_type'] == 'AutoShape']
        }
        
        return consolidated_entity


