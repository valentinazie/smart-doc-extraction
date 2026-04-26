"""
Reading order computation, line/section detection, and text flow summaries.
"""

import os
import json
import time
from pathlib import Path
from datetime import datetime

from utils.geometry import (
    calculate_group_bounds,
    calculate_shapes_bounds,
    calculate_spatial_overlap,
    calculate_text_box_overlap,
    get_group_center,
)


class ReadingOrderMixin:
    """Methods for computing reading order, sections, and hierarchical flow."""

    def apply_hierarchical_group_reading_order(self, smart_groups, line_elements, slide_data):
        """Apply grid-based reading order using lines to create section boundaries"""
        print(f"   📖 Applying grid-based sectional reading order for {len(smart_groups)} groups")
        
        # STEP 1: Identify horizontal and vertical lines (explicit + spatial boundaries)
        horizontal_lines = []
        vertical_lines = []
        
        # Add explicit line shapes
        for line in line_elements:
            pos = line['position']
            width = max(pos['width'], 1)
            height = max(pos['height'], 1)
            aspect_ratio = width / height
            
            if aspect_ratio > 2.0:  # Horizontal line (가로)
                horizontal_lines.append(line)
                print(f"         📏 Explicit horizontal divider: {line['box_id']} (ratio={aspect_ratio:.2f})")
            elif aspect_ratio < 3.0:  # Vertical line (세로)
                vertical_lines.append(line)
                print(f"         📏 Explicit vertical divider: {line['box_id']} (ratio={aspect_ratio:.2f})")
        
        # STEP 2: No spatial boundaries detection - use only explicit lines
        
        # STEP 3: Create section grid using all lines as boundaries
        section_grid = self.create_section_grid(smart_groups, horizontal_lines, vertical_lines)
        
        # STEP 4: Read the section grid naturally (top-to-bottom rows, left-to-right within rows)
        final_group_order = []
        for row_idx in range(section_grid['rows']):
            for col_idx in range(section_grid['cols']):
                section_key = f"R{row_idx}C{col_idx}"
                if section_key in section_grid['sections']:
                    section = section_grid['sections'][section_key]
                    section_groups = section['groups']
                    
                    # Sort groups within this section by position (ROWS first with 80-120% height tolerance, then LEFT-TO-RIGHT within rows)
                    sorted_section_groups = self.sort_groups_by_simple_reading_order(section_groups)
                    final_group_order.extend(sorted_section_groups)
                    
                    print(f"      📍 Section R{row_idx+1}C{col_idx+1}: {[g['group_id'] for g in sorted_section_groups]}")
        
        print(f"      ✅ Grid-based reading order: {len(final_group_order)} groups across {section_grid['rows']}×{section_grid['cols']} grid")
        return final_group_order



    def create_section_grid(self, smart_groups, horizontal_lines, vertical_lines):
        """Create a section grid using lines as boundaries"""
        if not smart_groups:
            return {'sections': {}, 'rows': 0, 'cols': 0}
        
        # Get overall bounds
        all_group_centers = [get_group_center(g) for g in smart_groups if get_group_center(g)]
        if not all_group_centers:
            return {'sections': {}, 'rows': 0, 'cols': 0}
        
        min_x = min(center[0] for center in all_group_centers)
        max_x = max(center[0] for center in all_group_centers)
        min_y = min(center[1] for center in all_group_centers)
        max_y = max(center[1] for center in all_group_centers)
        
        # Create Y boundaries from horizontal lines (or auto-detect natural boundaries)
        y_boundaries = [min_y - 100000]  # Top boundary
        
        if horizontal_lines:
            # Use explicit horizontal lines
            sorted_h_lines = sorted(horizontal_lines, key=lambda l: l['position']['top'])
            for h_line in sorted_h_lines:
                line_center_y = h_line['position']['top'] + h_line['position']['height'] // 2
                y_boundaries.append(line_center_y)
                print(f"            📐 Row boundary at Y={line_center_y} (from {h_line['box_id']})")
        else:
            # Auto-detect natural horizontal boundaries by clustering Y positions
            y_positions = sorted([center[1] for center in all_group_centers])
            
            # Find natural gaps in Y positions
            y_gaps = []
            for i in range(len(y_positions) - 1):
                gap_size = y_positions[i + 1] - y_positions[i]
                y_gaps.append((gap_size, (y_positions[i] + y_positions[i + 1]) / 2))
            
            # Use the largest gaps as natural row boundaries
            y_gaps.sort(reverse=True)  # Largest gaps first
            natural_boundaries = []
            
            # Take up to 2-3 largest gaps as natural row dividers
            for gap_size, gap_center in y_gaps[:3]:  # Max 3 natural boundaries
                if gap_size > (max_y - min_y) * 0.15:  # Only significant gaps (>15% of total height)
                    natural_boundaries.append(gap_center)
                    print(f"            📐 Natural row boundary at Y={gap_center:.0f} (gap={gap_size:.0f})")
            
            y_boundaries.extend(sorted(natural_boundaries))
        
        y_boundaries.append(max_y + 100000)  # Bottom boundary
        y_boundaries.sort()
        
        # Create X boundaries from vertical lines
        x_boundaries = [min_x - 100000]  # Left boundary
        
        if vertical_lines:
            sorted_v_lines = sorted(vertical_lines, key=lambda l: l['position']['left'])
            for v_line in sorted_v_lines:
                line_center_x = v_line['position']['left'] + v_line['position']['width'] // 2
                x_boundaries.append(line_center_x)
                print(f"            📐 Column boundary at X={line_center_x} (from {v_line['box_id']})")
        
        x_boundaries.append(max_x + 100000)  # Right boundary
        x_boundaries.sort()
        
        print(f"            🗂️  Created {len(y_boundaries)-1}×{len(x_boundaries)-1} section grid")
        
        # Create sections and assign groups
        sections = {}
        for row in range(len(y_boundaries) - 1):
            for col in range(len(x_boundaries) - 1):
                y_min, y_max = y_boundaries[row], y_boundaries[row + 1]
                x_min, x_max = x_boundaries[col], x_boundaries[col + 1]
                # Find groups that belong to this section
                section_groups = []
                for group in smart_groups:
                    group_center = get_group_center(group)
                    if group_center and x_min < group_center[0] <= x_max and y_min < group_center[1] <= y_max:
                        section_groups.append(group)
                if section_groups:  # Only create section if it has groups
                    section_key = f"R{row}C{col}"
                    sections[section_key] = {
                        'section_id': f'Section-R{row+1}C{col+1}',
                        'row': row,
                        'col': col,
                        'groups': section_groups,
                        'bounds': {
                            'x_min': x_min, 'x_max': x_max,
                            'y_min': y_min, 'y_max': y_max
                        }
                    }
                    print(f"               📦 Section R{row+1}C{col+1}: {[g['group_id'] for g in section_groups]}")
        
        return {
            'sections': sections,
            'rows': len(y_boundaries) - 1,
            'cols': len(x_boundaries) - 1,
            'y_boundaries': y_boundaries,
            'x_boundaries': x_boundaries
        }



    def auto_detect_row_levels(self, smart_groups):
        """Auto-detect row levels based on Y-position clustering when no horizontal lines exist"""
        if not smart_groups:
            return []
        
        # Calculate group centers and their Y positions
        groups_with_centers = []
        for group in smart_groups:
            center = get_group_center(group)
            if center:
                groups_with_centers.append({
                    'group': group,
                    'center_y': center[1],
                    'center_x': center[0]
                })
        
        if not groups_with_centers:
            return []
        
        # Calculate row tolerance based on Y-position spread
        y_positions = [g['center_y'] for g in groups_with_centers]
        y_span = max(y_positions) - min(y_positions)
        row_tolerance = max(y_span * 0.25, 300000)  # 25% of Y-span or 300k units
        
        print(f"            📏 Auto-detect row tolerance: {row_tolerance:.0f} units")
        
        # Cluster groups by Y position into rows
        groups_with_centers.sort(key=lambda x: x['center_y'])
        
        row_levels = []
        current_level_groups = []
        current_level_y = None
        
        for group_data in groups_with_centers:
            group_y = group_data['center_y']
            
            # Start new level if Y position differs significantly
            if current_level_y is None or abs(group_y - current_level_y) <= row_tolerance:
                current_level_groups.append(group_data)
                current_level_y = group_y if current_level_y is None else current_level_y
            else:
                # Save current level and start new level
                if current_level_groups:
                    level_y_min = min(g['center_y'] for g in current_level_groups) - row_tolerance/2
                    level_y_max = max(g['center_y'] for g in current_level_groups) + row_tolerance/2
                    row_levels.append({
                        'level_id': f'AutoLevel{len(row_levels)+1}',
                        'y_range': (level_y_min, level_y_max),
                        'groups': [g['group'] for g in current_level_groups]
                    })
                current_level_groups = [group_data]
                current_level_y = group_y
        
        # Don't forget the last level
        if current_level_groups:
            level_y_min = min(g['center_y'] for g in current_level_groups) - row_tolerance/2
            level_y_max = max(g['center_y'] for g in current_level_groups) + row_tolerance/2
            row_levels.append({
                'level_id': f'AutoLevel{len(row_levels)+1}',
                'y_range': (level_y_min, level_y_max),
                'groups': [g['group'] for g in current_level_groups]
            })
        
        print(f"            📊 Auto-detected {len(row_levels)} row levels:")
        for i, level in enumerate(row_levels):
            level_y_center = (level['y_range'][0] + level['y_range'][1]) / 2
            print(f"               Level {i+1}: {len(level['groups'])} groups at Y≈{level_y_center:.0f}")
        
        return row_levels



    def sort_level_groups_left_to_right(self, level_groups, vertical_lines, y_range):
        """Sort groups within a level from left to right, respecting vertical line boundaries"""
        if not level_groups:
            return []
        
        if len(level_groups) <= 1:
            return level_groups
        
        # Filter vertical lines that intersect with this level's Y range
        relevant_v_lines = []
        y_min, y_max = y_range
        level_y_center = (y_min + y_max) / 2
        
        for v_line in vertical_lines:
            line_top = v_line['position']['top']
            line_bottom = v_line['position']['top'] + v_line['position']['height']
            
            # Check if vertical line intersects with this level
            if (line_top <= y_max and line_bottom >= y_min):
                line_center_x = v_line['position']['left'] + v_line['position']['width'] // 2
                relevant_v_lines.append({
                    'line': v_line,
                    'center_x': line_center_x
                })
                print(f"            🔹 Vertical line {v_line['box_id']} affects this level at x={line_center_x}")
        
        if not relevant_v_lines:
            # No vertical lines affect this level, simple left-to-right sort
            return self.sort_groups_by_left_to_right(level_groups)
        
        # Create X segments based on vertical lines
        x_boundaries = [-float('inf')]  # Leftmost boundary
        for v_line in sorted(relevant_v_lines, key=lambda x: x['center_x']):
            x_boundaries.append(v_line['center_x'])
        x_boundaries.append(float('inf'))  # Rightmost boundary
        
        # Group by X segments and sort within each segment
        final_order = []
        for i in range(len(x_boundaries) - 1):
            x_min, x_max = x_boundaries[i], x_boundaries[i + 1]
            
            # Find groups in this X segment
            segment_groups = []
            for group in level_groups:
                group_center = get_group_center(group)
                if group_center and x_min < group_center[0] <= x_max:
                    segment_groups.append(group)
            
            if segment_groups:
                # Sort within segment by X position (left to right)
                segment_sorted = self.sort_groups_by_left_to_right(segment_groups)
                final_order.extend(segment_sorted)
                print(f"               📍 Segment {i+1}: {[g['group_id'] for g in segment_sorted]}")
        
        return final_order



    def organize_groups_into_rows(self, smart_groups):
        """Organize groups into horizontal rows based on their vertical positions"""
        if not smart_groups:
            return []
        
        # Calculate group centers and their Y positions
        groups_with_centers = []
        for group in smart_groups:
            center = get_group_center(group)
            if center:
                groups_with_centers.append({
                    'group': group,
                    'center_y': center[1],
                    'center_x': center[0]
                })
        
        if not groups_with_centers:
            return []
        
        # Calculate row tolerance based on average group heights
        avg_y_span = (max(g['center_y'] for g in groups_with_centers) - 
                      min(g['center_y'] for g in groups_with_centers)) / len(groups_with_centers)
        row_tolerance = max(avg_y_span * 0.3, 200000)  # 30% of average span or 200k units
        
        print(f"         📏 Row tolerance: {row_tolerance:.0f} units")
        
        # Group by rows using Y position clustering
        groups_with_centers.sort(key=lambda x: x['center_y'])
        
        rows = []
        current_row_groups = []
        current_row_y = None
        
        for group_data in groups_with_centers:
            group_y = group_data['center_y']
            
            # Start new row if Y position differs significantly
            if current_row_y is None or abs(group_y - current_row_y) <= row_tolerance:
                current_row_groups.append(group_data)
                current_row_y = group_y if current_row_y is None else current_row_y
            else:
                # Save current row and start new row
                if current_row_groups:
                    rows.append({
                        'row_y': current_row_y,
                        'groups': [g['group'] for g in current_row_groups]
                    })
                current_row_groups = [group_data]
                current_row_y = group_y
        
        # Don't forget the last row
        if current_row_groups:
            rows.append({
                'row_y': current_row_y,
                'groups': [g['group'] for g in current_row_groups]
            })
        
        print(f"         📊 Organized into {len(rows)} horizontal rows")
        for i, row in enumerate(rows):
            print(f"            Row {i+1}: {len(row['groups'])} groups at Y≈{row['row_y']:.0f}")
        
        return rows



    def sort_row_groups_with_line_awareness(self, row_groups, vertical_lines):
        """Sort groups within a row, respecting vertical line boundaries"""
        if not vertical_lines or len(row_groups) <= 1:
            return self.sort_groups_by_left_to_right(row_groups)
        
        # Create segments based on vertical lines
        groups_with_positions = []
        for group in row_groups:
            center = get_group_center(group)
            if center:
                groups_with_positions.append({
                    'group': group,
                    'center_x': center[0],
                    'center_y': center[1]
                })
        
        if not groups_with_positions:
            return row_groups
        
        # Create X boundaries from vertical lines
        x_boundaries = [-float('inf')]  # Start boundary
        for v_line in vertical_lines:
            # Check if this vertical line intersects with this row's Y range
            row_y_center = sum(g['center_y'] for g in groups_with_positions) / len(groups_with_positions)
            row_y_tolerance = 300000  # Generous tolerance for row intersection
            
            if (v_line['top'] <= row_y_center + row_y_tolerance and 
                v_line['bottom'] >= row_y_center - row_y_tolerance):
                x_boundaries.append(v_line['center_x'])
                print(f"            🔹 Line {v_line['line']['box_id']} affects this row")
        
        x_boundaries.append(float('inf'))  # End boundary
        x_boundaries.sort()
        
        # Group by segments and sort within each segment
        segments = []
        for i in range(len(x_boundaries) - 1):
            segment_groups = []
            x_min, x_max = x_boundaries[i], x_boundaries[i + 1]
            
            for group_data in groups_with_positions:
                if x_min < group_data['center_x'] <= x_max:
                    segment_groups.append(group_data['group'])
            
            if segment_groups:
                # Sort within segment by X position
                segment_sorted = self.sort_groups_by_left_to_right(segment_groups)
                segments.extend(segment_sorted)
        
        return segments



    def sort_groups_by_left_to_right(self, groups):
        """Sort groups by their X position (left to right)"""
        groups_with_x = []
        for group in groups:
            center = get_group_center(group)
            if center:
                groups_with_x.append((group, center[0]))
        
        # Sort by X position
        groups_with_x.sort(key=lambda x: x[1])
        return [group for group, x in groups_with_x]



    def create_sections_from_lines(self, smart_groups, line_elements):
        """Create sections using both horizontal and vertical lines as boundaries"""
        print(f"         🗺️  Creating sections from {len(line_elements)} lines")
        
        # Identify horizontal and vertical lines
        horizontal_lines = []
        vertical_lines = []
        
        for line in line_elements:
            pos = line['position']
            width = max(pos['width'], 1)
            height = max(pos['height'], 1)
            aspect_ratio = width / height
            
            if aspect_ratio > 2.0:  # Horizontal line
                horizontal_lines.append(line)
                print(f"            📏 Horizontal boundary: {line['box_id']} (ratio={aspect_ratio:.2f})")
            elif aspect_ratio < 0.5:  # Vertical line  
                vertical_lines.append(line)
                print(f"            📏 Vertical boundary: {line['box_id']} (ratio={aspect_ratio:.2f})")
        
        # Create grid sections based on line boundaries
        sections = self.create_line_boundary_grid(smart_groups, horizontal_lines, vertical_lines)
        
        return sections



    def create_line_boundary_grid(self, smart_groups, horizontal_lines, vertical_lines):
        """Create a grid of sections based on horizontal and vertical line boundaries"""
        
        # Get overall bounds for the slide
        all_group_centers = []
        for group in smart_groups:
            if group['members']:
                center = get_group_center(group)
                all_group_centers.append(center)
        
        if not all_group_centers:
            return [{'section_id': 'main', 'groups': smart_groups, 'bounds': {}}]
        
        min_x = min(c[0] for c in all_group_centers)
        max_x = max(c[0] for c in all_group_centers)
        min_y = min(c[1] for c in all_group_centers)
        max_y = max(c[1] for c in all_group_centers)
        
        # Create Y boundaries from horizontal lines
        y_boundaries = [min_y - 100000]  # Start boundary
        if horizontal_lines:
            for h_line in sorted(horizontal_lines, key=lambda l: l['position']['top']):
                y_center = h_line['position']['top'] + h_line['position']['height'] // 2
                y_boundaries.append(y_center)
        y_boundaries.append(max_y + 100000)  # End boundary
        
        # Create X boundaries from vertical lines  
        x_boundaries = [min_x - 100000]  # Start boundary
        if vertical_lines:
            for v_line in sorted(vertical_lines, key=lambda l: l['position']['left']):
                x_center = v_line['position']['left'] + v_line['position']['width'] // 2
                x_boundaries.append(x_center)
        x_boundaries.append(max_x + 100000)  # End boundary
        
        print(f"            🗂️  Grid: {len(y_boundaries)-1} rows × {len(x_boundaries)-1} columns")
        
        # Create sections and assign groups
        sections = []
        section_id = 1
        
        for row in range(len(y_boundaries) - 1):
            for col in range(len(x_boundaries) - 1):
                y_min, y_max = y_boundaries[row], y_boundaries[row + 1]
                x_min, x_max = x_boundaries[col], x_boundaries[col + 1]
                # Find groups that belong to this section
                section_groups = []
                for group in smart_groups:
                    group_center = get_group_center(group)
                    if group_center and x_min <= group_center[0] <= x_max and y_min <= group_center[1] <= y_max:
                        section_groups.append(group)
                if section_groups:  # Only create section if it has groups
                    sections.append({
                        'section_id': f'S{row+1}C{col+1}',
                        'row': row,
                        'col': col,
                        'groups': section_groups,
                        'bounds': {
                            'x_min': x_min, 'x_max': x_max,
                            'y_min': y_min, 'y_max': y_max
                        }
                    })
                    print(f"               📦 Section S{row+1}C{col+1}: {len(section_groups)} groups")
        
        return sections



    def sort_sections_by_reading_order(self, sections):
        """Sort sections by ACTUAL CONTENT POSITION, not artificial grid coordinates"""
        def get_section_reading_key(section):
            # FIXED: Sort by actual content position, not grid coordinates
            if not section.get('groups'):
                return (0, 0)  # Empty section
            # Calculate the average position of all groups in this section
            total_y = 0
            total_x = 0
            group_count = 0
            
            for group in section['groups']:
                if group.get('members'):
                    # Calculate group TOP and LEFT for row-wise reading
                    all_positions = []
                    for member in group['members']:
                        pos = member['position']
                        all_positions.append({
                            'top': pos['top'],
                            'left': pos['left'],
                            'bottom': pos['top'] + pos['height'],
                            'right': pos['left'] + pos['width']
                        })
                    
                    # ROW-WISE READING: Use TOP and LEFT edges (where content STARTS)
                    group_top = min(p['top'] for p in all_positions)
                    group_left = min(p['left'] for p in all_positions)
                    
                    total_y += group_top
                    total_x += group_left
                    group_count += 1
            
            if group_count == 0:
                return (0, 0)
            # Section position = average of all group TOP and LEFT positions in section
            section_top = total_y // group_count
            section_left = total_x // group_count
            
            # Use same row tolerance as groups for consistency
            avg_group_height = 200000  # Approximate average
            # FIXED: Use predictable tolerance instead of content-dependent calculation
            consistent_row_tolerance = 300000  # Fixed 300k EMU tolerance (~10.5mm) - content independent
            row_group = int(section_top / consistent_row_tolerance) * consistent_row_tolerance
            
            return (row_group, section_left)  # TOP→BOTTOM rows, LEFT→RIGHT within rows
        
        sorted_sections = sorted(sections, key=get_section_reading_key)
        
        print(f"         📖 Section reading order (by ACTUAL content position):")
        for idx, section in enumerate(sorted_sections):
            key = get_section_reading_key(section)
            print(f"            {idx+1}. {section['section_id']} (row_group={key[0]}, center_x={key[1]})")
        
        return sorted_sections



    def sort_groups_by_simple_reading_order(self, smart_groups):
        """Sort groups by ROW first (80-120% height tolerance), then LEFT-TO-RIGHT within rows"""
        
        # FIXED: Calculate consistent row tolerance across ALL groups
        if not smart_groups:
            return smart_groups
            
        # Calculate average group height across ALL groups for consistent tolerance
        all_group_heights = []
        for group in smart_groups:
            if group['members']:
                all_positions = []
                for member in group['members']:
                    pos = member['position']
                    all_positions.append({
                        'top': pos['top'],
                        'bottom': pos['top'] + pos['height']
                    })
                group_height = max(p['bottom'] for p in all_positions) - min(p['top'] for p in all_positions)
                all_group_heights.append(group_height)
        
        # Use consistent row tolerance based on AVERAGE height of all groups
        avg_group_height = sum(all_group_heights) / len(all_group_heights) if all_group_heights else 200000
        # FIXED: Use predictable tolerance instead of content-dependent calculation  
        consistent_row_tolerance = 300000  # Fixed 300k EMU tolerance (~10.5mm) - content independent
        
        def get_group_center_reading_key(group):
            # Calculate the actual center of the entire group (not just first member)
            if group['members']:
                # Calculate group bounds from all members
                all_positions = []
                for member in group['members']:
                    pos = member['position']
                    all_positions.append({
                        'left': pos['left'],
                        'top': pos['top'],
                        'right': pos['left'] + pos['width'],
                        'bottom': pos['top'] + pos['height']
                    })
                # Calculate group bounds for row-wise reading
                min_left = min(p['left'] for p in all_positions)
                min_top = min(p['top'] for p in all_positions)
                max_right = max(p['right'] for p in all_positions)
                max_bottom = max(p['bottom'] for p in all_positions)
                # ROW-WISE READING: Use TOP edge for row determination, LEFT edge for within-row ordering
                group_top = min_top      # Use TOP edge (where content STARTS)
                group_left = min_left    # Use LEFT edge for left-to-right ordering
                group_height = max_bottom - min_top  # Actual group height
                # Use TOP edge to determine row grouping (not center!)
                row_group = int(group_top / consistent_row_tolerance) * consistent_row_tolerance
                print(f"         📍 Group {group['group_id']}: top={group_top}, left={group_left}, height={group_height}, row_group={row_group}, tolerance={consistent_row_tolerance}")
                return (
                    row_group,    # PRIMARY: Row grouping based on TOP edge (where content starts)
                    group_left    # SECONDARY: LEFT-TO-RIGHT within each row
                )
            return (0, 0)
        
        sorted_groups = sorted(smart_groups, key=get_group_center_reading_key)
        print(f"         📖 Sorted {len(smart_groups)} groups by ROWS first (80-120% height tolerance), then LEFT→RIGHT within rows")
        
        # Debug: Show the final order
        for idx, group in enumerate(sorted_groups):
            print(f"            {idx+1}. Group {group['group_id']} ({len(group['members'])} members)")
        
        return sorted_groups



    def apply_group_line_interruption(self, current_group_order, line_element):
        """Apply line interruption logic to group reading order"""
        pos = line_element['position']
        width = max(pos['width'], 1)
        height = max(pos['height'], 1)
        aspect_ratio = width / height
        
        # Only process clear vertical lines for group interruption
        if aspect_ratio >= 0.7:
            print(f"         📏 {line_element['box_id']}: Not a clear vertical line (ratio={aspect_ratio:.2f}), skipping group interruption")
            return current_group_order
        
        line_center_x = pos['left'] + width // 2
        line_top = pos['top']
        line_bottom = pos['top'] + height
        
        print(f"         📏 Processing group line interruption for {line_element['box_id']} at x={line_center_x}")
        
        # Define group influence area (larger than component influence)
        group_influence_area = {
            'left': line_center_x - width * 20,  # Larger group influence 
            'right': line_center_x + width * 20,
            'top': max(0, line_top - height * 0.2),
            'bottom': line_bottom + height * 0.2
        }
        
        # Find groups within the line's influence area
        influenced_groups = []
        non_influenced_groups = []
        
        for group in current_group_order:
            if not group['members']:
                non_influenced_groups.append(group)
                continue
            # Use group center to determine influence
            root_member = group['members'][0]
            group_center_x = root_member['position']['left'] + root_member['position']['width'] // 2
            group_center_y = root_member['position']['top'] + root_member['position']['height'] // 2
            
            # Check if group is within line's influence area
            if (group_influence_area['left'] <= group_center_x <= group_influence_area['right'] and
                group_influence_area['top'] <= group_center_y <= group_influence_area['bottom']):
                influenced_groups.append(group)
            else:
                non_influenced_groups.append(group)
        
        if len(influenced_groups) < 2:
            print(f"            ⚠️  Only {len(influenced_groups)} groups in influence area, no group interruption needed")
            return current_group_order
        
        print(f"            📍 Found {len(influenced_groups)} groups in line influence area")
        
        # Split influenced groups by line boundary
        left_influenced_groups = []
        right_influenced_groups = []
        
        for group in influenced_groups:
            if not group['members']:
                continue
            root_member = group['members'][0]
            group_center_x = root_member['position']['left'] + root_member['position']['width'] // 2
            
            if group_center_x < line_center_x:
                left_influenced_groups.append(group)
            else:
                right_influenced_groups.append(group)
        
        print(f"            📍 Group line influence: LEFT({len(left_influenced_groups)}) | RIGHT({len(right_influenced_groups)})")
        
        if not left_influenced_groups or not right_influenced_groups:
            print(f"            ⚠️  No group boundary split needed")
            return current_group_order
        
        # Sort each side by group reading order
        left_groups_sorted = self.sort_groups_by_simple_reading_order(left_influenced_groups)
        right_groups_sorted = self.sort_groups_by_simple_reading_order(right_influenced_groups)
        
        # Reconstruct group reading order with line interruption
        result = []
        influenced_groups_reordered = left_groups_sorted + right_groups_sorted
        influenced_index = 0
        
        for group in current_group_order:
            if group in influenced_groups:
                # Replace with reordered influenced group (LEFT groups first, then RIGHT groups)
                if influenced_index < len(influenced_groups_reordered):
                    result.append(influenced_groups_reordered[influenced_index])
                    influenced_index += 1
            else:
                # Keep non-influenced group in original position
                result.append(group)
        
        print(f"            ✅ Applied group line interruption: LEFT({len(left_influenced_groups)}) → RIGHT({len(right_influenced_groups)})")
        return result



    def sort_group_components_with_nearby_vertical_lines(self, components, group_id, line_elements):
        """Sort components within a group, using nearby vertical lines to split left-right"""
        if len(components) <= 1:
            return components
        
        # Calculate group bounds
        group_positions = [c['position'] for c in components]
        group_left = min(p['left'] for p in group_positions)
        group_right = max(p['left'] + p['width'] for p in group_positions)
        group_top = min(p['top'] for p in group_positions)
        group_bottom = max(p['top'] + p['height'] for p in group_positions)
        group_center_x = (group_left + group_right) // 2
        
        print(f"            🔍 Checking {group_id} for nearby vertical lines (bounds: {group_left//12700}-{group_right//12700}pt)")
        
        # Find vertical lines that intersect or are near this group
        relevant_vertical_lines = []
        
        for line in line_elements:
            pos = line['position']
            width = max(pos['width'], 1)
            height = max(pos['height'], 1)
            aspect_ratio = width / height
            
            # Check if this is a vertical line
            if aspect_ratio < 0.5:  # Vertical line
                line_x = pos['left'] + pos['width'] // 2
                line_top = pos['top']
                line_bottom = pos['top'] + pos['height']
                # Check if line intersects with group vertically and is within group horizontally
                vertical_overlap = not (line_bottom <= group_top or line_top >= group_bottom)
                horizontal_within = group_left <= line_x <= group_right
                if vertical_overlap and horizontal_within:
                    relevant_vertical_lines.append((line, line_x))
                    print(f"            📏 Found relevant vertical line {line['box_id']} at X={line_x//12700}pt for {group_id}")
        
        if not relevant_vertical_lines:
            print(f"            📖 No vertical lines found for {group_id}, using simple reading order")
            return self.sort_shapes_by_simple_reading_order(components)
        
        # Use the most central vertical line
        best_line = None
        best_x = None
        best_distance = float('inf')
        
        for line, line_x in relevant_vertical_lines:
            distance = abs(line_x - group_center_x)
            if distance < best_distance:
                best_distance = distance
                best_line = line
                best_x = line_x
        
        if best_line:
            print(f"            📏 Using vertical line {best_line['box_id']} at X={best_x//12700}pt to split {group_id}")
            
            # Split components into left and right sides
            left_components = []
            right_components = []
            
            print(f"            🔍 Splitting {group_id} components by vertical line at X={best_x//12700}pt:")
            
            for component in components:
                comp_left = component['position']['left']
                comp_width = component['position']['width']
                comp_right = comp_left + comp_width
                # ROW-WISE READING: Use LEFT edge for component positioning (consistent with group-level sorting)
                comp_left_edge = comp_left
                comp_x_pt = comp_left_edge // 12700
                line_x_pt = best_x // 12700
                # Special handling for wide tables that span across the vertical line
                shape_type = component.get('shape_type', 'unknown')
                is_spanning_table = (shape_type == 'Table' and 
                                   comp_left < best_x and comp_right > best_x and
                                   comp_width > (best_x - comp_left) * 2)  # Table is significantly wide
                if is_spanning_table:
                    # For spanning tables, classify based on where the LEFT EDGE appears (consistent reading order)
                    # This prevents visual back-and-forth in the reading flow
                    side = "LEFT" if comp_left_edge < best_x else "RIGHT"
                    side_reason = "SPANNING-LEFTEDGE"
                    debug_info = f"type={shape_type}, width={comp_width//12700}pt, {side_reason}"
                else:
                    side = "LEFT" if comp_left_edge < best_x else "RIGHT"
                    debug_info = f"type={shape_type}"
                    if shape_type == 'Table':
                        debug_info += f", width={comp_width//12700}pt"
                if side == "LEFT":
                    left_components.append(component)
                else:
                    right_components.append(component)
                print(f"               {component['box_id']}: left={comp_left//12700}pt, leftedge={comp_x_pt}pt, line={line_x_pt}pt ({debug_info}) → {side}")
            
                            # Sort each side by reading order (ROWS first with 80-120% height tolerance, then LEFT-TO-RIGHT within rows)
            left_sorted = self.sort_shapes_by_simple_reading_order(left_components)
            right_sorted = self.sort_shapes_by_simple_reading_order(right_components)
            
            print(f"            📖 {group_id} final order:")
            print(f"               LEFT: {[c['box_id'] for c in left_sorted]}")
            print(f"               RIGHT: {[c['box_id'] for c in right_sorted]}")
            print(f"               COMBINED: {[c['box_id'] for c in left_sorted + right_sorted]}")
            
            # Combine: left side first, then right side
            result = left_sorted + right_sorted
            
            # VERIFICATION: Double-check the final order with proper spanning table logic
            print(f"            ✅ VERIFICATION - Final {group_id} order:")
            for i, comp in enumerate(result, 1):
                comp_left = comp['position']['left']
                comp_width = comp['position']['width']
                comp_right = comp_left + comp_width
                comp_left_edge = comp_left  # Use LEFT edge for consistency
                comp_x_pt = comp_left_edge // 12700
                # Use same spanning table logic as assignment
                shape_type = comp.get('shape_type', 'unknown')
                is_spanning_table = (shape_type == 'Table' and 
                                   comp_left < best_x and comp_right > best_x and
                                   comp_width > (best_x - comp_left) * 2)
                if is_spanning_table:
                    base_side = "LEFT" if comp_left_edge < best_x else "RIGHT"
                    side = f"{base_side} (SPANNING-LEFTEDGE)"
                else:
                    side = "LEFT" if comp_left_edge < best_x else "RIGHT"
                print(f"               {i:2d}. {comp['box_id']} (leftedge={comp_x_pt}pt) - {side}")
            
            return result
        
        # Fallback to simple reading order
        print(f"            📖 {group_id} fallback to simple reading order")
        return self.sort_shapes_by_simple_reading_order(components)



    def flatten_hierarchical_groups_to_components(self, ordered_groups, line_elements):
        """Flatten groups to component list, maintaining hierarchical reading order"""
        all_components = []
        
        for group_idx, group in enumerate(ordered_groups, 1):
            # Sort components within this group with simple LEFT→RIGHT reading (LINE SECTIONING DISABLED)
            group_components = group['members']
            sorted_group_components = self.sort_shapes_by_simple_reading_order(group_components)
            # DISABLED: self.sort_group_components_with_nearby_vertical_lines(group_components, group['group_id'], line_elements)
            print(f"            📖 {group['group_id']}: Using simple LEFT→RIGHT reading (line sectioning disabled)")
            
            # Add group hierarchy info to each component
            for comp_idx, component in enumerate(sorted_group_components, 1):
                component['hierarchical_info'] = {
                    'group_order': group_idx,
                    'component_order_in_group': comp_idx,
                    'group_id': group['group_id'],
                    'total_groups': len(ordered_groups),
                    'total_components_in_group': len(sorted_group_components)
                }
                all_components.append(component)
        
        print(f"      📋 Flattened {len(ordered_groups)} groups to {len(all_components)} components with hierarchy")
        return all_components



    def create_hierarchical_sections(self, ordered_groups, line_elements):
        """Create sections representing hierarchical group-based reading"""
        return [{
            'section_id': 'hierarchical_group_based',
            'section_type': 'hierarchical_group_reading',
            'bounds': self.calculate_groups_bounds(ordered_groups),
            'groups': ordered_groups,
            'reading_order': 1,
            'group_line_interruptions_applied': len([l for l in line_elements if max(l['position']['width'], 1) / max(l['position']['height'], 1) < 0.7]),
            'note': f'Hierarchical reading: {len(ordered_groups)} groups with {len(line_elements)} line interruptions'
        }]



    def calculate_groups_bounds(self, groups):
        """Calculate bounding box for a collection of groups"""
        all_components = []
        for group in groups:
            all_components.extend(group['members'])
        return calculate_shapes_bounds(all_components)



    def apply_precise_line_boundary_reading_order(self, content_elements, line_elements):
        """Apply reading order with local line interruptions (not global sectioning)"""
        print(f"   📖 Applying local line interruptions for {len(line_elements)} lines")
        
        # Start with normal reading order as the foundation
        reading_order = self.sort_shapes_by_simple_reading_order(content_elements)
        
        # Apply local line interruptions
        for line in line_elements:
            reading_order = self.apply_local_line_interruption(reading_order, line)
        
        print(f"      ✅ Local line interruptions applied to {len(reading_order)} elements")
        return reading_order



    def apply_local_line_interruption(self, current_order, line_element):
        """Apply local line interruption to reading order within line's influence area"""
        pos = line_element['position']
        width = max(pos['width'], 1)
        height = max(pos['height'], 1)
        aspect_ratio = width / height
        
        # Only process clear vertical lines
        if aspect_ratio >= 0.7:
            print(f"         📏 {line_element['box_id']}: Not a clear vertical line (ratio={aspect_ratio:.2f}), skipping")
            return current_order
        
        line_center_x = pos['left'] + width // 2
        line_top = pos['top']
        line_bottom = pos['top'] + height
        
        print(f"         📏 Processing vertical line {line_element['box_id']} at x={line_center_x}")
        
        # Define LOCAL influence area around the line (REDUCED from width*10 to width*2)
        proximity_threshold = 50 * 12700  # 50pt in EMUs (~1.76cm) - only affect very close components
        
        influence_area = {
            'left': line_center_x - max(width * 2, proximity_threshold),  # Much smaller horizontal expansion
            'right': line_center_x + max(width * 2, proximity_threshold),
            'top': max(0, line_top - height * 0.1),  # Small vertical expansion
            'bottom': line_bottom + height * 0.1
        }
        
        print(f"         📏 Line influence area: X={line_center_x-max(width*2, proximity_threshold):.0f} to {line_center_x+max(width*2, proximity_threshold):.0f} (±{max(width*2, proximity_threshold)/12700:.1f}pt from line)")
        
        # Find elements within the line's LOCAL influence area
        influenced_elements = []
        non_influenced_elements = []
        
        for element in current_order:
            element_center_x = element['position']['left'] + element['position']['width'] // 2
            element_center_y = element['position']['top'] + element['position']['height'] // 2
            
            # COMPONENT-BOUND CHECK: Only affect elements that are VERY close to the line
            distance_to_line = abs(element_center_x - line_center_x)
            is_vertically_aligned = (influence_area['top'] <= element_center_y <= influence_area['bottom'])
            is_close_enough = distance_to_line <= max(width * 2, proximity_threshold)
            
            if is_close_enough and is_vertically_aligned:
                influenced_elements.append(element)
                print(f"            📍 {element.get('box_id', 'unknown')} (X={element_center_x/12700:.0f}pt) - INFLUENCED (distance: {distance_to_line/12700:.1f}pt)")
            else:
                non_influenced_elements.append(element)
                if distance_to_line > max(width * 2, proximity_threshold):
                    print(f"            ⚪ {element.get('box_id', 'unknown')} (X={element_center_x/12700:.0f}pt) - TOO FAR (distance: {distance_to_line/12700:.1f}pt)")
        
        if len(influenced_elements) < 2:
            print(f"            ⚠️  Only {len(influenced_elements)} elements in influence area, no interruption needed")
            return current_order
        
        print(f"            📍 Found {len(influenced_elements)} elements in line influence area")
        
        # Split influenced elements by line boundary
        left_influenced = []
        right_influenced = []
        
        for element in influenced_elements:
            element_center_x = element['position']['left'] + element['position']['width'] // 2
            if element_center_x < line_center_x:
                left_influenced.append(element)
            else:
                right_influenced.append(element)
        
        print(f"            📍 Line influence: LEFT({len(left_influenced)}) | RIGHT({len(right_influenced)})")
        
        if not left_influenced or not right_influenced:
            print(f"            ⚠️  No boundary split needed")
            return current_order
        
        # Sort each side by reading order
        left_sorted = self.sort_shapes_by_simple_reading_order(left_influenced)
        right_sorted = self.sort_shapes_by_simple_reading_order(right_influenced)
        
        # Reconstruct reading order with local line interruption
        result = []
        influenced_reordered = left_sorted + right_sorted
        influenced_index = 0
        
        for element in current_order:
            if element in influenced_elements:
                # Replace with reordered influenced element (LEFT first, then RIGHT)
                if influenced_index < len(influenced_reordered):
                    result.append(influenced_reordered[influenced_index])
                    influenced_index += 1
            else:
                # Keep non-influenced element in original reading order position
                result.append(element)
        
        print(f"            ✅ Applied local line interruption: LEFT({len(left_influenced)}) → RIGHT({len(right_influenced)})")
        return result



    def create_line_boundary_sections(self, content_elements, line_elements):
        """Create a single section representing locally line-interrupted reading order"""
        return [{
            'section_id': 'local_line_interrupted',
            'section_type': 'local_line_interrupted_reading',
            'bounds': calculate_shapes_bounds(content_elements),
            'shapes': content_elements,  # Will be reordered by apply_precise_line_boundary_reading_order
            'reading_order': 1,
            'line_interruptions_applied': len([l for l in line_elements if max(l['position']['width'], 1) / max(l['position']['height'], 1) < 0.7]),
            'note': f'Normal reading order with {len(line_elements)} local line interruptions'
        }]



    def sort_shapes_by_simple_reading_order(self, shapes):
        """Sort shapes with proper row grouping - components at same Y level read left-to-right"""
        if not shapes:
            return shapes
        
        print(f"         📖 Sorting {len(shapes)} shapes with row-aware reading order...")
        
        # STEP 1: Group components that are at similar vertical levels (rows)
        rows = []
        row_tolerance = 50000  # ~4pt tolerance for same row (very tight)
        
        # Sort by Y coordinate first to process top-to-bottom
        shapes_by_y = sorted(shapes, key=lambda s: s['position']['top'])
        
        for shape in shapes_by_y:
            shape_top = shape['position']['top']
            
            # Try to find existing row this shape belongs to
            added_to_existing_row = False
            for row in rows:
                # Check if this shape's Y is close to any shape in the existing row
                for existing_shape in row:
                    existing_top = existing_shape['position']['top']
                    if abs(shape_top - existing_top) <= row_tolerance:
                        row.append(shape)
                        added_to_existing_row = True
                        break
                if added_to_existing_row:
                    break
            
            # If not added to existing row, create new row
            if not added_to_existing_row:
                rows.append([shape])
        
        # STEP 2: Sort each row left-to-right, then combine
        final_sorted = []
        for i, row in enumerate(rows):
            # Sort this row by X coordinate (left to right)
            row.sort(key=lambda s: s['position']['left'])
            final_sorted.extend(row)
            
            # Debug: Show row contents
            row_y_values = [s['position']['top']//12700 for s in row]
            row_x_values = [s['position']['left']//12700 for s in row]
            row_ids = [s['box_id'] for s in row]
            
            if len(row) > 1:  # Only show multi-component rows
                print(f"         📖 Row {i+1}: {' → '.join(row_ids)} (Y: {min(row_y_values)}-{max(row_y_values)}pt)")
            else:
                print(f"         📖 Row {i+1}: {row_ids[0]} (single)")
        
        # DEBUG: Show final order
        if len(final_sorted) <= 15:
            print(f"         📖 Final row-aware reading order:")
            for i, shape in enumerate(final_sorted, 1):
                pos = shape['position']
                shape_type = shape.get('shape_type', 'unknown')
                print(f"            {i}. {shape['box_id']}: Y={pos['top']//12700}pt, X={pos['left']//12700}pt ({shape_type})")
        
        print(f"         ✅ Processed {len(shapes)} shapes into {len(rows)} rows")
        
        return final_sorted



    def get_reading_order_for_slide(self, slide_number):
        """Get reading order information for a specific slide from text structure"""
        if 'text_structure' not in self.comprehensive_data:
            return None
        
        for slide in self.comprehensive_data['text_structure']['slides']:
            if slide['slide_number'] == slide_number:
                return slide['content']
        return None



    def map_boxes_to_reading_order(self, slide_boxes, reading_order_info):
        """Map spatial boxes to reading order components"""
        if not reading_order_info:
            # If no reading order info, add default reading order based on position
            for i, box in enumerate(slide_boxes):
                box['reading_order'] = i + 1
                box['reading_order_text'] = box.get('text', '')
            return slide_boxes
        
        # Create a mapping of positions to reading order
        reading_order_positions = {}
        for ro_item in reading_order_info:
            pos = ro_item['position']
            reading_order_positions[f"{pos['top']}_{pos['left']}"] = {
                'reading_order': ro_item['reading_order'],
                'text': ro_item['text'],
                'is_title': ro_item['is_title']
            }
        
        # Map boxes to reading order
        for box in slide_boxes:
            box_pos = box['position']
            pos_key = f"{box_pos['top']}_{box_pos['left']}"
            
            if pos_key in reading_order_positions:
                ro_info = reading_order_positions[pos_key]
                box['reading_order'] = ro_info['reading_order']
                box['reading_order_text'] = ro_info['text']
                box['is_title'] = ro_info['is_title']
            else:
                # Try to find closest match by position
                closest_order = self._find_closest_reading_order(box_pos, reading_order_info)
                box['reading_order'] = closest_order
                box['reading_order_text'] = box.get('text', '')
                box['is_title'] = False
        
        return slide_boxes



    def _find_closest_reading_order(self, box_pos, reading_order_info):
        """Find closest reading order for a box based on position"""
        if not reading_order_info:
            return 999
        
        min_distance = float('inf')
        closest_order = 999
        
        box_center_x = box_pos['left'] + box_pos['width'] / 2
        box_center_y = box_pos['top'] + box_pos['height'] / 2
        
        for ro_item in reading_order_info:
            ro_pos = ro_item['position']
            ro_center_x = ro_pos['left'] + ro_pos['width'] / 2
            ro_center_y = ro_pos['top'] + ro_pos['height'] / 2
            
            distance = ((box_center_x - ro_center_x) ** 2 + (box_center_y - ro_center_y) ** 2) ** 0.5
            
            if distance < min_distance:
                min_distance = distance
                closest_order = ro_item['reading_order']
        
        return closest_order



    def _get_reading_order_span(self, root_box, members):
        """Calculate the reading order span of a group"""
        if not members:
            root_order = root_box.get('reading_order', 999)
            return {'min': root_order, 'max': root_order, 'span': 1}
        
        all_orders = [root_box.get('reading_order', 999)]
        all_orders.extend([m['box'].get('reading_order', 999) for m in members])
        
        valid_orders = [o for o in all_orders if o != 999]
        if not valid_orders:
            return {'min': 999, 'max': 999, 'span': 1}
        
        return {
            'min': min(valid_orders),
            'max': max(valid_orders),
            'span': max(valid_orders) - min(valid_orders) + 1
        }



    def create_enhanced_group_content_summary(self, all_slide_analyses):
        """Create enhanced comprehensive summary showing actual content for each smart group"""
        print(f"   📋 Creating enhanced group content summary...")
        
        content_summary_path = self.smart_groups_dir / "enhanced_group_content_summary.txt"
        
        with open(content_summary_path, 'w', encoding='utf-8') as f:
            f.write("ENHANCED SMART GROUP CONTENT SUMMARY\n")
            f.write("=" * 70 + "\n\n")
            f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Approach: Enhanced Smart Grouping with Spatial Containment\n")
            f.write(f"Foundation: Spatial containment relationships\n")
            f.write(f"Overlap Threshold: {self.overlap_threshold:.0%} (for group independence)\n")
            f.write(f"Containment Threshold: {self.containment_threshold:.0%} (for assignment)\n\n")
            
            f.write("ENHANCED FEATURES:\n")
            f.write("• 🎨 Dual visualization (Assignment + Structure views)\n")
            f.write("• 📝 Content preview in visual maps\n")
            f.write("• 🔗 Assignment quality indicators (Full/Partial)\n")
            f.write("• 📊 Detailed member type analysis\n")
            f.write("• 🧠 Spatial containment as primary organizing principle\n\n")
            
            # Overall statistics
            total_slides = len(all_slide_analyses)
            total_boxes = sum(s.get('total_boxes', 0) for s in all_slide_analyses)
            total_groups = sum(len(s.get('smart_groups', {})) for s in all_slide_analyses)
            total_assignments = sum(len(s.get('assignments', [])) for s in all_slide_analyses)
            
            f.write("OVERALL STATISTICS:\n")
            f.write(f"• Total Slides: {total_slides}\n")
            f.write(f"• Total Components: {total_boxes}\n")
            f.write(f"• Total Smart Groups: {total_groups}\n")
            f.write(f"• Total Assignments: {total_assignments}\n")
            f.write(f"• Average Groups per Slide: {total_groups/total_slides:.1f}\n\n")
            
            for analysis in all_slide_analyses:
                slide_num = analysis['slide_number']
                f.write(f"\n{'='*25} SLIDE {slide_num} {'='*25}\n\n")
                smart_groups = analysis['smart_groups']
                if not smart_groups:
                    f.write("No groups found on this slide.\n")
                    continue
                f.write(f"📊 SLIDE SUMMARY:\n")
                f.write(f"   • Total Components: {analysis.get('total_boxes', 0)}\n")
                f.write(f"   • Smart Groups: {len(smart_groups)}\n")
                f.write(f"   • Assignments Made: {len(analysis.get('assignments', []))}\n\n")
                # Process each smart group with enhanced details (original spatial approach)
                for group_name, group_data in smart_groups.items():
                    f.write(f"🏠 {group_name} (Enhanced Smart Group)\n")
                    f.write("-" * 50 + "\n")
                    
                    # Write root component info with enhanced details
                    root_box = group_data['root_component']
                    f.write(f"  🏠 ROOT COMPONENT:\n")
                    self._write_enhanced_component_content(f, root_box, "     ", slide_num)
                    
                    # Write member components with assignment details
                    members = group_data['members']
                    if members:
                        f.write(f"\n  📦 CONTAINS {len(members)} MEMBERS:\n")
                        for i, member in enumerate(members, 1):
                            containment_pct = member['containment_percentage']
                            assignment_type = member['assignment_type']
                            type_icon = "🔗" if assignment_type == "full" else "🔸"
                            
                            f.write(f"     {type_icon} MEMBER {i} ({containment_pct:.1%} contained, {assignment_type}):\n")
                            self._write_enhanced_component_content(f, member['box'], "        ", slide_num)
                    else:
                        f.write(f"\n  📄 STANDALONE COMPONENT (no members)\n")
                    
                    # Write enhanced member type summary
                    member_types = group_data.get('member_types', {})
                    if member_types:
                        f.write(f"\n  📊 MEMBER TYPE BREAKDOWN:\n")
                        for type_name, count in member_types.items():
                            f.write(f"     • {type_name}: {count} component{'s' if count > 1 else ''}\n")
                    
                    # Assignment quality analysis
                    if members:
                        full_assignments = len([m for m in members if m['assignment_type'] == 'full'])
                        partial_assignments = len([m for m in members if m['assignment_type'] == 'partial'])
                        
                        f.write(f"\n  🎯 ASSIGNMENT QUALITY:\n")
                        f.write(f"     • Full Containment (95%+): {full_assignments}\n")
                        f.write(f"     • Partial Containment (50-95%): {partial_assignments}\n")
                        
                        avg_containment = sum(m['containment_percentage'] for m in members) / len(members)
                        f.write(f"     • Average Containment: {avg_containment:.1%}\n")
                    
                    f.write("\n")
        
        print(f"      📋 Enhanced group content summary: {content_summary_path}")



    def create_reading_order_based_groups(self):
        """Create reading-order-based groups using local line sectioning data"""
        print(f"\n📖 CREATING READING ORDER BASED GROUPS (Local Line Sectioning)")
        print("=" * 50)
        
        if 'spatial_analysis' not in self.comprehensive_data:
            print("❌ No spatial analysis available for reading order groups")
            return False
        
        reading_order_groups_analysis = {
            'slides': [],
            'approach': 'local_line_sectioning_reading_order',
            'based_on': 'local_sections_with_individual_shapes',
            'reading_pattern': 'left_to_right_within_rows_top_to_bottom_between_rows',
            'table_extractions': self._get_table_extractions_info()  # NEW: Table references
        }
        
        # Process each slide using local sectioning data
        for slide_data in self.comprehensive_data['spatial_analysis']['slides']:
            print(f"   📖 Processing slide {slide_data['slide_number']} for reading order groups...")
            reading_order_slide = self.create_reading_order_from_local_sections(slide_data)
            if reading_order_slide:
                reading_order_groups_analysis['slides'].append(reading_order_slide)
        
        # Store the reading order groups analysis
        self.comprehensive_data['reading_order_groups'] = reading_order_groups_analysis
        
        # Save reading order groups analysis
        reading_order_groups_path = self.reading_order_groups_dir / "reading_order_groups_analysis.json"
        with open(reading_order_groups_path, 'w', encoding='utf-8') as f:
            json.dump(reading_order_groups_analysis, f, ensure_ascii=False, indent=2, default=str)
        
        # Create reading order groups summary using proper reading order groups data
        self.create_hierarchical_reading_order_summary(reading_order_groups_analysis)
        
        # Create hierarchical visualization (using spatial analysis data which has hierarchical_groups) - if enabled
        if self.config['generate_hierarchical_flow_visualizations'] and 'spatial_analysis' in self.comprehensive_data:
            self.create_reading_order_flow_visualization(self.comprehensive_data['spatial_analysis']['slides'])
        elif self.config['generate_hierarchical_flow_visualizations']:
            print("⚠️  No spatial analysis data available for hierarchical visualization")
        
        print(f"✅ Reading order based groups created: {len(reading_order_groups_analysis['slides'])} slides")
        return True



    def create_hierarchical_reading_order_summary(self, reading_order_groups_analysis):
        """Create reading order summary using proper reading order groups data"""
        summary_path = self.reading_order_groups_dir / "reading_order_groups_summary.txt"
        
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write("HIERARCHICAL READING ORDER GROUPS SUMMARY\n")
            f.write("=" * 65 + "\n\n")
            f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("Approach: Hierarchical Group-based Reading Order with Vertical Line Awareness\n")
            f.write("Reading Pattern: TOP→BOTTOM between groups, LEFT→RIGHT within groups\n")
            f.write("Line Handling: Vertical lines split groups into LEFT then RIGHT sides\n\n")
            
            # Process each slide using reading order groups data
            total_groups = 0
            total_components = 0
            
            for slide_data in reading_order_groups_analysis['slides']:
                slide_num = slide_data['slide_number']
                reading_order_groups = slide_data.get('reading_order_groups', [])

                # Pull hierarchical_groups up here so the non-fallback branch
                # below can also reference it for the SLIDE SUMMARY block.
                hierarchical_groups = []
                if 'spatial_analysis' in self.comprehensive_data:
                    for spatial in self.comprehensive_data['spatial_analysis']['slides']:
                        if spatial['slide_number'] == slide_num:
                            hierarchical_groups = spatial.get('hierarchical_groups', [])
                            break

                if not reading_order_groups:
                    # Fallback to spatial analysis data if reading order groups are empty
                    if 'spatial_analysis' in self.comprehensive_data:
                        spatial_slide = None
                        for spatial in self.comprehensive_data['spatial_analysis']['slides']:
                            if spatial['slide_number'] == slide_num:
                                spatial_slide = spatial
                                break
                        if spatial_slide and spatial_slide.get('hierarchical_groups'):
                            all_components = []
                            for box in spatial_slide.get('boxes', []):
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
                        else:
                            continue
                    else:
                        continue
                else:
                    # Use proper reading order groups data
                    all_components = []
                    for group in reading_order_groups:
                        all_components.extend(group.get('components', []))
                
                f.write(f"{'='*25} SLIDE {slide_num} {'='*25}\n\n")
                # Statistics
                f.write(f"📊 SLIDE SUMMARY:\n")
                f.write(f"   • Groups: {len(hierarchical_groups)}\n")
                f.write(f"   • Content Components: {len(all_components)} (layout containers excluded)\n\n")
                # Group by group
                current_group = None
                group_components = []
                for component in all_components:
                    group_id = component['hierarchical_info']['group_id']
                    
                    if current_group != group_id:
                        # Write previous group
                        if current_group and group_components:
                            self._write_group_summary(f, current_group, group_components, slide_num)
                        
                        # Start new group
                        current_group = group_id
                        group_components = []
                    
                    group_components.append(component)
                # Write last group
                if current_group and group_components:
                    self._write_group_summary(f, current_group, group_components, slide_num)
                f.write("\n")
                total_groups += len(hierarchical_groups)
                total_components += len(all_components)
            
            # Overall statistics
            if total_groups > 0:
                f.write(f"\n📈 OVERALL STATISTICS:\n")
                f.write(f"   • Total Groups: {total_groups}\n")
                f.write(f"   • Content Components: {total_components} (layout containers excluded)\n")
                f.write(f"   • Average Components per Group: {total_components/total_groups:.1f}\n")
        
        print(f"📋 Hierarchical reading order summary saved: {summary_path}")
        return True



    def _write_group_summary(self, f, group_id, components, slide_number):
        """Write summary for a single group"""
        f.write(f"📦 GROUP {group_id} ({len(components)} components):\n")
        
        # Add group location if available
        if components:
            # Calculate group bounds from components
            group_bounds = calculate_group_bounds(components)
            if group_bounds:
                f.write(f"   📍 Location: Top={group_bounds['top']:.0f}, Bottom={group_bounds['bottom']:.0f}, Left={group_bounds['left']:.0f}, Right={group_bounds['right']:.0f} (W×H: {group_bounds['width']:.0f}×{group_bounds['height']:.0f})\n")
        
        for i, component in enumerate(components, 1):
            # Check if this is a table - if so, use hierarchical display
            if component.get('shape_type') == 'Table' and 'cell_contents' in component:
                self._write_hierarchical_table_summary(f, component, i, slide_number)
            # Check if this is a UnifiedGroup - if so, use hierarchical display
            elif component.get('shape_type') == 'UnifiedGroup' and ('component_images' in component or 'component_texts' in component):
                self._write_hierarchical_unified_group_summary(f, component, i, slide_number)
            # Check if this is a SmartVisualGroup that contains tables
            elif component.get('shape_type') == 'SmartVisualGroup' and 'component_visuals' in component:
                self._write_smart_visual_group_summary(f, component, i, slide_number)
            else:
                self._write_standard_component_summary(f, component, i, slide_number)
        
        f.write("\n")



    def _write_hierarchical_table_summary(self, f, component, index, slide_number):
        """Write hierarchical table display showing container + individual cells"""
        box_id = component['box_id']
        table_dimensions = component.get('table_dimensions', 'unknown')
        
        # Check for visual capture
        visual_capture = self.find_visual_capture_file(box_id, slide_number)
        visual_info = f" → {visual_capture}" if visual_capture else ""
        
        # Add position information
        pos = component.get('position', {})
        position_info = ""
        if pos:
            position_info = f" [T={pos.get('top', 0):.0f}, L={pos.get('left', 0):.0f}, W={pos.get('width', 0):.0f}×H={pos.get('height', 0):.0f}]"
        
        # Write table container
        f.write(f"   {index:2d}. {box_id} (Table): [Table Container {table_dimensions}]{visual_info}{position_info}\n")
        
        # Write individual cells
        cell_contents = component.get('cell_contents', [])
        print(f"      🔍 DEBUG - Table {box_id} has {len(cell_contents)} cells:")
        for cell in cell_contents:
            print(f"         Cell({cell['row']},{cell['col']}): has_content={cell['has_content']}, text='{cell.get('text', '')[:20]}...', display='{cell.get('display_text', 'N/A')}'")
        if cell_contents:
            # Group cells by row for better display
            rows = {}
            for cell in cell_contents:
                if cell['has_content']:  # Only show cells with content
                    row_idx = cell['row']
                    if row_idx not in rows:
                        rows[row_idx] = []
                    rows[row_idx].append(cell)
            
            # Display cells row by row
            for row_idx in sorted(rows.keys()):
                for cell in sorted(rows[row_idx], key=lambda c: c['col']):
                    # Use display_text which includes "[IMAGE CELL]" for image cells
                    # Show FULL text (no truncation as requested by user)
                    display_text = cell.get('display_text', cell.get('text', '')).replace('\n', ' ').replace('\r', ' ')
                    
                    # Render the cell. A cell may have:
                    #   (a) only text                     → "Cell(r,c): text"
                    #   (b) only visual content          → image-cell style + Contains: lines
                    #   (c) BOTH text and visual content → text first, then Contains: lines
                    has_visual = cell.get('has_visual_content', False)
                    cell_text = cell.get('text', '').strip()

                    if has_visual and not cell_text:
                        # Pure visual cell – legacy rendering path.
                        visual_emoji = "🖼️"  # default for images
                        if 'shapes' in cell and cell['shapes']:
                            visual_types = [shape['type'] for shape in cell['shapes']]
                            if 'chart' in visual_types:
                                visual_emoji = "📈"
                            elif 'graphic' in visual_types:
                                visual_emoji = "🎨"
                        elif 'overlapping_shapes' in cell and cell['overlapping_shapes']:
                            overlapping_types = [shape['shape_type'] for shape in cell['overlapping_shapes']]
                            if 'Chart' in overlapping_types:
                                visual_emoji = "📈"
                        f.write(f"       ├── Cell({cell['row']},{cell['col']}): {display_text} {visual_emoji}\n")
                    else:
                        # Either text-only or text + visual content. Always show the text first.
                        f.write(f"       ├── Cell({cell['row']},{cell['col']}): \"{display_text}\"\n")

                    # Whenever the cell carries visual content (image / table / chart / shape),
                    # append Contains: lines so the embedded shapes are not dropped from the
                    # reading-order summary, regardless of whether text was present.
                    if has_visual and 'overlapping_shapes' in cell and cell['overlapping_shapes']:
                        for shape in cell['overlapping_shapes']:
                            visual_path = f" → {shape['visual_capture']}" if shape.get('visual_capture') else ""
                            overlap_info = f" ({shape['overlap_percentage']:.1f}% overlap)" if shape.get('overlap_percentage') else ""
                            text_info = f" - \"{shape['text_content']}\"" if shape.get('text_content') else ""
                            f.write(f"       │   └── Contains: {shape['box_id']} ({shape['shape_type']}){overlap_info}{text_info}{visual_path}\n")
        else:
            f.write(f"       └── [Empty table or no cell data available]\n")



    def _write_hierarchical_unified_group_summary(self, f, component, index, slide_number):
        """Write hierarchical unified group display showing container + individual components"""
        box_id = component['box_id']
        
        # Check for visual capture
        visual_capture = self.find_visual_capture_file(box_id, slide_number)
        visual_info = f" → {visual_capture}" if visual_capture else ""
        
        # Add position information
        pos = component.get('position', {})
        position_info = ""
        if pos:
            position_info = f" [T={pos.get('top', 0):.0f}, L={pos.get('left', 0):.0f}, W={pos.get('width', 0):.0f}×H={pos.get('height', 0):.0f}]"
        
        # Get text content for the unified group
        text_content = component.get('text', '').strip()
        text_display = f'"{text_content}"' if text_content else "[Combined Visual Group]"
        
        # Write unified group container
        f.write(f"   {index:2d}. {box_id} (UnifiedGroup): {text_display}{visual_info}{position_info}\n")
        
        # Count total components
        component_images = component.get('component_images', [])
        component_texts = component.get('component_texts', [])
        component_lines = component.get('component_lines', [])
        total_components = len(component_images) + len(component_texts) + len(component_lines)
        
        f.write(f"       📦 Contains {total_components} visual components:\n")
        
        # Write individual images
        for img in component_images:
            img_text = img.get('text', '').strip()
            img_display = f'"{img_text}"' if img_text else "[No text]"
            f.write(f"       ├── {img['box_id']} ({img.get('shape_type', 'Unknown')}): {img_display}\n")
        
        # Write individual texts  
        for txt in component_texts:
            txt_text = txt.get('text', '').strip()
            txt_display = f'"{txt_text}"' if txt_text else "[No text]"
            f.write(f"       ├── {txt['box_id']} ({txt.get('shape_type', 'Unknown')}): {txt_display}\n")
        
        # Write individual lines (if any)
        for line in component_lines:
            f.write(f"       ├── {line['box_id']} (Line): [Structural line]\n")



    def _write_smart_visual_group_summary(self, f, component, index, slide_number):
        """Write SmartVisualGroup summary showing component tables hierarchically"""
        box_id = component['box_id']
        component_visuals = component.get('component_visuals', [])
        
        # Check for visual capture
        visual_capture = self.find_visual_capture_file(box_id, slide_number)
        visual_info = f" → {visual_capture}" if visual_capture else ""
        
        # Add position information
        pos = component.get('position', {})
        position_info = ""
        if pos:
            position_info = f" [T={pos.get('top', 0):.0f}, L={pos.get('left', 0):.0f}, W={pos.get('width', 0):.0f}×H={pos.get('height', 0):.0f}]"
        
        # Get combined text
        combined_text = component.get('text', '').strip()
        if combined_text:
            text_preview = combined_text.replace('\n', ' ').replace('\r', ' ')
            if len(text_preview) > 60:
                text_preview = text_preview[:60] + "..."
            text_display = f'"{text_preview}"'
        else:
            text_display = "[No text]"
        
        # Write group container
        f.write(f"   {index:2d}. {box_id} (SmartVisualGroup): {text_display}{visual_info}{position_info}\n")
        f.write(f"       📦 Contains {len(component_visuals)} visual components:\n")
        
        # Display each component in the group
        for j, visual_component in enumerate(component_visuals, 1):
            comp_id = visual_component.get('box_id', f'Component_{j}')
            comp_type = visual_component.get('shape_type', 'Unknown')
            comp_text = visual_component.get('text', '').strip()
            
            # Check if this visual component is a table with cell contents
            if comp_type == 'Table' and 'cell_contents' in visual_component:
                # Display table hierarchically
                table_dimensions = visual_component.get('table_dimensions', 'unknown')
                f.write(f"       ├── {comp_id} ({comp_type}): [Table Container {table_dimensions}]\n")
                
                # Show cell contents
                cell_contents = visual_component.get('cell_contents', [])
                if cell_contents:
                    # Group cells by row for better display
                    rows = {}
                    for cell in cell_contents:
                        if cell['has_content']:  # Only show cells with content
                            row_idx = cell['row']
                            if row_idx not in rows:
                                rows[row_idx] = []
                            rows[row_idx].append(cell)
                    
                    # Display cells row by row
                    for row_idx in sorted(rows.keys()):
                        for cell in sorted(rows[row_idx], key=lambda c: c['col']):
                            cell_text = cell['text'].replace('\n', ' ').replace('\r', ' ')
                            if len(cell_text) > 40:
                                cell_text = cell_text[:40] + "..."
                            f.write(f"       │   ├── Cell({cell['row']},{cell['col']}): \"{cell_text}\"\n")
                else:
                    f.write(f"       │   └── [Empty table or no cell data]\n")
            else:
                # Display non-table components normally
                if comp_text:
                    comp_text_preview = comp_text.replace('\n', ' ').replace('\r', ' ')
                    if len(comp_text_preview) > 40:
                        comp_text_preview = comp_text_preview[:40] + "..."
                    comp_text_display = f'"{comp_text_preview}"'
                else:
                    comp_text_display = "[No text]"
                f.write(f"       ├── {comp_id} ({comp_type}): {comp_text_display}\n")



    def _write_standard_component_summary(self, f, component, index, slide_number):
        """Write standard component summary (non-table components)"""
        # Get component info
        box_id = component['box_id']
        shape_type = component.get('shape_type', 'Unknown')
        text_content = component.get('text', '').strip()
        
        # Format text preview
        if text_content:
            text_preview = text_content.replace('\n', ' ').replace('\r', ' ')
            if len(text_preview) > 60:
                text_preview = text_preview[:60] + "..."
            text_display = f'"{text_preview}"'
        else:
            text_display = "[No text]"
        
        # Check for visual capture
        visual_capture = self.find_visual_capture_file(box_id, slide_number)
        visual_info = f" → {visual_capture}" if visual_capture else ""
        
        # Add position information
        pos = component.get('position', {})
        position_info = ""
        if pos:
            position_info = f" [T={pos.get('top', 0):.0f}, L={pos.get('left', 0):.0f}, W={pos.get('width', 0):.0f}×H={pos.get('height', 0):.0f}]"
        
        # Base component line
        base_line = f"   {index:2d}. {box_id} ({shape_type}): {text_display}{visual_info}{position_info}"
        
        # Add VLM caption if it's a visual component with no text
        if not text_content and visual_capture:
            vlm_caption = self.find_vlm_caption_for_visual_file(visual_capture)
            if vlm_caption:
                f.write(f"{base_line} -> 📸 **{vlm_caption['title']}** - {vlm_caption['description'][:100]}{'...' if len(vlm_caption['description']) > 100 else ''}\n")
                f.write(f"       🏷️ Korean Filename: `{vlm_caption['korean_filename']}`\n")
            else:
                f.write(f"{base_line}\n")
        else:
            f.write(f"{base_line}\n")



    def create_reading_order_from_local_sections(self, slide_data):
        """Create reading order groups from local sectioning data - complete each section before moving to next"""
        slide_number = slide_data['slide_number']
        local_sections = slide_data.get('local_sections', [])
        line_dividers = slide_data.get('line_dividers', [])
        
        if not local_sections:
            return None
        
        # Step 1: Order sections themselves by reading order (left-to-right, top-to-bottom)
        ordered_sections = self.order_sections_by_reading_order(local_sections)
        
        # Step 2: Process each section in order and assign global reading numbers
        reading_order_groups = []
        overall_reading_order = 1
        
        for section_idx, section in enumerate(ordered_sections):
            section_shapes = section.get('shapes', [])
            if not section_shapes:
                continue
            
            # Create a reading order group for this section
            section_group = {
                'group_id': f"Section_{section_idx + 1}",
                'section_info': {
                    'section_id': section['section_id'],
                    'section_type': section['section_type'],
                    'divider_line': section.get('divider_line', 'none'),
                    'bounds': section.get('bounds', {}),
                    'section_reading_order': section_idx + 1  # Order of this section among all sections
                },
                'components': [],
                'total_components': len(section_shapes),
                'reading_order_start': overall_reading_order,
                'reading_order_end': overall_reading_order + len(section_shapes) - 1
            }
            
            # Add each component in reading order within this section
            for comp_idx, shape in enumerate(section_shapes):
                component = {
                    'component_id': shape['box_id'],
                    'reading_order': overall_reading_order,
                    'section_position': comp_idx + 1,
                    'section_order': section_idx + 1,  # Which section this belongs to
                    'shape_type': shape.get('shape_type', 'unknown'),
                    'box_type': shape.get('box_type', 'unknown'),
                    'has_text': shape.get('has_text', False),
                    'text_content': shape.get('text', '').strip(),
                    'position': shape['position'],
                    'visual_capture_file': self.find_visual_capture_file(shape['box_id'], slide_number) if shape.get('box_type') in ['picture', 'group', 'unified_group'] else None
                }
                section_group['components'].append(component)
                overall_reading_order += 1
            
            reading_order_groups.append(section_group)
        
        return {
            'slide_number': slide_number,
            'title': slide_data.get('title', ''),
            'total_line_dividers': len(line_dividers),
            'total_sections': len(local_sections),
            'total_components': sum(len(group['components']) for group in reading_order_groups),
            'reading_order_groups': reading_order_groups,
            'line_dividers': line_dividers,
            'section_ordering_applied': True,
            'analysis_timestamp': datetime.now().isoformat()
        }



    def order_sections_by_reading_order(self, local_sections):
        """Order sections themselves by reading order: left-to-right, top-to-bottom"""
        if not local_sections:
            return local_sections
        
        print(f"         📋 Ordering {len(local_sections)} sections by reading order...")
        
        # Calculate center point of each section for ordering
        sections_with_centers = []
        for section in local_sections:
            bounds = section.get('bounds', {})
            if bounds:
                center_x = (bounds.get('left', 0) + bounds.get('right', 0)) / 2
                center_y = (bounds.get('top', 0) + bounds.get('bottom', 0)) / 2
            else:
                # If no bounds, use average position of shapes in section
                shapes = section.get('shapes', [])
                if shapes:
                    avg_left = sum(s['position']['left'] for s in shapes) / len(shapes)
                    avg_top = sum(s['position']['top'] for s in shapes) / len(shapes)
                    avg_width = sum(s['position']['width'] for s in shapes) / len(shapes)
                    avg_height = sum(s['position']['height'] for s in shapes) / len(shapes)
                    center_x = avg_left + avg_width / 2
                    center_y = avg_top + avg_height / 2
                else:
                    center_x = center_y = 0
            
            sections_with_centers.append({
                'section': section,
                'center_x': center_x,
                'center_y': center_y
            })
        
                    # Sort sections by reading order: maintain grid-based flow (sections are already positioned correctly)
        # Use row grouping similar to shape sorting
        if len(sections_with_centers) > 1:
            avg_height = sum(abs(bounds.get('bottom', 0) - bounds.get('top', 0)) for bounds in [s['section'].get('bounds', {}) for s in sections_with_centers]) / len(sections_with_centers)
            row_tolerance = max(avg_height * 0.3, 100)  # 30% of average section height
            
            def get_section_reading_order_key(section_data):
                center_y = section_data['center_y']
                center_x = section_data['center_x']
                # Create row groups for sections
                row_group = int(center_y / row_tolerance) * row_tolerance
                return (
                    row_group,  # Primary: Row (top-to-bottom)
                    center_x    # Secondary: Column (left-to-right within row)
                )
            
            sections_with_centers.sort(key=get_section_reading_order_key)
            print(f"           ✅ Sections ordered: {[s['section']['section_id'] for s in sections_with_centers]}")
        
        return [s['section'] for s in sections_with_centers]



    def create_reading_order_components_summary(self, all_slide_analyses):
        """Create detailed summary showing individual components in reading order"""
        print(f"   📋 Creating reading order components summary...")
        
        summary_path = self.reading_order_groups_dir / "reading_order_groups_summary.txt"
        
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write("READING ORDER GROUPS SUMMARY (Local Line Sectioning)\n")
            f.write("=" * 70 + "\n\n")
            f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Approach: Local Line Sectioning with Individual Component Reading Order\n")
            f.write(f"Reading Pattern: LEFT→RIGHT within rows, TOP→BOTTOM between rows\n\n")
            
            f.write("📖 APPROACH:\n")
            f.write("1. Identify line shapes as local dividers\n")
            f.write("2. Create local sections around each line\n") 
            f.write("3. Order sections by reading order (top-to-bottom, left-to-right)\n")
            f.write("4. COMPLETE each section entirely before moving to next section\n")
            f.write("5. Within each section: sort components by standard reading order\n")
            f.write("6. Number all components sequentially across sections\n")
            f.write("7. Reading flow: Complete LEFT section, then RIGHT section (respects line dividers)\n\n")
            
            total_slides = len(all_slide_analyses)
            total_components = sum(s.get('total_components', 0) for s in all_slide_analyses)
            total_sections = sum(s.get('total_sections', 0) for s in all_slide_analyses)
            total_lines = sum(s.get('total_line_dividers', 0) for s in all_slide_analyses)
            
            f.write("OVERALL STATISTICS:\n")
            f.write(f"• Total Slides: {total_slides}\n")
            f.write(f"• Total Line Dividers: {total_lines}\n")
            f.write(f"• Total Local Sections: {total_sections}\n")
            f.write(f"• Total Components: {total_components}\n")
            f.write(f"• Average Components per Slide: {total_components/total_slides:.1f}\n\n")
            
            for analysis in all_slide_analyses:
                slide_num = analysis['slide_number']
                f.write(f"\n{'='*25} SLIDE {slide_num} {'='*25}\n\n")
                reading_order_groups = analysis['reading_order_groups']
                if not reading_order_groups:
                    f.write("No reading order groups found on this slide.\n")
                    continue
                f.write(f"📊 SLIDE SUMMARY:\n")
                f.write(f"   • Title: {analysis.get('title', 'Untitled')}\n")
                f.write(f"   • Line Dividers: {analysis.get('total_line_dividers', 0)}\n")
                f.write(f"   • Local Sections: {analysis.get('total_sections', 0)}\n")
                f.write(f"   • Total Components: {analysis.get('total_components', 0)}\n\n")
                # Show line dividers
                if analysis.get('line_dividers'):
                    f.write(f"📏 LINE DIVIDERS:\n")
                    for line in analysis['line_dividers']:
                        f.write(f"   • {line['box_id']}: {line.get('shape_type', 'Line')} divider\n")
                    f.write("\n")
                # Process each section with its components
                for group in reading_order_groups:
                    f.write(f"📍 {group['group_id']} - {group['section_info']['section_id']}\n")
                    f.write("-" * 60 + "\n")
                    f.write(f"  📋 Section Type: {group['section_info']['section_type']}\n")
                    if group['section_info']['divider_line'] != 'none':
                        f.write(f"  📏 Divider Line: {group['section_info']['divider_line']}\n")
                    f.write(f"  🔢 Components: {group['total_components']}\n")
                    f.write(f"  📖 Reading Order Range: #{group['reading_order_start']} - #{group['reading_order_end']}\n\n")
                    
                    f.write(f"  📝 COMPONENTS IN READING ORDER:\n")
                    for component in group['components']:
                        # Component header with section info
                        f.write(f"     #{component['reading_order']} - {component['component_id']} ")
                        f.write(f"(Section {component['section_order']}, Position #{component['section_position']})\n")
                        
                        # Component details
                        if component['box_type'] == 'table':
                            f.write(f"        📊 TABLE\n")
                        elif component['box_type'] == 'picture':
                            f.write(f"        🖼️  IMAGE\n")
                        elif component['box_type'] == 'group':
                            f.write(f"        📦 GROUP\n")
                        elif component['box_type'] == 'chart':
                            f.write(f"        📈 CHART\n")
                        else:
                            f.write(f"        📝 {component['box_type'].upper()}\n")
                        
                        # Add text content
                        if component['has_text'] and component['text_content']:
                            text_lines = component['text_content'].split('\n')
                            if len(text_lines) == 1 and len(component['text_content']) <= 80:
                                f.write(f"        Content: \"{component['text_content']}\"\n")
                            else:
                                f.write(f"        Content:\n")
                                for i, line in enumerate(text_lines[:2]):  # Show first 2 lines
                                    clean_line = line.strip()
                                    if clean_line:
                                        f.write(f"           \"{clean_line}\"\n")
                                if len(text_lines) > 2:
                                    f.write(f"           ... ({len(text_lines)} total lines)\n")
                        else:
                            f.write(f"        Content: [no text content]\n")
                        
                        # Add visual capture file reference
                        if component['visual_capture_file']:
                            f.write(f"        📸 Visual capture: {component['visual_capture_file']}\n")
                        
                        # Add position info
                        pos = component['position']
                        f.write(f"        Position: ({pos.get('left', 0)}, {pos.get('top', 0)}) "
                                f"Size: {pos.get('width', 0)}×{pos.get('height', 0)}\n")
                        f.write("\n")
                    
                    f.write("\n")
        
        print(f"      📋 Reading order components summary: {summary_path}")
        return summary_path



    def _write_enhanced_component_content(self, file_handle, box, prefix="", slide_number=None):
        """Write component content with enhanced formatting and details"""
        box_type = box.get('box_type', 'unknown')
        box_id = box.get('box_id', 'unknown')
        
        # Enhanced type-specific formatting
        if box_type == 'table':
            file_handle.write(f"{prefix}📊 TABLE ({box_id})\n")
            file_handle.write(f"{prefix}   Type: Data table component\n")
        elif box_type == 'picture':
            file_handle.write(f"{prefix}🖼️  IMAGE ({box_id})\n")
            file_handle.write(f"{prefix}   Type: Visual content/image\n")
        elif box_type == 'group':
            file_handle.write(f"{prefix}📦 GROUP ({box_id})\n")
            file_handle.write(f"{prefix}   Type: Grouped elements\n")
        elif box_type == 'consolidated_image':
            file_handle.write(f"{prefix}🖼️  CONSOLIDATED IMAGE ({box_id})\n")
            file_handle.write(f"{prefix}   Type: {box.get('consolidated_images_count', 0)} overlapping images merged\n")
            # Show constituent images
            constituent_images = box.get('constituent_images', [])
            if constituent_images:
                file_handle.write(f"{prefix}   Constituent Images: {', '.join([img['box_id'] for img in constituent_images])}\n")
        elif box_type == 'image_autoshape_combo':
            file_handle.write(f"{prefix}🖼️🔗 IMAGE+AUTOSHAPE COMBO ({box_id})\n")
            file_handle.write(f"{prefix}   Type: Image with overlapping AutoShapes\n")
            # Show primary image
            primary_image = box.get('primary_image', {})
            if primary_image:
                file_handle.write(f"{prefix}   Primary Image: {primary_image.get('box_id', 'N/A')}\n")
            # Show overlapping AutoShapes
            autoshapes = box.get('overlapping_autoshapes', [])
            if autoshapes:
                autoshape_ids = [ash['box_id'] for ash in autoshapes]
                file_handle.write(f"{prefix}   Overlapping AutoShapes: {', '.join(autoshape_ids)}\n")
        elif box_type == 'chart':
            file_handle.write(f"{prefix}📈 CHART ({box_id})\n")
            file_handle.write(f"{prefix}   Type: Data visualization\n")
        else:
            # Enhanced text content handling
            if box.get('has_text', False) and box.get('text', '').strip():
                text_content = box['text'].strip()
                text_lines = text_content.split('\n')
                file_handle.write(f"{prefix}📝 TEXT ({box_id}):\n")
                file_handle.write(f"{prefix}   Type: {box_type.replace('_', ' ').title()}\n")
                # Write text content with proper formatting
                if len(text_lines) == 1 and len(text_content) <= 100:
                    file_handle.write(f"{prefix}   Content: \"{text_content}\"\n")
                else:
                    file_handle.write(f"{prefix}   Content:\n")
                    for i, line in enumerate(text_lines[:3]):  # Show first 3 lines
                        clean_line = line.strip()
                        if clean_line:
                            file_handle.write(f"{prefix}      \"{clean_line}\"\n")
                    if len(text_lines) > 3:
                        file_handle.write(f"{prefix}      ... ({len(text_lines)} total lines)\n")
            else:
                file_handle.write(f"{prefix}📄 {box_type.upper()} ({box_id})\n")
                file_handle.write(f"{prefix}   Type: {box_type.replace('_', ' ').title()}\n")
                file_handle.write(f"{prefix}   Content: [no text content]\n")
        
        # Add position information
        pos = box.get('position', {})
        if pos:
            file_handle.write(f"{prefix}   Position: ({pos.get('left', 0)}, {pos.get('top', 0)}) "
                            f"Size: {pos.get('width', 0)}×{pos.get('height', 0)}\n")
        
        # Add visual capture file mapping if available
        if slide_number and box_type in ['picture', 'group', 'unified_group']:
            visual_capture_file = self.find_visual_capture_file(box_id, slide_number)
            if visual_capture_file:
                file_handle.write(f"{prefix}   📸 Visual capture: {visual_capture_file}\n")



    def create_reading_order_integration(self):
        """Integrate text structure with spatial groups using reading order"""
        print(f"\n📖 CREATING READING ORDER INTEGRATION")
        print("=" * 50)
        
        if 'text_structure' not in self.comprehensive_data or 'smart_groups' not in self.comprehensive_data:
            print("❌ Missing text structure or smart groups for reading order integration")
            return False
        
        reading_order_analysis = {
            'slides': []
        }
        
        # Process each slide
        for slide_idx, text_slide in enumerate(self.comprehensive_data['text_structure']['slides']):
            slide_number = text_slide['slide_number']
            
            print(f"   📖 Processing slide {slide_number} reading order integration...")
            
            # Find corresponding smart groups slide
            smart_groups_slide = None
            for sg_slide in self.comprehensive_data['smart_groups']['slides']:
                if sg_slide['slide_number'] == slide_number:
                    smart_groups_slide = sg_slide
                    break
            
            if smart_groups_slide:
                integrated_slide = self.integrate_slide_reading_order(text_slide, smart_groups_slide)
                reading_order_analysis['slides'].append(integrated_slide)
        
        self.comprehensive_data['reading_order'] = reading_order_analysis
        
        # Save reading order integration
        reading_order_path = self.reading_order_dir / "reading_order_integration.json"
        with open(reading_order_path, 'w', encoding='utf-8') as f:
            json.dump(reading_order_analysis, f, ensure_ascii=False, indent=2, default=str)
        
        # Create comprehensive summary
        self.create_comprehensive_summary()
        
        print(f"✅ Reading order integration created: {len(reading_order_analysis['slides'])} slides")
        return True



    def integrate_slide_reading_order(self, text_slide, smart_groups_slide):
        """Integrate text structure with smart groups for a single slide"""
        integrated_slide = {
            'slide_number': text_slide['slide_number'],
            'title': text_slide['title'],
            'layout_name': text_slide['layout_name'],
            'reading_order_groups': [],
            'text_content': text_slide['content'],
            'smart_groups': smart_groups_slide['smart_groups'],
            'notes': text_slide['notes']
        }
        
        # Map text content to spatial groups based on position
        text_content = text_slide['content']
        smart_groups = smart_groups_slide['smart_groups']
        
        # Create reading order groups by mapping text to spatial groups
        for text_item in text_content:
            if text_item['is_title']:
                continue  # Skip title as it's handled separately
            
            # Find which spatial group this text belongs to
            best_group = self.find_text_spatial_group_match(text_item, smart_groups)
            
            reading_order_group = {
                'reading_order': text_item['reading_order'],
                'text_content': text_item['text'],
                'text_position': text_item.get('position', None),  # FIXED: Handle watsonx items without position
                'mapped_spatial_group': best_group['group_id'] if best_group else 'no_group',
                'spatial_components': []
            }
            
            # Add spatial components from the matched group
            if best_group:
                # Add root component
                reading_order_group['spatial_components'].append({
                    'component_type': 'root',
                    'box_id': best_group['root_component']['box_id'],
                    'box_type': best_group['root_component']['box_type'],
                    'has_text': best_group['root_component'].get('has_text', False),
                    'text': best_group['root_component'].get('text', '')
                })
                # Add member components
                for member in best_group['members']:
                    reading_order_group['spatial_components'].append({
                        'component_type': 'member',
                        'box_id': member['box_id'],
                        'box_type': member['box']['box_type'],
                        'has_text': member['box'].get('has_text', False),
                        'text': member['box'].get('text', ''),
                        'containment_percentage': member['containment_percentage']
                    })
            
            integrated_slide['reading_order_groups'].append(reading_order_group)
        
        return integrated_slide



    def find_text_spatial_group_match(self, text_item, smart_groups):
        """Find the best spatial group match for a text item"""
        # Check if text item has position information (native extraction vs watsonx)
        if 'position' not in text_item:
            # watsonx content items don't have position data - use text-based matching
            return self.find_text_group_match_by_content(text_item, smart_groups)
        
        text_pos = text_item['position']
        
        best_group = None
        best_overlap = 0.0
        
        for group_id, group_data in smart_groups.items():
            # Check overlap with root component
            root_pos = group_data['root_component']['position']
            
            # Simple overlap calculation
            overlap = calculate_text_box_overlap(text_pos, root_pos)
            
            if overlap > best_overlap:
                best_overlap = overlap
                best_group = group_data
        
        return best_group if best_overlap > 0.1 else None  # 10% minimum overlap



    def find_text_group_match_by_content(self, text_item, smart_groups):
        """Find spatial group match for watsonx text items using content similarity"""
        if not smart_groups:
            return None
            
        text_content = text_item.get('text', '').strip().lower()
        if not text_content:
            return None
        
        # For watsonx content, return the first available group as a reasonable fallback
        # since we don't have spatial coordinates to do precise matching
        first_group = next(iter(smart_groups.values()), None)
        
        # In a more sophisticated implementation, we could:
        # - Use text similarity scoring between watsonx text and group content
        # - Use reading order position as a proxy for spatial position
        # - Apply NLP-based matching techniques
        
        return first_group



    def save_local_sectioning_analysis(self, spatial_slides):
        """Save detailed analysis of local line sectioning with reading order info"""
        local_sectioning_data = {
            'analysis_type': 'local_line_sectioning',
            'approach': 'Lines divide only their immediate vicinity',
            'reading_order': 'Top-to-bottom, left-to-right within each section',
            'timestamp': datetime.now().isoformat(),
            'slides': []
        }
        
        total_lines = 0
        total_sections = 0
        
        for slide in spatial_slides:
            line_dividers = slide.get('line_dividers', [])
            local_sections = slide.get('local_sections', [])
            
            if line_dividers or local_sections:
                slide_sectioning = {
                    'slide_number': slide['slide_number'],
                    'title': slide.get('title', ''),
                    'total_shapes': slide.get('total_shapes', 0),
                    'line_dividers': len(line_dividers),
                    'local_sections': len(local_sections),
                    'sectioning_details': {
                        'lines': line_dividers,
                        'sections': local_sections
                    },
                    'reading_order_applied': True
                }
                local_sectioning_data['slides'].append(slide_sectioning)
                total_lines += len(line_dividers)
                total_sections += len(local_sections)
        
        # Add summary statistics
        local_sectioning_data['summary'] = {
            'total_slides_with_sectioning': len(local_sectioning_data['slides']),
            'total_line_dividers': total_lines,
            'total_local_sections': total_sections,
            'average_sections_per_slide': total_sections / len(local_sectioning_data['slides']) if local_sectioning_data['slides'] else 0
        }
        
        # Save to spatial analysis directory
        sectioning_path = self.spatial_analysis_dir / "local_sectioning_analysis.json"
        with open(sectioning_path, 'w', encoding='utf-8') as f:
            json.dump(local_sectioning_data, f, ensure_ascii=False, indent=2)
        
        # Create human-readable summary
        summary_path = self.spatial_analysis_dir / "local_sectioning_summary.txt"
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write("LOCAL LINE SECTIONING ANALYSIS\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Approach: Lines divide only their immediate vicinity\n")
            f.write(f"Reading Order: Left→Right within rows, Top→Bottom between rows (standard reading)\n\n")
            
            f.write("📊 OVERALL SUMMARY:\n")
            f.write(f"• Slides with line sectioning: {local_sectioning_data['summary']['total_slides_with_sectioning']}\n")
            f.write(f"• Total line dividers: {local_sectioning_data['summary']['total_line_dividers']}\n")
            f.write(f"• Total local sections: {local_sectioning_data['summary']['total_local_sections']}\n")
            f.write(f"• Average sections per slide: {local_sectioning_data['summary']['average_sections_per_slide']:.1f}\n\n")
            
            f.write("📖 READING ORDER APPROACH:\n")
            f.write("1. Identify line shapes (horizontal, vertical, diagonal)\n")
            f.write("2. Find local area around each line (not entire slide)\n")
            f.write("3. Group content shapes within each line's local area\n") 
            f.write("4. Divide local content based on line orientation:\n")
            f.write("   • Horizontal lines → Above/Below sections\n")
            f.write("   • Vertical lines → Left/Right sections\n")
            f.write("   • Diagonal lines → Single area section\n")
            f.write("5. Within each section: Sort by ROWS first (80-120% height tolerance), then LEFT→RIGHT within rows\n")
            f.write("6. Add unprocessed shapes (not near any lines) at end\n\n")
            
            for slide_data in local_sectioning_data['slides']:
                f.write(f"{'='*15} SLIDE {slide_data['slide_number']} {'='*15}\n")
                f.write(f"Title: {slide_data['title']}\n")
                f.write(f"Line Dividers: {slide_data['line_dividers']}\n")
                f.write(f"Local Sections: {slide_data['local_sections']}\n\n")
                for section in slide_data['sectioning_details']['sections']:
                    f.write(f"📍 Section: {section['section_id']}\n")
                    f.write(f"   Type: {section['section_type']}\n")
                    
                    # Handle hierarchical structure - sections now have 'groups' instead of 'shapes'
                    if 'groups' in section:
                        total_components = sum(len(group.get('members', [])) for group in section['groups'])
                        f.write(f"   Groups: {len(section['groups'])}\n")
                        f.write(f"   Total Components: {total_components}\n")
                    elif 'shapes' in section:
                        f.write(f"   Shapes: {len(section['shapes'])}\n")
                    else:
                        f.write(f"   Components: Unknown structure\n")
                        
                    if 'divider_line' in section:
                        f.write(f"   Divider: {section['divider_line']}\n")
                    f.write(f"   Reading Order: {section['reading_order']}\n")
                    
                    # Show hierarchical content structure
                    if 'groups' in section and section['groups']:
                        f.write("   Hierarchical Content (Group → Components):\n")
                        for group_idx, group in enumerate(section['groups'][:2]):  # Show first 2 groups
                            f.write(f"      Group {group['group_id']}:\n")
                            members = group.get('members', [])
                            for comp_idx, component in enumerate(members[:3]):  # Show first 3 components
                                comp_text = component.get('text', '').strip()
                                text_preview = comp_text[:30] + '...' if len(comp_text) > 30 else comp_text
                                f.write(f"         {comp_idx+1}. {component['box_id']}: {text_preview}\n")
                            if len(members) > 3:
                                f.write(f"         ... and {len(members) - 3} more components\n")
                        if len(section['groups']) > 2:
                            f.write(f"      ... and {len(section['groups']) - 2} more groups\n")
                    elif 'shapes' in section and section['shapes']:
                        f.write("   Content (reading order):\n")
                        for i, shape in enumerate(section['shapes'][:3]):  # Show first 3
                            shape_text = shape.get('text', '').strip()
                            text_preview = shape_text[:40] + '...' if len(shape_text) > 40 else shape_text
                            f.write(f"      {i+1}. {shape['box_id']}: {text_preview}\n")
                        if len(section['shapes']) > 3:
                            f.write(f"      ... and {len(section['shapes']) - 3} more shapes\n")
                    f.write("\n")
        
        if total_lines > 0:
            print(f"      📋 Local sectioning analysis saved: {sectioning_path}")
            print(f"      📋 Local sectioning summary saved: {summary_path}")
        
        return sectioning_path



    def sort_shapes_with_local_line_awareness(self, content_shapes, line_shapes):
        """Sort shapes by normal reading order, but respect line boundaries within local areas"""
        if not line_shapes or not content_shapes:
            return self.sort_shapes_by_simple_reading_order(content_shapes)
        
        print(f"      🔍 Applying local line-aware reading for {len(line_shapes)} lines...")
        
        # Start with normal reading order
        initially_sorted = self.sort_shapes_by_simple_reading_order(content_shapes)
        
        # Apply line-aware adjustments within local areas
        for line in line_shapes:
            initially_sorted = self.apply_local_line_adjustment(initially_sorted, line)
        
        print(f"         ✅ Local line-aware reading complete: {len(initially_sorted)} shapes")
        return initially_sorted



    def apply_local_line_adjustment(self, sorted_shapes, line_shape):
        """Apply line-aware reading adjustment within the line's local area"""
        line_pos = line_shape['position']
        line_width = max(line_pos['width'], 1)
        line_height = max(line_pos['height'], 1)
        aspect_ratio = line_width / line_height
        
        # Only process clear vertical lines for now
        is_vertical = aspect_ratio < 0.7
        if not is_vertical:
            print(f"         📏 {line_shape['box_id']}: Not a clear vertical line (ratio={aspect_ratio:.2f}), skipping")
            return sorted_shapes
        
        line_center_x = line_pos['left'] + line_width / 2
        line_top = line_pos['top']
        line_bottom = line_pos['top'] + line_height
        
        # Define local area around the line (vertical expansion)
        local_area_expansion = line_height * 0.1  # 10% of line height
        local_top = max(0, line_top - local_area_expansion)
        local_bottom = line_bottom + local_area_expansion
        
        print(f"         📏 Processing vertical line {line_shape['box_id']} at x={line_center_x:.0f}")
        print(f"            Local area: y={local_top:.0f} to {local_bottom:.0f}")
        
        # Find shapes in the line's local area
        local_shapes = []
        non_local_shapes = []
        
        for shape in sorted_shapes:
            shape_center_y = shape['position']['top'] + shape['position']['height'] / 2
            
            # Check if shape is in the line's vertical range (local area)
            if local_top <= shape_center_y <= local_bottom:
                # Check if shape is reasonably close horizontally to the line
                shape_center_x = shape['position']['left'] + shape['position']['width'] / 2
                horizontal_distance = abs(shape_center_x - line_center_x)
                max_horizontal_distance = line_width * 20  # Allow shapes within 20x line width
                if horizontal_distance <= max_horizontal_distance:
                    local_shapes.append(shape)
                else:
                    non_local_shapes.append(shape)
            else:
                non_local_shapes.append(shape)
        
        if len(local_shapes) < 2:
            print(f"            ⚠️  Only {len(local_shapes)} shapes in local area, no adjustment needed")
            return sorted_shapes
        
        print(f"            📍 Found {len(local_shapes)} shapes in local area")
        
        # Separate local shapes by left/right of line
        left_local = []
        right_local = []
        
        for shape in local_shapes:
            shape_center_x = shape['position']['left'] + shape['position']['width'] / 2
            if shape_center_x < line_center_x:
                left_local.append(shape)
            else:
                right_local.append(shape)
        
        print(f"            📍 LEFT side: {len(left_local)} shapes")
        print(f"            📍 RIGHT side: {len(right_local)} shapes")
        
        # Sort each side by reading order independently
        left_sorted = self.sort_shapes_by_simple_reading_order(left_local)
        right_sorted = self.sort_shapes_by_simple_reading_order(right_local)
        
        # Reconstruct the full list maintaining overall order but with local line adjustment
        result = []
        local_reordered = left_sorted + right_sorted
        local_index = 0
        
        for shape in sorted_shapes:
            if shape in local_shapes:
                # Replace with reordered local shape
                if local_index < len(local_reordered):
                    result.append(local_reordered[local_index])
                    local_index += 1
            else:
                # Keep non-local shape in its original position
                result.append(shape)
        
        print(f"            ✅ Applied local line adjustment: LEFT({len(left_local)}) → RIGHT({len(right_local)})")
        return result



    def create_sections_from_local_line_aware_order(self, ordered_shapes, line_shapes):
        """Create a single section representing the local line-aware reading order"""
        return [{
            'section_id': 'local_line_aware',
            'section_type': 'local_line_aware_reading',
            'bounds': calculate_shapes_bounds(ordered_shapes),
            'shapes': ordered_shapes,
            'reading_order': 1,
            'line_dividers_processed': len(line_shapes),
            'note': 'Normal reading order with local line-aware adjustments'
        }]



    def reorder_overlapping_text_and_images(self, sorted_shapes):
        """Reorder overlapping text and image elements so text comes first"""
        if len(sorted_shapes) < 2:
            return sorted_shapes
        
        reordered = list(sorted_shapes)
        swaps_made = 0
        
        # Check each pair of adjacent elements for overlaps
        i = 0
        while i < len(reordered) - 1:
            current = reordered[i]
            next_shape = reordered[i + 1]
            
            # Check if we have a text/image pair
            current_type = current.get('shape_type', 'unknown')
            next_type = next_shape.get('shape_type', 'unknown')
            
            is_text_image_pair = (
                (current_type in ['Picture', 'Group', 'Chart'] and next_type in ['TextBox']) or
                (current_type in ['TextBox'] and next_type in ['Picture', 'Group', 'Chart'])
            )
            
            if is_text_image_pair:
                # Check if they spatially overlap
                overlap_pct = calculate_spatial_overlap(current['position'], next_shape['position'])
                if overlap_pct > 30.0:  # 30% overlap threshold
                    # If image comes before text, swap them
                    if current_type in ['Picture', 'Group', 'Chart'] and next_type in ['TextBox']:
                        reordered[i], reordered[i + 1] = reordered[i + 1], reordered[i]
                        swaps_made += 1
                        print(f"         🔄 Swapped overlapping {current['box_id']} ({current_type}) and {next_shape['box_id']} ({next_type}) - text now first (overlap: {overlap_pct:.1f}%)")
                        # Don't increment i, check this position again
                        continue
            
            i += 1
        
        if swaps_made > 0:
            print(f"         📝 Made {swaps_made} swaps to prioritize text over overlapping images")
        
        return reordered


