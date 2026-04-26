"""
Geometry utility functions for spatial calculations.
"""

def calculate_overlap_percentage(box1, box2):
    """Calculate overlap percentage"""
    def normalize_coords(pos):
        # DISABLED: Keep EMU coordinates for accurate overlap calculation
        return pos  # No conversion - use original EMU coordinates
        # OLD CODE BELOW (disabled):
        if isinstance(pos.get('left', 0), int) and pos.get('left', 0) > 10000:
            return {
                'left': pos['left'] / 12700,
                'top': pos['top'] / 12700,
                'width': pos['width'] / 12700,
                'height': pos['height'] / 12700
            }
        return pos
    
    pos1 = normalize_coords(box1['position'])
    pos2 = normalize_coords(box2['position'])
    
    # Calculate boundaries
    x1_left, y1_top = pos1['left'], pos1['top']
    x1_right, y1_bottom = x1_left + pos1['width'], y1_top + pos1['height']
    
    x2_left, y2_top = pos2['left'], pos2['top']
    x2_right, y2_bottom = x2_left + pos2['width'], y2_top + pos2['height']
    
    # Calculate intersection
    intersect_left = max(x1_left, x2_left)
    intersect_top = max(y1_top, y2_top)
    intersect_right = min(x1_right, x2_right)
    intersect_bottom = min(y1_bottom, y2_bottom)
    
    if intersect_right <= intersect_left or intersect_bottom <= intersect_top:
        return 0.0
    
    intersect_area = (intersect_right - intersect_left) * (intersect_bottom - intersect_top)
    box1_area = pos1['width'] * pos1['height']
    
    if box1_area == 0:
        return 0.0
    
    return min(intersect_area / box1_area, 1.0)


def calculate_spatial_overlap(pos1, pos2):
    """Calculate percentage overlap between two rectangular positions - SIMPLE VERSION like working code"""
    try:
        # Calculate overlap area
        left = max(pos1['left'], pos2['left'])
        top = max(pos1['top'], pos2['top'])
        right = min(pos1['left'] + pos1['width'], pos2['left'] + pos2['width'])
        bottom = min(pos1['top'] + pos1['height'], pos2['top'] + pos2['height'])
        
        if left < right and top < bottom:
            overlap_area = (right - left) * (bottom - top)
            area1 = pos1['width'] * pos1['height']
            area2 = pos2['width'] * pos2['height']
            min_area = min(area1, area2)
            if min_area > 0:
                # Return as percentage of smaller box (like working code)
                return (overlap_area / min_area) * 100.0
        
        return 0.0
        
    except Exception as e:
        print(f"⚠️  Error calculating spatial overlap: {e}")
        return 0.0


def calculate_proximity(pos1, pos2):
    """Calculate minimum distance between two rectangles in EMUs"""
    try:
        # Convert to right/bottom coordinates
        left1, top1 = pos1['left'], pos1['top']
        right1, bottom1 = left1 + pos1['width'], top1 + pos1['height']
        
        left2, top2 = pos2['left'], pos2['top']
        right2, bottom2 = left2 + pos2['width'], top2 + pos2['height']
        
        # Calculate horizontal and vertical distances
        horizontal_distance = 0
        if right1 < left2:  # pos1 is to the left of pos2
            horizontal_distance = left2 - right1
        elif right2 < left1:  # pos2 is to the left of pos1
            horizontal_distance = left1 - right2
        # else: rectangles overlap horizontally, distance = 0
        
        vertical_distance = 0
        if bottom1 < top2:  # pos1 is above pos2
            vertical_distance = top2 - bottom1
        elif bottom2 < top1:  # pos2 is above pos1
            vertical_distance = top1 - bottom2
        # else: rectangles overlap vertically, distance = 0
        
        # Return Euclidean distance
        return (horizontal_distance ** 2 + vertical_distance ** 2) ** 0.5
        
    except Exception as e:
        print(f"⚠️  Error calculating proximity: {e}")
        return float('inf')
    except Exception as e:
        print(f"⚠️  Error calculating spatial overlap: {e}")
        return 0.0


def calculate_group_center(components):
    """Calculate the center position of a group of components"""
    try:
        if not components:
            return None
        
        total_x = 0
        total_y = 0
        valid_positions = 0
        
        for component in components:
            position = component.get('position', {})
            if position and 'left' in position and 'top' in position:
                # Calculate center of this component
                left = position.get('left', 0)
                top = position.get('top', 0)
                width = position.get('width', 0)
                height = position.get('height', 0)
                
                center_x = left + width / 2
                center_y = top + height / 2
                
                total_x += center_x
                total_y += center_y
                valid_positions += 1
        
        if valid_positions > 0:
            return {
                'x': total_x / valid_positions,
                'y': total_y / valid_positions
            }
        
        return None
        
    except Exception as e:
        print(f"⚠️ Error calculating group center: {e}")
        return None


def calculate_group_bounds(components):
    """Calculate the full bounding box (top, bottom, left, right) of a group of components"""
    try:
        if not components:
            return None
        
        min_left = float('inf')
        min_top = float('inf')
        max_right = float('-inf')
        max_bottom = float('-inf')
        valid_positions = 0
        
        for component in components:
            position = component.get('position', {})
            if position and 'left' in position and 'top' in position:
                left = position.get('left', 0)
                top = position.get('top', 0)
                width = position.get('width', 0)
                height = position.get('height', 0)
                
                right = left + width
                bottom = top + height
                
                min_left = min(min_left, left)
                min_top = min(min_top, top)
                max_right = max(max_right, right)
                max_bottom = max(max_bottom, bottom)
                valid_positions += 1
        
        if valid_positions > 0:
            return {
                'left': min_left,
                'top': min_top,
                'right': max_right,
                'bottom': max_bottom,
                'width': max_right - min_left,
                'height': max_bottom - min_top
            }
        
        return None
        
    except Exception as e:
        print(f"⚠️  Error calculating group bounds: {e}")
        return None


def calculate_shapes_bounds(shapes):
    """Calculate the bounding box of a group of shapes"""
    if not shapes:
        return {'left': 0, 'top': 0, 'right': 0, 'bottom': 0}
    
    min_left = min(s['position']['left'] for s in shapes)
    min_top = min(s['position']['top'] for s in shapes)
    max_right = max(s['position']['left'] + s['position']['width'] for s in shapes)
    max_bottom = max(s['position']['top'] + s['position']['height'] for s in shapes)
    
    return {
        'left': min_left,
        'top': min_top,
        'right': max_right,
        'bottom': max_bottom
    }


def calculate_merged_boundary(positions):
    """Calculate the merged bounding box from multiple positions"""
    if not positions:
        return {'left': 0, 'top': 0, 'width': 0, 'height': 0}
    
    print(f"      🔍 DEBUG - Calculating merged boundary from {len(positions)} positions:")
    for i, pos in enumerate(positions):
        left, top, width, height = pos['left'], pos['top'], pos['width'], pos['height']
        right, bottom = left + width, top + height
        print(f"         Position {i+1}: left={left}, top={top}, right={right}, bottom={bottom}")
    
    # Find overall boundaries (convert width/height to right/bottom)
    min_left = min(pos['left'] for pos in positions)
    min_top = min(pos['top'] for pos in positions)  
    max_right = max(pos['left'] + pos['width'] for pos in positions)
    max_bottom = max(pos['top'] + pos['height'] for pos in positions)
    
    print(f"         📐 Merged: left={min_left}, top={min_top}, right={max_right}, bottom={max_bottom}")
    
    # Return in the same format as input positions
    merged = {
        'left': min_left,
        'top': min_top,
        'width': max_right - min_left,
        'height': max_bottom - min_top
    }
    print(f"         ✅ Final merged boundary: {merged}")
    return merged


def calculate_text_box_overlap(text_pos, box_pos):
    """Calculate overlap between text and box positions"""
    # Normalize coordinates
    def norm(pos):
        # DISABLED: Keep EMU coordinates for consistency
        return pos  # No conversion
        # OLD CODE BELOW (disabled):
        return {
            'left': pos['left'] / 12700 if pos['left'] > 10000 else pos['left'],
            'top': pos['top'] / 12700 if pos['top'] > 10000 else pos['top'],
            'width': pos['width'] / 12700 if pos['width'] > 10000 else pos['width'],
            'height': pos['height'] / 12700 if pos['height'] > 10000 else pos['height']
        }
    
    text_norm = norm(text_pos)
    box_norm = norm(box_pos)
    
    # Calculate overlap
    x_overlap = max(0, min(text_norm['left'] + text_norm['width'], box_norm['left'] + box_norm['width']) - 
                     max(text_norm['left'], box_norm['left']))
    y_overlap = max(0, min(text_norm['top'] + text_norm['height'], box_norm['top'] + box_norm['height']) - 
                     max(text_norm['top'], box_norm['top']))
    
    overlap_area = x_overlap * y_overlap
    text_area = text_norm['width'] * text_norm['height']
    
    return overlap_area / text_area if text_area > 0 else 0.0


def get_group_center(group):
    """Calculate the center of a group from all its members"""
    if not group['members']:
        return None
        
    all_positions = []
    for member in group['members']:
        pos = member['position']
        all_positions.append({
            'left': pos['left'],
            'top': pos['top'],
            'right': pos['left'] + pos['width'],
            'bottom': pos['top'] + pos['height']
        })
    
    min_left = min(p['left'] for p in all_positions)
    min_top = min(p['top'] for p in all_positions)
    max_right = max(p['right'] for p in all_positions)
    max_bottom = max(p['bottom'] for p in all_positions)
    
    center_x = (min_left + max_right) // 2
    center_y = (min_top + max_bottom) // 2
    
    return (center_x, center_y)


def simple_boxes_overlap(pos1, pos2, overlap_threshold=0.1):
    """Check if two boxes overlap with a minimum threshold (simple script approach)"""
    # Calculate overlap area
    left = max(pos1['left'], pos2['left'])
    top = max(pos1['top'], pos2['top'])
    right = min(pos1['left'] + pos1['width'], pos2['left'] + pos2['width'])
    bottom = min(pos1['top'] + pos1['height'], pos2['top'] + pos2['height'])
    
    if left < right and top < bottom:
        overlap_area = (right - left) * (bottom - top)
        area1 = pos1['width'] * pos1['height']
        area2 = pos2['width'] * pos2['height']
        min_area = min(area1, area2)
        
        # Check if overlap is significant (at least 10% of smaller box)
        return overlap_area >= (min_area * overlap_threshold)
    
    return False


def calculate_spatial_containment(component_pos, table_pos):
    """Calculate what percentage of component is contained within table bounds"""
    try:
        # Component bounds
        comp_left = component_pos['left']
        comp_top = component_pos['top']
        comp_right = comp_left + component_pos['width']
        comp_bottom = comp_top + component_pos['height']
        
        # Table bounds
        table_left = table_pos['left']
        table_top = table_pos['top']
        table_right = table_left + table_pos['width']
        table_bottom = table_top + table_pos['height']
        
        # Calculate intersection
        intersect_left = max(comp_left, table_left)
        intersect_top = max(comp_top, table_top)
        intersect_right = min(comp_right, table_right)
        intersect_bottom = min(comp_bottom, table_bottom)
        
        # Check if there's any intersection
        if intersect_left >= intersect_right or intersect_top >= intersect_bottom:
            return 0.0
        
        # Calculate areas
        intersect_area = (intersect_right - intersect_left) * (intersect_bottom - intersect_top)
        component_area = component_pos['width'] * component_pos['height']
        
        if component_area == 0:
            return 0.0
        
        containment_percentage = (intersect_area / component_area) * 100.0
        return containment_percentage
        
    except Exception as e:
        print(f"⚠️  Error calculating spatial containment: {e}")
        return 0.0


def calculate_combined_boundary(positions):
    """Calculate the combined bounding box for multiple positions"""
    if not positions:
        return {}
    
    # Find the encompassing boundary
    min_left = min(pos['left'] for pos in positions)
    min_top = min(pos['top'] for pos in positions)
    max_right = max(pos['left'] + pos['width'] for pos in positions)
    max_bottom = max(pos['top'] + pos['height'] for pos in positions)
    
    return {
        'left': min_left,
        'top': min_top,
        'width': max_right - min_left,
        'height': max_bottom - min_top
    }


