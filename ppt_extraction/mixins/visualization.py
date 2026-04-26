"""
Matplotlib-based visualization for groups, reading order, and sections.
"""

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from pathlib import Path

from utils.geometry import (
    calculate_shapes_bounds,
)



class VisualizationMixin:
    """Methods for generating matplotlib visualizations."""

    def create_enhanced_group_visualization(self, slide_analysis):
        """Create enhanced visualization showing smart containment groups with detailed mapping"""
        slide_num = slide_analysis['slide_number']
        smart_groups = slide_analysis['smart_groups']
        
        print(f"      🎨 Creating enhanced group visualization for slide {slide_num}...")
        
        # Create figure with dual visualization
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(28, 14))
        fig.suptitle(f'Slide {slide_num} - Enhanced Smart Containment Groups', fontsize=18, fontweight='bold')
        
        # Get slide dimensions
        if self.comprehensive_data and 'metadata' in self.comprehensive_data:
            if 'slide_width' in self.comprehensive_data['metadata']:
                slide_width = self.comprehensive_data['metadata']['slide_width'] / 12700
                slide_height = self.comprehensive_data['metadata']['slide_height'] / 12700
            else:
                slide_width, slide_height = 720, 540
        else:
            slide_width, slide_height = 720, 540
        
        # Left plot: Component → Group Assignments with content
        ax1.set_title('Component → Group Assignments (with Content)', fontsize=14, fontweight='bold')
        
        # Right plot: Smart groups with enhanced details
        ax2.set_title('Smart Groups with Component Details', fontsize=14, fontweight='bold')
        
        # Draw slide boundaries
        for ax, title in [(ax1, 'Assignment View'), (ax2, 'Group Structure View')]:
            slide_rect = patches.Rectangle((0, 0), slide_width, slide_height,
                                         linewidth=3, edgecolor='black',
                                         facecolor='lightgray', alpha=0.1)
            ax.add_patch(slide_rect)
            
            # Add view title
            ax.text(slide_width/2, -30, title, ha='center', va='center', 
                   fontsize=12, fontweight='bold', color='darkblue')
        
        # Generate distinct colors for each group
        group_colors = plt.cm.Set3(np.linspace(0, 1, len(smart_groups)))
        group_color_map = {}
        for i, group_name in enumerate(smart_groups.keys()):
            group_color_map[group_name] = group_colors[i]
        
        # Create group name mapping for display
        group_display_names = {}
        for group_name in smart_groups.keys():
            original_id = smart_groups[group_name].get('original_box_id', group_name)
            group_display_names[group_name] = f"{group_name} ({original_id})"
        
        # Draw groups and assignments with enhanced details
        group_legend_info = []
        
        for group_name, group_data in smart_groups.items():
            group_color = group_color_map[group_name]
            
            # Draw root component
            root_box = group_data['root_component']
            self._draw_enhanced_box_with_content(ax1, root_box, group_color, is_root=True, group_name=group_name)
            self._draw_enhanced_box_with_content(ax2, root_box, group_color, is_root=True, group_name=group_name)
            
            # Collect group info for legend
            member_count = len(group_data['members'])
            total_members = group_data['total_members']
            member_types = group_data.get('member_types', {})
            type_summary = ", ".join([f"{count} {type_name}" for type_name, count in member_types.items()])
            
            group_legend_info.append({
                'name': group_name,
                'color': group_color,
                'members': member_count,
                'total': total_members,
                'types': type_summary
            })
            
            # Draw members and assignment arrows
            for member in group_data['members']:
                member_box = member['box']
                containment_pct = member['containment_percentage']
                assignment_type = member['assignment_type']
                # Draw member boxes with content info
                self._draw_enhanced_box_with_content(ax1, member_box, group_color, is_root=False, 
                                                   group_name=group_name, containment_pct=containment_pct)
                self._draw_enhanced_box_with_content(ax2, member_box, group_color, is_root=False, 
                                                   group_name=group_name, containment_pct=containment_pct)
                # Draw assignment arrow on left plot
                self._draw_enhanced_assignment_arrow(ax1, member_box, root_box, containment_pct, assignment_type)
        
        # Add comprehensive legend
        self._add_enhanced_group_legend(fig, group_legend_info, slide_width, slide_height)
        
        # Set axis properties
        for ax in [ax1, ax2]:
            ax.set_xlim(-80, slide_width + 80)
            ax.set_ylim(-80, slide_height + 80)
            ax.invert_yaxis()
            ax.set_aspect('equal')
            ax.grid(True, alpha=0.3)
            ax.set_xlabel('Width (points)', fontsize=11)
            ax.set_ylabel('Height (points)', fontsize=11)
        
        plt.tight_layout()
        
        # Save enhanced visualization
        viz_path = self.smart_groups_dir / f"slide_{slide_num:02d}_enhanced_smart_groups.png"
        plt.savefig(viz_path, dpi=100, facecolor='white')
        print(f"         💾 Saved enhanced group visualization: {viz_path}")
        plt.close('all')
        
        return viz_path



    def _draw_enhanced_box_with_content(self, ax, box, group_color, is_root=False, group_name="", containment_pct=None):
        """Draw a box with enhanced group-based coloring and content information"""
        pos = box['position']
        
        # Normalize position
        if pos['left'] > 10000:  # EMU units
            left, top = pos['left'] / 12700, pos['top'] / 12700
            width, height = pos['width'] / 12700, pos['height'] / 12700
        else:
            left, top, width, height = pos['left'], pos['top'], pos['width'], pos['height']
        
        line_width = 5 if is_root else 3
        alpha = 0.8 if is_root else 0.5
        
        # Enhanced styling based on containment quality
        if containment_pct is not None:
            if containment_pct >= 0.95:
                line_style = '-'  # Solid for full containment
            elif containment_pct >= 0.7:
                line_style = '--'  # Dashed for partial containment
            else:
                line_style = ':'  # Dotted for weak containment
        else:
            line_style = '-'
        
        # Draw box with enhanced styling
        rect = patches.Rectangle((left, top), width, height,
                               linewidth=line_width, edgecolor=group_color,
                               facecolor=group_color, alpha=alpha, linestyle=line_style)
        ax.add_patch(rect)
        
        # Add box ID with enhanced styling
        font_size = 14 if is_root else 11
        font_weight = 'bold' if is_root else 'normal'
        
        ax.text(left + 8, top + 8, box['box_id'],
                fontsize=font_size, weight=font_weight, color='black',
                bbox=dict(boxstyle="round,pad=0.4", facecolor='white', alpha=0.95, edgecolor=group_color))
        
        # Add content preview for text boxes
        if box.get('has_text', False) and box.get('text', '').strip():
            text_content = box['text'].strip()
            text_preview = text_content[:25] + "..." if len(text_content) > 25 else text_content
            text_preview = text_preview.replace('\n', ' ').replace('\r', ' ')
            
            # Position text preview
            text_y = top + height - 15 if height > 30 else top + height + 5
            ax.text(left + 8, text_y, f'"{text_preview}"',
                   fontsize=9, style='italic', color='darkblue',
                   bbox=dict(boxstyle="round,pad=0.3", facecolor='lightyellow', alpha=0.9))
        
        # Add group label for root components
        if is_root:
            ax.text(left + width/2, top - 20, f"🏠 GROUP {group_name}",
                   ha='center', va='bottom', fontsize=13, weight='bold', color='darkblue',
                   bbox=dict(boxstyle="round,pad=0.4", facecolor=group_color, alpha=0.9))
        
        # Add containment percentage for members
        elif containment_pct is not None:
            ax.text(left + width - 8, top + 8, f"{containment_pct:.0%}",
                   ha='right', va='top', fontsize=10, weight='bold', color='white',
                   bbox=dict(boxstyle="round,pad=0.3", facecolor=group_color, alpha=0.9))



    def _draw_enhanced_assignment_arrow(self, ax, member_box, root_box, containment_pct, assignment_type):
        """Draw enhanced arrow showing assignment relationship with quality indicators"""
        def get_center(box):
            pos = box['position']
            if pos['left'] > 10000:
                left, top = pos['left'] / 12700, pos['top'] / 12700
                width, height = pos['width'] / 12700, pos['height'] / 12700
            else:
                left, top, width, height = pos['left'], pos['top'], pos['width'], pos['height']
            return left + width/2, top + height/2
        
        member_x, member_y = get_center(member_box)
        root_x, root_y = get_center(root_box)
        
        # Enhanced arrow styling based on containment quality
        if containment_pct >= 0.95:
            arrow_color = 'green'
            arrow_style = '->'
            line_width = 4
        elif containment_pct >= 0.7:
            arrow_color = 'orange'
            arrow_style = '->'
            line_width = 3
        else:
            arrow_color = 'red'
            arrow_style = '->'
            line_width = 2
        
        # Draw enhanced arrow
        ax.annotate('', xy=(root_x, root_y), xytext=(member_x, member_y),
                    arrowprops=dict(arrowstyle=arrow_style, color=arrow_color, 
                                  lw=line_width, alpha=0.8))
        
        # Add containment percentage with assignment type
        mid_x, mid_y = (member_x + root_x) / 2, (member_y + root_y) / 2
        assignment_icon = "🔗" if assignment_type == "full" else "🔸"
        ax.text(mid_x, mid_y, f"{assignment_icon}{containment_pct:.0%}",
                fontsize=10, ha='center', va='center', weight='bold',
                bbox=dict(boxstyle="round,pad=0.3", facecolor='white', alpha=0.95, 
                         edgecolor=arrow_color, linewidth=2))



    def _add_enhanced_group_legend(self, fig, group_legend_info, slide_width, slide_height):
        """Add comprehensive legend with group information"""
        legend_text = "GROUP LEGEND:\n"
        legend_text += "=" * 30 + "\n"
        
        for group_info in group_legend_info:
            legend_text += f"🏠 {group_info['name']}: {group_info['members']} members"
            if group_info['types']:
                legend_text += f" ({group_info['types']})"
            legend_text += "\n"
        
        legend_text += "\nASSIGNMENT QUALITY:\n"
        legend_text += "🔗 Full (95%+) | 🔸 Partial (50-95%)"
        
        # Add legend to figure
        fig.text(0.02, 0.02, legend_text, fontsize=10, 
                bbox=dict(boxstyle="round,pad=0.5", facecolor='lightcyan', alpha=0.9),
                verticalalignment='bottom')



    def create_reading_order_flow_visualization(self, all_slide_analyses):
        """Create hierarchical visualization showing group-level and component-level reading order"""
        print(f"   🎨 Creating hierarchical reading order flow visualizations...")
        
        for analysis in all_slide_analyses:
            slide_num = analysis['slide_number']
            
            # Get hierarchical groups from spatial analysis for this slide
            hierarchical_groups = analysis.get('hierarchical_groups', [])
            line_dividers = analysis.get('line_dividers', [])
            
            if not hierarchical_groups:
                print(f"      ⚠️  No hierarchical groups found for slide {slide_num}, skipping")
                continue
            
            print(f"      🎨 Creating hierarchical flow visualization for slide {slide_num}")
            
            # Create visualization showing both group and component flow
            fig, ax = plt.subplots(1, 1, figsize=(24, 16))
            fig.suptitle(f'Slide {slide_num} - Sectional Boundary Reading Order Flow\n(Sections created by horizontal & vertical lines | Thick arrows: Section flow | Thin arrows: Component flow)', 
                        fontsize=18, fontweight='bold')
            
            # Get slide dimensions
            if self.comprehensive_data and 'metadata' in self.comprehensive_data:
                slide_width = self.comprehensive_data['metadata'].get('slide_width', 9144000) / 12700
                slide_height = self.comprehensive_data['metadata'].get('slide_height', 6858000) / 12700
            else:
                slide_width, slide_height = 720, 540
            
            ax.set_title('Sectional Reading Order: Sections → Groups → Components', fontsize=16, fontweight='bold')
            
            # Draw slide boundary
            slide_rect = patches.Rectangle((0, 0), slide_width, slide_height,
                                         linewidth=3, edgecolor='black',
                                         facecolor='lightgray', alpha=0.1)
            ax.add_patch(slide_rect)
            
            # Draw line dividers first (show boundaries that create sections)
            for line in line_dividers:
                pos = line['position']
                left = pos['left'] / 12700 if pos['left'] > 10000 else pos['left']
                top = pos['top'] / 12700 if pos['top'] > 10000 else pos['top']
                width = pos['width'] / 12700 if pos['width'] > 10000 else pos['width']
                height = pos['height'] / 12700 if pos['height'] > 10000 else pos['height']
                # Determine line type for better visualization
                aspect_ratio = max(width, 1) / max(height, 1)
                if aspect_ratio > 2.0:
                    line_color = 'red'
                    line_type = "H-BOUNDARY"
                elif aspect_ratio < 0.5:
                    line_color = 'blue'
                    line_type = "V-BOUNDARY"
                else:
                    line_color = 'orange'
                    line_type = "DIAGONAL"
                # Enhanced line visualization for zero-height/zero-width lines
                if aspect_ratio > 2.0:  # Horizontal line
                    # Draw horizontal line as actual line (not rectangle)
                    line_y = top + height/2 if height > 0 else top
                    ax.plot([left, left + width], [line_y, line_y], 
                           color=line_color, linewidth=10, alpha=0.9, solid_capstyle='round')
                    
                    # Add visual thickness with rectangle
                    visual_height = max(height, 5)  # Minimum 5pt visual height
                    line_rect = patches.Rectangle((left, top), width, visual_height,
                                                linewidth=2, edgecolor=line_color,
                                                facecolor=line_color, alpha=0.4)
                elif aspect_ratio < 0.5:  # Vertical line
                    # Draw vertical line as actual line (not rectangle)
                    line_x = left + width/2 if width > 0 else left
                    ax.plot([line_x, line_x], [top, top + height], 
                           color=line_color, linewidth=10, alpha=0.9, solid_capstyle='round')
                    
                    # Add visual thickness with rectangle
                    visual_width = max(width, 5)  # Minimum 5pt visual width
                    line_rect = patches.Rectangle((left, top), visual_width, height,
                                                linewidth=2, edgecolor=line_color,
                                                facecolor=line_color, alpha=0.4)
                else:  # Diagonal line
                    line_rect = patches.Rectangle((left, top), width, height,
                                                linewidth=4, edgecolor=line_color,
                                                facecolor=line_color, alpha=0.7)
                ax.add_patch(line_rect)
                ax.text(left + width/2, top - 25, f"📏 {line_type} {line['box_id']}",
                       ha='center', va='bottom', fontsize=12, weight='bold', color=line_color,
                       bbox=dict(boxstyle="round,pad=0.5", facecolor='white', alpha=0.9))
            
            # Get section information if available
            spatial_sections = analysis.get('local_sections', [])
            section_colors = plt.cm.Set3(np.linspace(0, 1, max(len(spatial_sections), 1)))
            
            # Draw section boundaries if available
            section_centers = []
            for section_idx, section in enumerate(spatial_sections):
                if 'bounds' in section and section['bounds']:
                    bounds = section['bounds']
                    section_color = section_colors[section_idx % len(section_colors)]
                    
                    # Convert bounds to screen coordinates
                    sect_left = bounds.get('x_min', 0) / 12700 if bounds.get('x_min', 0) > 10000 else bounds.get('x_min', 0)
                    sect_top = bounds.get('y_min', 0) / 12700 if bounds.get('y_min', 0) > 10000 else bounds.get('y_min', 0)
                    sect_right = bounds.get('x_max', slide_width) / 12700 if bounds.get('x_max', slide_width) > 10000 else bounds.get('x_max', slide_width)
                    sect_bottom = bounds.get('y_max', slide_height) / 12700 if bounds.get('y_max', slide_height) > 10000 else bounds.get('y_max', slide_height)
                    
                    sect_width = sect_right - sect_left
                    sect_height = sect_bottom - sect_top
                    
                    section_center = (sect_left + sect_width/2, sect_top + sect_height/2)
                    section_centers.append(section_center)
                    
                    # Draw section boundary
                    section_rect = patches.Rectangle((sect_left, sect_top), sect_width, sect_height,
                                                   linewidth=4, edgecolor=section_color,
                                                   facecolor=section_color, alpha=0.1)
                    ax.add_patch(section_rect)
                    
                    # Section label
                    section_label = f"SECTION {section.get('section_id', section_idx+1)}"
                    ax.text(sect_left + 5, sect_top + 5, section_label,
                           ha='left', va='top', fontsize=14, weight='bold', color=section_color,
                           bbox=dict(boxstyle="round,pad=0.5", facecolor='white', edgecolor=section_color, linewidth=2))
            
            # Generate colors for groups
            group_colors = plt.cm.Set1(np.linspace(0, 1, max(len(hierarchical_groups), 1)))
            
            # LEVEL 1: Draw groups with boundaries and group flow arrows
            group_centers = []
            for group_idx, group in enumerate(hierarchical_groups):
                group_color = group_colors[group_idx % len(group_colors)]
                members = group['members']
                if not members:
                    continue
                # Calculate group bounds
                group_bounds = calculate_shapes_bounds(members)
                left = group_bounds['left'] / 12700 if group_bounds['left'] > 10000 else group_bounds['left']
                top = group_bounds['top'] / 12700 if group_bounds['top'] > 10000 else group_bounds['top']
                right = group_bounds['right'] / 12700 if group_bounds['right'] > 10000 else group_bounds['right']
                bottom = group_bounds['bottom'] / 12700 if group_bounds['bottom'] > 10000 else group_bounds['bottom']
                width = right - left
                height = bottom - top
                group_center = (left + width/2, top + height/2)
                group_centers.append(group_center)
                # Draw group boundary rectangle
                group_rect = patches.Rectangle((left - 10, top - 10), width + 20, height + 20,
                                             linewidth=3, edgecolor=group_color,
                                             facecolor=group_color, alpha=0.15)
                ax.add_patch(group_rect)
                # Group label
                group_label = f"GROUP {group['group_id']}"
                ax.text(left - 5, top - 25, group_label,
                       ha='left', va='bottom', fontsize=14, weight='bold', color=group_color,
                       bbox=dict(boxstyle="round,pad=0.5", facecolor='white', edgecolor=group_color, linewidth=2))
                # LEVEL 2: Draw components within group and component flow arrows
                component_centers = []
                for member_idx, member in enumerate(members):
                    pos = member['position']
                    comp_left = pos['left'] / 12700 if pos['left'] > 10000 else pos['left']
                    comp_top = pos['top'] / 12700 if pos['top'] > 10000 else pos['top']
                    comp_width = pos['width'] / 12700 if pos['width'] > 10000 else pos['width']
                    comp_height = pos['height'] / 12700 if pos['height'] > 10000 else pos['height']
                    
                    comp_center = (comp_left + comp_width/2, comp_top + comp_height/2)
                    component_centers.append(comp_center)
                    
                    # Draw component rectangle
                    comp_rect = patches.Rectangle((comp_left, comp_top), comp_width, comp_height,
                                                linewidth=2, edgecolor=group_color,
                                                facecolor=group_color, alpha=0.3)
                    ax.add_patch(comp_rect)
                    
                    # Component label with hierarchy info
                    hierarchy_info = member.get('hierarchical_info', {})
                    comp_order = hierarchy_info.get('component_order_in_group', member_idx + 1)
                    comp_label = f"{member['box_id']}\n#{comp_order}"
                    
                    ax.text(comp_center[0], comp_center[1], comp_label,
                           ha='center', va='center', fontsize=10, weight='bold',
                           bbox=dict(boxstyle="circle,pad=0.3", facecolor='white', edgecolor=group_color, linewidth=2))
                # Draw THIN arrows for component flow within group
                for i in range(len(component_centers) - 1):
                    from_center = component_centers[i]
                    to_center = component_centers[i + 1]
                    
                    # Thin arrow for component flow
                    ax.annotate('', xy=to_center, xytext=from_center,
                              arrowprops=dict(arrowstyle='->', lw=2, color=group_color, alpha=0.7,
                                            connectionstyle="arc3,rad=0.1"))
            
            # Draw THICK arrows for group flow (sectional order)
            for i in range(len(group_centers) - 1):
                from_group_center = group_centers[i]
                to_group_center = group_centers[i + 1]
                # Thick arrow for group flow
                ax.annotate('', xy=to_group_center, xytext=from_group_center,
                          arrowprops=dict(arrowstyle='->', lw=6, color='darkblue', alpha=0.8,
                                        connectionstyle="arc3,rad=0.2"))
                # Group flow step number
                mid_x = (from_group_center[0] + to_group_center[0]) / 2
                mid_y = (from_group_center[1] + to_group_center[1]) / 2 + 30
                ax.text(mid_x, mid_y, f"G{i+1}→G{i+2}",
                       ha='center', va='center', fontsize=12, weight='bold', color='darkblue',
                       bbox=dict(boxstyle="round,pad=0.5", facecolor='lightblue', alpha=0.9))
            
            # Draw SUPER THICK arrows for section flow if available
            for i in range(len(section_centers) - 1):
                from_section_center = section_centers[i]
                to_section_center = section_centers[i + 1]
                # Super thick arrow for section flow
                ax.annotate('', xy=to_section_center, xytext=from_section_center,
                          arrowprops=dict(arrowstyle='->', lw=10, color='darkred', alpha=0.9,
                                        connectionstyle="arc3,rad=0.3"))
                # Section flow step number
                mid_x = (from_section_center[0] + to_section_center[0]) / 2
                mid_y = (from_section_center[1] + to_section_center[1]) / 2 + 50
                ax.text(mid_x, mid_y, f"SECT{i+1}→SECT{i+2}",
                       ha='center', va='center', fontsize=14, weight='bold', color='darkred',
                       bbox=dict(boxstyle="round,pad=0.7", facecolor='lightyellow', alpha=0.95))
            
            # Add enhanced legend and summary
            legend_text = f"📊 SECTIONAL BOUNDARY READING ORDER\n"
            legend_text += f"🔴 {len(section_centers)} Sections (super thick red arrows)\n"
            legend_text += f"🔵 {len(hierarchical_groups)} Groups (thick blue arrows)\n"
            total_components = sum(len(g['members']) for g in hierarchical_groups)
            legend_text += f"🔸 {total_components} Components (thin colored arrows)\n"
            legend_text += f"📏 {len(line_dividers)} Line Boundaries\n\n"
            legend_text += f"📖 SECTIONAL READING FLOW:\n"
            legend_text += f"1. Lines create SECTION boundaries\n"
            legend_text += f"2. Follow RED arrows for section order\n"
            legend_text += f"3. Within sections, follow BLUE arrows for groups\n"
            legend_text += f"4. Within groups, follow COLORED arrows for components"
            
            ax.text(0.02, 0.98, legend_text, transform=ax.transAxes, fontsize=11,
                   verticalalignment='top', bbox=dict(boxstyle="round,pad=0.8", facecolor='lightyellow', alpha=0.95))
            
            # Set axis properties
            ax.set_xlim(-50, slide_width + 50)
            ax.set_ylim(-50, slide_height + 50)
            ax.invert_yaxis()  # PowerPoint coordinates
            ax.set_aspect('equal')
            ax.grid(True, alpha=0.3)
            ax.set_xlabel('X Position', fontsize=12)
            ax.set_ylabel('Y Position', fontsize=12)
            
            plt.tight_layout()
            
            # Save visualization
            output_file = self.reading_order_groups_dir / f"slide_{slide_num:02d}_hierarchical_flow.png"
            try:
                plt.savefig(output_file, dpi=100)
                plt.close('all')
                print(f"         💾 Hierarchical flow saved: {output_file.absolute()}")
                # Verify file was actually created
                if output_file.exists():
                    file_size = output_file.stat().st_size
                    print(f"         ✅ File verified: {file_size} bytes")
                else:
                    print(f"         ❌ File does not exist after saving!")
                    
            except Exception as e:
                print(f"         ❌ Error saving hierarchical flow for slide {slide_num}: {e}")
                plt.close('all')
        
        print(f"   ✅ Hierarchical reading order flow visualizations complete!")



    def _draw_reading_order_box(self, ax, box, color, group_label, reading_order, is_root=False):
        """Draw box with reading order emphasis"""
        pos = box['position']
        
        # Normalize position
        if pos['left'] > 10000:  # EMU units
            left, top = pos['left'] / 12700, pos['top'] / 12700
            width, height = pos['width'] / 12700, pos['height'] / 12700
        else:
            left, top, width, height = pos['left'], pos['top'], pos['width'], pos['height']
        
        line_width = 5 if is_root else 3
        alpha = 0.7 if is_root else 0.4
        
        # Draw box
        rect = patches.Rectangle((left, top), width, height,
                               linewidth=line_width, edgecolor=color,
                               facecolor=color, alpha=alpha)
        ax.add_patch(rect)
        
        # Add labels
        if is_root:
            ax.text(left + width/2, top - 15, group_label,
                   ha='center', va='bottom', fontsize=12, weight='bold', color='darkblue',
                   bbox=dict(boxstyle="round,pad=0.3", facecolor=color, alpha=0.9))
            
            ax.text(left + 5, top + 5, f"#{reading_order}",
                   ha='left', va='top', fontsize=14, weight='bold', color='white',
                   bbox=dict(boxstyle="round,pad=0.3", facecolor='darkblue', alpha=0.9))



    def _draw_reading_flow_arrow(self, ax, from_box, to_box, step_num):
        """Draw arrow showing reading flow"""
        def get_center(box):
            pos = box['position']
            if pos['left'] > 10000:
                left, top = pos['left'] / 12700, pos['top'] / 12700
                width, height = pos['width'] / 12700, pos['height'] / 12700
            else:
                left, top, width, height = pos['left'], pos['top'], pos['width'], pos['height']
            return left + width/2, top + height/2
        
        from_x, from_y = get_center(from_box)
        to_x, to_y = get_center(to_box)
        
        # Draw flow arrow
        ax.annotate('', xy=(to_x, to_y), xytext=(from_x, from_y),
                    arrowprops=dict(arrowstyle='->', color='red', lw=3, alpha=0.8))
        
        # Add step number
        mid_x, mid_y = (from_x + to_x) / 2, (from_y + to_y) / 2
        ax.text(mid_x, mid_y, f"Step {step_num}",
                fontsize=10, ha='center', va='center', weight='bold',
                bbox=dict(boxstyle="round,pad=0.3", facecolor='yellow', alpha=0.9))



    def create_local_sectioning_visualization(self, spatial_slides):
        """Create visualization showing local line sectioning with reading order flow"""
        print(f"   🎨 Creating local line sectioning visualizations...")
        
        for slide in spatial_slides:
            line_dividers = slide.get('line_dividers', [])
            local_sections = slide.get('local_sections', [])
            slide_number = slide['slide_number']
            
            if not line_dividers and not local_sections:
                continue
            print(f"      🎨 Creating visualization for slide {slide_number} ({len(line_dividers)} lines, {len(local_sections)} sections)")
            
            # Create figure with single plot for local sectioning
            fig, ax = plt.subplots(1, 1, figsize=(20, 14))
            fig.suptitle(f'Slide {slide_number} - Local Line Sectioning & Reading Order Flow', 
                        fontsize=18, fontweight='bold')
            
            # Get slide dimensions
            if self.comprehensive_data and 'metadata' in self.comprehensive_data:
                slide_width = self.comprehensive_data['metadata'].get('slide_width', 9144000) / 12700
                slide_height = self.comprehensive_data['metadata'].get('slide_height', 6858000) / 12700
            else:
                slide_width, slide_height = 720, 540
            
            ax.set_title('Local Sections with Reading Order (Left→Right within rows, Top→Bottom between rows)', 
                        fontsize=14, fontweight='bold')
            
            # Draw slide boundary
            slide_rect = patches.Rectangle((0, 0), slide_width, slide_height,
                                         linewidth=3, edgecolor='black',
                                         facecolor='lightgray', alpha=0.1)
            ax.add_patch(slide_rect)
            
            # Generate distinct colors for each section
            section_colors = plt.cm.Set3(np.linspace(0, 1, max(len(local_sections), 1)))
            
            # Draw line dividers first (as thick red lines)
            for line in line_dividers:
                pos = line['position']
                left = pos['left'] / 12700 if pos['left'] > 10000 else pos['left']
                top = pos['top'] / 12700 if pos['top'] > 10000 else pos['top']
                width = pos['width'] / 12700 if pos['width'] > 10000 else pos['width']
                height = pos['height'] / 12700 if pos['height'] > 10000 else pos['height']
                # Enhanced line visualization for zero-height/zero-width lines
                # Determine line orientation
                if height == 0 or (width > 0 and width / max(height, 1) > 2.0):  # Horizontal line
                    # Draw horizontal line as actual line (not rectangle)
                    line_y = top + height/2 if height > 0 else top
                    ax.plot([left, left + width], [line_y, line_y], 
                           color='red', linewidth=12, alpha=0.9, solid_capstyle='round')
                    
                    # Add visual thickness with rectangle
                    visual_height = max(height, 6)  # Minimum 6pt visual height
                    line_rect = patches.Rectangle((left, top), width, visual_height,
                                                linewidth=3, edgecolor='red',
                                                facecolor='red', alpha=0.4)
                elif width == 0 or (height > 0 and max(width, 1) / height < 0.5):  # Vertical line
                    # Draw vertical line as actual line (not rectangle)
                    line_x = left + width/2 if width > 0 else left
                    ax.plot([line_x, line_x], [top, top + height], 
                           color='red', linewidth=12, alpha=0.9, solid_capstyle='round')
                    
                    # Add visual thickness with rectangle
                    visual_width = max(width, 6)  # Minimum 6pt visual width
                    line_rect = patches.Rectangle((left, top), visual_width, height,
                                                linewidth=3, edgecolor='red',
                                                facecolor='red', alpha=0.4)
                else:  # Diagonal or other line
                    line_rect = patches.Rectangle((left, top), width, height,
                                                linewidth=4, edgecolor='red',
                                                facecolor='red', alpha=0.8)
                ax.add_patch(line_rect)
                # Add line label
                ax.text(left + width/2, top - 10, f"📏 {line['box_id']} (LINE DIVIDER)",
                       ha='center', va='bottom', fontsize=10, weight='bold', color='red',
                       bbox=dict(boxstyle="round,pad=0.3", facecolor='white', alpha=0.9))
            
            # Draw each local section with its shapes in reading order
            for section_idx, section in enumerate(local_sections):
                section_color = section_colors[section_idx % len(section_colors)]
                section_shapes = section.get('shapes', [])
                if not section_shapes:
                    continue
                # Draw section boundary (optional - can be commented out if too cluttered)
                section_bounds = section.get('bounds')
                if section_bounds:
                    bounds_left = section_bounds['left'] / 12700 if section_bounds['left'] > 10000 else section_bounds['left']
                    bounds_top = section_bounds['top'] / 12700 if section_bounds['top'] > 10000 else section_bounds['top']
                    bounds_width = (section_bounds['right'] - section_bounds['left']) / 12700 if section_bounds['right'] > 10000 else (section_bounds['right'] - section_bounds['left'])
                    bounds_height = (section_bounds['bottom'] - section_bounds['top']) / 12700 if section_bounds['bottom'] > 10000 else (section_bounds['bottom'] - section_bounds['top'])
                    
                    section_boundary = patches.Rectangle((bounds_left, bounds_top), bounds_width, bounds_height,
                                                       linewidth=2, edgecolor=section_color,
                                                       facecolor='none', alpha=0.7, linestyle='--')
                    ax.add_patch(section_boundary)
                # Add section label
                if section_bounds:
                    ax.text(bounds_left + bounds_width/2, bounds_top - 5, 
                           f"📍 {section['section_id']} ({len(section_shapes)} shapes)",
                           ha='center', va='bottom', fontsize=11, weight='bold', color=section_color,
                           bbox=dict(boxstyle="round,pad=0.4", facecolor=section_color, alpha=0.3))
                # Draw shapes in reading order with sequence numbers
                prev_shape = None
                for shape_idx, shape in enumerate(section_shapes):
                    pos = shape['position']
                    left = pos['left'] / 12700 if pos['left'] > 10000 else pos['left']
                    top = pos['top'] / 12700 if pos['top'] > 10000 else pos['top']
                    width = pos['width'] / 12700 if pos['width'] > 10000 else pos['width']
                    height = pos['height'] / 12700 if pos['height'] > 10000 else pos['height']
                    
                    # Draw shape rectangle
                    shape_rect = patches.Rectangle((left, top), width, height,
                                                 linewidth=2, edgecolor=section_color,
                                                 facecolor=section_color, alpha=0.4)
                    ax.add_patch(shape_rect)
                    
                    # Add shape ID and reading order number
                    ax.text(left + 5, top + 5, f"{shape['box_id']}\n#{shape_idx + 1}",
                           fontsize=9, weight='bold', color='black',
                           bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.9))
                    
                    # Add text content preview if available
                    if shape.get('has_text') and shape.get('text'):
                        text_preview = shape['text'][:20] + '...' if len(shape['text']) > 20 else shape['text']
                        text_preview = text_preview.replace('\n', ' ')
                        ax.text(left + width/2, top + height - 5, f'"{text_preview}"',
                               ha='center', va='bottom', fontsize=8, style='italic', color='darkblue',
                               bbox=dict(boxstyle="round,pad=0.2", facecolor='lightyellow', alpha=0.8))
                    
                    # Draw reading order arrow to next shape
                    if prev_shape and shape_idx > 0:
                        self._draw_reading_order_arrow(ax, prev_shape, shape, shape_idx, section_color)
                    
                    prev_shape = shape
            
            # Set axis properties
            ax.set_xlim(-50, slide_width + 50)
            ax.set_ylim(-50, slide_height + 50)
            ax.invert_yaxis()
            ax.set_aspect('equal')
            ax.grid(True, alpha=0.3)
            ax.set_xlabel('Width (points)', fontsize=11)
            ax.set_ylabel('Height (points)', fontsize=11)
            
            # Add legend
            legend_text = f"📏 {len(line_dividers)} Line Dividers | 📍 {len(local_sections)} Local Sections\n"
            legend_text += "Reading Order: Left→Right within rows, Top→Bottom between rows\n"
            legend_text += "Arrows show reading flow within each section"
            
            ax.text(0.02, 0.98, legend_text, transform=ax.transAxes, fontsize=10,
                   verticalalignment='top', bbox=dict(boxstyle="round,pad=0.5", facecolor='lightcyan', alpha=0.9))
            
            plt.tight_layout()
            
            # Save local sectioning visualization
            viz_path = self.spatial_analysis_dir / f"slide_{slide_number:02d}_local_sectioning.png"
            plt.savefig(viz_path, dpi=100, facecolor='white')
            plt.close('all')
            
            print(f"         💾 Saved local sectioning visualization: {viz_path}")
        
        print(f"      ✅ Local sectioning visualizations complete")



    def _draw_reading_order_arrow(self, ax, from_shape, to_shape, step_num, color):
        """Draw arrow showing reading order flow between shapes"""
        def get_shape_center(shape):
            pos = shape['position']
            left = pos['left'] / 12700 if pos['left'] > 10000 else pos['left']
            top = pos['top'] / 12700 if pos['top'] > 10000 else pos['top']
            width = pos['width'] / 12700 if pos['width'] > 10000 else pos['width']
            height = pos['height'] / 12700 if pos['height'] > 10000 else pos['height']
            return left + width/2, top + height/2
        
        from_x, from_y = get_shape_center(from_shape)
        to_x, to_y = get_shape_center(to_shape)
        
        # Draw reading order arrow
        ax.annotate('', xy=(to_x, to_y), xytext=(from_x, from_y),
                    arrowprops=dict(arrowstyle='->', color=color, lw=2, alpha=0.7))
        
        # Add step number on the arrow
        mid_x, mid_y = (from_x + to_x) / 2, (from_y + to_y) / 2
        ax.text(mid_x, mid_y, f"{step_num}",
                fontsize=8, ha='center', va='center', weight='bold', color='white',
                bbox=dict(boxstyle="circle,pad=0.2", facecolor=color, alpha=0.9))


