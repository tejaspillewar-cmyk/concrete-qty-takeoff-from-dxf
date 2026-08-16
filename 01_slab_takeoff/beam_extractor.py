import re
from dataclasses import dataclass, field
import ezdxf

from dxf_helpers import load_dxf, extract_slab_polygons
from geometry import (
    distance, line_angle, are_lines_parallel, 
    point_line_distance, check_lines_overlap,
    point_in_polygon, distance_pt_to_polygon,
    polygon_area_sqmm
)

@dataclass
class BeamText:
    raw_text: str
    name: str
    width: float
    depth: float
    cx: float
    cy: float

@dataclass
class VirtualBeam:
    p1: tuple[float, float]
    p2: tuple[float, float]  # Centerline endpoints
    calc_width: float        # CAD geometric width
    calc_length: float       # CAD centerline length
    is_rect: bool
    layer: str = ""
    
@dataclass
class BeamEntry:
    name: str
    width: float
    depth: float
    cad_length: float      # Original geometric span
    clear_length: float    # Span after subtracting column/wall overlaps
    volume: float          # clear_length * width * depth
    cx: float
    cy: float
    angle: float           # Centerline angle in degrees
    source: str = "text_label"
    layer: str = ""

@dataclass
class BeamReport:
    beams: list[BeamEntry] = field(default_factory=list)


def parse_beam_text(text: str) -> tuple[str, float, float]:
    """
    Parses strings like 'B 300X1350/850' or 'PB1 230x450'
    Returns (Name, Width, Depth). Max depth is taken if multiple are specified.
    """
    # Regex to find patterns like 'name WxDxD'
    text = text.upper().strip()
    # E.g. find B, PB, CB followed by numbers and X
    # Let's use a simpler heuristic: find the part containing 'X'
    
    parts = text.split()
    name = "BEAM"
    dims_str = ""
    
    if len(parts) >= 2:
        name = parts[0]
        dims_str = parts[1]
    elif len(parts) == 1:
        dims_str = parts[0]
        
    # Extract numbers from dims_str (e.g. "300X1350/850")
    # Split by X or *
    if 'X' in dims_str:
        dim_parts = dims_str.split('X')
    elif '*' in dims_str:
        dim_parts = dims_str.split('*')
    else:
        return name, 0.0, 0.0
        
    if len(dim_parts) < 2:
        return name, 0.0, 0.0
        
    try:
        width = float(re.sub(r'[^0-9.]', '', dim_parts[0]))
        
        # Depth part might have multiple separated by /
        depth_str = dim_parts[1]
        depths = [float(re.sub(r'[^0-9.]', '', d)) for d in depth_str.split('/') if re.sub(r'[^0-9.]', '', d)]
        
        depth = max(depths) if depths else 0.0
        return name, width, depth
    except Exception:
        return name, 0.0, 0.0


def extract_beam_texts(msp) -> list[BeamText]:
    """Find and parse all beam text labels."""
    beam_texts = []
    # Beams labels often exist on STR-TYP-TXTNUM or STR-BEAM layers
    for e in msp:
        if e.dxftype() in ('TEXT', 'MTEXT'):
            txt = e.dxf.text.strip()
            if 'X' in txt.upper() or '*' in txt:
                name, w, d = parse_beam_text(txt)
                if w > 0 and d > 0:
                    # Get position
                    ins = e.dxf.insert
                    beam_texts.append(BeamText(txt, name, w, d, ins.x, ins.y))
    return beam_texts


def build_virtual_beams(msp, allowed_layers: list[str] = None) -> list[VirtualBeam]:
    """
    Process LINE and LWPOLYLINE entities on beam layers to form VirtualBeam objects.
    """
    lines = []
    polys = []
    
    for e in msp:
        layer = e.dxf.layer
        if allowed_layers is not None:
            if layer not in allowed_layers:
                continue
        else:
            if 'BEAM' not in layer.upper():
                continue
                
        if e.dxftype() == 'LINE':
            lines.append(e)
        elif e.dxftype() == 'LWPOLYLINE':
            polys.append(e)
                
    virtual_beams = []
    
    # 1. Process Rectangles (LWPOLYLINEs)
    for poly in polys:
        pts = poly.get_points(format='xy')
        if len(pts) >= 4 and poly.closed:
            # Assuming rectangle
            p0, p1, p2, p3 = pts[0], pts[1], pts[2], pts[3]
            d1 = distance(p0, p1)
            d2 = distance(p1, p2)
            
            if d1 > d2:
                length, width = d1, d2
                # Centerline is midpoint of short edges
                cl_p1 = ((p0[0]+p3[0])/2, (p0[1]+p3[1])/2)
                cl_p2 = ((p1[0]+p2[0])/2, (p1[1]+p2[1])/2)
            else:
                length, width = d2, d1
                cl_p1 = ((p0[0]+p1[0])/2, (p0[1]+p1[1])/2)
                cl_p2 = ((p2[0]+p3[0])/2, (p2[1]+p3[1])/2)
                
            virtual_beams.append(VirtualBeam(cl_p1, cl_p2, width, length, is_rect=True, layer=poly.dxf.layer))
            
    # 2. Process Parallel Lines
    # O(N^2) pairing algorithm
    used_lines = set()
    
    # Pre-filter lines to ignore tiny segments (cross marks, column outlines, etc)
    long_lines = []
    for l in lines:
        p1a = (l.dxf.start.x, l.dxf.start.y)
        p1b = (l.dxf.end.x, l.dxf.end.y)
        if distance(p1a, p1b) >= 500: # Beams are typically > 500mm
            long_lines.append((l, p1a, p1b, distance(p1a, p1b)))
            
    for i, (l1, p1a, p1b, L1) in enumerate(long_lines):
        if i in used_lines:
            continue
        
        best_partner = -1
        best_overlap = 0
        best_width = 0
        
        for j, (l2, p2a, p2b, L2) in enumerate(long_lines):
            if i == j or j in used_lines:
                continue
            
            if are_lines_parallel(p1a, p1b, p2a, p2b, tol_deg=2.0):
                overlap = check_lines_overlap(p1a, p1b, p2a, p2b)
                # Mutual overlap check: must overlap at least 70% of BOTH lines
                if overlap > L1 * 0.7 and overlap > L2 * 0.7:
                    width = point_line_distance(p2a, p1a, p1b)
                    if 50 < width < 1500: # Reasonable beam width in mm
                        if overlap > best_overlap:
                            best_overlap = overlap
                            best_partner = j
                            best_width = width
                            
        if best_partner != -1:
            # Pair found!
            used_lines.add(i)
            used_lines.add(best_partner)
            
            # Create centerline
            # Midpoint between p1a and its projection on l2
            l2_partner, p2a, p2b, L2 = long_lines[best_partner]
            
            cl_p1 = ((p1a[0] + p2a[0])/2, (p1a[1] + p2a[1])/2)
            cl_p2 = ((p1b[0] + p2b[0])/2, (p1b[1] + p2b[1])/2)
            
            virtual_beams.append(VirtualBeam(cl_p1, cl_p2, best_width, best_overlap, is_rect=False, layer=l1.dxf.layer))

    return virtual_beams


def get_clear_span(vb: VirtualBeam, structural_walls: list[list[tuple[float, float]]]) -> float:
    """
    Subtracts column/wall overlaps from the geometric length.
    If a centerline endpoint falls inside a structural wall polygon,
    we subtract the distance from the endpoint to the polygon edge.
    """
    clear_length = vb.calc_length
    
    # Check Endpoint 1
    for wall in structural_walls:
        if point_in_polygon(vb.p1[0], vb.p1[1], wall):
            dist = distance_pt_to_polygon(vb.p1, wall)
            clear_length -= dist
            break # Only subtract once per endpoint
            
    # Check Endpoint 2
    for wall in structural_walls:
        if point_in_polygon(vb.p2[0], vb.p2[1], wall):
            dist = distance_pt_to_polygon(vb.p2, wall)
            clear_length -= dist
            break
            
    return max(0.0, clear_length)


def extract_beams(dxf_path: str, allowed_layers: list[str] = None, str_wall_layers: list[str] = None) -> BeamReport:
    """Main pipeline for Beam quantity takeoff."""
    doc, msp = load_dxf(dxf_path)
    
    # 1. Parse Texts
    beam_texts = extract_beam_texts(msp)
    
    # 2. Build Virtual Beams from geometry
    virtual_beams = build_virtual_beams(msp, allowed_layers=allowed_layers)
    
    # 3. Load structural walls for clear-span intersections
    str_walls = []
    if str_wall_layers is None:
        str_wall_layers = ["STR-WALL-REG"]
        
    for e in msp:
        if e.dxftype() == 'LWPOLYLINE' and e.dxf.get('layer', '') in str_wall_layers:
            if e.closed:
                str_walls.append(e.get_points(format='xy'))
                
    # 4. Associate Texts to Virtual Beams
    final_beams = []
    for vb in virtual_beams:
        cx = (vb.p1[0] + vb.p2[0]) / 2
        cy = (vb.p1[1] + vb.p2[1]) / 2
        
        # Find closest text that roughly matches the calculated width
        best_txt = None
        best_dist = float('inf')
        
        for txt in beam_texts:
            dist = distance((cx, cy), (txt.cx, txt.cy))
            # Width must be within 50mm tolerance
            if abs(txt.width - vb.calc_width) < 50:
                if dist < best_dist:
                    best_dist = dist
                    best_txt = txt
                    
        if best_txt:
            # Calculate clear span
            clear_len = get_clear_span(vb, str_walls)
            
            # Convert units: mm -> m
            w_m = best_txt.width / 1000.0
            d_m = best_txt.depth / 1000.0
            cad_L_m = vb.calc_length / 1000.0
            clr_L_m = clear_len / 1000.0
            
            vol = clr_L_m * w_m * d_m
            
            angle = line_angle(vb.p1, vb.p2)
            
            final_beams.append(BeamEntry(
                name=best_txt.name,
                width=w_m,
                depth=d_m,
                cad_length=cad_L_m,
                clear_length=clr_L_m,
                volume=vol,
                cx=cx,
                cy=cy,
                angle=angle,
                source="text_label",
                layer=vb.layer
            ))
        else:
            # Fallback to layer name
            name, w, d = parse_beam_text(vb.layer)
            if w > 0 and d > 0:
                clear_len = get_clear_span(vb, str_walls)
                
                w_m = w / 1000.0
                d_m = d / 1000.0
                cad_L_m = vb.calc_length / 1000.0
                clr_L_m = clear_len / 1000.0
                
                vol = clr_L_m * w_m * d_m
                angle = line_angle(vb.p1, vb.p2)
                
                final_beams.append(BeamEntry(
                    name=name,
                    width=w_m,
                    depth=d_m,
                    cad_length=cad_L_m,
                    clear_length=clr_L_m,
                    volume=vol,
                    cx=cx,
                    cy=cy,
                    angle=angle,
                    source="layer_name",
                    layer=vb.layer
                ))
            
    return BeamReport(beams=final_beams)
