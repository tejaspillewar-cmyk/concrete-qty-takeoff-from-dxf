from dataclasses import dataclass
from datetime import datetime
import os
import ezdxf

from geometry import polygon_area_sqmm, sqmm_to_sqm, polygon_centroid
from dxf_helpers import load_dxf

@dataclass
class WallEntry:
    name: str
    layer: str
    thickness_mm: int
    area_sqm: float
    length_m: float
    centroid_x: float
    centroid_y: float
    vertices: list
    is_structural: bool

@dataclass
class WallSummaryRow:
    thickness_mm: int
    count: int
    total_area_sqm: float
    total_length_m: float

@dataclass
class WallReport:
    file_name: str
    file_path: str
    generated_at: str
    structural_walls: list[WallEntry]
    non_structural_walls: list[WallEntry]
    structural_summary: list[WallSummaryRow]
    non_structural_summary: list[WallSummaryRow]
    warnings: list[str]

def calculate_perimeter_m(vertices):
    """Calculate the perimeter of a polygon in meters (input in mm)."""
    import math
    perimeter_mm = 0.0
    for i in range(len(vertices)):
        p1 = vertices[i]
        p2 = vertices[(i + 1) % len(vertices)]
        perimeter_mm += math.dist(p1, p2)
    return perimeter_mm / 1000.0

def extract_walls(dxf_path: str, str_layers: list[str] = None, ns_layers: list[str] = None) -> WallReport:
    doc, msp = load_dxf(dxf_path)
    file_name = os.path.basename(dxf_path)
    warnings = []

    if str_layers is None:
        str_layers = ["STR-WALL-REG"]
    if ns_layers is None:
        ns_layers = ["STR-WALL-NS-100", "STR-WALL-NS-150", "STR-WALL-NS-200"]

    all_walls = []
    
    # Simple extraction of LWPOLYLINEs
    for entity in msp:
        if entity.dxftype() != "LWPOLYLINE" or not entity.closed:
            continue
        
        layer = entity.dxf.get("layer", "")
        if layer not in str_layers and layer not in ns_layers:
            continue
            
        vertices = list(entity.get_points(format="xy"))
        area_sqm = sqmm_to_sqm(polygon_area_sqmm(vertices))
        perimeter_m = calculate_perimeter_m(vertices)
        
        # Estimate length (perimeter / 2) as walls are usually drawn as thin rectangles
        length_m = perimeter_m / 2.0
        
        cx, cy = polygon_centroid(vertices)
        is_str = layer in str_layers
        
        # Determine thickness from layer name if possible
        thickness = 0
        import re
        numbers = re.findall(r'\d+', layer)
        if numbers:
            thickness = int(numbers[-1])
        
        all_walls.append({
            "layer": layer,
            "thickness_mm": thickness,
            "area_sqm": area_sqm,
            "length_m": length_m,
            "cx": cx,
            "cy": cy,
            "vertices": vertices,
            "is_str": is_str
        })

    # Build objects
    str_entries = []
    ns_entries = []
    
    str_count = 1
    ns_counters = {}
    
    for w in all_walls:
        if w["is_str"]:
            name = f"SW-{str_count}"
            str_count += 1
            str_entries.append(WallEntry(
                name=name, layer=w["layer"], thickness_mm=w["thickness_mm"],
                area_sqm=round(w["area_sqm"], 3), length_m=round(w["length_m"], 3),
                centroid_x=round(w["cx"], 1), centroid_y=round(w["cy"], 1),
                vertices=w["vertices"], is_structural=True
            ))
        else:
            t = w["thickness_mm"]
            if t not in ns_counters: ns_counters[t] = 1
            name = f"NSW{t}-{ns_counters[t]}"
            ns_counters[t] += 1
            ns_entries.append(WallEntry(
                name=name, layer=w["layer"], thickness_mm=t,
                area_sqm=round(w["area_sqm"], 3), length_m=round(w["length_m"], 3),
                centroid_x=round(w["cx"], 1), centroid_y=round(w["cy"], 1),
                vertices=w["vertices"], is_structural=False
            ))

    # Build summaries
    def build_summary(entries):
        summary_dict = {}
        for e in entries:
            t = e.thickness_mm
            if t not in summary_dict:
                summary_dict[t] = {"count": 0, "area": 0.0, "length": 0.0}
            summary_dict[t]["count"] += 1
            summary_dict[t]["area"] += e.area_sqm
            summary_dict[t]["length"] += e.length_m
            
        return [
            WallSummaryRow(
                thickness_mm=t, count=d["count"], 
                total_area_sqm=round(d["area"], 3), total_length_m=round(d["length"], 3)
            ) for t, d in sorted(summary_dict.items())
        ]

    return WallReport(
        file_name=file_name,
        file_path=dxf_path,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        structural_walls=str_entries,
        non_structural_walls=ns_entries,
        structural_summary=build_summary(str_entries),
        non_structural_summary=build_summary(ns_entries),
        warnings=warnings
    )
