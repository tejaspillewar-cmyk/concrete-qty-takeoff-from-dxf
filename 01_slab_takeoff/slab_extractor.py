"""
Slab Quantity Take-Off Extractor
================================
Core logic: matches slab polygons with their thickness labels,
computes areas and volumes, and produces a structured report.
"""
from dataclasses import dataclass, field
from datetime import datetime
import os

from geometry import polygon_area_sqmm, sqmm_to_sqm, polygon_centroid, point_in_polygon
from dxf_helpers import load_dxf, extract_slab_polygons, extract_thickness_labels, extract_cutout_polygons


@dataclass
class SlabEntry:
    """One slab polygon with its computed quantities."""
    name: str                # e.g., "S1-150"
    layer: str               # e.g., "STR-SLAB-REG150"
    thickness_mm: int        # e.g., 150
    thickness_source: str    # "text_label" or "layer_name"
    label_text: str          # original label e.g., "150 THK" or "(from layer)"
    area_sqm: float          # m²
    volume_cum: float        # m³
    centroid_x: float        # mm (drawing coordinates)
    centroid_y: float        # mm (drawing coordinates)
    vertices: list = None    # raw polygon vertices [(x,y), ...] for visualization


@dataclass
class SlabSummaryRow:
    """Summary for one thickness group."""
    thickness_mm: int
    count: int
    total_area_sqm: float
    total_volume_cum: float


@dataclass
class SlabReport:
    """Complete slab take-off report."""
    file_name: str
    file_path: str
    generated_at: str
    total_slabs: int
    slabs: list[SlabEntry]
    summary: list[SlabSummaryRow]
    grand_total_area: float
    grand_total_volume: float
    warnings: list[str]
    # Stats
    matched_by_text: int
    matched_by_layer: int


def extract_slabs(dxf_path: str, allowed_layers: list[str] = None) -> SlabReport:
    """
    Main extraction function.
    
    1. Loads the DXF file
    2. Extracts slab polygons from STR-SLAB-REG* layers
    3. Extracts THK text labels from STR-TYP-TXTNUM
    4. Matches each slab polygon to its thickness label (point-in-polygon)
    5. Falls back to layer name if no text label found inside
    6. Computes area (m²) and volume (m³) for each slab
    7. Returns a structured SlabReport
    
    Args:
        dxf_path: Path to the DXF file.
        allowed_layers: Optional list of specific layers to extract from.
    
    Returns:
        SlabReport with all slab data.
    """
    warnings = []
    
    # ── Step 1: Load DXF ──────────────────────────────────────
    doc, msp = load_dxf(dxf_path)
    file_name = os.path.basename(dxf_path)
    
    # ── Step 2: Extract slab polygons ─────────────────────────
    slab_polys = extract_slab_polygons(msp, allowed_layers=allowed_layers)
    if not slab_polys:
        warnings.append("No slab polygons found on STR-SLAB-REG* layers!")
        return SlabReport(
            file_name=file_name, file_path=dxf_path,
            generated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            total_slabs=0, slabs=[], summary=[],
            grand_total_area=0, grand_total_volume=0,
            warnings=warnings, matched_by_text=0, matched_by_layer=0,
        )
    
    # ── Step 3: Extract THK text labels ───────────────────────
    thk_labels = extract_thickness_labels(msp)
    
    # ── Step 4 & 5: Match each slab to its thickness ──────────
    slab_entries = []
    matched_by_text = 0
    matched_by_layer = 0
    
    # Track used labels to avoid double-matching
    used_labels = set()
    
    # Group slabs by thickness for auto-naming
    thickness_counters = {}
    
    for slab in slab_polys:
        vertices = slab["vertices"]
        layer = slab["layer"]
        
        # Try to find a THK text label inside this polygon
        thickness_mm = None
        thickness_source = None
        label_text = ""
        
        for idx, label in enumerate(thk_labels):
            if idx in used_labels:
                continue
            lx, ly = label["position"]
            if point_in_polygon(lx, ly, vertices):
                thickness_mm = label["thickness_mm"]
                thickness_source = "text_label"
                label_text = label["text"]
                used_labels.add(idx)
                matched_by_text += 1
                break
        
        # Fallback: use thickness from layer name
        if thickness_mm is None:
            if slab["thickness_from_layer"] is not None:
                thickness_mm = slab["thickness_from_layer"]
                thickness_source = "layer_name"
                label_text = f"(from layer: {layer})"
                matched_by_layer += 1
            else:
                thickness_mm = 0
                thickness_source = "unknown"
                label_text = "UNKNOWN"
                warnings.append(
                    f"Slab on layer '{layer}' (handle={slab['handle']}): "
                    f"no THK label found and layer name has no thickness."
                )
        
        # Cross-check: if text label thickness doesn't match layer thickness
        if (thickness_source == "text_label" 
                and slab["thickness_from_layer"] is not None
                and thickness_mm != slab["thickness_from_layer"]):
            warnings.append(
                f"Slab on layer '{layer}' (handle={slab['handle']}): "
                f"text says {thickness_mm}mm but layer says {slab['thickness_from_layer']}mm. "
                f"Using text label value."
            )
        
        # ── Step 6: Compute quantities ────────────────────────
        area_sqmm = polygon_area_sqmm(vertices)
        area_sqm = sqmm_to_sqm(area_sqmm)
        volume_cum = area_sqm * (thickness_mm / 1000.0)  # m³
        cx, cy = polygon_centroid(vertices)
        
        # Auto-name: S1-150, S2-150, etc.
        if thickness_mm not in thickness_counters:
            thickness_counters[thickness_mm] = 0
        thickness_counters[thickness_mm] += 1
        slab_name = f"S{thickness_counters[thickness_mm]}-{thickness_mm}"
        
        slab_entries.append(SlabEntry(
            name=slab_name,
            layer=layer,
            thickness_mm=thickness_mm,
            thickness_source=thickness_source,
            label_text=label_text,
            area_sqm=round(area_sqm, 3),
            volume_cum=round(volume_cum, 4),
            centroid_x=round(cx, 1),
            centroid_y=round(cy, 1),
            vertices=vertices,
        ))
    
    # ── Step 7: Summary by thickness ──────────────────────────
    summary_dict = {}
    for s in slab_entries:
        t = s.thickness_mm
        if t not in summary_dict:
            summary_dict[t] = {"count": 0, "area": 0.0, "volume": 0.0}
        summary_dict[t]["count"] += 1
        summary_dict[t]["area"] += s.area_sqm
        summary_dict[t]["volume"] += s.volume_cum
    
    summary = [
        SlabSummaryRow(
            thickness_mm=t,
            count=d["count"],
            total_area_sqm=round(d["area"], 3),
            total_volume_cum=round(d["volume"], 4),
        )
        for t, d in sorted(summary_dict.items())
    ]
    
    grand_area = sum(s.total_area_sqm for s in summary)
    grand_volume = sum(s.total_volume_cum for s in summary)
    
    return SlabReport(
        file_name=file_name,
        file_path=dxf_path,
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        total_slabs=len(slab_entries),
        slabs=slab_entries,
        summary=summary,
        grand_total_area=round(grand_area, 3),
        grand_total_volume=round(grand_volume, 4),
        warnings=warnings,
        matched_by_text=matched_by_text,
        matched_by_layer=matched_by_layer,
    )
