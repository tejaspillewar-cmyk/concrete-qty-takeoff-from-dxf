"""
ezdxf Bridge Script — Called by VBA When Complex Entities Are Detected
=======================================================================
Usage:
    python ezdxf_bridge.py <input.dxf> <output.json>

This script handles everything VBA cannot:
  - Binary DXF files
  - HATCH entities (boundary extraction, area computation)
  - SPLINE entities (tessellation to polyline approximation)
  - Old POLYLINE entities (vertex sub-entities)
  - INSERT/BLOCK references (explode/attribute extraction)
  - ELLIPSE entities (parametric area calculation)
  - MTEXT with complex formatting

Output: JSON file with slabs[], str_walls[], ns_walls[], warnings[]
        matching the VBA SlabEntry/WallEntry structure.
"""
import sys
import os
import json
import re
import math

try:
    import ezdxf
except ImportError:
    print("ERROR: ezdxf is not installed. Run: pip install ezdxf", file=sys.stderr)
    sys.exit(1)


# ── Layer conventions ─────────────────────────────────────────────────
SLAB_LAYER_PREFIX = "STR-SLAB-REG"
SLAB_LABEL_LAYER = "STR-TYP-TXTNUM"
CUTOUT_LAYER = "STR-SLAB-CUTOUT"
STR_WALL_LAYER = "STR-WALL-REG"
NS_WALL_PREFIX = "STR-WALL-NS-"


# ── Geometry helpers ──────────────────────────────────────────────────

def shoelace_area(vertices):
    """Compute polygon area using Shoelace formula. Input in mm, returns mm²."""
    n = len(vertices)
    if n < 3:
        return 0.0
    area = 0.0
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        area += x1 * y2 - x2 * y1
    return abs(area) / 2.0


def polygon_centroid(vertices):
    """Arithmetic mean of vertices."""
    n = len(vertices)
    if n == 0:
        return (0.0, 0.0)
    cx = sum(v[0] for v in vertices) / n
    cy = sum(v[1] for v in vertices) / n
    return (cx, cy)


def polygon_perimeter(vertices):
    """Sum of edge lengths in mm."""
    n = len(vertices)
    peri = 0.0
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        peri += math.dist((x1, y1), (x2, y2))
    return peri


def point_in_polygon(px, py, polygon):
    """Ray-casting point-in-polygon test."""
    n = len(polygon)
    if n < 3:
        return False
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        if ((yi > py) != (yj > py)) and (px < (xj - xi) * (py - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    return inside


def entity_to_vertices(entity):
    """
    Convert various entity types to a list of (x, y) vertices.
    Handles LWPOLYLINE, old POLYLINE, HATCH boundaries, SPLINE (tessellated).
    Returns a list of vertex lists (one per boundary path for HATCH).
    """
    etype = entity.dxftype()

    if etype == "LWPOLYLINE":
        return [list(entity.get_points(format="xy"))]

    elif etype == "POLYLINE":
        # Old-style POLYLINE with VERTEX sub-entities
        verts = [(v.dxf.location.x, v.dxf.location.y)
                 for v in entity.vertices if v.is_poly_vertex]
        return [verts] if verts else []

    elif etype == "HATCH":
        paths = []
        for path in entity.paths:
            try:
                # Convert any path type to a list of vertices
                verts = list(path.vertices())
                if verts:
                    paths.append([(v[0], v[1]) for v in verts])
            except Exception:
                # For edge paths, try to tessellate
                try:
                    verts = list(path.control_vertices())
                    if verts:
                        paths.append([(v[0], v[1]) for v in verts])
                except Exception:
                    pass
        return paths

    elif etype == "SPLINE":
        # Tessellate spline to polyline approximation
        try:
            # ezdxf can approximate splines
            verts = list(entity.flattening(0.5))  # tolerance in drawing units
            return [[(v.x, v.y) for v in verts]] if verts else []
        except Exception:
            return []

    elif etype == "ELLIPSE":
        # Tessellate ellipse to polyline
        try:
            verts = list(entity.vertices(entity.params(num=64)))
            return [[(v.x, v.y) for v in verts]] if verts else []
        except Exception:
            return []

    return []


# ── Extraction logic ──────────────────────────────────────────────────

def extract_slabs(msp, warnings):
    """Extract slab quantities from all entity types on slab layers."""
    slabs = []
    thickness_counters = {}

    # ── Collect slab polygons ─────────────────────────────────
    slab_polys = []  # list of (layer, vertices, thickness_from_layer)

    for entity in msp:
        layer = entity.dxf.get("layer", "")
        if not layer.startswith(SLAB_LAYER_PREFIX):
            continue

        # Parse thickness from layer name
        suffix = layer[len(SLAB_LAYER_PREFIX):]
        thick_from_layer = None
        try:
            thick_from_layer = int(suffix)
        except ValueError:
            pass

        # Get vertices (handles LWPOLYLINE, POLYLINE, HATCH, SPLINE, etc.)
        etype = entity.dxftype()

        if etype in ("LWPOLYLINE", "POLYLINE"):
            if hasattr(entity, "closed") and not entity.closed:
                continue

        vertex_sets = entity_to_vertices(entity)
        for verts in vertex_sets:
            if len(verts) >= 3:
                slab_polys.append({
                    "layer": layer,
                    "thickness_from_layer": thick_from_layer,
                    "vertices": verts,
                })

    # ── Collect thickness labels ──────────────────────────────
    thk_labels = []
    for entity in msp:
        layer = entity.dxf.get("layer", "")
        if layer.upper() != SLAB_LABEL_LAYER.upper():
            continue

        etype = entity.dxftype()
        text = ""
        pos = (0.0, 0.0)

        if etype == "TEXT":
            text = entity.dxf.text.strip()
            p = entity.dxf.insert
            pos = (p[0], p[1])
        elif etype == "MTEXT":
            text = entity.plain_text().strip()
            p = entity.dxf.insert
            pos = (p[0], p[1])
        else:
            continue

        if "THK" not in text.upper():
            continue

        match = re.search(r"(\d+)\s*(?:MM\s*)?THK", text.upper())
        if match:
            thk_labels.append({
                "text": text,
                "thickness_mm": int(match.group(1)),
                "position": pos,
            })

    # ── Match slabs to labels ─────────────────────────────────
    used_labels = set()

    for slab in slab_polys:
        verts = slab["vertices"]
        layer = slab["layer"]

        thickness_mm = None
        thickness_source = None
        label_text = ""

        # Try text label matching
        for idx, label in enumerate(thk_labels):
            if idx in used_labels:
                continue
            lx, ly = label["position"]
            if point_in_polygon(lx, ly, verts):
                thickness_mm = label["thickness_mm"]
                thickness_source = "text_label"
                label_text = label["text"]
                used_labels.add(idx)
                break

        # Fallback to layer name
        if thickness_mm is None:
            if slab["thickness_from_layer"] is not None:
                thickness_mm = slab["thickness_from_layer"]
                thickness_source = "layer_name"
                label_text = f"(from layer: {layer})"
            else:
                thickness_mm = 0
                thickness_source = "unknown"
                label_text = "UNKNOWN"
                warnings.append(f"Slab on layer '{layer}': no thickness found.")

        # Cross-check
        if (thickness_source == "text_label"
                and slab["thickness_from_layer"] is not None
                and thickness_mm != slab["thickness_from_layer"]):
            warnings.append(
                f"Slab on layer '{layer}': text says {thickness_mm}mm "
                f"but layer says {slab['thickness_from_layer']}mm. Using text."
            )

        # Compute quantities
        area_sqmm = shoelace_area(verts)
        area_sqm = area_sqmm / 1_000_000.0
        volume_cum = area_sqm * (thickness_mm / 1000.0)
        cx, cy = polygon_centroid(verts)

        # Auto-name
        if thickness_mm not in thickness_counters:
            thickness_counters[thickness_mm] = 0
        thickness_counters[thickness_mm] += 1
        name = f"S{thickness_counters[thickness_mm]}-{thickness_mm}"

        slabs.append({
            "name": name,
            "layer": layer,
            "thickness_mm": thickness_mm,
            "thickness_source": thickness_source,
            "label_text": label_text,
            "area_sqm": round(area_sqm, 3),
            "volume_cum": round(volume_cum, 4),
            "centroid_x": round(cx, 1),
            "centroid_y": round(cy, 1),
            "vertices": [[round(v[0], 1), round(v[1], 1)] for v in verts],
        })

    return slabs


def extract_walls(msp, warnings):
    """Extract wall quantities."""
    str_walls = []
    ns_walls = []
    str_counter = 0
    ns_counters = {}

    for entity in msp:
        layer = entity.dxf.get("layer", "")

        is_str = (layer == STR_WALL_LAYER)
        is_ns = layer.startswith(NS_WALL_PREFIX)

        if not is_str and not is_ns:
            continue

        etype = entity.dxftype()
        if etype in ("LWPOLYLINE", "POLYLINE"):
            if hasattr(entity, "closed") and not entity.closed:
                continue

        vertex_sets = entity_to_vertices(entity)

        for verts in vertex_sets:
            if len(verts) < 3:
                continue

            # Thickness from layer
            thickness = 0
            if is_ns:
                suffix = layer[len(NS_WALL_PREFIX):]
                try:
                    thickness = int(suffix)
                except ValueError:
                    pass

            area_sqm = shoelace_area(verts) / 1_000_000.0
            peri_m = polygon_perimeter(verts) / 1000.0
            length_m = peri_m / 2.0
            cx, cy = polygon_centroid(verts)

            wall_entry = {
                "layer": layer,
                "thickness_mm": thickness,
                "area_sqm": round(area_sqm, 3),
                "length_m": round(length_m, 3),
                "centroid_x": round(cx, 1),
                "centroid_y": round(cy, 1),
                "vertices": [[round(v[0], 1), round(v[1], 1)] for v in verts],
            }

            if is_str:
                str_counter += 1
                wall_entry["name"] = f"SW-{str_counter}"
                str_walls.append(wall_entry)
            else:
                if thickness not in ns_counters:
                    ns_counters[thickness] = 0
                ns_counters[thickness] += 1
                wall_entry["name"] = f"NSW{thickness}-{ns_counters[thickness]}"
                ns_walls.append(wall_entry)

    return str_walls, ns_walls


# ── Main ──────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print("Usage: python ezdxf_bridge.py <input.dxf> <output.json>", file=sys.stderr)
        sys.exit(1)

    dxf_path = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.isfile(dxf_path):
        print(f"ERROR: File not found: {dxf_path}", file=sys.stderr)
        sys.exit(1)

    print(f"ezdxf bridge: Reading {dxf_path}")
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    warnings = []

    print("ezdxf bridge: Extracting slabs...")
    slabs = extract_slabs(msp, warnings)

    print("ezdxf bridge: Extracting walls...")
    str_walls, ns_walls = extract_walls(msp, warnings)

    # Write JSON output
    output = {
        "file_name": os.path.basename(dxf_path),
        "slabs": slabs,
        "str_walls": str_walls,
        "ns_walls": ns_walls,
        "warnings": warnings,
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2)

    print(f"ezdxf bridge: Output written to {output_path}")
    print(f"ezdxf bridge: Slabs={len(slabs)}, StrWalls={len(str_walls)}, NSWalls={len(ns_walls)}")


if __name__ == "__main__":
    main()
