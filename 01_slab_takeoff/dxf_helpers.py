"""
DXF entity extraction helpers.
Reads ezdxf entities and returns plain Python dicts — no DXF objects leak out.
"""
import re
import ezdxf


# Default layer patterns (your team's convention)
SLAB_LAYER_PREFIX = "STR-SLAB-REG"
SLAB_LABEL_LAYER = "STR-TYP-TXTNUM"
CUTOUT_LAYER = "STR-SLAB-CUTOUT"


def load_dxf(filepath: str):
    """
    Load and validate a DXF file.
    
    Returns:
        (doc, msp) tuple — the ezdxf document and its modelspace.
    
    Raises:
        FileNotFoundError, ezdxf.DXFError on invalid files.
    """
    doc = ezdxf.readfile(filepath)
    msp = doc.modelspace()
    return doc, msp


def extract_slab_polygons(msp, layer_prefix: str = SLAB_LAYER_PREFIX, allowed_layers: list[str] = None) -> list[dict]:
    """
    Extract all closed LWPOLYLINE entities from slab layers.
    
    If allowed_layers is provided, only those layers are processed.
    Otherwise, scans all layers whose name starts with `layer_prefix`.
    
    Returns:
        List of dicts:
        {
            "layer": str,
            "thickness_from_layer": int or None,
            "vertices": [(x, y), ...],
            "handle": str,
        }
    """
    slabs = []
    
    for entity in msp:
        if entity.dxftype() != "LWPOLYLINE":
            continue
        layer = entity.dxf.get("layer", "")
        
        if allowed_layers is not None:
            if layer not in allowed_layers:
                continue
        else:
            if not layer.startswith(layer_prefix):
                continue
                
        if not entity.closed:
            continue
        
        # Parse thickness
        thickness = None
        if allowed_layers is not None:
            # Use regex to find the last number in the custom layer name
            numbers = re.findall(r'\d+', layer)
            if numbers:
                thickness = int(numbers[-1])
        else:
            suffix = layer[len(layer_prefix):]
            try:
                thickness = int(suffix)
            except ValueError:
                pass
        
        vertices = list(entity.get_points(format="xy"))
        
        slabs.append({
            "layer": layer,
            "thickness_from_layer": thickness,
            "vertices": vertices,
            "handle": entity.dxf.get("handle", ""),
        })
    
    return slabs


def extract_thickness_labels(msp, label_layer: str = SLAB_LABEL_LAYER) -> list[dict]:
    """
    Extract TEXT entities that contain thickness info (e.g., "150 THK").
    
    Scans TEXT entities on `label_layer`, filters for those containing "THK",
    and parses the numeric thickness.
    
    Returns:
        List of dicts:
        {
            "text": str,           # original text, e.g., "150 THK"
            "thickness_mm": int,   # parsed value, e.g., 150
            "position": (x, y),    # insertion point
        }
    """
    labels = []
    
    for entity in msp.query(f'TEXT[layer=="{label_layer}"]'):
        text = entity.dxf.text.strip()
        if "THK" not in text.upper():
            continue
        
        # Parse thickness: "150 THK" → 150, "( 125MM THK- ...)" → 125
        match = re.search(r"(\d+)\s*(?:MM\s*)?THK", text.upper())
        if not match:
            continue
        
        thickness = int(match.group(1))
        pos = entity.dxf.insert
        
        labels.append({
            "text": text,
            "thickness_mm": thickness,
            "position": (pos[0], pos[1]),
        })
    
    return labels


def extract_cutout_polygons(msp, cutout_layer: str = CUTOUT_LAYER) -> list[dict]:
    """
    Extract closed LWPOLYLINE entities from the cutout layer.
    
    Returns:
        List of dicts:
        {
            "layer": str,
            "vertices": [(x, y), ...],
            "handle": str,
        }
    """
    cutouts = []
    
    for entity in msp.query(f'LWPOLYLINE[layer=="{cutout_layer}"]'):
        if not entity.closed:
            continue
        vertices = list(entity.get_points(format="xy"))
        cutouts.append({
            "layer": cutout_layer,
            "vertices": vertices,
            "handle": entity.dxf.get("handle", ""),
        })
    
    return cutouts


def get_layer_names(doc) -> list[str]:
    """Return all layer names in the DXF document."""
    return [layer.dxf.name for layer in doc.layers]


def get_slab_layer_names(doc, prefix: str = SLAB_LAYER_PREFIX) -> list[str]:
    """Return layer names that match the slab prefix."""
    return [name for name in get_layer_names(doc) if name.startswith(prefix)]
