"""
Slab Visualization
==================
Renders a labeled PNG map of all slab polygons, color-coded by thickness,
with auto-generated slab names placed at each polygon's centroid.
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Polygon as MplPolygon
from matplotlib.collections import PatchCollection
import numpy as np

from slab_extractor import SlabReport


# Color palette per thickness (dark background friendly)
THICKNESS_COLORS = {
    125: "#00BFFF",   # cyan / sky blue
    150: "#32CD32",   # lime green
    175: "#FFD700",   # gold
    200: "#FF6347",   # tomato red
}
DEFAULT_COLOR = "#AAAAAA"

BACKGROUND_COLOR = "#1a1a2e"
GRID_COLOR = "#2a2a4e"
LABEL_COLOR = "#FFFFFF"
BORDER_COLOR = "#FFFFFF"


def render_slab_map(report: SlabReport, output_path: str, dpi: int = 200):
    """
    Render a PNG image showing all slab polygons, color-coded by thickness,
    with slab names labeled at each polygon's centroid.
    
    Args:
        report: SlabReport from slab_extractor.
        output_path: Path to save the PNG file.
        dpi: Image resolution.
    """
    if not report.slabs:
        return
    
    fig, ax = plt.subplots(1, 1, figsize=(24, 17), facecolor=BACKGROUND_COLOR)
    ax.set_facecolor(BACKGROUND_COLOR)
    
    all_x = []
    all_y = []
    
    # ── Draw polygons ──────────────────────────────────────────
    for slab in report.slabs:
        verts = slab.vertices
        thickness = slab.thickness_mm
        color = THICKNESS_COLORS.get(thickness, DEFAULT_COLOR)
        
        if not verts:
            continue
            
        # Collect for bounds
        xs = [v[0] for v in verts]
        ys = [v[1] for v in verts]
        all_x.extend(xs)
        all_y.extend(ys)
        
        # Draw filled polygon with transparency
        polygon = MplPolygon(
            verts,
            closed=True,
            facecolor=color,
            edgecolor=BORDER_COLOR,
            linewidth=0.8,
            alpha=0.35,
        )
        ax.add_patch(polygon)
        
        # Draw border with full opacity
        polygon_border = MplPolygon(
            verts,
            closed=True,
            facecolor="none",
            edgecolor=color,
            linewidth=1.2,
            alpha=0.9,
        )
        ax.add_patch(polygon_border)
    
    # ── Label each slab at centroid ────────────────────────────
    for slab in report.slabs:
        cx, cy = slab.centroid_x, slab.centroid_y
        name = slab.name
        area = slab.area_sqm
        thickness = slab.thickness_mm
        color = THICKNESS_COLORS.get(thickness, DEFAULT_COLOR)
        
        # Adaptive font size based on area
        if area < 2.0:
            fontsize = 5
        elif area < 8.0:
            fontsize = 6.5
        elif area < 20.0:
            fontsize = 8
        else:
            fontsize = 9.5
        
        # Slab name (bold, top line)
        ax.text(
            cx, cy + (fontsize * 25),  # offset up slightly
            name,
            fontsize=fontsize,
            fontweight="bold",
            color=LABEL_COLOR,
            ha="center", va="center",
            bbox=dict(
                boxstyle="round,pad=0.2",
                facecolor=color,
                edgecolor="none",
                alpha=0.7,
            ),
        )
        
        # Area text (smaller, below name)
        ax.text(
            cx, cy - (fontsize * 25),
            f"{area:.2f} m\u00b2",
            fontsize=fontsize * 0.75,
            color=color,
            ha="center", va="center",
            alpha=0.9,
        )
    
    # ── Axis formatting ────────────────────────────────────────
    if all_x and all_y:
        margin_x = (max(all_x) - min(all_x)) * 0.05
        margin_y = (max(all_y) - min(all_y)) * 0.05
        ax.set_xlim(min(all_x) - margin_x, max(all_x) + margin_x)
        ax.set_ylim(min(all_y) - margin_y, max(all_y) + margin_y)
    
    ax.set_aspect("equal")
    ax.grid(True, color=GRID_COLOR, linewidth=0.3, alpha=0.5)
    ax.tick_params(colors="#555555", labelsize=6)
    for spine in ax.spines.values():
        spine.set_color(GRID_COLOR)
    
    # ── Title ──────────────────────────────────────────────────
    ax.set_title(
        "SLAB QUANTITY TAKE-OFF MAP",
        fontsize=16,
        fontweight="bold",
        color=LABEL_COLOR,
        pad=20,
    )
    
    # ── Legend ─────────────────────────────────────────────────
    legend_handles = []
    for thickness in sorted(THICKNESS_COLORS.keys()):
        color = THICKNESS_COLORS[thickness]
        patch = mpatches.Patch(
            facecolor=color,
            edgecolor=BORDER_COLOR,
            alpha=0.6,
            label=f"{thickness} mm slab",
        )
        legend_handles.append(patch)
    
    ax.legend(
        handles=legend_handles,
        loc="upper right",
        fontsize=10,
        facecolor="#16213e",
        edgecolor="#444444",
        labelcolor=LABEL_COLOR,
        framealpha=0.9,
    )
    
    # ── Save ───────────────────────────────────────────────────
    fig.savefig(output_path, dpi=dpi, facecolor=fig.get_facecolor(),
                bbox_inches="tight", pad_inches=0.3)
    plt.close(fig)
