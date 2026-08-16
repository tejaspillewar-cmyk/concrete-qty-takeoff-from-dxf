import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Polygon as MplPolygon

from wall_extractor import WallReport

BACKGROUND_COLOR = "#1a1a2e"
GRID_COLOR = "#2a2a4e"
LABEL_COLOR = "#FFFFFF"
BORDER_COLOR = "#FFFFFF"

def render_wall_map(report: WallReport, output_path: str, is_structural: bool, dpi: int = 200):
    
    walls = report.structural_walls if is_structural else report.non_structural_walls
    if not walls:
        return
        
    fig, ax = plt.subplots(1, 1, figsize=(24, 17), facecolor=BACKGROUND_COLOR)
    ax.set_facecolor(BACKGROUND_COLOR)
    
    all_x = []
    all_y = []
    
    color = "#FF4500" if is_structural else "#1E90FF"
    
    for w in walls:
        verts = w.vertices
        if not verts: continue
            
        xs = [v[0] for v in verts]
        ys = [v[1] for v in verts]
        all_x.extend(xs)
        all_y.extend(ys)
        
        polygon = MplPolygon(verts, closed=True, facecolor=color, edgecolor=BORDER_COLOR, linewidth=0.8, alpha=0.5)
        ax.add_patch(polygon)
        
        polygon_border = MplPolygon(verts, closed=True, facecolor="none", edgecolor=color, linewidth=1.5, alpha=0.9)
        ax.add_patch(polygon_border)
        
        # Label
        cx, cy = w.centroid_x, w.centroid_y
        area = w.area_sqm
        fontsize = 7
        
        ax.text(
            cx, cy,
            w.name,
            fontsize=fontsize, fontweight="bold", color=LABEL_COLOR, ha="center", va="center",
            bbox=dict(boxstyle="round,pad=0.2", facecolor=color, edgecolor="none", alpha=0.7)
        )

    if all_x and all_y:
        margin_x = (max(all_x) - min(all_x)) * 0.05
        margin_y = (max(all_y) - min(all_y)) * 0.05
        ax.set_xlim(min(all_x) - margin_x, max(all_x) + margin_x)
        ax.set_ylim(min(all_y) - margin_y, max(all_y) + margin_y)
    
    ax.set_aspect("equal")
    ax.grid(True, color=GRID_COLOR, linewidth=0.3, alpha=0.5)
    
    title = "STRUCTURAL WALL MAP" if is_structural else "NON-STRUCTURAL WALL MAP"
    ax.set_title(title, fontsize=16, fontweight="bold", color=LABEL_COLOR, pad=20)
    
    fig.savefig(output_path, dpi=dpi, facecolor=fig.get_facecolor(), bbox_inches="tight", pad_inches=0.3)
    plt.close(fig)
