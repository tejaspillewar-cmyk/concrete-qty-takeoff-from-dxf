import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.patches as mpatches
import numpy as np

from beam_extractor import BeamReport

BACKGROUND_COLOR = "#1a1a2e"
GRID_COLOR       = "#2a2a4e"
LABEL_COLOR      = "#FFFFFF"


def generate_beam_map(report: BeamReport, output_path: str):
    """
    Renders a color-coded map of all extracted beams.
    Color is based on beam depth. Dark background matching slab/wall visualizers.
    """
    fig, ax = plt.subplots(figsize=(24, 17), dpi=200, facecolor=BACKGROUND_COLOR)
    ax.set_facecolor(BACKGROUND_COLOR)
    
    if not report.beams:
        ax.text(0.5, 0.5, "NO BEAMS FOUND", color=LABEL_COLOR, fontsize=20, ha="center")
        plt.savefig(output_path, bbox_inches='tight', facecolor=fig.get_facecolor())
        plt.close()
        return

    # Color map for depths
    depths = sorted(list(set([b.depth for b in report.beams])))
    cmap = matplotlib.colormaps['viridis'].resampled(max(1, len(depths)))
    
    # Plot beams
    for b in report.beams:
        c_idx = depths.index(b.depth)
        color = cmap(c_idx)
        
        # Draw a rotated rectangle for the beam
        L_mm = b.cad_length * 1000
        W_mm = b.width * 1000
        
        # Calculate bottom-left corner of the UNROTATED rectangle relative to centroid
        bx = b.cx - L_mm / 2
        by = b.cy - W_mm / 2
        
        rect = patches.Rectangle(
            (bx, by), L_mm, W_mm,
            angle=b.angle,
            rotation_point='center',
            facecolor=color, alpha=0.6, edgecolor='white', linewidth=1.0
        )
        ax.add_patch(rect)
        
        label = f"{b.name}\n{int(b.width*1000)}x{int(b.depth*1000)}\nL={b.clear_length:.1f}m"
        ax.text(
            b.cx, b.cy, label,
            color=LABEL_COLOR, fontsize=4, ha="center", va="center",
            bbox=dict(boxstyle="round,pad=0.15", facecolor="#0d1117", edgecolor="none", alpha=0.6)
        )

    ax.autoscale_view()
    ax.set_aspect('equal')
    ax.grid(True, color=GRID_COLOR, linewidth=0.3, alpha=0.5)
    ax.tick_params(colors="#555555", labelsize=6)
    for spine in ax.spines.values():
        spine.set_color(GRID_COLOR)

    # Legend for beam depths
    depths = sorted(list(set([b.depth for b in report.beams])))
    cmap = matplotlib.colormaps['viridis'].resampled(max(1, len(depths)))
    legend_handles = [
        mpatches.Patch(facecolor=cmap(depths.index(d)), edgecolor="white", alpha=0.7,
                       label=f"D={int(d*1000)} mm")
        for d in depths
    ]
    if legend_handles:
        ax.legend(handles=legend_handles, loc="upper right", fontsize=9,
                  facecolor="#16213e", edgecolor="#444444",
                  labelcolor=LABEL_COLOR, framealpha=0.9)

    ax.set_title(
        f"BEAM MAP — {len(report.beams)} Beams Extracted",
        color=LABEL_COLOR, fontsize=16, fontweight="bold", pad=20
    )

    fig.savefig(output_path, dpi=200, facecolor=fig.get_facecolor(),
                bbox_inches='tight', pad_inches=0.3)
    plt.close(fig)
