import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib
import numpy as np

matplotlib.use('Agg')

from beam_extractor import BeamReport


def generate_beam_map(report: BeamReport, output_path: str):
    """
    Renders a color-coded map of all extracted beams.
    Color is based on beam depth.
    """
    fig, ax = plt.subplots(figsize=(16, 12), dpi=200)
    ax.set_facecolor("#1e1e1e") # Dark background
    
    if not report.beams:
        ax.text(0.5, 0.5, "NO BEAMS FOUND", color="white", fontsize=20, ha="center")
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
        ax.text(b.cx, b.cy, label, color="white", fontsize=4, ha="center", va="center")

    ax.autoscale_view()
    ax.set_aspect('equal')
    ax.set_title(f"Extracted Beams: {len(report.beams)}", color="white", fontsize=16)
    ax.axis("off")
    
    plt.tight_layout()
    plt.savefig(output_path, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)
