"""
Slab Quantity Take-Off — Entry Point
=====================================
Usage:
    python run.py "C:\\path\\to\\your\\file.dxf"

Output:
    1. Formatted table printed to console
    2. Excel file saved next to the DXF: <filename>_slab_takeoff.xlsx
"""
import sys
import os

# Add this directory to path so imports work
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from slab_extractor import extract_slabs
from excel_report import write_excel_report
from slab_visualizer import render_slab_map


def print_report(report):
    """Print a formatted console summary of the slab take-off."""
    
    sep = "=" * 78
    thin_sep = "-" * 78
    
    print()
    print(sep)
    print("  SLAB QUANTITY TAKE-OFF")
    print(sep)
    print(f"  File     : {report.file_name}")
    print(f"  Generated: {report.generated_at}")
    print(f"  Slabs    : {report.total_slabs}")
    print(f"  Matching : {report.matched_by_text} by text label, "
          f"{report.matched_by_layer} by layer name")
    print(sep)
    print()
    
    # ── Detail table header ───────────────────────────────────
    fmt = "  {:<10} {:<22} {:>8} {:>10} {:>12} {:>12}"
    print(fmt.format(
        "Slab", "Layer", "Thk(mm)", "Source", "Area(m2)", "Volume(m3)"
    ))
    print(f"  {thin_sep}")
    
    # ── Detail rows ───────────────────────────────────────────
    for slab in report.slabs:
        src_short = "text" if slab.thickness_source == "text_label" else "layer"
        print(fmt.format(
            slab.name,
            slab.layer,
            slab.thickness_mm,
            src_short,
            f"{slab.area_sqm:.3f}",
            f"{slab.volume_cum:.4f}",
        ))
    
    # ── Summary ───────────────────────────────────────────────
    print()
    print(sep)
    print("  SUMMARY BY THICKNESS")
    print(sep)
    
    sum_fmt = "  {:>8} mm  |  {:>5} slabs  |  Area: {:>12} m2  |  Volume: {:>12} m3"
    for row in report.summary:
        print(sum_fmt.format(
            row.thickness_mm,
            row.count,
            f"{row.total_area_sqm:.3f}",
            f"{row.total_volume_cum:.4f}",
        ))
    
    print(f"  {thin_sep}")
    print(sum_fmt.format(
        "TOTAL",
        report.total_slabs,
        f"{report.grand_total_area:.3f}",
        f"{report.grand_total_volume:.4f}",
    ))
    print(sep)
    
    # ── Warnings ──────────────────────────────────────────────
    if report.warnings:
        print()
        print("  WARNINGS:")
        for w in report.warnings:
            print(f"    - {w}")
        print()


def main():
    # ── Parse argument ────────────────────────────────────────
    if len(sys.argv) < 2:
        print("Usage: python run.py <path_to_dxf>")
        print('Example: python run.py "C:\\Users\\OMEN\\OneDrive\\Desktop\\Sample-1.dxf"')
        sys.exit(1)
    
    dxf_path = sys.argv[1]
    
    if not os.path.isfile(dxf_path):
        print(f"ERROR: File not found: {dxf_path}")
        sys.exit(1)
    
    # ── Extract ───────────────────────────────────────────────
    print(f"\nReading: {dxf_path}")
    print("Extracting slab quantities...")
    
    report = extract_slabs(dxf_path)
    
    # ── Print console report ──────────────────────────────────
    print_report(report)
    
    # ── Generate Visual Map ───────────────────────────────────
    dxf_dir = os.path.dirname(os.path.abspath(dxf_path))
    base_name = os.path.splitext(os.path.basename(dxf_path))[0]
    
    print("Generating visual map...")
    img_path = os.path.join(dxf_dir, f"{base_name}_slab_map.png")
    render_slab_map(report, img_path)
    
    # ── Save Excel ────────────────────────────────────────────
    excel_path = os.path.join(dxf_dir, f"{base_name}_slab_takeoff_v2.xlsx")
    
    print("Writing Excel report...")
    write_excel_report(report, excel_path, image_path=img_path)
    print(f"\n  Excel saved: {excel_path}")
    print(f"  Map saved:   {img_path}")
    print()


if __name__ == "__main__":
    main()
