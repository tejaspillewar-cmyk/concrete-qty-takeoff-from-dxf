# Slab Quantity Take-Off Tool

## What This Does

Reads a DXF structural floor plan and automatically extracts **slab quantities**:
- Identifies every slab polygon (closed boundary on `STR-SLAB-REG*` layers)
- Reads the thickness label text box inside each slab (e.g., "150 THK")
- Computes **area (m2)** and **volume (m3)** for each slab
- Outputs a formatted **Excel report** with details + summary

## How to Run

```powershell
python run.py "C:\path\to\your\file.dxf"
```

**Example:**
```powershell
python run.py "C:\Users\OMEN\OneDrive\Desktop\Sample-1.dxf"
```

**Output:**
- Console table with all slab quantities
- Excel file saved next to the DXF: `Sample-1_slab_takeoff.xlsx`

## Required DXF Layer Conventions

For this tool to work, the DXF must follow these layer naming rules:

| Element | Layer Name | Entity Type | Example |
|:---|:---|:---|:---|
| Slab boundaries | `STR-SLAB-REG{thickness}` | Closed LWPOLYLINE | `STR-SLAB-REG150` |
| Thickness labels | `STR-TYP-TXTNUM` | TEXT | "150 THK" |
| Slab cutouts | `STR-SLAB-CUTOUT` | Closed LWPOLYLINE + LINE (X-cross) | — |

### Rules:
1. Each slab must be drawn as a **closed polyline** on its thickness layer
2. Each slab should have a **TEXT entity** with `"{thickness} THK"` placed **inside** the polyline boundary
3. If a text label is missing, the tool falls back to reading the thickness from the layer name
4. Drawing units must be **millimeters**

## Excel Output

### Sheet 1: "Slab Details"
Every individual slab with:
- Slab Name (auto-generated: S1-150, S2-150, etc.)
- Layer name
- Thickness in mm
- Source (how thickness was determined: "text_label" or "layer_name")
- Area in m2
- Volume in m3
- Centroid coordinates

### Sheet 2: "Summary"
Totals grouped by thickness:
- Count of slabs per thickness
- Total area per thickness
- Total volume per thickness
- Grand totals

## Requirements

- Python 3.10+
- `ezdxf` — for reading DXF files
- `openpyxl` — for writing Excel files

Install:
```powershell
pip install ezdxf openpyxl
```

## File Structure

```
01_slab_takeoff/
  run.py              <- Entry point (run this)
  slab_extractor.py   <- Core extraction logic
  dxf_helpers.py      <- DXF entity reading helpers
  geometry.py         <- Polygon math (area, point-in-polygon)
  excel_report.py     <- Excel report generation
  README.md           <- This file
```
