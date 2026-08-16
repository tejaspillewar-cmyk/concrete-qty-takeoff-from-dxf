# DXF Quantity Take-Off — Excel VBA Template

## What This Is

A single `.xlsm` Excel workbook that extracts **slab** and **wall** quantities from a DXF structural floor plan. It uses a **Hybrid Dual-Engine Architecture**:
- **Primary Engine (VBA)**: Fast, zero-dependency processing for simple ASCII DXF files.
- **Fallback Engine (Python + ezdxf)**: Automatically kicks in for binary DXF files or when complex entities (HATCH, SPLINE, INSERT) are detected.

Click a button → select a DXF file → the workbook detects the best engine, processes the file, and gives you formatted reports + visual maps. All inside Excel.

## Building the Template

### Prerequisites

1. **Python 3.x** with `pywin32` installed
2. **Microsoft Excel** installed on your machine
3. **Trust access** to VBA project object model enabled:
   - Open Excel → File → Options → Trust Center → Trust Center Settings
   - Macro Settings → ✅ "Trust access to the VBA project object model"

### Build

```powershell
cd e:\experiments\04_excel_vba_takeoff
python build_workbook.py
```

This creates `DXF_Takeoff.xlsm` in the same directory.

## Using the Template

1. Open `DXF_Takeoff.xlsm` and **Enable Macros**
2. Set the **Floor-to-Floor Height** on the Dashboard (default: 3.0 m)
3. Click **Run Take-Off**
4. Select your DXF file
5. Results populate across 6 sheet tabs:
   - **Slab Details** — every slab with area and volume
   - **Slab Summary** — totals grouped by thickness
   - **Slab Map** — color-coded polygon visualization
   - **Structural Walls** — quantities with volume formulas
   - **Non-Structural Walls** — quantities with volume formulas
   - **Wall Map** — structural (red) + non-structural (blue) visualization

## DXF Requirements

The system handles both simple and complex DXF features automatically.

**VBA Engine handles:**
- Slab boundaries: `STR-SLAB-REG{thickness}` (Closed LWPOLYLINE, CIRCLE, ARC)
- Thickness labels: `STR-TYP-TXTNUM` (TEXT, MTEXT with formatting stripped)
- Structural walls: `STR-WALL-REG` (Closed LWPOLYLINE)
- Non-structural walls: `STR-WALL-NS-{100|150|200}` (Closed LWPOLYLINE)

**ezdxf Fallback Engine kicks in automatically if it detects:**
- **Binary DXF** format
- **HATCH** entities (e.g. solid fills representing slabs)
- **SPLINE** boundaries (NURBS curves)
- **INSERT** (Block References, e.g. for columns)
- **POLYLINE** (legacy 2D/3D polylines)
- **ELLIPSE**

If ezdxf is needed but not installed, the system will offer a "graceful degradation" option to run a partial VBA extraction skipping the complex entities.

## Manual VBA Import (Fallback)

If the build script can't inject VBA (permissions issue), you can import manually:

1. Create a new blank `.xlsm` workbook
2. Press `Alt+F11` to open the VBA editor
3. For each `.bas` file in `vba_code/`:
   - File → Import File → select the `.bas` file
4. Add sheet tabs: Dashboard, Slab Details, Slab Summary, Slab Map, Structural Walls, Non-Structural Walls, Wall Map
5. Add a button on Dashboard and assign the `RunTakeoff` macro

## File Structure

```text
04_excel_vba_takeoff/
├── build_workbook.py           # Python builder (creates .xlsm)
├── DXF_Takeoff.xlsm            # Output template
├── README.md
├── vba_code/
│   ├── mod_Types.bas            # Shared type definitions
│   ├── mod_DXFParser.bas        # ASCII DXF parser (VBA engine)
│   ├── mod_Geometry.bas         # Math functions
│   ├── mod_SlabExtractor.bas    # Slab extraction (VBA engine)
│   ├── mod_WallExtractor.bas    # Wall extraction (VBA engine)
│   ├── mod_ReportWriter.bas     # Sheet formatting
│   ├── mod_Visualizer.bas       # Excel shape drawing
│   ├── mod_EntityScanner.bas    # Quick-scan for entity types & routing logic
│   ├── mod_EzdxfBridge.bas      # VBA→Python bridge orchestrator
│   └── mod_Main.bas             # Dashboard UI & decision gates
└── ezdxf_bridge/
    ├── ezdxf_bridge.py          # Python worker for complex entities
    └── requirements.txt         # ezdxf dependency
```
