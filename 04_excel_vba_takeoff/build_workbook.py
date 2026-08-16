"""
Build Script: DXF Quantity Take-Off Excel Template
====================================================
Creates a macro-enabled .xlsm workbook with:
  - 7 formatted sheets (Dashboard + 6 report/map sheets)
  - All VBA modules imported
  - A "Run Take-Off" button on the Dashboard

Usage:
    python build_workbook.py

Requirements:
    - Python 3.x
    - pywin32 (pip install pywin32)
    - Microsoft Excel installed

IMPORTANT: You must enable "Trust access to the VBA project object model"
           in Excel's Trust Center settings before running this script.
           (File > Options > Trust Center > Trust Center Settings >
            Macro Settings > check the box)
"""
import sys
import os
import subprocess
import time


def ensure_pywin32():
    """Install pywin32 if not available."""
    try:
        import win32com.client
        return True
    except ImportError:
        print("pywin32 not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
        print("pywin32 installed successfully.")
        return True


def main():
    ensure_pywin32()
    import win32com.client as win32

    script_dir = os.path.dirname(os.path.abspath(__file__))
    vba_dir = os.path.join(script_dir, "vba_code")
    output_path = os.path.join(script_dir, "DXF_Takeoff.xlsm")

    # Verify VBA files exist
    vba_modules = [
        "mod_Types.bas",
        "mod_DXFParser.bas",
        "mod_Geometry.bas",
        "mod_SlabExtractor.bas",
        "mod_WallExtractor.bas",
        "mod_ReportWriter.bas",
        "mod_Visualizer.bas",
        "mod_EntityScanner.bas",
        "mod_EzdxfBridge.bas",
        "mod_Main.bas",
    ]
    for mod in vba_modules:
        path = os.path.join(vba_dir, mod)
        if not os.path.exists(path):
            print(f"ERROR: Missing VBA module: {path}")
            sys.exit(1)

    print("=" * 60)
    print("  DXF Quantity Take-Off — Excel Template Builder")
    print("=" * 60)
    print()

    # ── Launch Excel ──────────────────────────────────────────
    print("Starting Excel...")
    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    try:
        # ── Create workbook ───────────────────────────────────
        print("Creating workbook...")
        wb = xl.Workbooks.Add()

        # ── Create sheets ─────────────────────────────────────
        # Default workbook has Sheet1. Rename and add more.
        sheet_names = [
            "Dashboard",
            "Slab Details",
            "Slab Summary",
            "Slab Map",
            "Structural Walls",
            "Non-Structural Walls",
            "Wall Map",
        ]

        # Rename first sheet
        wb.Sheets(1).Name = "Dashboard"

        # Add remaining sheets
        for name in sheet_names[1:]:
            ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
            ws.Name = name

        # ── Format Dashboard ──────────────────────────────────
        print("Setting up Dashboard...")
        ws = wb.Sheets("Dashboard")

        # Title
        ws.Range("A2:G2").Merge()
        cell = ws.Range("A2")
        cell.Value = "DXF QUANTITY TAKE-OFF"
        cell.Font.Name = "Calibri"
        cell.Font.Bold = True
        cell.Font.Size = 22
        cell.Font.Color = 0x64381F  # Dark blue (BGR)

        ws.Range("A3:G3").Merge()
        cell = ws.Range("A3")
        cell.Value = "Automated structural slab & wall quantity extraction from DXF files"
        cell.Font.Name = "Calibri"
        cell.Font.Size = 11
        cell.Font.Color = 0x666666

        # Separator
        ws.Range("A5:G5").Merge()
        ws.Range("A5").Interior.Color = 0xF0E4D6  # Light blue bar

        # Labels
        labels = {
            "C7": "File:",
            "C8": "Status:",
            "C9": "Results:",
            "C11": "Settings:",
            "C12": "Floor-to-Floor Height (m):",
        }
        for addr, text in labels.items():
            c = ws.Range(addr)
            c.Value = text
            c.Font.Name = "Calibri"
            c.Font.Bold = True
            c.Font.Size = 11
            c.Font.Color = 0x64381F

        # Default values
        ws.Range("D7").Value = "(No file selected)"
        ws.Range("D7").Font.Color = 0x999999
        ws.Range("D8").Value = "Ready"
        ws.Range("D8").Font.Color = 0x009900

        # Height input cell (yellow, editable)
        height_cell = ws.Range("D12")
        height_cell.Value = 3.0
        height_cell.Font.Bold = True
        height_cell.Font.Size = 12
        height_cell.Interior.Color = 0x00FFFF  # Yellow (BGR)
        height_cell.HorizontalAlignment = -4108  # xlCenter
        height_cell.Borders.LineStyle = 1
        height_cell.Borders.Color = 0xCCCCCC

        # Instructions
        ws.Range("A15:G15").Merge()
        ws.Range("A15").Value = "HOW TO USE"
        ws.Range("A15").Font.Bold = True
        ws.Range("A15").Font.Size = 13
        ws.Range("A15").Font.Color = 0x64381F

        instructions = [
            '1.  Set the "Floor-to-Floor Height" above (used for wall volume calculation)',
            '2.  Click the "Run Take-Off" button to select your DXF file',
            "3.  Wait for processing — progress is shown in the status bar",
            "4.  Results appear in the sheet tabs below (Slab Details, Summary, Maps, Walls)",
            "",
            "SUPPORTED DXF FORMAT:",
            "  - ASCII DXF only (not binary). Most CAD software exports ASCII by default.",
            "  - Slab layers: STR-SLAB-REG{thickness}  (e.g., STR-SLAB-REG150)",
            "  - Thickness labels: TEXT entities on STR-TYP-TXTNUM layer with 'THK'",
            "  - Structural walls: STR-WALL-REG layer",
            "  - Non-structural walls: STR-WALL-NS-{100|150|200} layers",
            "  - Drawing units must be millimeters",
        ]
        for idx, line in enumerate(instructions):
            row = 16 + idx
            ws.Range(f"A{row}:G{row}").Merge()
            c = ws.Range(f"A{row}")
            c.Value = line
            c.Font.Name = "Calibri"
            c.Font.Size = 10
            if line.startswith("SUPPORTED"):
                c.Font.Bold = True
                c.Font.Color = 0x64381F
            else:
                c.Font.Color = 0x444444

        # Column widths
        ws.Columns("A").ColumnWidth = 4
        ws.Columns("B").ColumnWidth = 6
        ws.Columns("C").ColumnWidth = 30
        ws.Columns("D").ColumnWidth = 50

        # ── Add Button ────────────────────────────────────────
        print("Adding Run Take-Off button...")
        # Using Form Control button (Buttons collection)
        btn = ws.Buttons().Add(30, 87, 180, 42)   # Left, Top, Width, Height (points)
        btn.OnAction = "RunTakeoff"
        btn.Characters.Text = "Run Take-Off"
        btn.Font.Size = 14
        btn.Font.Bold = True
        btn.Name = "btnRunTakeoff"

        # ── Import VBA modules ────────────────────────────────
        print("Importing VBA modules...")
        try:
            vbp = wb.VBProject
            for mod_file in vba_modules:
                mod_path = os.path.join(vba_dir, mod_file)
                vbp.VBComponents.Import(mod_path)
                mod_name = os.path.splitext(mod_file)[0]
                print(f"  + {mod_name}")
        except Exception as e:
            error_msg = str(e)
            if "programmatic access" in error_msg.lower() or "trust" in error_msg.lower() or "denied" in error_msg.lower():
                print()
                print("!" * 60)
                print("  ERROR: VBA project access is not trusted.")
                print()
                print("  To fix this:")
                print("  1. Open Excel")
                print("  2. File > Options > Trust Center")
                print("  3. Click 'Trust Center Settings'")
                print("  4. Go to 'Macro Settings'")
                print("  5. Check 'Trust access to the VBA project")
                print("     object model'")
                print("  6. Click OK, close Excel, and re-run this script")
                print("!" * 60)
            else:
                print(f"\n  ERROR importing VBA: {e}")
                print("  The workbook will be saved WITHOUT VBA macros.")
                print("  You can manually import the .bas files from vba_code/")
            
            # Still save the workbook (just without VBA)
            print()

        # ── Save as .xlsm ─────────────────────────────────────
        print(f"\nSaving: {output_path}")
        # FileFormat 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        wb.SaveAs(os.path.abspath(output_path), FileFormat=52)
        wb.Close(False)

        print()
        print("=" * 60)
        print(f"  SUCCESS! Template created:")
        print(f"  {output_path}")
        print("=" * 60)
        print()
        print("  Open the .xlsm file, enable macros, and click")
        print("  'Run Take-Off' to get started!")
        print()

    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        import traceback
        traceback.print_exc()

    finally:
        xl.DisplayAlerts = True
        try:
            xl.Quit()
        except:
            pass
        # Release COM objects
        del xl
        time.sleep(1)


if __name__ == "__main__":
    main()
