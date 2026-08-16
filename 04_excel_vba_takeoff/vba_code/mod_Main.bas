Attribute VB_Name = "mod_Main"
'======================================================================
' mod_Main — Orchestration & Decision Gate Routing
'======================================================================
' Entry point for the DXF Quantity Take-Off.
' Called from the Dashboard "Run Take-Off" button.
'
' Decision Gates:
'   GATE 1: ASCII or Binary DXF?
'   GATE 2: Simple or complex entities?
'   GATE 3: Complex entities on quantity-relevant layers?
'   GATE 4: Python + ezdxf available?
'
' Two execution paths:
'   VBA PATH:  Parse → Extract → Report → Visualize (all in VBA)
'   EZDXF PATH: Shell to Python → Read JSON → Report → Visualize
'======================================================================
Option Explicit


'----------------------------------------------------------------------
' RunTakeoff — Main entry point (assigned to the Dashboard button)
'----------------------------------------------------------------------
Public Sub RunTakeoff()
    
    On Error GoTo ErrorHandler
    
    ' ── 1. File Selection ─────────────────────────────────────
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        FileFilter:="DXF Files (*.dxf), *.dxf", _
        title:="Select DXF File for Quantity Take-Off")
    
    If filePath = False Then Exit Sub   ' user cancelled
    
    Dim filePathStr As String: filePathStr = CStr(filePath)
    Dim fileName As String: fileName = Dir(filePathStr)
    
    ' ── Update Dashboard ──────────────────────────────────────
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    wsDash.Range("D7").value = fileName
    wsDash.Range("D8").value = "Scanning..."
    wsDash.Range("D8").Font.Color = RGB(255, 165, 0)   ' orange
    wsDash.Range("D9").value = ""
    wsDash.Range("D10").value = ""
    DoEvents
    
    ' ── Performance settings ──────────────────────────────────
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim startTime As Double: startTime = Timer
    
    ' ══════════════════════════════════════════════════════════
    ' GATE 1: ASCII or Binary DXF?
    ' ══════════════════════════════════════════════════════════
    Application.StatusBar = "Gate 1: Checking DXF format..."
    DoEvents
    
    Dim isASCII As Boolean
    isASCII = IsASCIIDXF(filePathStr)
    
    If Not isASCII Then
        ' ── BINARY DXF → Must use ezdxf ──────────────────────
        wsDash.Range("D10").value = "Binary DXF detected → routing to ezdxf"
        Call RunEzdxfPath(filePathStr, fileName, wsDash, startTime)
        Exit Sub
    End If
    
    ' ══════════════════════════════════════════════════════════
    ' GATE 2 + 3: Scan entity inventory
    ' ══════════════════════════════════════════════════════════
    Application.StatusBar = "Gate 2: Scanning entity types..."
    DoEvents
    
    Dim scanResult As ScanResult
    scanResult = ScanDXFFile(filePathStr)
    
    wsDash.Range("D10").value = "Entities: " & scanResult.TotalEntities & _
                                 " total, " & scanResult.SimpleCount & " simple, " & _
                                 scanResult.ComplexCount & " complex (" & _
                                 scanResult.ComplexOnQuantity & " on qty layers)"
    DoEvents
    
    If scanResult.NeedsEzdxf Then
        ' ── COMPLEX ENTITIES ON QUANTITY LAYERS → ezdxf path ──
        wsDash.Range("D10").value = wsDash.Range("D10").value & _
                                     " → routing to ezdxf"
        Call RunEzdxfPath(filePathStr, fileName, wsDash, startTime)
        Exit Sub
    End If
    
    ' ══════════════════════════════════════════════════════════
    ' VBA PATH — All entities are simple, proceed with VBA
    ' ══════════════════════════════════════════════════════════
    wsDash.Range("D10").value = wsDash.Range("D10").value & _
                                 " → VBA engine (no ezdxf needed)"
    Call RunVBAPath(filePathStr, fileName, wsDash, startTime)
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "An error occurred:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Please check that your DXF file is valid.", _
           vbCritical, "Error"
    
    On Error Resume Next
    Dim wsDashErr As Worksheet
    Set wsDashErr = ThisWorkbook.Sheets("Dashboard")
    wsDashErr.Range("D8").value = "Error!"
    wsDashErr.Range("D8").Font.Color = RGB(200, 0, 0)
    wsDashErr.Range("D9").value = "Error " & Err.Number & ": " & Err.Description
    On Error GoTo 0
End Sub


'======================================================================
' VBA PATH — Full VBA-only processing pipeline
'======================================================================
Private Sub RunVBAPath(ByVal filePathStr As String, ByVal fileName As String, _
                        ByVal wsDash As Worksheet, ByVal startTime As Double)
    
    On Error GoTo VBAError
    
    ' ── Step 1: Parse DXF ─────────────────────────────────────
    Application.StatusBar = "VBA Engine [1/5]: Parsing DXF file..."
    DoEvents
    
    Dim polys() As DXFPolyline, polyCount As Long
    Dim texts() As DXFText, textCount As Long
    Call ParseDXFFile(filePathStr, polys, polyCount, texts, textCount)
    
    ' ── Step 2: Extract Slabs ─────────────────────────────────
    Application.StatusBar = "VBA Engine [2/5]: Extracting slabs..."
    DoEvents
    
    Dim slabs() As SlabEntry, slabCount As Long
    Dim slabWarnings() As String, slabWarnCount As Long
    Call ExtractSlabs(polys, polyCount, texts, textCount, _
                      slabs, slabCount, slabWarnings, slabWarnCount)
    
    ' Count matching stats
    Dim matchedText As Long, matchedLayer As Long
    matchedText = 0: matchedLayer = 0
    Dim si As Long
    For si = 0 To slabCount - 1
        If slabs(si).ThicknessSource = "text_label" Then matchedText = matchedText + 1
        If slabs(si).ThicknessSource = "layer_name" Then matchedLayer = matchedLayer + 1
    Next si
    
    ' ── Step 3: Extract Walls ─────────────────────────────────
    Application.StatusBar = "VBA Engine [3/5]: Extracting walls..."
    DoEvents
    
    Dim strWalls() As WallEntry, strCount As Long
    Dim nsWalls() As WallEntry, nsCount As Long
    Dim wallWarnings() As String, wallWarnCount As Long
    Call ExtractWalls(polys, polyCount, strWalls, strCount, _
                      nsWalls, nsCount, wallWarnings, wallWarnCount)
    
    ' ── Get floor height from Dashboard ───────────────────────
    Dim floorHeight As Double: floorHeight = 3#
    On Error Resume Next
    floorHeight = CDbl(wsDash.Range("D12").value)
    On Error GoTo VBAError
    If floorHeight <= 0 Then floorHeight = 3#
    
    ' ── Step 4: Write Reports ─────────────────────────────────
    Application.StatusBar = "VBA Engine [4/5]: Writing reports..."
    DoEvents
    
    Call WriteSlabDetails(ThisWorkbook.Sheets("Slab Details"), _
                          slabs, slabCount, fileName, _
                          slabWarnings, slabWarnCount, matchedText, matchedLayer)
    
    Call WriteSlabSummary(ThisWorkbook.Sheets("Slab Summary"), _
                          slabs, slabCount, fileName)
    
    Call WriteWallSheet(ThisWorkbook.Sheets("Structural Walls"), _
                        strWalls, strCount, True, floorHeight)
    
    Call WriteWallSheet(ThisWorkbook.Sheets("Non-Structural Walls"), _
                        nsWalls, nsCount, False, floorHeight)
    
    ' ── Step 5: Draw Maps ─────────────────────────────────────
    Application.StatusBar = "VBA Engine [5/5]: Drawing visual maps..."
    DoEvents
    
    Call DrawSlabMap(ThisWorkbook.Sheets("Slab Map"), slabs, slabCount)
    Call DrawWallMap(ThisWorkbook.Sheets("Wall Map"), strWalls, strCount, nsWalls, nsCount)
    
    ' ── Done ──────────────────────────────────────────────────
    Call FinishTakeoff(wsDash, "VBA", slabCount, strCount, nsCount, startTime)
    Exit Sub
    
VBAError:
    Call HandlePathError("VBA", wsDash, startTime)
End Sub


'======================================================================
' EZDXF PATH — Python/ezdxf fallback pipeline
'======================================================================
Private Sub RunEzdxfPath(ByVal filePathStr As String, ByVal fileName As String, _
                          ByVal wsDash As Worksheet, ByVal startTime As Double)
    
    On Error GoTo EzdxfError
    
    ' ══════════════════════════════════════════════════════════
    ' GATE 4: Is Python + ezdxf available?
    ' ══════════════════════════════════════════════════════════
    Application.StatusBar = "Gate 4: Checking Python + ezdxf availability..."
    DoEvents
    
    Dim pythonAvailable As Boolean
    pythonAvailable = CheckPythonAvailable()
    
    If Not pythonAvailable Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        
        ' ── GATE 4 FAILED: Offer partial VBA run ─────────────
        Dim userChoice As VbMsgBoxResult
        userChoice = MsgBox( _
            "This DXF file contains complex entities that require " & _
            "Python + ezdxf to process." & vbCrLf & vbCrLf & _
            "Python + ezdxf was NOT found on this machine." & vbCrLf & vbCrLf & _
            "Options:" & vbCrLf & _
            "  YES = Run VBA-only (partial results — complex entities skipped)" & vbCrLf & _
            "  NO = Cancel (install Python + ezdxf first)" & vbCrLf & vbCrLf & _
            "To install: pip install ezdxf", _
            vbYesNo + vbExclamation, _
            "Python + ezdxf Not Available")
        
        If userChoice = vbYes Then
            ' Fall back to VBA-only with a warning
            wsDash.Range("D10").value = wsDash.Range("D10").value & _
                                         " → VBA fallback (partial results)"
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' Only proceed with VBA if the file is ASCII
            If IsASCIIDXF(filePathStr) Then
                Call RunVBAPath(filePathStr, fileName, wsDash, startTime)
            Else
                Application.StatusBar = False
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                Application.EnableEvents = True
                MsgBox "Cannot process: Binary DXF files require Python + ezdxf." & vbCrLf & _
                       "Please install Python and run: pip install ezdxf", _
                       vbCritical, "Cannot Process"
                wsDash.Range("D8").value = "Failed"
                wsDash.Range("D8").Font.Color = RGB(200, 0, 0)
            End If
        Else
            wsDash.Range("D8").value = "Cancelled"
            wsDash.Range("D8").Font.Color = RGB(150, 150, 150)
        End If
        Exit Sub
    End If
    
    ' ── Python + ezdxf IS available ───────────────────────────
    Application.StatusBar = "ezdxf Engine: Running Python bridge..."
    wsDash.Range("D8").value = "Running ezdxf..."
    wsDash.Range("D8").Font.Color = RGB(255, 165, 0)
    DoEvents
    
    ' ── Step 1: Call Python bridge ────────────────────────────
    Dim outputPath As String
    outputPath = Environ("TEMP") & "\dxf_takeoff_result.json"
    
    ' Delete previous output
    On Error Resume Next
    Kill outputPath
    On Error GoTo EzdxfError
    
    Dim bridgeOK As Boolean
    bridgeOK = CallEzdxfBridge(filePathStr, outputPath)
    
    If Not bridgeOK Then
        wsDash.Range("D8").value = "ezdxf bridge failed"
        wsDash.Range("D8").Font.Color = RGB(200, 0, 0)
        Call RestoreAppSettings
        Exit Sub
    End If
    
    ' ── Step 2: Read bridge results ───────────────────────────
    Application.StatusBar = "ezdxf Engine: Reading results..."
    DoEvents
    
    Dim slabs() As SlabEntry, slabCount As Long
    Dim strWalls() As WallEntry, strCount As Long
    Dim nsWalls() As WallEntry, nsCount As Long
    Dim warnings() As String, warnCount As Long
    
    Call ReadBridgeResults(outputPath, slabs, slabCount, _
                           strWalls, strCount, nsWalls, nsCount, _
                           warnings, warnCount)
    
    ' ── Get floor height ──────────────────────────────────────
    Dim floorHeight As Double: floorHeight = 3#
    On Error Resume Next
    floorHeight = CDbl(wsDash.Range("D12").value)
    On Error GoTo EzdxfError
    If floorHeight <= 0 Then floorHeight = 3#
    
    ' ── Step 3: Write Reports ─────────────────────────────────
    Application.StatusBar = "ezdxf Engine: Writing reports..."
    DoEvents
    
    ' Count matching stats
    Dim matchedText As Long, matchedLayer As Long
    matchedText = 0: matchedLayer = 0
    Dim si As Long
    For si = 0 To slabCount - 1
        If slabs(si).ThicknessSource = "text_label" Then matchedText = matchedText + 1
        If slabs(si).ThicknessSource = "layer_name" Then matchedLayer = matchedLayer + 1
    Next si
    
    Call WriteSlabDetails(ThisWorkbook.Sheets("Slab Details"), _
                          slabs, slabCount, fileName, _
                          warnings, warnCount, matchedText, matchedLayer)
    
    Call WriteSlabSummary(ThisWorkbook.Sheets("Slab Summary"), _
                          slabs, slabCount, fileName)
    
    Call WriteWallSheet(ThisWorkbook.Sheets("Structural Walls"), _
                        strWalls, strCount, True, floorHeight)
    
    Call WriteWallSheet(ThisWorkbook.Sheets("Non-Structural Walls"), _
                        nsWalls, nsCount, False, floorHeight)
    
    ' ── Step 4: Draw Maps ─────────────────────────────────────
    Application.StatusBar = "ezdxf Engine: Drawing visual maps..."
    DoEvents
    
    Call DrawSlabMap(ThisWorkbook.Sheets("Slab Map"), slabs, slabCount)
    Call DrawWallMap(ThisWorkbook.Sheets("Wall Map"), strWalls, strCount, nsWalls, nsCount)
    
    ' ── Done ──────────────────────────────────────────────────
    Call FinishTakeoff(wsDash, "ezdxf", slabCount, strCount, nsCount, startTime)
    
    ' Clean up temp file
    On Error Resume Next
    Kill outputPath
    On Error GoTo 0
    Exit Sub
    
EzdxfError:
    Call HandlePathError("ezdxf", wsDash, startTime)
End Sub


'======================================================================
' Shared Helpers
'======================================================================

Private Sub FinishTakeoff(ByVal wsDash As Worksheet, ByVal engine As String, _
                           ByVal slabCount As Long, ByVal strCount As Long, _
                           ByVal nsCount As Long, ByVal startTime As Double)
    Dim elapsed As Double: elapsed = Timer - startTime
    
    Call RestoreAppSettings
    
    wsDash.Activate
    wsDash.Range("D8").value = "Complete " & ChrW(10004) & " (" & engine & " engine)"
    wsDash.Range("D8").Font.Color = RGB(0, 150, 0)
    wsDash.Range("D9").value = "Slabs: " & slabCount & _
                               "  |  Str Walls: " & strCount & _
                               "  |  NS Walls: " & nsCount & _
                               "  |  Time: " & Format(elapsed, "0.0") & "s"
    
    MsgBox "Take-off complete!" & vbCrLf & vbCrLf & _
           "Engine used: " & engine & vbCrLf & _
           "Slabs found: " & slabCount & vbCrLf & _
           "Structural Walls: " & strCount & vbCrLf & _
           "Non-Structural Walls: " & nsCount & vbCrLf & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & vbCrLf & _
           "Check the sheet tabs for detailed results and maps.", _
           vbInformation, "DXF Quantity Take-Off"
End Sub

Private Sub HandlePathError(ByVal engine As String, ByVal wsDash As Worksheet, _
                             ByVal startTime As Double)
    Call RestoreAppSettings
    
    MsgBox "An error occurred in the " & engine & " engine:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, engine & " Engine Error"
    
    On Error Resume Next
    wsDash.Range("D8").value = "Error! (" & engine & ")"
    wsDash.Range("D8").Font.Color = RGB(200, 0, 0)
    wsDash.Range("D9").value = "Error " & Err.Number & ": " & Err.Description
    On Error GoTo 0
End Sub

Private Sub RestoreAppSettings()
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
