Attribute VB_Name = "mod_EntityScanner"
'======================================================================
' mod_EntityScanner — DXF Entity Inventory Scanner
'======================================================================
' Quick-scans an ASCII DXF file to build an inventory of which entity
' types exist on which layers. This is used by the decision gates to
' determine whether VBA can handle the file or ezdxf is needed.
'
' The scan is FAST because it only reads group codes 0 (entity type)
' and 8 (layer name). It does NOT read vertex data or other attributes.
'
' Entity classification:
'   SIMPLE (VBA can handle):  LWPOLYLINE, TEXT, LINE, CIRCLE, ARC, POINT, MTEXT
'   COMPLEX (needs ezdxf):    HATCH, SPLINE, POLYLINE (old), INSERT, ELLIPSE
'======================================================================
Option Explicit

' ── Quantity-relevant layer prefixes ──────────────────────────────────
Private Const SLAB_PREFIX   As String = "STR-SLAB-"
Private Const WALL_PREFIX   As String = "STR-WALL-"
Private Const LABEL_LAYER   As String = "STR-TYP-TXTNUM"
Private Const CUTOUT_LAYER  As String = "STR-SLAB-CUTOUT"

' ── Complex entity types that VBA cannot reliably parse ───────────────
Private Const COMPLEX_TYPES As String = ",HATCH,SPLINE,POLYLINE,INSERT,ELLIPSE,"


'----------------------------------------------------------------------
' ScanResult — Holds the scan output
'----------------------------------------------------------------------
Public Type ScanResult
    IsASCII             As Boolean    ' False if binary DXF
    TotalEntities       As Long       ' Total entity count
    SimpleCount         As Long       ' Entities VBA can handle
    ComplexCount         As Long       ' Entities needing ezdxf
    ComplexOnQuantity    As Long       ' Complex entities on relevant layers
    NeedsEzdxf          As Boolean    ' Final verdict
    ' Detail arrays for reporting
    ComplexDetails       As String    ' Human-readable summary
    EntityTypes          As String    ' Comma-separated list of all types found
End Type


'----------------------------------------------------------------------
' ScanDXFFile — Quick-scan a DXF file and return entity inventory
'----------------------------------------------------------------------
Public Function ScanDXFFile(ByVal filePath As String) As ScanResult
    Dim result As ScanResult
    
    ' ── Gate 1: ASCII check ───────────────────────────────────
    result.IsASCII = IsASCIIDXF(filePath)
    If Not result.IsASCII Then
        result.NeedsEzdxf = True
        result.ComplexDetails = "Binary DXF format detected. Requires ezdxf."
        ScanDXFFile = result
        Exit Function
    End If
    
    ' ── Scan ENTITIES section ─────────────────────────────────
    Dim f As Integer
    Dim codeLine As String, valueLine As String
    Dim code As Long, value As String
    Dim inEntities As Boolean, sectionPending As Boolean
    Dim currentType As String, currentLayer As String
    
    ' Track entity types (using concatenated string as poor-man's set)
    Dim allTypes As String: allTypes = ","
    ' Track complex entities on quantity layers
    Dim complexDetailsList As String: complexDetailsList = ""
    
    f = FreeFile
    Open filePath For Input As #f
    
    inEntities = False
    sectionPending = False
    
    Do While Not EOF(f)
        If EOF(f) Then Exit Do
        Line Input #f, codeLine
        If EOF(f) Then Exit Do
        Line Input #f, valueLine
        
        code = Val(Trim(codeLine))
        value = Trim(valueLine)
        
        ' Section tracking
        If code = 0 And value = "SECTION" Then
            sectionPending = True
            GoTo NextLine
        End If
        If code = 2 And sectionPending Then
            sectionPending = False
            If value = "ENTITIES" Then inEntities = True
            GoTo NextLine
        End If
        If sectionPending Then sectionPending = False
        
        If Not inEntities Then GoTo NextLine
        
        ' Entity boundary
        If code = 0 Then
            ' Process previous entity
            If Len(currentType) > 0 Then
                result.TotalEntities = result.TotalEntities + 1
                
                ' Track unique types
                If InStr(1, allTypes, "," & currentType & ",") = 0 Then
                    allTypes = allTypes & currentType & ","
                End If
                
                ' Classify
                If IsComplexType(currentType) Then
                    result.ComplexCount = result.ComplexCount + 1
                    
                    ' Is this on a quantity-relevant layer?
                    If IsQuantityLayer(currentLayer) Then
                        result.ComplexOnQuantity = result.ComplexOnQuantity + 1
                        complexDetailsList = complexDetailsList & _
                            currentType & " on layer '" & currentLayer & "'" & vbLf
                    End If
                Else
                    result.SimpleCount = result.SimpleCount + 1
                End If
            End If
            
            ' Start new entity
            If value = "ENDSEC" Then Exit Do
            currentType = value
            currentLayer = ""
            GoTo NextLine
        End If
        
        ' Capture layer (code 8) — only needed for classification
        If code = 8 And inEntities Then
            currentLayer = value
        End If
        
NextLine:
    Loop
    
    ' Process last entity
    If Len(currentType) > 0 And currentType <> "ENDSEC" Then
        result.TotalEntities = result.TotalEntities + 1
        If IsComplexType(currentType) Then
            result.ComplexCount = result.ComplexCount + 1
            If IsQuantityLayer(currentLayer) Then
                result.ComplexOnQuantity = result.ComplexOnQuantity + 1
                complexDetailsList = complexDetailsList & _
                    currentType & " on layer '" & currentLayer & "'" & vbLf
            End If
        Else
            result.SimpleCount = result.SimpleCount + 1
        End If
    End If
    
    Close #f
    
    ' ── Final verdict ─────────────────────────────────────────
    result.NeedsEzdxf = (result.ComplexOnQuantity > 0)
    result.ComplexDetails = complexDetailsList
    result.EntityTypes = Mid(allTypes, 2)   ' strip leading comma
    
    ScanDXFFile = result
End Function


'----------------------------------------------------------------------
' IsComplexType — Check if entity type requires ezdxf
'----------------------------------------------------------------------
Private Function IsComplexType(ByVal entityType As String) As Boolean
    IsComplexType = (InStr(1, COMPLEX_TYPES, "," & UCase(entityType) & ",") > 0)
End Function


'----------------------------------------------------------------------
' IsQuantityLayer — Check if a layer is relevant for quantity extraction
'----------------------------------------------------------------------
Private Function IsQuantityLayer(ByVal layerName As String) As Boolean
    Dim uLayer As String: uLayer = UCase(layerName)
    
    If Left(uLayer, Len(SLAB_PREFIX)) = UCase(SLAB_PREFIX) Then
        IsQuantityLayer = True
    ElseIf Left(uLayer, Len(WALL_PREFIX)) = UCase(WALL_PREFIX) Then
        IsQuantityLayer = True
    ElseIf StrComp(uLayer, UCase(LABEL_LAYER), vbTextCompare) = 0 Then
        IsQuantityLayer = True
    ElseIf StrComp(uLayer, UCase(CUTOUT_LAYER), vbTextCompare) = 0 Then
        IsQuantityLayer = True
    Else
        IsQuantityLayer = False
    End If
End Function


'----------------------------------------------------------------------
' CheckPythonAvailable — Check if Python + ezdxf are accessible
'----------------------------------------------------------------------
Public Function CheckPythonAvailable() As Boolean
    On Error GoTo NotAvailable
    
    Dim objShell As Object
    Dim objExec As Object
    Dim result As String
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("python -c ""import ezdxf; print('OK')""")
    
    ' Wait for completion (max 10 seconds)
    Dim waitCount As Long: waitCount = 0
    Do While objExec.Status = 0 And waitCount < 100
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 0) + 0.0001
        waitCount = waitCount + 1
    Loop
    
    If objExec.Status <> 0 Then
        result = ""
    End If
    
    result = Trim(objExec.StdOut.ReadAll)
    CheckPythonAvailable = (result = "OK")
    Exit Function
    
NotAvailable:
    CheckPythonAvailable = False
End Function


'----------------------------------------------------------------------
' FormatScanReport — Human-readable scan summary for Dashboard
'----------------------------------------------------------------------
Public Function FormatScanReport(ByRef scan As ScanResult) As String
    Dim rpt As String
    
    rpt = "Entity Scan Results:" & vbLf
    rpt = rpt & "  Total entities: " & scan.TotalEntities & vbLf
    rpt = rpt & "  Simple (VBA OK): " & scan.SimpleCount & vbLf
    rpt = rpt & "  Complex: " & scan.ComplexCount & vbLf
    rpt = rpt & "  Complex on quantity layers: " & scan.ComplexOnQuantity & vbLf
    rpt = rpt & "  Types found: " & scan.EntityTypes & vbLf
    rpt = rpt & vbLf
    
    If scan.NeedsEzdxf Then
        rpt = rpt & "VERDICT: ezdxf required for:" & vbLf
        rpt = rpt & scan.ComplexDetails
    Else
        rpt = rpt & "VERDICT: VBA can handle all entities."
    End If
    
    FormatScanReport = rpt
End Function
