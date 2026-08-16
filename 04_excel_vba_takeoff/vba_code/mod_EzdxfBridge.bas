Attribute VB_Name = "mod_EzdxfBridge"
'======================================================================
' mod_EzdxfBridge — VBA → Python/ezdxf Bridge
'======================================================================
' When VBA cannot handle certain DXF entities (HATCH, SPLINE, binary
' DXF, etc.), this module shells out to a Python script that uses ezdxf
' to do the heavy lifting.
'
' Flow:
'   1. VBA calls CallEzdxfBridge(dxfPath, outputPath)
'   2. Python script reads DXF with ezdxf, extracts all entities
'   3. Python writes results to a JSON file
'   4. VBA reads the JSON file and populates SlabEntry/WallEntry arrays
'
' The JSON format is a simple structure that both sides agree on.
'======================================================================
Option Explicit

Private Const BRIDGE_SCRIPT As String = "ezdxf_bridge\ezdxf_bridge.py"
Private Const TIMEOUT_SECONDS As Long = 120   ' max wait for Python


'----------------------------------------------------------------------
' CallEzdxfBridge — Execute the Python ezdxf bridge script
'----------------------------------------------------------------------
' Returns True if the script ran successfully and output file was created.
'----------------------------------------------------------------------
Public Function CallEzdxfBridge(ByVal dxfPath As String, _
                                 ByVal outputPath As String) As Boolean
    On Error GoTo BridgeError
    
    ' ── Locate the bridge script ──────────────────────────────
    Dim scriptPath As String
    scriptPath = ThisWorkbook.Path & "\" & BRIDGE_SCRIPT
    
    ' Also check relative to the workbook's parent (for dev layout)
    If Dir(scriptPath) = "" Then
        scriptPath = ThisWorkbook.Path & "\..\" & BRIDGE_SCRIPT
    End If
    If Dir(scriptPath) = "" Then
        MsgBox "Cannot find ezdxf bridge script:" & vbCrLf & _
               BRIDGE_SCRIPT & vbCrLf & vbCrLf & _
               "Expected at: " & ThisWorkbook.Path & "\" & BRIDGE_SCRIPT, _
               vbCritical, "Bridge Script Not Found"
        CallEzdxfBridge = False
        Exit Function
    End If
    
    ' ── Build command ─────────────────────────────────────────
    Dim cmd As String
    cmd = "python """ & scriptPath & """ """ & dxfPath & """ """ & outputPath & """"
    
    ' ── Execute via WScript.Shell ─────────────────────────────
    Dim objShell As Object
    Dim objExec As Object
    Set objShell = CreateObject("WScript.Shell")
    
    Application.StatusBar = "Running ezdxf bridge (Python)..."
    DoEvents
    
    Set objExec = objShell.Exec(cmd)
    
    ' Wait for completion with timeout
    Dim startWait As Double: startWait = Timer
    Do While objExec.Status = 0
        DoEvents
        If Timer - startWait > TIMEOUT_SECONDS Then
            objExec.Terminate
            MsgBox "Python script timed out after " & TIMEOUT_SECONDS & " seconds.", _
                   vbExclamation, "ezdxf Bridge Timeout"
            CallEzdxfBridge = False
            Exit Function
        End If
        ' Brief pause to prevent CPU spin
        Application.Wait Now + TimeSerial(0, 0, 0) + 0.05
    Loop
    
    ' ── Check results ─────────────────────────────────────────
    If objExec.ExitCode <> 0 Then
        Dim errOutput As String
        errOutput = objExec.StdErr.ReadAll
        If Len(errOutput) = 0 Then errOutput = objExec.StdOut.ReadAll
        MsgBox "Python ezdxf bridge failed:" & vbCrLf & vbCrLf & _
               Left(errOutput, 500), vbCritical, "ezdxf Bridge Error"
        CallEzdxfBridge = False
        Exit Function
    End If
    
    ' Verify output file exists
    If Dir(outputPath) = "" Then
        MsgBox "Python script completed but output file was not created:" & vbCrLf & _
               outputPath, vbCritical, "ezdxf Bridge Error"
        CallEzdxfBridge = False
        Exit Function
    End If
    
    CallEzdxfBridge = True
    Exit Function
    
BridgeError:
    MsgBox "Error calling ezdxf bridge:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure Python is installed and in your PATH.", _
           vbCritical, "ezdxf Bridge Error"
    CallEzdxfBridge = False
End Function


'----------------------------------------------------------------------
' ReadBridgeResults — Parse the JSON output from the Python bridge
'----------------------------------------------------------------------
' Reads a simplified JSON structure and populates SlabEntry/WallEntry arrays.
'
' Expected JSON format:
' {
'   "slabs": [
'     {"name":"S1-150","layer":"STR-SLAB-REG150","thickness_mm":150,
'      "thickness_source":"text_label","label_text":"150 THK",
'      "area_sqm":12.345,"volume_cum":1.852,"centroid_x":8500.0,
'      "centroid_y":6200.0,"vertices":[[x1,y1],[x2,y2],...]}
'   ],
'   "str_walls": [ ... ],
'   "ns_walls": [ ... ],
'   "warnings": ["warning 1", "warning 2"]
' }
'----------------------------------------------------------------------
Public Sub ReadBridgeResults(ByVal jsonPath As String, _
                              ByRef slabs() As SlabEntry, ByRef slabCount As Long, _
                              ByRef strWalls() As WallEntry, ByRef strCount As Long, _
                              ByRef nsWalls() As WallEntry, ByRef nsCount As Long, _
                              ByRef warnings() As String, ByRef warnCount As Long)
    
    ' Read entire JSON file
    Dim f As Integer
    Dim jsonText As String
    Dim line As String
    
    f = FreeFile
    Open jsonPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        jsonText = jsonText & line
    Loop
    Close #f
    
    ' ── Parse slabs ───────────────────────────────────────────
    slabCount = 0
    ReDim slabs(0 To 499)
    Call ParseSlabsFromJSON(jsonText, slabs, slabCount)
    
    ' ── Parse structural walls ────────────────────────────────
    strCount = 0
    ReDim strWalls(0 To 499)
    Call ParseWallsFromJSON(jsonText, "str_walls", strWalls, strCount, True)
    
    ' ── Parse non-structural walls ────────────────────────────
    nsCount = 0
    ReDim nsWalls(0 To 499)
    Call ParseWallsFromJSON(jsonText, "ns_walls", nsWalls, nsCount, False)
    
    ' ── Parse warnings ────────────────────────────────────────
    warnCount = 0
    ReDim warnings(0 To 49)
    Call ParseWarningsFromJSON(jsonText, warnings, warnCount)
    
    ' Trim arrays
    If slabCount > 0 Then ReDim Preserve slabs(0 To slabCount - 1)
    If strCount > 0 Then ReDim Preserve strWalls(0 To strCount - 1)
    If nsCount > 0 Then ReDim Preserve nsWalls(0 To nsCount - 1)
    If warnCount > 0 Then ReDim Preserve warnings(0 To warnCount - 1)
End Sub


'======================================================================
' JSON Parsing Helpers (Minimal — no external JSON library)
'======================================================================
' These are simplified parsers that work with the specific JSON structure
' output by ezdxf_bridge.py. They are NOT general-purpose JSON parsers.

Private Sub ParseSlabsFromJSON(ByVal json As String, _
                                ByRef slabs() As SlabEntry, ByRef slabCount As Long)
    ' Find the "slabs" array
    Dim arrStart As Long
    arrStart = InStr(1, json, """slabs""")
    If arrStart = 0 Then Exit Sub
    
    ' Find the opening bracket
    arrStart = InStr(arrStart, json, "[")
    If arrStart = 0 Then Exit Sub
    
    ' Parse each object in the array
    Dim objStart As Long, objEnd As Long
    Dim objJSON As String
    
    objStart = InStr(arrStart, json, "{")
    Do While objStart > 0
        objEnd = InStr(objStart, json, "}")
        If objEnd = 0 Then Exit Do
        
        objJSON = Mid(json, objStart, objEnd - objStart + 1)
        
        If slabCount > UBound(slabs) Then ReDim Preserve slabs(0 To slabCount * 2)
        
        With slabs(slabCount)
            .Name = ExtractJSONString(objJSON, "name")
            .Layer = ExtractJSONString(objJSON, "layer")
            .ThicknessMM = CLng(ExtractJSONNumber(objJSON, "thickness_mm"))
            .ThicknessSource = ExtractJSONString(objJSON, "thickness_source")
            .LabelText = ExtractJSONString(objJSON, "label_text")
            .AreaSqm = ExtractJSONNumber(objJSON, "area_sqm")
            .VolumeCum = ExtractJSONNumber(objJSON, "volume_cum")
            .CentroidX = ExtractJSONNumber(objJSON, "centroid_x")
            .CentroidY = ExtractJSONNumber(objJSON, "centroid_y")
            .NumVerts = 0  ' Vertices from bridge are optional for visualization
        End With
        slabCount = slabCount + 1
        
        ' Find next object
        objStart = InStr(objEnd + 1, json, "{")
        ' Stop if we've left the slabs array (hit "]")
        Dim bracketEnd As Long
        bracketEnd = InStr(objEnd + 1, json, "]")
        If bracketEnd > 0 And bracketEnd < objStart Then Exit Do
    Loop
End Sub

Private Sub ParseWallsFromJSON(ByVal json As String, ByVal arrayName As String, _
                                ByRef walls() As WallEntry, ByRef wallCount As Long, _
                                ByVal isStructural As Boolean)
    Dim arrStart As Long
    arrStart = InStr(1, json, """" & arrayName & """")
    If arrStart = 0 Then Exit Sub
    
    arrStart = InStr(arrStart, json, "[")
    If arrStart = 0 Then Exit Sub
    
    Dim objStart As Long, objEnd As Long
    Dim objJSON As String
    
    objStart = InStr(arrStart, json, "{")
    Do While objStart > 0
        objEnd = InStr(objStart, json, "}")
        If objEnd = 0 Then Exit Do
        
        objJSON = Mid(json, objStart, objEnd - objStart + 1)
        
        If wallCount > UBound(walls) Then ReDim Preserve walls(0 To wallCount * 2)
        
        With walls(wallCount)
            .Name = ExtractJSONString(objJSON, "name")
            .Layer = ExtractJSONString(objJSON, "layer")
            .ThicknessMM = CLng(ExtractJSONNumber(objJSON, "thickness_mm"))
            .AreaSqm = ExtractJSONNumber(objJSON, "area_sqm")
            .LengthM = ExtractJSONNumber(objJSON, "length_m")
            .CentroidX = ExtractJSONNumber(objJSON, "centroid_x")
            .CentroidY = ExtractJSONNumber(objJSON, "centroid_y")
            .IsStructural = isStructural
            .NumVerts = 0
        End With
        wallCount = wallCount + 1
        
        objStart = InStr(objEnd + 1, json, "{")
        Dim bracketEnd As Long
        bracketEnd = InStr(objEnd + 1, json, "]")
        If bracketEnd > 0 And bracketEnd < objStart Then Exit Do
    Loop
End Sub

Private Sub ParseWarningsFromJSON(ByVal json As String, _
                                   ByRef warnings() As String, ByRef warnCount As Long)
    Dim arrStart As Long
    arrStart = InStr(1, json, """warnings""")
    If arrStart = 0 Then Exit Sub
    
    arrStart = InStr(arrStart, json, "[")
    If arrStart = 0 Then Exit Sub
    
    Dim arrEnd As Long
    arrEnd = InStr(arrStart, json, "]")
    If arrEnd = 0 Then Exit Sub
    
    Dim segment As String
    segment = Mid(json, arrStart + 1, arrEnd - arrStart - 1)
    
    ' Split by comma between quoted strings
    Dim pos As Long, qStart As Long, qEnd As Long
    pos = 1
    Do
        qStart = InStr(pos, segment, """")
        If qStart = 0 Then Exit Do
        qEnd = InStr(qStart + 1, segment, """")
        If qEnd = 0 Then Exit Do
        
        If warnCount > UBound(warnings) Then ReDim Preserve warnings(0 To warnCount * 2)
        warnings(warnCount) = Mid(segment, qStart + 1, qEnd - qStart - 1)
        warnCount = warnCount + 1
        pos = qEnd + 1
    Loop
End Sub

' Extract a string value for a given key from a JSON object string
Private Function ExtractJSONString(ByVal json As String, ByVal key As String) As String
    Dim searchKey As String: searchKey = """" & key & """:"
    Dim pos As Long: pos = InStr(1, json, searchKey)
    If pos = 0 Then
        ExtractJSONString = ""
        Exit Function
    End If
    
    pos = pos + Len(searchKey)
    ' Skip whitespace
    Do While Mid(json, pos, 1) = " ": pos = pos + 1: Loop
    
    If Mid(json, pos, 1) = """" Then
        Dim valEnd As Long
        valEnd = InStr(pos + 1, json, """")
        If valEnd > 0 Then
            ExtractJSONString = Mid(json, pos + 1, valEnd - pos - 1)
        End If
    Else
        ExtractJSONString = ""
    End If
End Function

' Extract a numeric value for a given key from a JSON object string
Private Function ExtractJSONNumber(ByVal json As String, ByVal key As String) As Double
    Dim searchKey As String: searchKey = """" & key & """:"
    Dim pos As Long: pos = InStr(1, json, searchKey)
    If pos = 0 Then
        ExtractJSONNumber = 0#
        Exit Function
    End If
    
    pos = pos + Len(searchKey)
    ' Skip whitespace
    Do While Mid(json, pos, 1) = " ": pos = pos + 1: Loop
    
    ' Read number characters
    Dim numStr As String: numStr = ""
    Dim ch As String
    Do
        ch = Mid(json, pos, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "-" Or ch = "e" Or ch = "E" Or ch = "+" Then
            numStr = numStr & ch
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    
    If Len(numStr) > 0 Then
        ExtractJSONNumber = Val(numStr)
    Else
        ExtractJSONNumber = 0#
    End If
End Function
