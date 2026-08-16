Attribute VB_Name = "mod_SlabExtractor"
'======================================================================
' mod_SlabExtractor — Slab Extraction Logic
'======================================================================
' Mirrors the Python slab_extractor.py logic:
'   1. Filters polylines by STR-SLAB-REG* layer prefix (closed only)
'   2. Filters text labels on STR-TYP-TXTNUM layer containing "THK"
'   3. Matches each slab polygon to its thickness label (point-in-polygon)
'   4. Falls back to layer-name thickness if no label found inside
'   5. Cross-checks text vs layer thickness (generates warnings)
'   6. Computes area (m²) and volume (m³)
'   7. Auto-names slabs: S1-150, S2-150, S1-200, etc.
'======================================================================
Option Explicit

Private Const SLAB_LAYER_PREFIX As String = "STR-SLAB-REG"
Private Const SLAB_LABEL_LAYER As String = "STR-TYP-TXTNUM"


'----------------------------------------------------------------------
' ExtractSlabs — Main slab extraction
'----------------------------------------------------------------------
Public Sub ExtractSlabs(ByRef polys() As DXFPolyline, ByVal polyCount As Long, _
                        ByRef texts() As DXFText, ByVal textCount As Long, _
                        ByRef slabs() As SlabEntry, ByRef slabCount As Long, _
                        ByRef warnings() As String, ByRef warnCount As Long)
    
    Dim i As Long, j As Long, k As Long
    Dim thickFromLayer As Long
    Dim suffix As String
    Dim matched As Boolean
    
    ' ── Initialize outputs ────────────────────────────────────
    slabCount = 0
    warnCount = 0
    ReDim slabs(0 To 499)
    ReDim warnings(0 To 99)
    
    ' ── Thickness counters for auto-naming ────────────────────
    Dim thickCounters As Object
    Set thickCounters = CreateObject("Scripting.Dictionary")
    
    ' ── Pre-filter THK labels ─────────────────────────────────
    ' Build an index of text entities that are on the label layer and contain "THK"
    Dim thkIndices() As Long
    Dim thkCount As Long
    thkCount = 0
    ReDim thkIndices(0 To 499)
    
    For i = 0 To textCount - 1
        If StrComp(texts(i).Layer, SLAB_LABEL_LAYER, vbTextCompare) = 0 Then
            If InStr(1, UCase(texts(i).Content), "THK") > 0 Then
                If thkCount > UBound(thkIndices) Then
                    ReDim Preserve thkIndices(0 To thkCount * 2)
                End If
                thkIndices(thkCount) = i
                thkCount = thkCount + 1
            End If
        End If
    Next i
    
    ' Track which labels have been used (avoid double-matching)
    Dim usedLabels() As Boolean
    If textCount > 0 Then
        ReDim usedLabels(0 To textCount - 1)
        For i = 0 To textCount - 1
            usedLabels(i) = False
        Next i
    Else
        ReDim usedLabels(0 To 0)
    End If
    
    ' ── Process each polyline ─────────────────────────────────
    For i = 0 To polyCount - 1
        ' Filter: must be on a slab layer and closed
        If Left(polys(i).Layer, Len(SLAB_LAYER_PREFIX)) <> SLAB_LAYER_PREFIX Then GoTo NextPoly
        If Not polys(i).IsClosed Then GoTo NextPoly
        If polys(i).NumVerts < 3 Then GoTo NextPoly
        
        ' Parse thickness from layer name suffix
        suffix = Mid(polys(i).Layer, Len(SLAB_LAYER_PREFIX) + 1)
        thickFromLayer = 0
        If Len(suffix) > 0 And IsNumeric(suffix) Then
            thickFromLayer = CLng(suffix)
        End If
        
        ' ── Try to match a THK text label inside this polygon ─
        Dim thickMM As Long
        Dim thickSource As String
        Dim labelText As String
        thickMM = 0: thickSource = "": labelText = ""
        matched = False
        
        For j = 0 To thkCount - 1
            k = thkIndices(j)
            If usedLabels(k) Then GoTo NextLabel
            
            ' Is this label inside the slab polygon?
            If PointInPolygon(texts(k).X, texts(k).Y, _
                              polys(i).X, polys(i).Y, polys(i).NumVerts) Then
                
                thickMM = ParseThicknessFromText(texts(k).Content)
                If thickMM > 0 Then
                    thickSource = "text_label"
                    labelText = texts(k).Content
                    usedLabels(k) = True
                    matched = True
                    Exit For
                End If
            End If
NextLabel:
        Next j
        
        ' ── Fallback: use layer name thickness ────────────────
        If Not matched Then
            If thickFromLayer > 0 Then
                thickMM = thickFromLayer
                thickSource = "layer_name"
                labelText = "(from layer: " & polys(i).Layer & ")"
            Else
                thickMM = 0
                thickSource = "unknown"
                labelText = "UNKNOWN"
                Call AddWarning(warnings, warnCount, _
                    "Slab on layer '" & polys(i).Layer & "': " & _
                    "no THK label found and layer name has no thickness.")
            End If
        End If
        
        ' ── Cross-check: text vs layer thickness ──────────────
        If thickSource = "text_label" And thickFromLayer > 0 Then
            If thickMM <> thickFromLayer Then
                Call AddWarning(warnings, warnCount, _
                    "Slab on layer '" & polys(i).Layer & "': " & _
                    "text says " & thickMM & "mm but layer says " & _
                    thickFromLayer & "mm. Using text label value.")
            End If
        End If
        
        ' ── Compute quantities ────────────────────────────────
        Dim areaSqmm As Double, areaSqm As Double, volumeCum As Double
        Dim cx As Double, cy As Double
        
        areaSqmm = ShoelaceArea(polys(i).X, polys(i).Y, polys(i).NumVerts)
        areaSqm = SqmmToSqm(areaSqmm)
        volumeCum = areaSqm * (thickMM / 1000#)
        Call PolygonCentroid(polys(i).X, polys(i).Y, polys(i).NumVerts, cx, cy)
        
        ' ── Auto-name: S1-150, S2-150, etc. ──────────────────
        Dim counter As Long
        Dim thickKey As String
        thickKey = CStr(thickMM)
        If Not thickCounters.Exists(thickKey) Then
            thickCounters.Add thickKey, CLng(0)
        End If
        thickCounters(thickKey) = thickCounters(thickKey) + 1
        counter = thickCounters(thickKey)
        
        ' ── Save slab entry ───────────────────────────────────
        If slabCount > UBound(slabs) Then
            ReDim Preserve slabs(0 To slabCount * 2)
        End If
        
        With slabs(slabCount)
            .Name = "S" & counter & "-" & thickMM
            .Layer = polys(i).Layer
            .ThicknessMM = thickMM
            .ThicknessSource = thickSource
            .LabelText = labelText
            .AreaSqm = Round(areaSqm, 3)
            .VolumeCum = Round(volumeCum, 4)
            .CentroidX = Round(cx, 1)
            .CentroidY = Round(cy, 1)
            .NumVerts = polys(i).NumVerts
            ReDim .X(0 To polys(i).NumVerts - 1)
            ReDim .Y(0 To polys(i).NumVerts - 1)
            For k = 0 To polys(i).NumVerts - 1
                .X(k) = polys(i).X(k)
                .Y(k) = polys(i).Y(k)
            Next k
        End With
        slabCount = slabCount + 1
        
NextPoly:
    Next i
    
    ' Trim outputs
    If slabCount > 0 Then
        ReDim Preserve slabs(0 To slabCount - 1)
    Else
        ReDim slabs(0 To 0)
    End If
    If warnCount > 0 Then
        ReDim Preserve warnings(0 To warnCount - 1)
    Else
        ReDim warnings(0 To 0)
    End If
End Sub


'----------------------------------------------------------------------
' ParseThicknessFromText — Extract numeric thickness from text like "150 THK"
'----------------------------------------------------------------------
' Handles formats: "150 THK", "150THK", "( 125MM THK- ...)", "(150 MM THK)"
' Returns 0 if no valid thickness found.
'----------------------------------------------------------------------
Private Function ParseThicknessFromText(ByVal txt As String) As Long
    Dim s As String
    Dim thkPos As Long
    Dim numStr As String
    Dim i As Long
    Dim ch As String
    
    s = UCase(Trim(txt))
    thkPos = InStr(1, s, "THK")
    If thkPos = 0 Then
        ParseThicknessFromText = 0
        Exit Function
    End If
    
    ' Walk backwards from "THK" to find digits
    numStr = ""
    For i = thkPos - 1 To 1 Step -1
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            numStr = ch & numStr
        Else
            ' Allow skipping spaces and "M" (from "MM")
            If Len(numStr) > 0 Then Exit For
            If ch <> " " And ch <> "M" Then
                Exit For
            End If
        End If
    Next i
    
    If Len(numStr) > 0 And IsNumeric(numStr) Then
        ParseThicknessFromText = CLng(numStr)
    Else
        ParseThicknessFromText = 0
    End If
End Function


'----------------------------------------------------------------------
' AddWarning — Append a warning string to the warnings array
'----------------------------------------------------------------------
Private Sub AddWarning(ByRef warnings() As String, ByRef warnCount As Long, _
                        ByVal msg As String)
    If warnCount > UBound(warnings) Then
        ReDim Preserve warnings(0 To warnCount * 2)
    End If
    warnings(warnCount) = msg
    warnCount = warnCount + 1
End Sub
