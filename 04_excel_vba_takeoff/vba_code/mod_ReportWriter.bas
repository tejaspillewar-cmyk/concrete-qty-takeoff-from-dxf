Attribute VB_Name = "mod_ReportWriter"
'======================================================================
' mod_ReportWriter — Excel Sheet Formatting & Data Writing
'======================================================================
' Populates the report sheets with formatted data tables:
'   - Slab Details:        Every slab with quantities
'   - Slab Summary:        Totals grouped by thickness
'   - Structural Walls:    With interactive height input
'   - Non-Structural Walls: With interactive height input
'======================================================================
Option Explicit

' ── Color constants ──────────────────────────────────────────────────
Private Const CLR_HEADER_BG  As Long = &H64381F   ' #1F3864 (dark blue, BGR)
Private Const CLR_HEADER_FT  As Long = &HFFFFFF   ' white
Private Const CLR_TITLE      As Long = &H64381F   ' dark blue
Private Const CLR_SUBTITLE   As Long = &H666666   ' gray
Private Const CLR_ALT_ROW    As Long = &HFBF7F2   ' #F2F7FB light blue (BGR)
Private Const CLR_TOTAL_BG   As Long = &HF0E4D6   ' #D6E4F0 (BGR)
Private Const CLR_WARNING    As Long = &H0000CC   ' red
Private Const CLR_YELLOW     As Long = &H00FFFF   ' yellow highlight (BGR)


'----------------------------------------------------------------------
' WriteSlabDetails — Sheet: "Slab Details"
'----------------------------------------------------------------------
Public Sub WriteSlabDetails(ByVal ws As Worksheet, _
                             ByRef slabs() As SlabEntry, ByVal slabCount As Long, _
                             ByVal fileName As String, _
                             ByRef warnings() As String, ByVal warnCount As Long, _
                             ByVal matchedByText As Long, ByVal matchedByLayer As Long)
    
    Dim r As Long, c As Long, i As Long
    
    Call ClearSheet(ws)
    
    ' ── Title block ───────────────────────────────────────────
    ws.Range("A1:H1").Merge
    ws.Range("A1").value = "SLAB QUANTITY TAKE-OFF"
    Call StyleTitle ws.Range("A1")
    
    ws.Range("A2:H2").Merge
    ws.Range("A2").value = "File: " & fileName & "  |  Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Call StyleSubtitle ws.Range("A2")
    
    ws.Range("A3:H3").Merge
    ws.Range("A3").value = "Total slabs: " & slabCount & "  |  " & _
                           "Matched by text label: " & matchedByText & "  |  " & _
                           "Matched by layer name: " & matchedByLayer
    Call StyleSubtitle ws.Range("A3")
    
    ' ── Headers ───────────────────────────────────────────────
    Dim headers As Variant
    headers = Array("Slab Name", "Layer", "Thickness (mm)", "Source", _
                    "Label Text", "Area (m" & ChrW(178) & ")", _
                    "Volume (m" & ChrW(179) & ")", "Centroid")
    Dim widths As Variant
    widths = Array(14, 22, 16, 14, 16, 14, 14, 24)
    
    Dim headerRow As Long: headerRow = 5
    For c = 0 To 7
        With ws.Cells(headerRow, c + 1)
            .value = headers(c)
            .Interior.Color = CLR_HEADER_BG
            .Font.Color = CLR_HEADER_FT
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            Call AddThinBorder(.Cells)
        End With
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next c
    
    ' ── Data rows ─────────────────────────────────────────────
    For i = 0 To slabCount - 1
        r = headerRow + 1 + i
        
        ws.Cells(r, 1).value = slabs(i).Name
        ws.Cells(r, 2).value = slabs(i).Layer
        ws.Cells(r, 3).value = slabs(i).ThicknessMM
        ws.Cells(r, 4).value = slabs(i).ThicknessSource
        ws.Cells(r, 5).value = slabs(i).LabelText
        ws.Cells(r, 6).value = slabs(i).AreaSqm
        ws.Cells(r, 7).value = slabs(i).VolumeCum
        ws.Cells(r, 8).value = "(" & Format(slabs(i).CentroidX, "0") & ", " & _
                                      Format(slabs(i).CentroidY, "0") & ")"
        
        ' Format cells
        For c = 1 To 8
            With ws.Cells(r, c)
                .Font.Name = "Calibri"
                .Font.Size = 10
                Call AddThinBorder(.Cells)
                If i Mod 2 = 0 Then .Interior.Color = CLR_ALT_ROW
            End With
        Next c
        
        ws.Cells(r, 3).HorizontalAlignment = xlCenter
        ws.Cells(r, 4).HorizontalAlignment = xlCenter
        ws.Cells(r, 6).NumberFormat = "#,##0.000"
        ws.Cells(r, 6).HorizontalAlignment = xlRight
        ws.Cells(r, 7).NumberFormat = "#,##0.0000"
        ws.Cells(r, 7).HorizontalAlignment = xlRight
    Next i
    
    ' ── Grand total row ───────────────────────────────────────
    Dim totalRow As Long: totalRow = headerRow + 1 + slabCount
    
    ws.Cells(totalRow, 1).value = "GRAND TOTAL"
    ws.Cells(totalRow, 5).value = slabCount & " slabs"
    ws.Cells(totalRow, 5).HorizontalAlignment = xlRight
    
    ' Sum formulas for area and volume
    Dim areaSum As Double, volSum As Double
    areaSum = 0: volSum = 0
    For i = 0 To slabCount - 1
        areaSum = areaSum + slabs(i).AreaSqm
        volSum = volSum + slabs(i).VolumeCum
    Next i
    ws.Cells(totalRow, 6).value = Round(areaSum, 3)
    ws.Cells(totalRow, 6).NumberFormat = "#,##0.000"
    ws.Cells(totalRow, 7).value = Round(volSum, 4)
    ws.Cells(totalRow, 7).NumberFormat = "#,##0.0000"
    
    For c = 1 To 8
        With ws.Cells(totalRow, c)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Interior.Color = CLR_TOTAL_BG
            Call AddThinBorder(.Cells)
        End With
    Next c
    
    ' ── Warnings ──────────────────────────────────────────────
    If warnCount > 0 Then
        Dim warnStart As Long: warnStart = totalRow + 2
        ws.Cells(warnStart, 1).value = "WARNINGS:"
        ws.Cells(warnStart, 1).Font.Bold = True
        ws.Cells(warnStart, 1).Font.Color = CLR_WARNING
        For i = 0 To warnCount - 1
            ws.Cells(warnStart + 1 + i, 1).value = "  - " & warnings(i)
            ws.Cells(warnStart + 1 + i, 1).Font.Color = CLR_WARNING
        Next i
    End If
    
    ws.Rows(headerRow + 1 & ":" & headerRow + 1).Select
    ActiveWindow.FreezePanes = True
End Sub


'----------------------------------------------------------------------
' WriteSlabSummary — Sheet: "Slab Summary"
'----------------------------------------------------------------------
Public Sub WriteSlabSummary(ByVal ws As Worksheet, _
                             ByRef slabs() As SlabEntry, ByVal slabCount As Long, _
                             ByVal fileName As String)
    
    Dim r As Long, c As Long, i As Long
    
    Call ClearSheet(ws)
    
    ' ── Build summary dictionary ──────────────────────────────
    Dim summDict As Object
    Set summDict = CreateObject("Scripting.Dictionary")
    
    For i = 0 To slabCount - 1
        Dim tk As String: tk = CStr(slabs(i).ThicknessMM)
        If Not summDict.Exists(tk) Then
            summDict.Add tk, Array(CLng(0), 0#, 0#)   ' count, area, volume
        End If
        Dim arr As Variant: arr = summDict(tk)
        arr(0) = arr(0) + 1
        arr(1) = arr(1) + slabs(i).AreaSqm
        arr(2) = arr(2) + slabs(i).VolumeCum
        summDict(tk) = arr
    Next i
    
    ' Sort keys
    Dim keys As Variant: keys = summDict.keys
    Call SortLongKeys(keys)
    
    ' ── Title ─────────────────────────────────────────────────
    ws.Range("A1:D1").Merge
    ws.Range("A1").value = "SLAB QUANTITY SUMMARY"
    Call StyleTitle ws.Range("A1")
    
    ws.Range("A2:D2").Merge
    ws.Range("A2").value = "File: " & fileName & "  |  Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Call StyleSubtitle ws.Range("A2")
    
    ' ── Headers ───────────────────────────────────────────────
    Dim headers As Variant
    headers = Array("Slab Thickness", "Count", _
                    "Total Area (m" & ChrW(178) & ")", _
                    "Total Volume (m" & ChrW(179) & ")")
    Dim widths As Variant
    widths = Array(20, 12, 18, 18)
    
    Dim headerRow As Long: headerRow = 4
    For c = 0 To 3
        With ws.Cells(headerRow, c + 1)
            .value = headers(c)
            .Interior.Color = CLR_HEADER_BG
            .Font.Color = CLR_HEADER_FT
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
            Call AddThinBorder(.Cells)
        End With
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next c
    
    ' ── Data rows ─────────────────────────────────────────────
    Dim totalCount As Long: totalCount = 0
    Dim totalArea As Double: totalArea = 0
    Dim totalVol As Double: totalVol = 0
    
    For i = 0 To UBound(keys)
        r = headerRow + 1 + i
        arr = summDict(CStr(keys(i)))
        
        ws.Cells(r, 1).value = keys(i) & " mm"
        ws.Cells(r, 2).value = arr(0)
        ws.Cells(r, 3).value = Round(CDbl(arr(1)), 3)
        ws.Cells(r, 4).value = Round(CDbl(arr(2)), 4)
        
        totalCount = totalCount + CLng(arr(0))
        totalArea = totalArea + CDbl(arr(1))
        totalVol = totalVol + CDbl(arr(2))
        
        For c = 1 To 4
            With ws.Cells(r, c)
                .Font.Name = "Calibri"
                .Font.Size = 10
                Call AddThinBorder(.Cells)
                If i Mod 2 = 0 Then .Interior.Color = CLR_ALT_ROW
            End With
        Next c
        ws.Cells(r, 2).HorizontalAlignment = xlCenter
        ws.Cells(r, 3).NumberFormat = "#,##0.000"
        ws.Cells(r, 3).HorizontalAlignment = xlRight
        ws.Cells(r, 4).NumberFormat = "#,##0.0000"
        ws.Cells(r, 4).HorizontalAlignment = xlRight
    Next i
    
    ' ── Grand total ───────────────────────────────────────────
    Dim totalRow As Long: totalRow = headerRow + 1 + UBound(keys) + 1
    ws.Cells(totalRow, 1).value = "GRAND TOTAL"
    ws.Cells(totalRow, 2).value = totalCount
    ws.Cells(totalRow, 2).HorizontalAlignment = xlCenter
    ws.Cells(totalRow, 3).value = Round(totalArea, 3)
    ws.Cells(totalRow, 3).NumberFormat = "#,##0.000"
    ws.Cells(totalRow, 3).HorizontalAlignment = xlRight
    ws.Cells(totalRow, 4).value = Round(totalVol, 4)
    ws.Cells(totalRow, 4).NumberFormat = "#,##0.0000"
    ws.Cells(totalRow, 4).HorizontalAlignment = xlRight
    
    For c = 1 To 4
        With ws.Cells(totalRow, c)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Interior.Color = CLR_TOTAL_BG
            Call AddThinBorder(.Cells)
        End With
    Next c
    
    ws.Rows(headerRow + 1 & ":" & headerRow + 1).Select
    ActiveWindow.FreezePanes = True
End Sub


'----------------------------------------------------------------------
' WriteWallSheet — Sheet: "Structural Walls" or "Non-Structural Walls"
'----------------------------------------------------------------------
Public Sub WriteWallSheet(ByVal ws As Worksheet, _
                           ByRef walls() As WallEntry, ByVal wallCount As Long, _
                           ByVal isStructural As Boolean, _
                           ByVal floorHeight As Double)
    
    Dim r As Long, c As Long, i As Long
    
    Call ClearSheet(ws)
    
    Dim sheetTitle As String
    If isStructural Then
        sheetTitle = "STRUCTURAL WALLS QUANTITY TAKE-OFF"
    Else
        sheetTitle = "NON-STRUCTURAL WALLS QUANTITY TAKE-OFF"
    End If
    
    ' ── Title ─────────────────────────────────────────────────
    ws.Range("A1:H1").Merge
    ws.Range("A1").value = sheetTitle
    Call StyleTitle ws.Range("A1")
    
    ' ── Interactive Height Input ──────────────────────────────
    ws.Range("A3").value = "Floor-to-Floor Height (m):"
    ws.Range("A3").Font.Bold = True
    ws.Range("A3").Font.Name = "Calibri"
    ws.Range("A3").Font.Size = 11
    ws.Range("A3").Font.Color = CLR_TITLE
    
    With ws.Range("B3")
        .value = floorHeight
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Interior.Color = CLR_YELLOW
        .HorizontalAlignment = xlCenter
        Call AddThinBorder(.Cells)
    End With
    
    ' ── Headers ───────────────────────────────────────────────
    Dim headers As Variant
    headers = Array("Wall Name", "Layer", "Thickness (mm)", _
                    "Length (m)", "Area (m" & ChrW(178) & ")", _
                    "Volume (m" & ChrW(179) & ")", "Centroid")
    Dim widths As Variant
    widths = Array(14, 22, 16, 14, 14, 14, 24)
    
    Dim headerRow As Long: headerRow = 5
    For c = 0 To 6
        With ws.Cells(headerRow, c + 1)
            .value = headers(c)
            .Interior.Color = CLR_HEADER_BG
            .Font.Color = CLR_HEADER_FT
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
            Call AddThinBorder(.Cells)
        End With
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next c
    
    ' ── Data rows ─────────────────────────────────────────────
    For i = 0 To wallCount - 1
        r = headerRow + 1 + i
        
        ws.Cells(r, 1).value = walls(i).Name
        ws.Cells(r, 2).value = walls(i).Layer
        ws.Cells(r, 3).value = walls(i).ThicknessMM
        ws.Cells(r, 4).value = walls(i).LengthM
        ws.Cells(r, 5).value = walls(i).AreaSqm
        ' Volume = Area * Height (formula referencing B3)
        ws.Cells(r, 6).Formula = "=E" & r & "*B$3"
        ws.Cells(r, 7).value = "(" & Format(walls(i).CentroidX, "0") & ", " & _
                                      Format(walls(i).CentroidY, "0") & ")"
        
        For c = 1 To 7
            With ws.Cells(r, c)
                .Font.Name = "Calibri"
                .Font.Size = 10
                Call AddThinBorder(.Cells)
                If i Mod 2 = 0 Then .Interior.Color = CLR_ALT_ROW
            End With
        Next c
        
        ws.Cells(r, 3).HorizontalAlignment = xlCenter
        ws.Cells(r, 4).NumberFormat = "#,##0.000"
        ws.Cells(r, 4).HorizontalAlignment = xlRight
        ws.Cells(r, 5).NumberFormat = "#,##0.000"
        ws.Cells(r, 5).HorizontalAlignment = xlRight
        ws.Cells(r, 6).NumberFormat = "#,##0.000"
        ws.Cells(r, 6).HorizontalAlignment = xlRight
    Next i
    
    ' ── Grand total row ───────────────────────────────────────
    Dim totalRow As Long: totalRow = headerRow + 1 + wallCount
    
    ws.Cells(totalRow, 1).value = "GRAND TOTAL"
    
    ' SUM formulas
    If wallCount > 0 Then
        ws.Cells(totalRow, 4).Formula = "=SUM(D" & (headerRow + 1) & ":D" & (totalRow - 1) & ")"
        ws.Cells(totalRow, 5).Formula = "=SUM(E" & (headerRow + 1) & ":E" & (totalRow - 1) & ")"
        ws.Cells(totalRow, 6).Formula = "=SUM(F" & (headerRow + 1) & ":F" & (totalRow - 1) & ")"
    End If
    
    For c = 1 To 7
        With ws.Cells(totalRow, c)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Interior.Color = CLR_TOTAL_BG
            Call AddThinBorder(.Cells)
        End With
    Next c
    ws.Cells(totalRow, 4).NumberFormat = "#,##0.000"
    ws.Cells(totalRow, 5).NumberFormat = "#,##0.000"
    ws.Cells(totalRow, 6).NumberFormat = "#,##0.000"
    
    ws.Rows(headerRow + 1 & ":" & headerRow + 1).Select
    ActiveWindow.FreezePanes = True
End Sub


'======================================================================
' Helper Subs
'======================================================================

Private Sub ClearSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    ' Remove all non-button shapes
    Dim shp As Shape
    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.Type <> 8 Then shp.Delete   ' 8 = msoFormControl
    Next shp
    On Error GoTo 0
    ActiveWindow.FreezePanes = False
End Sub

Private Sub StyleTitle(ByVal rng As Range)
    With rng
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = CLR_TITLE
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub StyleSubtitle(ByVal rng As Range)
    With rng
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = CLR_SUBTITLE
    End With
End Sub

Private Sub AddThinBorder(ByVal rng As Range)
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
        With rng.Borders(edge)
            .LineStyle = xlContinuous
            .Color = &HCCCCCC
            .Weight = xlThin
        End With
    Next edge
End Sub

' Simple insertion sort for an array of numeric-string keys
Private Sub SortLongKeys(ByRef keys As Variant)
    Dim i As Long, j As Long
    Dim tmp As Variant
    For i = 1 To UBound(keys)
        tmp = keys(i)
        j = i - 1
        Do While j >= 0
            If CLng(keys(j)) > CLng(tmp) Then
                keys(j + 1) = keys(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        keys(j + 1) = tmp
    Next i
End Sub
