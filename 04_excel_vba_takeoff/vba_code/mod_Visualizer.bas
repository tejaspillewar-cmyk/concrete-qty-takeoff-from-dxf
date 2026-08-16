Attribute VB_Name = "mod_Visualizer"
'======================================================================
' mod_Visualizer — Excel Shape Drawing for Slab & Wall Maps
'======================================================================
' Draws polygon shapes directly on Excel worksheets:
'   - Color-coded by thickness
'   - Labels at centroids
'   - Dark background with legend
'
' Uses Shapes.BuildFreeform for arbitrary polygon drawing.
' Coordinates are transformed from DXF mm → Excel points.
'======================================================================
Option Explicit

' ── Drawing area constants (in Excel points) ─────────────────────────
Private Const DRAW_LEFT   As Double = 60
Private Const DRAW_TOP    As Double = 60
Private Const DRAW_WIDTH  As Double = 700
Private Const DRAW_HEIGHT As Double = 500

' ── Background / label colors ────────────────────────────────────────
Private Const BG_COLOR    As Long = &H2E1A1A     ' #1A1A2E (BGR)
Private Const LABEL_WHITE As Long = &HFFFFFF


'----------------------------------------------------------------------
' DrawSlabMap — Render slab polygons on the "Slab Map" sheet
'----------------------------------------------------------------------
Public Sub DrawSlabMap(ByVal ws As Worksheet, _
                       ByRef slabs() As SlabEntry, ByVal slabCount As Long)
    
    If slabCount = 0 Then
        ws.Range("A1").value = "No slabs found to visualize."
        Exit Sub
    End If
    
    Call ClearShapes(ws)
    
    ' ── Compute bounds ────────────────────────────────────────
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    Call ComputeSlabBounds(slabs, slabCount, minX, minY, maxX, maxY)
    
    Dim rangeX As Double: rangeX = maxX - minX
    Dim rangeY As Double: rangeY = maxY - minY
    If rangeX < 1 Then rangeX = 1
    If rangeY < 1 Then rangeY = 1
    
    ' Equal aspect ratio
    Dim scaleX As Double: scaleX = DRAW_WIDTH / rangeX
    Dim scaleY As Double: scaleY = DRAW_HEIGHT / rangeY
    Dim scale As Double
    If scaleX < scaleY Then scale = scaleX Else scale = scaleY
    
    ' Center offset
    Dim offsetX As Double: offsetX = DRAW_LEFT + (DRAW_WIDTH - rangeX * scale) / 2
    Dim offsetY As Double: offsetY = DRAW_TOP + (DRAW_HEIGHT - rangeY * scale) / 2
    
    ' ── Background rectangle ─────────────────────────────────
    Dim bg As Shape
    Set bg = ws.Shapes.AddShape(1, DRAW_LEFT - 15, DRAW_TOP - 15, _
                                 DRAW_WIDTH + 30, DRAW_HEIGHT + 30)   ' 1 = msoShapeRectangle
    bg.Fill.ForeColor.RGB = BG_COLOR
    bg.Line.Visible = msoFalse
    bg.Name = "MapBG"
    
    ' ── Draw each slab polygon ────────────────────────────────
    Dim i As Long, j As Long
    For i = 0 To slabCount - 1
        If slabs(i).NumVerts < 3 Then GoTo NextSlab
        
        Dim fillClr As Long
        fillClr = ThicknessColor(slabs(i).ThicknessMM)
        
        ' Map first vertex
        Dim sx As Single, sy As Single
        sx = CSng(offsetX + (slabs(i).X(0) - minX) * scale)
        sy = CSng(offsetY + (maxY - slabs(i).Y(0)) * scale)  ' Y inverted
        
        ' Build freeform
        Dim ff As FreeformBuilder
        Set ff = ws.Shapes.BuildFreeform(0, sx, sy)  ' 0 = msoEditingCorner
        
        For j = 1 To slabs(i).NumVerts - 1
            sx = CSng(offsetX + (slabs(i).X(j) - minX) * scale)
            sy = CSng(offsetY + (maxY - slabs(i).Y(j)) * scale)
            ff.AddNodes 0, 1, sx, sy   ' 0 = msoSegmentLine, 1 = msoEditingAuto
        Next j
        
        ' Close polygon
        sx = CSng(offsetX + (slabs(i).X(0) - minX) * scale)
        sy = CSng(offsetY + (maxY - slabs(i).Y(0)) * scale)
        ff.AddNodes 0, 1, sx, sy
        
        Dim poly As Shape
        Set poly = ff.ConvertToShape
        poly.Fill.ForeColor.RGB = fillClr
        poly.Fill.Transparency = 0.5
        poly.Line.ForeColor.RGB = LABEL_WHITE
        poly.Line.Weight = 0.8
        poly.Name = "Slab_" & slabs(i).Name
        
        ' ── Centroid label ────────────────────────────────────
        Dim lcx As Single, lcy As Single
        lcx = CSng(offsetX + (slabs(i).CentroidX - minX) * scale)
        lcy = CSng(offsetY + (maxY - slabs(i).CentroidY) * scale)
        
        Dim lbl As Shape
        Set lbl = ws.Shapes.AddTextbox(1, lcx - 28, lcy - 10, 56, 20)  ' 1 = horizontal
        lbl.TextFrame2.TextRange.Text = slabs(i).Name & vbLf & _
                                         Format(slabs(i).AreaSqm, "0.00") & " m" & ChrW(178)
        lbl.TextFrame2.TextRange.Font.Size = 6
        lbl.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = LABEL_WHITE
        lbl.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  ' msoAlignCenter
        lbl.Fill.ForeColor.RGB = fillClr
        lbl.Fill.Transparency = 0.3
        lbl.Line.Visible = msoFalse
        lbl.Name = "LblSlab_" & slabs(i).Name
        
NextSlab:
    Next i
    
    ' ── Title ─────────────────────────────────────────────────
    Call AddMapTitle ws, "SLAB QUANTITY TAKE-OFF MAP"
    
    ' ── Legend ────────────────────────────────────────────────
    Call DrawSlabLegend ws
End Sub


'----------------------------------------------------------------------
' DrawWallMap — Render wall polygons on the "Wall Map" sheet
'----------------------------------------------------------------------
Public Sub DrawWallMap(ByVal ws As Worksheet, _
                       ByRef strWalls() As WallEntry, ByVal strCount As Long, _
                       ByRef nsWalls() As WallEntry, ByVal nsCount As Long)
    
    Dim totalCount As Long: totalCount = strCount + nsCount
    If totalCount = 0 Then
        ws.Range("A1").value = "No walls found to visualize."
        Exit Sub
    End If
    
    Call ClearShapes(ws)
    
    ' ── Compute bounds across both wall sets ──────────────────
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double
    Call ComputeWallBounds(strWalls, strCount, nsWalls, nsCount, minX, minY, maxX, maxY)
    
    Dim rangeX As Double: rangeX = maxX - minX
    Dim rangeY As Double: rangeY = maxY - minY
    If rangeX < 1 Then rangeX = 1
    If rangeY < 1 Then rangeY = 1
    
    Dim scaleX As Double: scaleX = DRAW_WIDTH / rangeX
    Dim scaleY As Double: scaleY = DRAW_HEIGHT / rangeY
    Dim scale As Double
    If scaleX < scaleY Then scale = scaleX Else scale = scaleY
    
    Dim offsetX As Double: offsetX = DRAW_LEFT + (DRAW_WIDTH - rangeX * scale) / 2
    Dim offsetY As Double: offsetY = DRAW_TOP + (DRAW_HEIGHT - rangeY * scale) / 2
    
    ' Background
    Dim bg As Shape
    Set bg = ws.Shapes.AddShape(1, DRAW_LEFT - 15, DRAW_TOP - 15, _
                                 DRAW_WIDTH + 30, DRAW_HEIGHT + 30)
    bg.Fill.ForeColor.RGB = BG_COLOR
    bg.Line.Visible = msoFalse
    bg.Name = "WallMapBG"
    
    ' ── Draw structural walls (red/orange) ────────────────────
    Dim strColor As Long: strColor = RGB(255, 69, 0)   ' OrangeRed
    Call DrawWallSet ws, strWalls, strCount, strColor, scale, offsetX, offsetY, minX, maxY
    
    ' ── Draw non-structural walls (blue) ──────────────────────
    Dim nsColor As Long: nsColor = RGB(30, 144, 255)    ' DodgerBlue
    Call DrawWallSet ws, nsWalls, nsCount, nsColor, scale, offsetX, offsetY, minX, maxY
    
    ' Title
    Call AddMapTitle ws, "WALL QUANTITY MAP"
    
    ' Legend
    Dim legX As Single: legX = CSng(DRAW_LEFT + DRAW_WIDTH - 140)
    Dim legY As Single: legY = CSng(DRAW_TOP + 5)
    
    Dim legBg As Shape
    Set legBg = ws.Shapes.AddShape(1, legX - 5, legY - 5, 150, 50)
    legBg.Fill.ForeColor.RGB = RGB(22, 33, 62)
    legBg.Fill.Transparency = 0.1
    legBg.Line.Visible = msoFalse
    legBg.Name = "WallLegBG"
    
    Dim sq1 As Shape
    Set sq1 = ws.Shapes.AddShape(1, legX, legY, 12, 12)
    sq1.Fill.ForeColor.RGB = strColor
    sq1.Line.Visible = msoFalse
    sq1.Name = "WallLegStr"
    
    Dim t1 As Shape
    Set t1 = ws.Shapes.AddTextbox(1, legX + 16, legY - 2, 120, 16)
    t1.TextFrame2.TextRange.Text = "Structural (" & strCount & ")"
    t1.TextFrame2.TextRange.Font.Size = 8
    t1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = LABEL_WHITE
    t1.Fill.Visible = msoFalse
    t1.Line.Visible = msoFalse
    t1.Name = "WallLegStrTxt"
    
    Dim sq2 As Shape
    Set sq2 = ws.Shapes.AddShape(1, legX, legY + 20, 12, 12)
    sq2.Fill.ForeColor.RGB = nsColor
    sq2.Line.Visible = msoFalse
    sq2.Name = "WallLegNS"
    
    Dim t2 As Shape
    Set t2 = ws.Shapes.AddTextbox(1, legX + 16, legY + 18, 120, 16)
    t2.TextFrame2.TextRange.Text = "Non-Structural (" & nsCount & ")"
    t2.TextFrame2.TextRange.Font.Size = 8
    t2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = LABEL_WHITE
    t2.Fill.Visible = msoFalse
    t2.Line.Visible = msoFalse
    t2.Name = "WallLegNSTxt"
End Sub


'======================================================================
' Private Helpers
'======================================================================

Private Sub DrawWallSet(ByVal ws As Worksheet, _
                         ByRef walls() As WallEntry, ByVal wallCount As Long, _
                         ByVal fillClr As Long, _
                         ByVal scale As Double, _
                         ByVal offsetX As Double, ByVal offsetY As Double, _
                         ByVal minX As Double, ByVal maxY As Double)
    Dim i As Long, j As Long
    For i = 0 To wallCount - 1
        If walls(i).NumVerts < 3 Then GoTo NextW
        
        Dim sx As Single, sy As Single
        sx = CSng(offsetX + (walls(i).X(0) - minX) * scale)
        sy = CSng(offsetY + (maxY - walls(i).Y(0)) * scale)
        
        Dim ff As FreeformBuilder
        Set ff = ws.Shapes.BuildFreeform(0, sx, sy)
        
        For j = 1 To walls(i).NumVerts - 1
            sx = CSng(offsetX + (walls(i).X(j) - minX) * scale)
            sy = CSng(offsetY + (maxY - walls(i).Y(j)) * scale)
            ff.AddNodes 0, 1, sx, sy
        Next j
        
        sx = CSng(offsetX + (walls(i).X(0) - minX) * scale)
        sy = CSng(offsetY + (maxY - walls(i).Y(0)) * scale)
        ff.AddNodes 0, 1, sx, sy
        
        Dim poly As Shape
        Set poly = ff.ConvertToShape
        poly.Fill.ForeColor.RGB = fillClr
        poly.Fill.Transparency = 0.45
        poly.Line.ForeColor.RGB = fillClr
        poly.Line.Weight = 1.2
        poly.Name = "Wall_" & walls(i).Name
        
        ' Label
        Dim lcx As Single, lcy As Single
        lcx = CSng(offsetX + (walls(i).CentroidX - minX) * scale)
        lcy = CSng(offsetY + (maxY - walls(i).CentroidY) * scale)
        
        Dim lbl As Shape
        Set lbl = ws.Shapes.AddTextbox(1, lcx - 22, lcy - 7, 44, 14)
        lbl.TextFrame2.TextRange.Text = walls(i).Name
        lbl.TextFrame2.TextRange.Font.Size = 5.5
        lbl.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = LABEL_WHITE
        lbl.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        lbl.Fill.ForeColor.RGB = fillClr
        lbl.Fill.Transparency = 0.3
        lbl.Line.Visible = msoFalse
        lbl.Name = "LblWall_" & walls(i).Name
NextW:
    Next i
End Sub

Private Sub ComputeSlabBounds(ByRef slabs() As SlabEntry, ByVal slabCount As Long, _
                               ByRef minX As Double, ByRef minY As Double, _
                               ByRef maxX As Double, ByRef maxY As Double)
    Dim i As Long, j As Long
    minX = 1E+99: minY = 1E+99: maxX = -1E+99: maxY = -1E+99
    For i = 0 To slabCount - 1
        For j = 0 To slabs(i).NumVerts - 1
            If slabs(i).X(j) < minX Then minX = slabs(i).X(j)
            If slabs(i).Y(j) < minY Then minY = slabs(i).Y(j)
            If slabs(i).X(j) > maxX Then maxX = slabs(i).X(j)
            If slabs(i).Y(j) > maxY Then maxY = slabs(i).Y(j)
        Next j
    Next i
End Sub

Private Sub ComputeWallBounds(ByRef strWalls() As WallEntry, ByVal strCount As Long, _
                               ByRef nsWalls() As WallEntry, ByVal nsCount As Long, _
                               ByRef minX As Double, ByRef minY As Double, _
                               ByRef maxX As Double, ByRef maxY As Double)
    Dim i As Long, j As Long
    minX = 1E+99: minY = 1E+99: maxX = -1E+99: maxY = -1E+99
    
    For i = 0 To strCount - 1
        For j = 0 To strWalls(i).NumVerts - 1
            If strWalls(i).X(j) < minX Then minX = strWalls(i).X(j)
            If strWalls(i).Y(j) < minY Then minY = strWalls(i).Y(j)
            If strWalls(i).X(j) > maxX Then maxX = strWalls(i).X(j)
            If strWalls(i).Y(j) > maxY Then maxY = strWalls(i).Y(j)
        Next j
    Next i
    For i = 0 To nsCount - 1
        For j = 0 To nsWalls(i).NumVerts - 1
            If nsWalls(i).X(j) < minX Then minX = nsWalls(i).X(j)
            If nsWalls(i).Y(j) < minY Then minY = nsWalls(i).Y(j)
            If nsWalls(i).X(j) > maxX Then maxX = nsWalls(i).X(j)
            If nsWalls(i).Y(j) > maxY Then maxY = nsWalls(i).Y(j)
        Next j
    Next i
End Sub

Private Function ThicknessColor(ByVal thickMM As Long) As Long
    Select Case thickMM
        Case 125: ThicknessColor = RGB(0, 191, 255)     ' Cyan
        Case 150: ThicknessColor = RGB(50, 205, 50)      ' Lime green
        Case 175: ThicknessColor = RGB(255, 215, 0)      ' Gold
        Case 200: ThicknessColor = RGB(255, 99, 71)      ' Tomato
        Case Else: ThicknessColor = RGB(170, 170, 170)   ' Gray
    End Select
End Function

Private Sub ClearShapes(ByVal ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.Type <> 8 Then shp.Delete   ' keep form control buttons
    Next shp
    On Error GoTo 0
End Sub

Private Sub AddMapTitle(ByVal ws As Worksheet, ByVal title As String)
    Dim t As Shape
    Set t = ws.Shapes.AddTextbox(1, DRAW_LEFT, DRAW_TOP - 45, 400, 28)
    t.TextFrame2.TextRange.Text = title
    t.TextFrame2.TextRange.Font.Size = 14
    t.TextFrame2.TextRange.Font.Bold = msoTrue
    t.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = LABEL_WHITE
    t.Fill.Visible = msoFalse
    t.Line.Visible = msoFalse
    t.Name = "MapTitle"
End Sub

Private Sub DrawSlabLegend(ByVal ws As Worksheet)
    Dim legX As Single: legX = CSng(DRAW_LEFT + DRAW_WIDTH - 140)
    Dim legY As Single: legY = CSng(DRAW_TOP + 5)
    
    ' Legend background
    Dim legBg As Shape
    Set legBg = ws.Shapes.AddShape(1, legX - 5, legY - 5, 150, 90)
    legBg.Fill.ForeColor.RGB = RGB(22, 33, 62)
    legBg.Fill.Transparency = 0.1
    legBg.Line.Visible = msoFalse
    legBg.Name = "SlabLegBG"
    
    ' Legend items
    Dim thicknesses As Variant: thicknesses = Array(125, 150, 175, 200)
    Dim labels As Variant: labels = Array("125 mm slab", "150 mm slab", "175 mm slab", "200 mm slab")
    Dim idx As Long
    
    For idx = 0 To 3
        Dim sq As Shape
        Set sq = ws.Shapes.AddShape(1, legX, legY + idx * 20, 12, 12)
        sq.Fill.ForeColor.RGB = ThicknessColor(CLng(thicknesses(idx)))
        sq.Fill.Transparency = 0.3
        sq.Line.ForeColor.RGB = LABEL_WHITE
        sq.Line.Weight = 0.5
        sq.Name = "SlabLeg_" & thicknesses(idx)
        
        Dim lt As Shape
        Set lt = ws.Shapes.AddTextbox(1, legX + 16, legY + idx * 20 - 2, 120, 16)
        lt.TextFrame2.TextRange.Text = CStr(labels(idx))
        lt.TextFrame2.TextRange.Font.Size = 8
        lt.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = LABEL_WHITE
        lt.Fill.Visible = msoFalse
        lt.Line.Visible = msoFalse
        lt.Name = "SlabLegTxt_" & thicknesses(idx)
    Next idx
End Sub
