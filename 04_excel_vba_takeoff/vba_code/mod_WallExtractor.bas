Attribute VB_Name = "mod_WallExtractor"
'======================================================================
' mod_WallExtractor — Wall Extraction Logic
'======================================================================
' Extracts structural and non-structural walls from DXF polylines.
'
' Layer conventions:
'   STR-WALL-REG        → Structural walls (thickness unknown = 0)
'   STR-WALL-NS-100     → Non-structural 100mm
'   STR-WALL-NS-150     → Non-structural 150mm
'   STR-WALL-NS-200     → Non-structural 200mm
'
' Wall length is estimated as perimeter / 2 (assumes thin rectangular shape).
'======================================================================
Option Explicit

Private Const STR_WALL_LAYER As String = "STR-WALL-REG"
Private Const NS_WALL_PREFIX As String = "STR-WALL-NS-"


'----------------------------------------------------------------------
' ExtractWalls — Main wall extraction
'----------------------------------------------------------------------
Public Sub ExtractWalls(ByRef polys() As DXFPolyline, ByVal polyCount As Long, _
                        ByRef strWalls() As WallEntry, ByRef strCount As Long, _
                        ByRef nsWalls() As WallEntry, ByRef nsCount As Long, _
                        ByRef warnings() As String, ByRef warnCount As Long)
    
    Dim i As Long, k As Long
    Dim layer As String
    Dim isStr As Boolean, isNS As Boolean
    
    ' ── Initialize outputs ────────────────────────────────────
    strCount = 0
    nsCount = 0
    warnCount = 0
    ReDim strWalls(0 To 499)
    ReDim nsWalls(0 To 499)
    ReDim warnings(0 To 49)
    
    ' Counters for auto-naming
    Dim strCounter As Long: strCounter = 0
    Dim nsCounters As Object
    Set nsCounters = CreateObject("Scripting.Dictionary")
    
    ' ── Process each polyline ─────────────────────────────────
    For i = 0 To polyCount - 1
        If Not polys(i).IsClosed Then GoTo NextWall
        If polys(i).NumVerts < 3 Then GoTo NextWall
        
        layer = polys(i).Layer
        isStr = (layer = STR_WALL_LAYER)
        isNS = (Left(layer, Len(NS_WALL_PREFIX)) = NS_WALL_PREFIX)
        
        If Not isStr And Not isNS Then GoTo NextWall
        
        ' ── Determine thickness ───────────────────────────────
        Dim thickMM As Long: thickMM = 0
        If isNS Then
            Dim nsSuffix As String
            nsSuffix = Mid(layer, Len(NS_WALL_PREFIX) + 1)
            If Len(nsSuffix) > 0 And IsNumeric(nsSuffix) Then
                thickMM = CLng(nsSuffix)
            End If
        End If
        ' For structural walls, thickness remains 0 (unknown from layer name)
        
        ' ── Compute quantities ────────────────────────────────
        Dim areaSqmm As Double, areaSqm As Double
        Dim perimeterMM As Double, lengthM As Double
        Dim cx As Double, cy As Double
        
        areaSqmm = ShoelaceArea(polys(i).X, polys(i).Y, polys(i).NumVerts)
        areaSqm = SqmmToSqm(areaSqmm)
        perimeterMM = PolygonPerimeter(polys(i).X, polys(i).Y, polys(i).NumVerts)
        lengthM = (perimeterMM / 2#) / 1000#   ' perimeter/2, mm → m
        Call PolygonCentroid(polys(i).X, polys(i).Y, polys(i).NumVerts, cx, cy)
        
        ' ── Build entry ───────────────────────────────────────
        If isStr Then
            ' Structural wall
            If strCount > UBound(strWalls) Then
                ReDim Preserve strWalls(0 To strCount * 2)
            End If
            strCounter = strCounter + 1
            
            With strWalls(strCount)
                .Name = "SW-" & strCounter
                .Layer = layer
                .ThicknessMM = thickMM
                .AreaSqm = Round(areaSqm, 3)
                .LengthM = Round(lengthM, 3)
                .CentroidX = Round(cx, 1)
                .CentroidY = Round(cy, 1)
                .IsStructural = True
                .NumVerts = polys(i).NumVerts
                ReDim .X(0 To polys(i).NumVerts - 1)
                ReDim .Y(0 To polys(i).NumVerts - 1)
                For k = 0 To polys(i).NumVerts - 1
                    .X(k) = polys(i).X(k)
                    .Y(k) = polys(i).Y(k)
                Next k
            End With
            strCount = strCount + 1
        Else
            ' Non-structural wall
            If nsCount > UBound(nsWalls) Then
                ReDim Preserve nsWalls(0 To nsCount * 2)
            End If
            
            Dim nsKey As String: nsKey = CStr(thickMM)
            If Not nsCounters.Exists(nsKey) Then
                nsCounters.Add nsKey, CLng(0)
            End If
            nsCounters(nsKey) = nsCounters(nsKey) + 1
            
            With nsWalls(nsCount)
                .Name = "NSW" & thickMM & "-" & nsCounters(nsKey)
                .Layer = layer
                .ThicknessMM = thickMM
                .AreaSqm = Round(areaSqm, 3)
                .LengthM = Round(lengthM, 3)
                .CentroidX = Round(cx, 1)
                .CentroidY = Round(cy, 1)
                .IsStructural = False
                .NumVerts = polys(i).NumVerts
                ReDim .X(0 To polys(i).NumVerts - 1)
                ReDim .Y(0 To polys(i).NumVerts - 1)
                For k = 0 To polys(i).NumVerts - 1
                    .X(k) = polys(i).X(k)
                    .Y(k) = polys(i).Y(k)
                Next k
            End With
            nsCount = nsCount + 1
        End If
        
NextWall:
    Next i
    
    ' Trim outputs
    If strCount > 0 Then
        ReDim Preserve strWalls(0 To strCount - 1)
    Else
        ReDim strWalls(0 To 0)
    End If
    If nsCount > 0 Then
        ReDim Preserve nsWalls(0 To nsCount - 1)
    Else
        ReDim nsWalls(0 To 0)
    End If
    If warnCount > 0 Then
        ReDim Preserve warnings(0 To warnCount - 1)
    Else
        ReDim warnings(0 To 0)
    End If
End Sub
