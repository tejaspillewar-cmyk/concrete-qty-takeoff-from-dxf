Attribute VB_Name = "mod_Geometry"
'======================================================================
' mod_Geometry — Pure Math Functions
'======================================================================
' Polygon area (Shoelace), point-in-polygon (ray casting), centroid,
' and perimeter calculations. All pure math — no DXF or Excel dependency.
'
' All coordinate inputs are in millimeters (DXF drawing units).
'======================================================================
Option Explicit

'----------------------------------------------------------------------
' ShoelaceArea — Polygon area using the Shoelace formula
'----------------------------------------------------------------------
' Returns absolute area in square millimeters.
'----------------------------------------------------------------------
Public Function ShoelaceArea(ByRef X() As Double, ByRef Y() As Double, _
                              ByVal n As Long) As Double
    Dim area As Double
    Dim i As Long, j As Long
    
    If n < 3 Then
        ShoelaceArea = 0#
        Exit Function
    End If
    
    area = 0#
    For i = 0 To n - 1
        j = (i + 1) Mod n
        area = area + X(i) * Y(j) - X(j) * Y(i)
    Next i
    
    ShoelaceArea = Abs(area) / 2#
End Function


'----------------------------------------------------------------------
' SqmmToSqm — Convert square millimeters to square meters
'----------------------------------------------------------------------
Public Function SqmmToSqm(ByVal areaSqmm As Double) As Double
    SqmmToSqm = areaSqmm / 1000000#
End Function


'----------------------------------------------------------------------
' PointInPolygon — Ray-casting algorithm
'----------------------------------------------------------------------
' Returns True if point (px, py) is inside the polygon.
'----------------------------------------------------------------------
Public Function PointInPolygon(ByVal px As Double, ByVal py As Double, _
                                ByRef X() As Double, ByRef Y() As Double, _
                                ByVal n As Long) As Boolean
    Dim inside As Boolean
    Dim i As Long, j As Long
    Dim xi As Double, yi As Double, xj As Double, yj As Double
    
    If n < 3 Then
        PointInPolygon = False
        Exit Function
    End If
    
    inside = False
    j = n - 1
    
    For i = 0 To n - 1
        xi = X(i): yi = Y(i)
        xj = X(j): yj = Y(j)
        
        If ((yi > py) <> (yj > py)) Then
            If px < (xj - xi) * (py - yi) / (yj - yi) + xi Then
                inside = Not inside
            End If
        End If
        
        j = i
    Next i
    
    PointInPolygon = inside
End Function


'----------------------------------------------------------------------
' PolygonCentroid — Arithmetic mean of vertices
'----------------------------------------------------------------------
' Returns centroid in (cx, cy) via ByRef parameters.
'----------------------------------------------------------------------
Public Sub PolygonCentroid(ByRef X() As Double, ByRef Y() As Double, _
                           ByVal n As Long, _
                           ByRef cx As Double, ByRef cy As Double)
    Dim i As Long
    cx = 0#: cy = 0#
    
    If n = 0 Then Exit Sub
    
    For i = 0 To n - 1
        cx = cx + X(i)
        cy = cy + Y(i)
    Next i
    
    cx = cx / CDbl(n)
    cy = cy / CDbl(n)
End Sub


'----------------------------------------------------------------------
' PolygonPerimeter — Sum of edge lengths
'----------------------------------------------------------------------
' Returns perimeter in millimeters.
'----------------------------------------------------------------------
Public Function PolygonPerimeter(ByRef X() As Double, ByRef Y() As Double, _
                                  ByVal n As Long) As Double
    Dim peri As Double
    Dim i As Long, j As Long
    Dim dx As Double, dy As Double
    
    peri = 0#
    For i = 0 To n - 1
        j = (i + 1) Mod n
        dx = X(j) - X(i)
        dy = Y(j) - Y(i)
        peri = peri + Sqr(dx * dx + dy * dy)
    Next i
    
    PolygonPerimeter = peri
End Function
