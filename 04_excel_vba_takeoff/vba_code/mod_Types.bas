Attribute VB_Name = "mod_Types"
'======================================================================
' mod_Types — Shared Type Definitions
'======================================================================
' All Public Types used by the DXF Quantity Take-Off system.
' These must live in a standard module so all other modules can see them.
'======================================================================
Option Explicit

' ── DXF Parsing Output ──────────────────────────────────────────────

Public Type DXFPolyline
    Layer       As String
    IsClosed    As Boolean
    NumVerts    As Long
    X()         As Double       ' vertex X coordinates (0-based)
    Y()         As Double       ' vertex Y coordinates (0-based)
End Type

Public Type DXFText
    Layer       As String
    Content     As String       ' raw text, e.g. "150 THK"
    X           As Double       ' insertion point X
    Y           As Double       ' insertion point Y
End Type

' ── Slab Data ────────────────────────────────────────────────────────

Public Type SlabEntry
    Name            As String   ' auto-name, e.g. "S1-150"
    Layer           As String   ' DXF layer, e.g. "STR-SLAB-REG150"
    ThicknessMM     As Long     ' thickness in mm
    ThicknessSource As String   ' "text_label" or "layer_name" or "unknown"
    LabelText       As String   ' original label text
    AreaSqm         As Double   ' area in m²
    VolumeCum       As Double   ' volume in m³
    CentroidX       As Double   ' centroid X (drawing mm)
    CentroidY       As Double   ' centroid Y (drawing mm)
    NumVerts        As Long
    X()             As Double   ' polygon vertices
    Y()             As Double
End Type

Public Type SlabSummaryRow
    ThicknessMM     As Long
    Count           As Long
    TotalAreaSqm    As Double
    TotalVolumeCum  As Double
End Type

' ── Wall Data ────────────────────────────────────────────────────────

Public Type WallEntry
    Name            As String   ' auto-name, e.g. "SW-1" or "NSW100-3"
    Layer           As String
    ThicknessMM     As Long
    AreaSqm         As Double   ' plan area in m²
    LengthM         As Double   ' estimated length in m
    CentroidX       As Double
    CentroidY       As Double
    IsStructural    As Boolean
    NumVerts        As Long
    X()             As Double
    Y()             As Double
End Type

Public Type WallSummaryRow
    ThicknessMM     As Long
    Count           As Long
    TotalAreaSqm    As Double
    TotalLengthM    As Double
End Type
