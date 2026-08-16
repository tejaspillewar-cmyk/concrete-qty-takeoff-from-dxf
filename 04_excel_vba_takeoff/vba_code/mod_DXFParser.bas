Attribute VB_Name = "mod_DXFParser"
'======================================================================
' mod_DXFParser — ASCII DXF File Parser
'======================================================================
' Reads an ASCII DXF file line-by-line and extracts:
'   - LWPOLYLINE entities (layer, closed flag, vertex coordinates)
'   - TEXT entities (layer, text content, insertion point)
'   - MTEXT entities (layer, text content with formatting stripped,
'                     insertion point)
'
' Only works with ASCII DXF files. Binary DXF will be detected and rejected.
' Uses Val() instead of CDbl() for locale-independent decimal parsing.
'======================================================================
Option Explicit

Private Const INITIAL_ARRAY_SIZE As Long = 500

'----------------------------------------------------------------------
' IsASCIIDXF — Quick check if a file is ASCII (not binary) DXF
'----------------------------------------------------------------------
Public Function IsASCIIDXF(ByVal filePath As String) As Boolean
    Dim f As Integer
    Dim firstLine As String
    
    f = FreeFile
    On Error GoTo NotValid
    Open filePath For Input As #f
    
    If EOF(f) Then
        Close #f
        IsASCIIDXF = False
        Exit Function
    End If
    
    Line Input #f, firstLine
    Close #f
    
    firstLine = Trim(firstLine)
    
    ' Binary DXF starts with "AutoCAD Binary DXF"
    If Left(firstLine, 18) = "AutoCAD Binary DXF" Then
        IsASCIIDXF = False
        Exit Function
    End If
    
    ' ASCII DXF first line should be a group code (integer)
    If IsNumeric(firstLine) Then
        IsASCIIDXF = True
    Else
        IsASCIIDXF = False
    End If
    Exit Function
    
NotValid:
    IsASCIIDXF = False
End Function


'----------------------------------------------------------------------
' ParseDXFFile — Main parser
'----------------------------------------------------------------------
' Reads the ENTITIES section and extracts LWPOLYLINE and TEXT entities.
'
' Parameters (all ByRef for output):
'   polys()    — array of DXFPolyline
'   polyCount  — number of polylines found
'   texts()    — array of DXFText
'   textCount  — number of text entities found
'----------------------------------------------------------------------
Public Sub ParseDXFFile(ByVal filePath As String, _
                        ByRef polys() As DXFPolyline, ByRef polyCount As Long, _
                        ByRef texts() As DXFText, ByRef textCount As Long)
    
    Dim f As Integer
    Dim codeLine As String, valueLine As String
    Dim code As Long, value As String
    Dim inEntities As Boolean
    Dim sectionPending As Boolean
    Dim currentEntity As String
    
    ' ── Polyline temp vars ────────────────────────────────────
    Dim pLayer As String
    Dim pFlags As Long
    Dim pX() As Double, pY() As Double
    Dim pVertIdx As Long
    Dim pMaxVerts As Long
    
    ' ── Text temp vars ────────────────────────────────────────
    Dim tLayer As String
    Dim tContent As String
    Dim tX As Double, tY As Double
    Dim tHasContent As Boolean
    
    ' ── Initialize outputs ────────────────────────────────────
    polyCount = 0
    textCount = 0
    ReDim polys(0 To INITIAL_ARRAY_SIZE - 1)
    ReDim texts(0 To INITIAL_ARRAY_SIZE - 1)
    
    ' ── Temp vertex buffer ────────────────────────────────────
    pMaxVerts = 500
    ReDim pX(0 To pMaxVerts - 1)
    ReDim pY(0 To pMaxVerts - 1)
    
    ' ── Open file ─────────────────────────────────────────────
    f = FreeFile
    Open filePath For Input As #f
    
    inEntities = False
    sectionPending = False
    currentEntity = ""
    pVertIdx = 0
    
    ' ── Main read loop ────────────────────────────────────────
    Do While Not EOF(f)
        ' Read group code
        If EOF(f) Then Exit Do
        Line Input #f, codeLine
        If EOF(f) Then Exit Do
        Line Input #f, valueLine
        
        code = Val(Trim(codeLine))
        value = Trim(valueLine)
        
        ' ── Section tracking ──────────────────────────────────
        If code = 0 And value = "SECTION" Then
            sectionPending = True
            GoTo ContinueLoop
        End If
        
        If code = 2 And sectionPending Then
            sectionPending = False
            If value = "ENTITIES" Then
                inEntities = True
            End If
            GoTo ContinueLoop
        End If
        
        If sectionPending And code <> 2 Then
            sectionPending = False
        End If
        
        ' Skip everything outside ENTITIES section
        If Not inEntities Then GoTo ContinueLoop
        
        ' ── Entity boundary (group code 0) ────────────────────
        If code = 0 Then
            ' Save previous entity
            Call SavePendingPolyline(currentEntity, pLayer, pFlags, _
                                     pX, pY, pVertIdx, _
                                     polys, polyCount)
            Call SavePendingText(currentEntity, tLayer, tContent, _
                                 tX, tY, tHasContent, _
                                 texts, textCount)
            
            ' End of entities section?
            If value = "ENDSEC" Then Exit Do
            
            ' Start new entity
            currentEntity = value
            pLayer = "": pFlags = 0: pVertIdx = 0
            tLayer = "": tContent = "": tX = 0: tY = 0: tHasContent = False
            
            GoTo ContinueLoop
        End If
        
        ' ── Collect entity attributes ─────────────────────────
        Select Case currentEntity
            Case "LWPOLYLINE"
                Select Case code
                    Case 8:  pLayer = value
                    Case 70: pFlags = Val(value)
                    Case 10
                        ' X coordinate — grow buffer if needed
                        If pVertIdx >= pMaxVerts Then
                            pMaxVerts = pMaxVerts * 2
                            ReDim Preserve pX(0 To pMaxVerts - 1)
                            ReDim Preserve pY(0 To pMaxVerts - 1)
                        End If
                        pX(pVertIdx) = Val(value)
                    Case 20
                        pY(pVertIdx) = Val(value)
                        pVertIdx = pVertIdx + 1
                End Select
                
            Case "TEXT"
                Select Case code
                    Case 8:  tLayer = value
                    Case 1:  tContent = value: tHasContent = True
                    Case 10: tX = Val(value)
                    Case 20: tY = Val(value)
                End Select
                
            Case "MTEXT"
                Select Case code
                    Case 8:  tLayer = value
                    Case 1:  tContent = value: tHasContent = True
                    Case 3   ' MTEXT continuation chunk — append to content
                        tContent = tContent & value
                    Case 10: tX = Val(value)
                    Case 20: tY = Val(value)
                End Select
        End Select
        
ContinueLoop:
    Loop
    
    ' Save any trailing entity
    Call SavePendingPolyline(currentEntity, pLayer, pFlags, _
                             pX, pY, pVertIdx, _
                             polys, polyCount)
    Call SavePendingText(currentEntity, tLayer, tContent, _
                         tX, tY, tHasContent, _
                         texts, textCount)
    
    Close #f
    
    ' Trim output arrays
    If polyCount > 0 Then
        ReDim Preserve polys(0 To polyCount - 1)
    Else
        ReDim polys(0 To 0)
    End If
    If textCount > 0 Then
        ReDim Preserve texts(0 To textCount - 1)
    Else
        ReDim texts(0 To 0)
    End If
End Sub


'----------------------------------------------------------------------
' SavePendingPolyline — Store a completed LWPOLYLINE into the output array
'----------------------------------------------------------------------
Private Sub SavePendingPolyline(ByVal entityType As String, _
                                 ByVal pLayer As String, _
                                 ByVal pFlags As Long, _
                                 ByRef pX() As Double, ByRef pY() As Double, _
                                 ByVal pVertIdx As Long, _
                                 ByRef polys() As DXFPolyline, ByRef polyCount As Long)
    If entityType <> "LWPOLYLINE" Then Exit Sub
    If pVertIdx < 2 Then Exit Sub   ' need at least 2 vertices
    
    ' Grow output array if needed
    If polyCount > UBound(polys) Then
        ReDim Preserve polys(0 To polyCount * 2)
    End If
    
    Dim i As Long
    polys(polyCount).Layer = pLayer
    polys(polyCount).IsClosed = ((pFlags And 1) = 1)
    polys(polyCount).NumVerts = pVertIdx
    ReDim polys(polyCount).X(0 To pVertIdx - 1)
    ReDim polys(polyCount).Y(0 To pVertIdx - 1)
    For i = 0 To pVertIdx - 1
        polys(polyCount).X(i) = pX(i)
        polys(polyCount).Y(i) = pY(i)
    Next i
    polyCount = polyCount + 1
End Sub


'----------------------------------------------------------------------
' SavePendingText — Store a completed TEXT entity into the output array
'----------------------------------------------------------------------
Private Sub SavePendingText(ByVal entityType As String, _
                             ByVal tLayer As String, _
                             ByVal tContent As String, _
                             ByVal tX As Double, ByVal tY As Double, _
                             ByVal tHasContent As Boolean, _
                             ByRef texts() As DXFText, ByRef textCount As Long)
    ' Accept both TEXT and MTEXT entities
    If entityType <> "TEXT" And entityType <> "MTEXT" Then Exit Sub
    If Not tHasContent Then Exit Sub
    
    ' Strip MTEXT formatting codes if present
    If entityType = "MTEXT" Then
        tContent = StripMTEXTFormatting(tContent)
    End If
    
    ' Grow output array if needed
    If textCount > UBound(texts) Then
        ReDim Preserve texts(0 To textCount * 2)
    End If
    
    texts(textCount).Layer = tLayer
    texts(textCount).Content = tContent
    texts(textCount).X = tX
    texts(textCount).Y = tY
    textCount = textCount + 1
End Sub


'----------------------------------------------------------------------
' StripMTEXTFormatting — Remove MTEXT formatting codes from content
'----------------------------------------------------------------------
' MTEXT can contain RTF-like formatting:
'   \P           → paragraph break (newline)
'   \f...;       → font change
'   {\C1;text}   → color override
'   {\Hvalue;text} → height override
'   \~           → non-breaking space
'   %%d           → degree symbol
'   etc.
'
' This function strips all formatting and returns plain text.
'----------------------------------------------------------------------
Private Function StripMTEXTFormatting(ByVal txt As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim inBrace As Long  ' track nesting depth of {}
    
    result = txt
    
    ' Replace \P with space (paragraph break)
    result = Replace(result, "\P", " ")
    
    ' Remove backslash commands: \f...; \H...; \W...; \Q...; \T...; \A..;
    ' Pattern: backslash + letter + content + semicolon
    Dim cleaned As String: cleaned = ""
    i = 1
    Do While i <= Len(result)
        ch = Mid(result, i, 1)
        
        If ch = "\" And i < Len(result) Then
            Dim nextCh As String: nextCh = Mid(result, i + 1, 1)
            Select Case UCase(nextCh)
                Case "P"
                    cleaned = cleaned & " "
                    i = i + 2
                Case "~"
                    cleaned = cleaned & " "
                    i = i + 2
                Case "F", "H", "W", "Q", "T", "A", "C", "L", "O", "K"
                    ' Skip until semicolon
                    Dim scPos As Long
                    scPos = InStr(i + 1, result, ";")
                    If scPos > 0 Then
                        i = scPos + 1
                    Else
                        i = i + 2
                    End If
                Case Else
                    cleaned = cleaned & ch
                    i = i + 1
            End Select
        ElseIf ch = "{" Or ch = "}" Then
            ' Skip braces (used for grouping formatting)
            i = i + 1
        Else
            cleaned = cleaned & ch
            i = i + 1
        End If
    Loop
    
    StripMTEXTFormatting = Trim(cleaned)
End Function
