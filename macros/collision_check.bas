Option Explicit
 
Dim swApp As Object
Dim swDoc As Object
Dim fNum  As Integer
 
Const MAX_DIMS  As Integer = 500
Const MAX_VIEWS As Integer = 50
 
' Distance (meters) under which two dimension positions are considered
' a collision. Tune this instead of relying on a real text box.
Const COLLISION_RADIUS As Double = 0.004   ' 4mm
Const NUDGE_STEP       As Double = 0.006   ' 6mm per nudge
 
Dim dOwnerView(MAX_DIMS)  As String
Dim dFullName(MAX_DIMS)   As String
Dim dX(MAX_DIMS)          As Double
Dim dY(MAX_DIMS)          As Double
Dim dAnnObj(MAX_DIMS)     As Object
Dim numDims As Integer
 
Dim viewName(MAX_VIEWS) As String
Dim viewXMin(MAX_VIEWS) As Double, viewYMin(MAX_VIEWS) As Double
Dim viewXMax(MAX_VIEWS) As Double, viewYMax(MAX_VIEWS) As Double
Dim numViews As Integer
 
Dim collisionCount As Integer
Dim strayCount     As Integer
 
 
Sub main()
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc

    Dim sp As Variant                     
    sp = swDoc.GetCurrentSheet().GetProperties()
    sheetW = sp(5)
    sheetH = sp(6)
 
    fNum = FreeFile()
    Open "{LOGPATH}" For Output As #fNum
    Print #fNum, "Position-Based Dimension Collision Check (single pass)"
    Print #fNum, "========================================================"
 
    numDims = 0
    numViews = 0
    collisionCount = 0
    strayCount = 0
 
    Call CollectDimensions
 
    Print #fNum, ""
    Print #fNum, "Collected " & numDims & " dimensions across " & numViews & " views"
 
    Print #fNum, ""
    Print #fNum, "===== BEFORE (sheet map) ====="
    Call PrintAsciiMap
 
    Print #fNum, ""
    Call CheckDimVsDim
    Call CheckDimVsForeignView
 
    Print #fNum, ""
    Print #fNum, "===== AFTER (sheet map) ====="
    Call PrintAsciiMap
 
    swDoc.ForceRebuild3 False
 
    Print #fNum, ""
    Print #fNum, "-- Summary --"
    Print #fNum, "Total dimensions checked : " & numDims
    Print #fNum, "Total views              : " & numViews
    Print #fNum, "Dim-vs-Dim collisions    : " & collisionCount
    Print #fNum, "Dim-vs-ForeignView strays: " & strayCount
    Print #fNum, ""
    Print #fNum, "Done."
    Close #fNum
End Sub
 
 
' ─────────────────────────────────────────────────────────────────────
' Collect every view outline + every dimension's POSITION (not box)
' ─────────────────────────────────────────────────────────────────────
Sub CollectDimensions()
    Dim view As Object
    Set view = swDoc.GetFirstView()
    Set view = view.GetNextView()   ' skip sheet pseudo-view
 
    Do While Not view Is Nothing
        Dim outline As Variant
        outline = view.GetOutline()
        viewName(numViews) = view.GetName2()
        viewXMin(numViews) = outline(0) : viewYMin(numViews) = outline(1)
        viewXMax(numViews) = outline(2) : viewYMax(numViews) = outline(3)
        numViews = numViews + 1
 
        Dim ann As Object
        Set ann = view.GetFirstAnnotation3()
 
        Dim seenInView As Integer : seenInView = 0
        Dim skippedTypes As String : skippedTypes = ""
 
        Do While Not ann Is Nothing
            seenInView = seenInView + 1
 
            Dim annType As Long
            annType = ann.GetType()   ' swAnnotationType_e — 16 = swDisplayDimension
 
            If annType = swDisplayDimension Then
                Dim dd As Object
                Set dd = ann.GetSpecificAnnotation()
 
                If dd Is Nothing Then
                    Print #fNum, "  [SKIP] GetSpecificAnnotation returned Nothing in " & view.GetName2()
                    GoTo NextAnn
                End If
 
                Dim d As Object
                Err.Clear
                On Error Resume Next
                Set d = dd.GetDimension()
                If Err.Number <> 0 Then
                    Print #fNum, "  [SKIP] dd.GetDimension() ERROR " & Err.Number & ": " & Err.Description & _
                                 "  (dd type=" & TypeName(dd) & ") in " & view.GetName2()
                    Err.Clear
                    On Error GoTo 0
                    GoTo NextAnn
                End If
                On Error GoTo 0
 
                If d Is Nothing Then
                    Print #fNum, "  [SKIP] dd.GetDimension() returned Nothing (dd type=" & TypeName(dd) & _
                                 ") in " & view.GetName2()
                    GoTo NextAnn
                End If
 
                ' -- Just the position. No box, no tolerance. --
                Dim pos As Variant
                Err.Clear
                On Error Resume Next
                pos = ann.GetPosition()
                If Err.Number <> 0 Then
                    Print #fNum, "  [SKIP] ann.GetPosition() ERROR " & Err.Number & ": " & Err.Description & _
                                 "  (ann type=" & TypeName(ann) & ") in " & view.GetName2()
                    Err.Clear
                    On Error GoTo 0
                    GoTo NextAnn
                End If
                On Error GoTo 0
 
                If IsEmpty(pos) Then
                    Print #fNum, "  [SKIP] ann.GetPosition() returned Empty (no error) in " & view.GetName2()
                    GoTo NextAnn
                End If
 
                If numDims < MAX_DIMS Then
                    dOwnerView(numDims) = view.GetName2()
                    dFullName(numDims)  = d.FullName
                    dX(numDims) = pos(0)
                    dY(numDims) = pos(1)
                    Set dAnnObj(numDims) = ann
                    numDims = numDims + 1
                End If
            Else
                skippedTypes = skippedTypes & annType & ","
            End If
NextAnn:
            Set ann = ann.GetNext2()
        Loop
 
        Print #fNum, "  " & view.GetName2() & ": scanned " & seenInView & _
                     " annotations, non-dimension types skipped=[" & skippedTypes & "]"
 
        Set view = view.GetNextView()
    Loop
 
    ' ── Print full inventory ──
    Print #fNum, ""
    Print #fNum, "-- All Views (outline in mm) --"
    Dim v As Integer
    For v = 0 To numViews - 1
        Print #fNum, "  " & Pad(viewName(v), 14) & _
                     "  x:[" & Format(viewXMin(v) * 1000, "0.0") & " , " & _
                     Format(viewXMax(v) * 1000, "0.0") & "]" & _
                     "  y:[" & Format(viewYMin(v) * 1000, "0.0") & " , " & _
                     Format(viewYMax(v) * 1000, "0.0") & "]"
    Next v
 
    Print #fNum, ""
    Print #fNum, "-- All Dimensions (position in mm) --"
    Dim i As Integer
    For i = 0 To numDims - 1
        Print #fNum, "  [" & Format(i, "000") & "] " & Pad(dFullName(i), 22) & _
                     "  view=" & Pad(dOwnerView(i), 12) & _
                     "  pos:(" & Format(dX(i) * 1000, "0.0") & "," & _
                     Format(dY(i) * 1000, "0.0") & ")"
    Next i
End Sub
 
 
' ─────────────────────────────────────────────────────────────────────
' Helpers
' ─────────────────────────────────────────────────────────────────────
Function Pad(s As String, n As Integer) As String
    If Len(s) >= n Then
        Pad = Left(s, n)
    Else
        Pad = s & Space(n - Len(s))
    End If
End Function
 
 
Function Dist(ax As Double, ay As Double, bx As Double, by As Double) As Double
    Dist = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function
 
 
Function PointInBox(px As Double, py As Double, _
                     bx0 As Double, by0 As Double, bx1 As Double, by1 As Double) As Boolean
    PointInBox = (px >= bx0 And px <= bx1 And py >= by0 And py <= by1)
End Function
 
 
Sub NudgeDimension(idx As Integer)
    Dim oldY As Double
    oldY = dY(idx)
 
    Dim newY As Double
    newY = dY(idx) - NUDGE_STEP
 
    On Error Resume Next
    dAnnObj(idx).SetPosition2 dX(idx), newY, 0
    On Error GoTo 0
 
    Dim pos As Variant
    On Error Resume Next
    pos = dAnnObj(idx).GetPosition()
    On Error GoTo 0
    If Not IsEmpty(pos) Then
        dX(idx) = pos(0)
        dY(idx) = pos(1)
    End If
 
    Print #fNum, "    -> NUDGED [" & Format(idx, "000") & "] " & dFullName(idx) & _
                 "  from y=" & Format(oldY * 1000, "0.0") & _
                 "  to y=" & Format(dY(idx) * 1000, "0.0")
End Sub
 
 
' ─────────────────────────────────────────────────────────────────────
' Checks
' ─────────────────────────────────────────────────────────────────────
Sub CheckDimVsDim()
    Dim i As Integer, j As Integer
    Print #fNum, "-- Dimension vs Dimension (proximity, radius=" & _
                 Format(COLLISION_RADIUS * 1000, "0.0") & "mm) --"
    For i = 0 To numDims - 1
        For j = i + 1 To numDims - 1
            If Dist(dX(i), dY(i), dX(j), dY(j)) < COLLISION_RADIUS Then
                collisionCount = collisionCount + 1
                Print #fNum, "COLLISION: " & dFullName(i) & " (" & dOwnerView(i) & _
                             ")  <->  " & dFullName(j) & " (" & dOwnerView(j) & ")"
                Call NudgeDimension(j)
            End If
        Next j
    Next i
    If collisionCount = 0 Then Print #fNum, "  (none found)"
End Sub
 
 
Sub CheckDimVsForeignView()
    Dim i As Integer, v As Integer
    Print #fNum, ""
    Print #fNum, "-- Dimension vs Foreign View Territory --"
    For i = 0 To numDims - 1
        For v = 0 To numViews - 1
            If viewName(v) <> dOwnerView(i) Then
                If PointInBox(dX(i), dY(i), viewXMin(v), viewYMin(v), viewXMax(v), viewYMax(v)) Then
                    strayCount = strayCount + 1
                    Print #fNum, "STRAY: " & dFullName(i) & " (belongs to " & dOwnerView(i) & _
                                 ") sits inside territory of " & viewName(v)
                    Call NudgeDimension(i)
                End If
            End If
        Next v
    Next i
    If strayCount = 0 Then Print #fNum, "  (none found)"
End Sub
 
 
' ─────────────────────────────────────────────────────────────────────
' Crude ASCII visualization of the sheet: lowercase letters = view
' outlines (a, b, c...), '#' = dimension position (single cell).
' ─────────────────────────────────────────────────────────────────────
Sub PrintAsciiMap()
    Const COLS As Integer = 100
    Const ROWS As Integer = 40
 
    Dim sp As Variant
    sp = swDoc.GetCurrentSheet().GetProperties()
    Dim sheetW As Double : sheetW = sp(5)
    Dim sheetH As Double : sheetH = sp(6)
 
    Dim grid(ROWS, COLS) As String
    Dim r As Integer, c As Integer
    For r = 0 To ROWS : For c = 0 To COLS : grid(r, c) = "." : Next c : Next r
 
    Dim v As Integer
    For v = 0 To numViews - 1
        Dim ch As String : ch = Chr(97 + (v Mod 26))   ' a, b, c...
        Dim c0 As Integer, c1 As Integer, r0 As Integer, r1 As Integer
        c0 = Int(viewXMin(v) / sheetW * COLS) : c1 = Int(viewXMax(v) / sheetW * COLS)
        r0 = ROWS - Int(viewYMax(v) / sheetH * ROWS) : r1 = ROWS - Int(viewYMin(v) / sheetH * ROWS)
        For r = r0 To r1
            For c = c0 To c1
                If r >= 0 And r <= ROWS And c >= 0 And c <= COLS Then
                    If grid(r, c) = "." Then grid(r, c) = ch
                End If
            Next c
        Next r
    Next v
 
    Dim i As Integer
    For i = 0 To numDims - 1
        c = Int(dX(i) / sheetW * COLS)
        r = ROWS - Int(dY(i) / sheetH * ROWS)
        If r >= 0 And r <= ROWS And c >= 0 And c <= COLS Then
            grid(r, c) = "#"
        End If
    Next i
 
    For r = 0 To ROWS
        Dim line As String : line = ""
        For c = 0 To COLS
            line = line & grid(r, c)
        Next c
        Print #fNum, line
    Next r
End Sub
