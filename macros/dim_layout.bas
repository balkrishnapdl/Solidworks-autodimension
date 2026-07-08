Option Explicit

Dim swApp As Object
Dim swDoc As Object
Dim fNum  As Integer

Const MAX_DIMS  As Integer = 500
Const MAX_VIEWS As Integer = 50

Const ZONE_PAD   As Double = 0.035   ' 25mm - view's dimension territory (both axes)
Const DIM_GAP    As Double = 0.009   ' 8mm  - spacing between stacked dims
Const MIN_OFFSET As Double = 0.006  ' 6mm  - gap from view edge to first dim

Dim dOwnerView(MAX_DIMS)   As String
Dim dFullName(MAX_DIMS)    As String
Dim dX(MAX_DIMS)           As Double
Dim dY(MAX_DIMS)           As Double
Dim dAnnObj(MAX_DIMS)      As Object
Dim dIsVertical(MAX_DIMS)  As Boolean   ' True = height-type dim (stack left/right)
                                        ' False = width-type dim (stack above/below)
Dim numDims As Integer

Dim viewName(MAX_VIEWS) As String
Dim viewXMin(MAX_VIEWS) As Double, viewYMin(MAX_VIEWS) As Double
Dim viewXMax(MAX_VIEWS) As Double, viewYMax(MAX_VIEWS) As Double
Dim numViews As Integer

Dim sheetW As Double, sheetH As Double


Sub main()
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc

    Dim sp As Variant
    sp = swDoc.GetCurrentSheet().GetProperties()
    sheetW = sp(5)
    sheetH = sp(6)

    fNum = FreeFile()
    Open "{LOGPATH}" For Output As #fNum
    Print #fNum, "Dimension Layout Pass (orientation-aware, dual-axis)"
    Print #fNum, "======================================================"
    Print #fNum, "Sheet: " & Format(sheetW*1000,"0.0") & " x " & Format(sheetH*1000,"0.0") & " mm"

    numDims = 0
    numViews = 0

    Call CollectViewsAndDims
    Call CheckZoneOverlap
    Call AssignSlots

    swDoc.ForceRebuild3 False
    Print #fNum, "Done."
    Close #fNum
End Sub


' ─────────────────────────────────────────────────────────────────────
' Collect views + dims, classifying each dim's orientation from its
' ORIGINAL position (before we move anything):
'   - if the label already sits to the left/right of its view -> vertical
'     (height) dimension -> we will stack these along X
'   - if the label already sits above/below its view          -> horizontal
'     (width) dimension  -> we will stack these along Y
' ─────────────────────────────────────────────────────────────────────
Sub CollectViewsAndDims()
    Dim view As Object
    Set view = swDoc.GetFirstView()
    Set view = view.GetNextView()

    Do While Not view Is Nothing
        Dim outline As Variant
        outline = view.GetOutline()
        viewName(numViews) = view.GetName2()
        viewXMin(numViews) = outline(0) : viewYMin(numViews) = outline(1)
        viewXMax(numViews) = outline(2) : viewYMax(numViews) = outline(3)

        Dim ann As Object
        Set ann = view.GetFirstAnnotation3()
        Do While Not ann Is Nothing
            If ann.GetType() = swDisplayDimension Then
                Dim dd As Object, d As Object, pos As Variant
                Set dd = ann.GetSpecificAnnotation()
                If Not dd Is Nothing Then
                    On Error Resume Next
                    Set d = dd.GetDimension()
                    pos = ann.GetPosition()
                    On Error GoTo 0
                    If Not d Is Nothing And Not IsEmpty(pos) And numDims < MAX_DIMS Then
                        dOwnerView(numDims) = view.GetName2()
                        dFullName(numDims)  = d.FullName
                        dX(numDims) = pos(0)
                        dY(numDims) = pos(1)
                        Set dAnnObj(numDims) = ann

                        ' Classify orientation using the ORIGINAL AutoDimension placement
                        If pos(0) < viewXMin(numViews) Or pos(0) > viewXMax(numViews) Then
                            dIsVertical(numDims) = True
                        Else
                            dIsVertical(numDims) = False
                        End If

                        numDims = numDims + 1
                    End If
                End If
            End If
            Set ann = ann.GetNext2()
        Loop

        numViews = numViews + 1
        Set view = view.GetNextView()
    Loop
End Sub


Sub CheckZoneOverlap()
    Dim a As Integer, b As Integer
    Print #fNum, "-- Zone overlap check (pad=" & Format(ZONE_PAD*1000,"0.0") & "mm) --"
    For a = 0 To numViews - 1
        For b = a + 1 To numViews - 1
            If Not (viewXMax(a) + ZONE_PAD < viewXMin(b) - ZONE_PAD Or _
                    viewXMin(a) - ZONE_PAD > viewXMax(b) + ZONE_PAD Or _
                    viewYMax(a) + ZONE_PAD < viewYMin(b) - ZONE_PAD Or _
                    viewYMin(a) - ZONE_PAD > viewYMax(b) + ZONE_PAD) Then
                Print #fNum, "  WARNING: zones of " & viewName(a) & " and " & viewName(b) & _
                             " overlap - views are too close for ZONE_PAD"
            End If
        Next b
    Next a
End Sub


' ─────────────────────────────────────────────────────────────────────
' Assign slots on TWO axes per view:
'   - horizontal-type dims (dIsVertical=False) -> stacked above/below,
'     split proportionally between the two sides by available capacity
'   - vertical-type dims (dIsVertical=True)     -> stacked left/right,
'     split proportionally the same way
' Every offset is capped at ZONE_PAD so nothing wanders past the
' view's territory, and further capped by physical room to the next
' view/sheet edge either way.
' ─────────────────────────────────────────────────────────────────────
Sub AssignSlots()
    Dim v As Integer, i As Integer, w As Integer
    Dim s As Integer, t As Integer, tmp As Integer

    Dim roomBelow As Double, roomAbove As Double
    Dim roomLeft  As Double, roomRight As Double
    Dim gapBelowRaw As Double, gapAboveRaw As Double
    Dim gapLeftRaw As Double, gapRightRaw As Double

    Dim effBelow As Double, effAbove As Double
    Dim effLeft  As Double, effRight As Double

    Dim capBelow As Integer, capAbove As Integer
    Dim capLeft  As Integer, capRight As Integer

    Dim idxs() As Integer
    Dim count As Integer

    Dim horizIdxs() As Integer, vertIdxs() As Integer
    Dim hCount As Integer, vCount As Integer

    Dim countBelow As Integer, countAbove As Integer
    Dim countLeft  As Integer, countRight As Integer
    Dim totalCapH As Integer, totalCapV As Integer

    Dim gapBelow As Double, gapAbove As Double
    Dim gapLeft  As Double, gapRight As Double

    Dim slot As Integer
    Dim targetY As Double, targetX As Double

    ReDim idxs(numDims)
    ReDim horizIdxs(numDims)
    ReDim vertIdxs(numDims)

    Print #fNum, ""
    Print #fNum, "-- Assigning slots (dual-axis, split above/below and left/right) --"

    For v = 0 To numViews - 1

        ' ===== Room available on each of the 4 sides =====
        roomBelow = viewYMin(v)
        For w = 0 To numViews - 1
            If w <> v Then
                If viewYMax(w) <= viewYMin(v) Then
                    gapBelowRaw = viewYMin(v) - viewYMax(w)
                    If gapBelowRaw < roomBelow Then roomBelow = gapBelowRaw
                End If
            End If
        Next w

        roomAbove = sheetH - viewYMax(v)
        For w = 0 To numViews - 1
            If w <> v Then
                If viewYMin(w) >= viewYMax(v) Then
                    gapAboveRaw = viewYMin(w) - viewYMax(v)
                    If gapAboveRaw < roomAbove Then roomAbove = gapAboveRaw
                End If
            End If
        Next w

        roomLeft = viewXMin(v)
        For w = 0 To numViews - 1
            If w <> v Then
                If viewXMax(w) <= viewXMin(v) Then
                    gapLeftRaw = viewXMin(v) - viewXMax(w)
                    If gapLeftRaw < roomLeft Then roomLeft = gapLeftRaw
                End If
            End If
        Next w

        roomRight = sheetW - viewXMax(v)
        For w = 0 To numViews - 1
            If w <> v Then
                If viewXMin(w) >= viewXMax(v) Then
                    gapRightRaw = viewXMin(w) - viewXMax(v)
                    If gapRightRaw < roomRight Then roomRight = gapRightRaw
                End If
            End If
        Next w

        ' ===== Cap every side to ZONE_PAD - never leave the view's territory =====
        effBelow = roomBelow : If effBelow > ZONE_PAD Then effBelow = ZONE_PAD
        effAbove = roomAbove : If effAbove > ZONE_PAD Then effAbove = ZONE_PAD
        effLeft  = roomLeft  : If effLeft  > ZONE_PAD Then effLeft  = ZONE_PAD
        effRight = roomRight : If effRight > ZONE_PAD Then effRight = ZONE_PAD

        ' ===== Capacity (slots that fit at normal DIM_GAP) on each side =====
        capBelow = Int((effBelow - MIN_OFFSET) / DIM_GAP) + 1 : If capBelow < 0 Then capBelow = 0
        capAbove = Int((effAbove - MIN_OFFSET) / DIM_GAP) + 1 : If capAbove < 0 Then capAbove = 0
        capLeft  = Int((effLeft  - MIN_OFFSET) / DIM_GAP) + 1 : If capLeft  < 0 Then capLeft  = 0
        capRight = Int((effRight - MIN_OFFSET) / DIM_GAP) + 1 : If capRight < 0 Then capRight = 0

        ' ===== Gather this view's dims, split by orientation =====
        count = 0
        For i = 0 To numDims - 1
            If dOwnerView(i) = viewName(v) Then
                idxs(count) = i : count = count + 1
            End If
        Next i
        For s = 0 To count - 2
            For t = 0 To count - 2 - s
                If dX(idxs(t)) > dX(idxs(t+1)) Then
                    tmp = idxs(t) : idxs(t) = idxs(t+1) : idxs(t+1) = tmp
                End If
            Next t
        Next s

        hCount = 0 : vCount = 0
        For s = 0 To count - 1
            i = idxs(s)
            If dIsVertical(i) Then
                vertIdxs(vCount) = i : vCount = vCount + 1
            Else
                horizIdxs(hCount) = i : hCount = hCount + 1
            End If
        Next s

        ' ===== Split horizontal-type dims between below/above =====
        totalCapH = capBelow + capAbove
        If totalCapH <= 0 Then
            countBelow = hCount : countAbove = 0
        ElseIf hCount <= totalCapH Then
            countBelow = Int(hCount * capBelow / totalCapH)
            If countBelow > capBelow Then countBelow = capBelow
            countAbove = hCount - countBelow
            If countAbove > capAbove Then
                countAbove = capAbove
                countBelow = hCount - countAbove
            End If
        Else
            countBelow = capBelow
            countAbove = hCount - countBelow
        End If

        gapBelow = DIM_GAP
        If countBelow > 0 And (MIN_OFFSET + countBelow * DIM_GAP) > effBelow Then
            gapBelow = (effBelow - MIN_OFFSET) / countBelow
            If gapBelow < 0.001 Then gapBelow = 0.001
        End If
        gapAbove = DIM_GAP
        If countAbove > 0 And (MIN_OFFSET + countAbove * DIM_GAP) > effAbove Then
            gapAbove = (effAbove - MIN_OFFSET) / countAbove
            If gapAbove < 0.001 Then gapAbove = 0.001
        End If

        ' ===== Split vertical-type dims between left/right =====
        totalCapV = capLeft + capRight
        If totalCapV <= 0 Then
            countRight = vCount : countLeft = 0
        ElseIf vCount <= totalCapV Then
            countRight = Int(vCount * capRight / totalCapV)
            If countRight > capRight Then countRight = capRight
            countLeft = vCount - countRight
            If countLeft > capLeft Then
                countLeft = capLeft
                countRight = vCount - countLeft
            End If
        Else
            countRight = capRight
            countLeft = vCount - countRight
        End If

        gapRight = DIM_GAP
        If countRight > 0 And (MIN_OFFSET + countRight * DIM_GAP) > effRight Then
            gapRight = (effRight - MIN_OFFSET) / countRight
            If gapRight < 0.001 Then gapRight = 0.001
        End If
        gapLeft = DIM_GAP
        If countLeft > 0 And (MIN_OFFSET + countLeft * DIM_GAP) > effLeft Then
            gapLeft = (effLeft - MIN_OFFSET) / countLeft
            If gapLeft < 0.001 Then gapLeft = 0.001
        End If

        Print #fNum, "  " & viewName(v) & ":  h=" & hCount & " (below=" & countBelow & _
                     " above=" & countAbove & ")   v=" & vCount & " (left=" & countLeft & _
                     " right=" & countRight & ")"

        ' ===== Place horizontal-type dims: below first, then above =====
        For s = 0 To hCount - 1
            i = horizIdxs(s)
            If s < countBelow Then
                slot = s
                targetY = viewYMin(v) - MIN_OFFSET - (slot * gapBelow)
            Else
                slot = s - countBelow
                targetY = viewYMax(v) + MIN_OFFSET + (slot * gapAbove)
            End If
            dAnnObj(i).SetPosition2 dX(i), targetY, 0
            Print #fNum, "    " & dFullName(i) & " -> y=" & Format(targetY*1000,"0.0") & _
                         IIf_S(s < countBelow, "  (below)", "  (above)")
        Next s

        ' ===== Place vertical-type dims: right first, then left =====
        For s = 0 To vCount - 1
            i = vertIdxs(s)
            If s < countRight Then
                slot = s
                targetX = viewXMax(v) + MIN_OFFSET + (slot * gapRight)
            Else
                slot = s - countRight
                targetX = viewXMin(v) - MIN_OFFSET - (slot * gapLeft)
            End If
            dAnnObj(i).SetPosition2 targetX, dY(i), 0
            Print #fNum, "    " & dFullName(i) & " -> x=" & Format(targetX*1000,"0.0") & _
                         IIf_S(s < countRight, "  (right)", "  (left)")
        Next s

    Next v
End Sub


' VBA's IIf always evaluates both branches, which is fine for numbers but
' annoying for readability with strings in a Print statement - small helper
' to keep the log lines above tidy.
Function IIf_S(cond As Boolean, a As String, b As String) As String
    If cond Then
        IIf_S = a
    Else
        IIf_S = b
    End If
End Function
