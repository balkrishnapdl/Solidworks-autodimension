Option Explicit

Dim swApp As Object
Dim swDoc As Object
Dim fNum  As Integer

Const FILLET_MAX_R    As Double = 0.05
Const CHAMFER_MAX_LEN As Double = 0.010
Const ANGLE_TOL       As Double = 10
Const DIM_GAP         As Double = 0.01
Const SHEET_MARGIN    As Double = 0.005
Const MAX_FEAT        As Integer = 50

Dim numFillets               As Integer
Dim filletR(MAX_FEAT)        As Double
Dim filletBestView(MAX_FEAT) As String
Dim filletBestCount(MAX_FEAT)  As Integer
Dim filletTotalCount(MAX_FEAT) As Integer
Dim filletCX(MAX_FEAT) As Double
Dim filletCY(MAX_FEAT) As Double
Dim filletCZ(MAX_FEAT) As Double

Dim numChamfers                 As Integer
Dim chamferSz(MAX_FEAT)         As Double
Dim chamferBestView(MAX_FEAT)   As String
Dim chamferBestCount(MAX_FEAT)  As Integer
Dim chamferTotalCount(MAX_FEAT) As Integer
Dim chamferMX(MAX_FEAT) As Double
Dim chamferMY(MAX_FEAT) As Double
Dim chamferMZ(MAX_FEAT) As Double

Dim sheetW As Double, sheetH As Double


Sub main()
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc

    fNum = FreeFile()
    Open "{LOGPATH}" For Output As #fNum
    Print #fNum, "Fillet + Chamfer Annotation (2-pass)"
    Print #fNum, "======================================="

    Dim sp As Variant
    sp = swDoc.GetCurrentSheet().GetProperties()
    sheetW = sp(5)
    sheetH = sp(6)
    Print #fNum, "Sheet: " & Format(sheetW*1000,"0.0") & " x " & Format(sheetH*1000,"0.0") & " mm"

    numFillets  = 0
    numChamfers = 0

    Print #fNum, ""
    Print #fNum, "-- Pass 1: scanning --"
    Dim view As Object
    Set view = swDoc.GetFirstView()
    Set view = view.GetNextView()

    Do While Not view Is Nothing
        Call ScanView(view, view.GetName2())
        Set view = view.GetNextView()
    Loop

    Dim i As Integer
    Print #fNum, "Fillets unique: " & numFillets
    For i = 0 To numFillets - 1
        Print #fNum, "  R=" & Format(filletR(i)*1000,"0.00") & _
                     "mm  bestView=" & filletBestView(i) & _
                     "(" & filletBestCount(i) & ")  total=" & filletTotalCount(i)
    Next i
    Print #fNum, "Chamfers unique: " & numChamfers
    For i = 0 To numChamfers - 1
        Print #fNum, "  C=" & Format(chamferSz(i)*1000,"0.00") & _
                     "mm  bestView=" & chamferBestView(i) & _
                     "(" & chamferBestCount(i) & ")  total=" & chamferTotalCount(i)
    Next i

    Print #fNum, ""
    Print #fNum, "-- Pass 2: annotating --"
    Set view = swDoc.GetFirstView()
    Set view = view.GetNextView()

    Do While Not view Is Nothing
        Dim vName As String
        vName = view.GetName2()

        Dim outline As Variant
        outline = view.GetOutline()
        Dim xMin As Double, yMin As Double, xMax As Double, yMax As Double
        xMin = outline(0) : yMin = outline(1) : xMax = outline(2) : yMax = outline(3)

        swDoc.ActivateView vName
        swDoc.ClearSelection2 True
        swDoc.ForceRebuild3 False

        Dim belowRoom As Double, aboveRoom As Double
        belowRoom = yMin - SHEET_MARGIN
        aboveRoom = (sheetH - SHEET_MARGIN) - yMax

        Dim hOff As Double : hOff = DIM_GAP

        For i = 0 To numFillets - 1
            If filletBestView(i) = vName Then
                Call PlaceFilletDim(view, i, xMin, yMin, xMax, yMax, _
                                    belowRoom, aboveRoom, hOff)
                hOff = hOff + DIM_GAP
            End If
        Next i

        For i = 0 To numChamfers - 1
            If chamferBestView(i) = vName Then
                Call PlaceChamferDim(view, i)
            End If
        Next i

        Set view = view.GetNextView()
    Loop

    swDoc.ForceRebuild3 False
    Print #fNum, ""
    Print #fNum, "Done."
    Close #fNum
End Sub


' ─────────────────────────────────────────────────────────────────────────────
' ScanView — unchanged from original
' ─────────────────────────────────────────────────────────────────────────────
Sub ScanView(view As Object, vName As String)
    Dim entities As Variant
    On Error Resume Next
    entities = view.GetVisibleEntities2(Nothing, 1)
    On Error GoTo 0
    If IsEmpty(entities) Then Exit Sub

    Dim lFN As Integer : lFN = 0
    Dim lFR(MAX_FEAT) As Double, lFC(MAX_FEAT) As Integer
    Dim lFCX(MAX_FEAT) As Double, lFCY(MAX_FEAT) As Double, lFCZ(MAX_FEAT) As Double

    Dim lCN As Integer : lCN = 0
    Dim lCS(MAX_FEAT) As Double, lCC(MAX_FEAT) As Integer
    Dim lCMX(MAX_FEAT) As Double, lCMY(MAX_FEAT) As Double, lCMZ(MAX_FEAT) As Double

    Dim i As Integer
    For i = 0 To UBound(entities)
        Dim ent As Object
        Set ent = entities(i)
        If ent Is Nothing Then GoTo NextEnt

        Dim crv As Object
        On Error Resume Next
        Set crv = ent.GetCurve()
        On Error GoTo 0
        If crv Is Nothing Then GoTo NextEnt

        If crv.IsCircle() Then
            Dim svC As Object, evC As Object
            On Error Resume Next
            Set svC = ent.GetStartVertex()
            Set evC = ent.GetEndVertex()
            On Error GoTo 0
            If svC Is Nothing And evC Is Nothing Then GoTo NextEnt

            Dim cp As Variant
            cp = crv.CircleParams()
            Dim r As Double : r = cp(6)
            r = Int(r * 100000 + 0.5) / 100000

            If r > 0.00001 And r <= FILLET_MAX_R Then
                Dim fi As Integer, foundF As Boolean : foundF = False
                For fi = 0 To lFN - 1
                    If Abs(lFR(fi) - r) < 0.000025 Then
                        lFC(fi) = lFC(fi) + 1 : foundF = True : Exit For
                    End If
                Next fi
                If Not foundF And lFN < MAX_FEAT Then
                    lFR(lFN) = r : lFC(lFN) = 1
                    lFCX(lFN) = cp(0) : lFCY(lFN) = cp(1) : lFCZ(lFN) = cp(2)
                    lFN = lFN + 1
                End If
            End If

        ElseIf crv.IsLine() Then
            Dim sv As Object, ev As Object
            On Error Resume Next
            Set sv = ent.GetStartVertex() : Set ev = ent.GetEndVertex()
            On Error GoTo 0
            If sv Is Nothing Or ev Is Nothing Then GoTo NextEnt

            Dim spt As Variant, ept As Variant
            spt = sv.GetPoint() : ept = ev.GetPoint()

            Dim dx As Double, dy As Double, elen As Double
            dx = ept(0)-spt(0) : dy = ept(1)-spt(1)
            elen = Sqr(dx*dx + dy*dy + (ept(2)-spt(2))^2)

            If elen >= 0.0001 And elen <= CHAMFER_MAX_LEN Then
                Dim ang As Double
                If Abs(dx) < 0.0000001 Then
                    ang = 90
                Else
                    ang = Abs(Atn(dy/dx)) * 180 / 3.14159265
                End If
                If ang >= ANGLE_TOL And ang <= (90 - ANGLE_TOL) Then
                    Dim cs As Double
                    cs = elen / Sqr(2)
                    cs = Int(cs * 20000 + 0.5) / 20000

                    Dim ci As Integer, foundC As Boolean : foundC = False
                    For ci = 0 To lCN - 1
                        If Abs(lCS(ci) - cs) < 0.000025 Then
                            lCC(ci) = lCC(ci) + 1 : foundC = True : Exit For
                        End If
                    Next ci
                    If Not foundC And lCN < MAX_FEAT Then
                        lCS(lCN) = cs : lCC(lCN) = 1
                        lCMX(lCN) = (spt(0)+ept(0))/2
                        lCMY(lCN) = (spt(1)+ept(1))/2
                        lCMZ(lCN) = (spt(2)+ept(2))/2
                        lCN = lCN + 1
                    End If
                End If
            End If
        End If
NextEnt:
    Next i

    Dim gi As Integer
    For fi = 0 To lFN - 1
        Dim gFoundF As Boolean : gFoundF = False
        For gi = 0 To numFillets - 1
            If Abs(filletR(gi) - lFR(fi)) < 0.000025 Then
                filletTotalCount(gi) = filletTotalCount(gi) + lFC(fi)
                If lFC(fi) > filletBestCount(gi) Then
                    filletBestCount(gi) = lFC(fi) : filletBestView(gi) = vName
                    filletCX(gi) = lFCX(fi) : filletCY(gi) = lFCY(fi) : filletCZ(gi) = lFCZ(fi)
                End If
                gFoundF = True : Exit For
            End If
        Next gi
        If Not gFoundF And numFillets < MAX_FEAT Then
            filletR(numFillets)          = lFR(fi)
            filletBestCount(numFillets)  = lFC(fi)
            filletBestView(numFillets)   = vName
            filletTotalCount(numFillets) = lFC(fi)
            filletCX(numFillets) = lFCX(fi)
            filletCY(numFillets) = lFCY(fi)
            filletCZ(numFillets) = lFCZ(fi)
            numFillets = numFillets + 1
        End If
    Next fi

    For ci = 0 To lCN - 1
        Dim gFoundC As Boolean : gFoundC = False
        For gi = 0 To numChamfers - 1
            If Abs(chamferSz(gi) - lCS(ci)) < 0.000025 Then
                chamferTotalCount(gi) = chamferTotalCount(gi) + lCC(ci)
                If lCC(ci) > chamferBestCount(gi) Then
                    chamferBestCount(gi) = lCC(ci) : chamferBestView(gi) = vName
                    chamferMX(gi) = lCMX(ci) : chamferMY(gi) = lCMY(ci) : chamferMZ(gi) = lCMZ(ci)
                End If
                gFoundC = True : Exit For
            End If
        Next gi
        If Not gFoundC And numChamfers < MAX_FEAT Then
            chamferSz(numChamfers)          = lCS(ci)
            chamferBestCount(numChamfers)   = lCC(ci)
            chamferBestView(numChamfers)    = vName
            chamferTotalCount(numChamfers)  = lCC(ci)
            chamferMX(numChamfers) = lCMX(ci)
            chamferMY(numChamfers) = lCMY(ci)
            chamferMZ(numChamfers) = lCMZ(ci)
            numChamfers = numChamfers + 1
        End If
    Next ci

    Print #fNum, "  " & vName & ": " & lFN & " fillet size(s), " & lCN & " chamfer size(s)"
End Sub


' ─────────────────────────────────────────────────────────────────────────────
' PlaceFilletDim — unchanged from original
' ─────────────────────────────────────────────────────────────────────────────
Sub PlaceFilletDim(view As Object, idx As Integer, _
                   xMin As Double, yMin As Double, _
                   xMax As Double, yMax As Double, _
                   belowRoom As Double, aboveRoom As Double, _
                   hOff As Double)

    Dim ptX As Double, ptY As Double, ptZ As Double
    ptX = filletCX(idx) + filletR(idx)
    ptY = filletCY(idx)
    ptZ = filletCZ(idx)

    Dim labelX As Double, labelY As Double
    labelX = ptX
    If belowRoom >= DIM_GAP Then
        labelY = yMin - hOff
    Else
        labelY = yMax + hOff
    End If
    If labelX < SHEET_MARGIN + 0.003 Then labelX = SHEET_MARGIN + 0.003
    If labelX > sheetW - SHEET_MARGIN - 0.003 Then labelX = sheetW - SHEET_MARGIN - 0.003
    If labelY < SHEET_MARGIN + 0.003 Then labelY = SHEET_MARGIN + 0.003
    If labelY > sheetH - SHEET_MARGIN - 0.003 Then labelY = sheetH - SHEET_MARGIN - 0.003

    swDoc.ClearSelection2 True
    Dim dimObj As Object
    Set dimObj = swDoc.AddDimension2(ptX, ptY, ptZ)

    If Not dimObj Is Nothing Then
        Dim ann As Object
        Set ann = dimObj.GetAnnotation()
        If Not ann Is Nothing Then ann.SetPosition2 labelX, labelY, 0
        Print #fNum, "  [OK] Fillet R=" & Format(filletR(idx)*1000,"0.00") & "mm -> " & filletBestView(idx)
    Else
        Print #fNum, "  [!!] Fillet R=" & Format(filletR(idx)*1000,"0.00") & "mm AddDimension2 failed"
    End If
End Sub


' ─────────────────────────────────────────────────────────────────────────────
' PlaceChamferDim — NEW: replaces PlaceChamferNote
' Mimics: Smart Dimension toolbar → click chamfer edge → click neighbour edge
' SolidWorks then produces the proper chamfer dimension (e.g. 2 x 45deg)
' ─────────────────────────────────────────────────────────────────────────────
Sub PlaceChamferDim(view As Object, idx As Integer)

    Const VTX_TOL As Double = 0.000002   ' 2 µm

    ' Re-query all edges visible in this view
    Dim entities As Variant
    On Error Resume Next
    entities = view.GetVisibleEntities2(Nothing, 1)
    On Error GoTo 0
    If IsEmpty(entities) Then
        Print #fNum, "  [!!] No entities - skip chamfer C=" & Format(chamferSz(idx)*1000,"0.##") & "mm"
        Exit Sub
    End If

    ' ── Step 1: find the chamfer edge object by matching its stored midpoint ──
    Dim chamfEnt As Object
    Dim cSX As Double, cSY As Double   ' start vertex of chamfer edge
    Dim cEX As Double, cEY As Double   ' end vertex of chamfer edge
    Dim i As Integer

    For i = 0 To UBound(entities)
        Dim ent As Object : Set ent = entities(i)
        If ent Is Nothing Then GoTo FindChamf_Next

        Dim crv As Object
        On Error Resume Next : Set crv = ent.GetCurve() : On Error GoTo 0
        If crv Is Nothing Then GoTo FindChamf_Next
        If Not crv.IsLine() Then GoTo FindChamf_Next

        Dim sv As Object, ev As Object
        On Error Resume Next
        Set sv = ent.GetStartVertex() : Set ev = ent.GetEndVertex()
        On Error GoTo 0
        If sv Is Nothing Or ev Is Nothing Then GoTo FindChamf_Next

        Dim spt As Variant, ept As Variant
        spt = sv.GetPoint() : ept = ev.GetPoint()

        ' Match midpoint to the one stored in ScanView
        If Abs((spt(0)+ept(0))/2 - chamferMX(idx)) < VTX_TOL And _
           Abs((spt(1)+ept(1))/2 - chamferMY(idx)) < VTX_TOL Then
            Set chamfEnt = ent
            cSX = spt(0) : cSY = spt(1)
            cEX = ept(0) : cEY = ept(1)
            Exit For
        End If
FindChamf_Next:
    Next i

    If chamfEnt Is Nothing Then
        Print #fNum, "  [!!] Chamfer edge not found for C=" & Format(chamferSz(idx)*1000,"0.##") & "mm"
        Exit Sub
    End If

    ' ── Step 2: find the neighbouring straight edge (shares a vertex, longer) ─
    Dim neighEnt As Object
    Dim neighPickX As Double, neighPickY As Double
    Dim bestLen As Double : bestLen = 999999

    For i = 0 To UBound(entities)
        Set ent = entities(i)
        If ent Is Nothing Then GoTo FindNeigh_Next
        If ent Is chamfEnt Then GoTo FindNeigh_Next

        On Error Resume Next : Set crv = ent.GetCurve() : On Error GoTo 0
        If crv Is Nothing Then GoTo FindNeigh_Next
        If Not crv.IsLine() Then GoTo FindNeigh_Next

        On Error Resume Next
        Set sv = ent.GetStartVertex() : Set ev = ent.GetEndVertex()
        On Error GoTo 0
        If sv Is Nothing Or ev Is Nothing Then GoTo FindNeigh_Next

        spt = sv.GetPoint() : ept = ev.GetPoint()
        Dim elen As Double
        elen = Sqr((ept(0)-spt(0))^2 + (ept(1)-spt(1))^2 + (ept(2)-spt(2))^2)

        ' Must be a real face edge, not another chamfer
        If elen <= CHAMFER_MAX_LEN Then GoTo FindNeigh_Next

        ' Must share exactly one vertex with the chamfer edge
        Dim sharedX As Double, sharedY As Double
        Dim farX As Double, farY As Double
        Dim adj As Boolean : adj = False

        If (Abs(spt(0)-cSX) < VTX_TOL And Abs(spt(1)-cSY) < VTX_TOL) Or _
           (Abs(spt(0)-cEX) < VTX_TOL And Abs(spt(1)-cEY) < VTX_TOL) Then
            adj = True : sharedX = spt(0) : sharedY = spt(1)
            farX = ept(0) : farY = ept(1)
        ElseIf (Abs(ept(0)-cSX) < VTX_TOL And Abs(ept(1)-cSY) < VTX_TOL) Or _
               (Abs(ept(0)-cEX) < VTX_TOL And Abs(ept(1)-cEY) < VTX_TOL) Then
            adj = True : sharedX = ept(0) : sharedY = ept(1)
            farX = spt(0) : farY = spt(1)
        End If

        ' Among all adjacent edges pick the shortest (closest parent face edge)
        If adj And elen < bestLen Then
            bestLen  = elen
            Set neighEnt = ent
            ' Pick safely at 33% from shared vertex - well away from the junction
            neighPickX = sharedX + 0.33 * (farX - sharedX)
            neighPickY = sharedY + 0.33 * (farY - sharedY)
        End If
FindNeigh_Next:
    Next i

    ' ── Step 3: select chamfer edge (mark=1) ─────────────────────────────────
    swDoc.ClearSelection2 True

    Dim ok1 As Boolean
    ok1 = swDoc.Extension.SelectByRay( _
        chamferMX(idx), chamferMY(idx), 0.001, _  ' ray hits chamfer midpoint
        0, 0, -1, 0.0004, _                        ' direction down, 0.4mm radius
        2, False, 1, 0)                            ' swSelEDGES  Append=False  Mark=1

    Print #fNum, "  Select chamfer edge  (mark=1): " & ok1

    ' ── Step 4: select neighbouring edge (mark=2, append) ────────────────────
    Dim ok2 As Boolean : ok2 = False
    If Not neighEnt Is Nothing Then
        ok2 = swDoc.Extension.SelectByRay( _
            neighPickX, neighPickY, 0.001, _
            0, 0, -1, 0.0004, _
            2, True, 2, 0)                         ' swSelEDGES  Append=True  Mark=2
        Print #fNum, "  Select neighbour edge (mark=2): " & ok2
    Else
        Print #fNum, "  WARNING: neighbour edge not found"
    End If

    ' ── Step 5: AddDimension2 - with chamfer+neighbour selected SW creates ────
    '            the proper chamfer dim, not a plain linear one
    If ok1 Then
        Dim dimObj As Object
        Set dimObj = swDoc.AddDimension2(chamferMX(idx), chamferMY(idx), chamferMZ(idx))
        If Not dimObj Is Nothing Then
            Print #fNum, "  [OK] Chamfer dim placed: C=" & Format(chamferSz(idx)*1000,"0.##") & "mm"
        Else
            Print #fNum, "  [!!] AddDimension2 returned Nothing"
        End If
    Else
        Print #fNum, "  [!!] Chamfer edge select failed - skipped"
    End If
End Sub
