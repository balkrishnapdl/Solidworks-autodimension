Option Explicit

Dim swApp As Object
Dim swDoc As Object
Dim fNum  As Integer

Sub main()
    Set swApp = Application.SldWorks
    Set swDoc = swApp.ActiveDoc

    fNum = FreeFile()
    Open "{LOGPATH}" For Output As #fNum
    Print #fNum, "Clamp Annotation Positions"
    Print #fNum, "=========================="

    ' Get sheet dimensions
    Dim sp As Variant
    sp = swDoc.GetCurrentSheet().GetProperties()
    Dim sheetW As Double : sheetW = sp(5)
    Dim sheetH As Double : sheetH = sp(6)
    Dim mg As Double : mg = 0.010  ' 10 mm margin from sheet edge
    Print #fNum, "Sheet: " & Format(sheetW*1000,"0.0") & " x " & Format(sheetH*1000,"0.0") & " mm"
    Print #fNum, "Margin: " & Format(mg*1000,"0.0") & " mm"

    Dim movedCount As Integer : movedCount = 0
    Dim totalCount As Integer : totalCount = 0

    ' Iterate every view (including sheet pseudo-view) and clamp each annotation
    Dim view As Object
    Set view = swDoc.GetFirstView()

    Do While Not view Is Nothing
        Dim ann As Object
        Set ann = view.GetFirstAnnotation2()

        Do While Not ann Is Nothing
            totalCount = totalCount + 1

            Dim pos As Variant
            On Error Resume Next
            pos = ann.GetPosition()
            On Error GoTo 0

            If Not IsEmpty(pos) Then
                Dim px As Double : px = pos(0)
                Dim py As Double : py = pos(1)
                Dim pz As Double : pz = pos(2)

                Dim newX As Double : newX = px
                Dim newY As Double : newY = py
                Dim changed As Boolean : changed = False

                ' Clamp X
                If px < mg Then
                    newX = mg : changed = True
                ElseIf px > sheetW - mg Then
                    newX = sheetW - mg : changed = True
                End If

                ' Clamp Y
                If py < mg Then
                    newY = mg : changed = True
                ElseIf py > sheetH - mg Then
                    newY = sheetH - mg : changed = True
                End If

                If changed Then
                    On Error Resume Next
                    ann.SetPosition2 newX, newY, pz
                    On Error GoTo 0
                    movedCount = movedCount + 1
                    Print #fNum, "  Moved: (" & Format(px*1000,"0.0") & "," & Format(py*1000,"0.0") & _
                                 ") -> (" & Format(newX*1000,"0.0") & "," & Format(newY*1000,"0.0") & ")"
                End If
            End If

            Set ann = ann.GetNext2()
        Loop

        Set view = view.GetNextView()
    Loop

    swDoc.ForceRebuild3 False
    Print #fNum, ""
    Print #fNum, "Total annotations: " & totalCount & "  Repositioned: " & movedCount
    Print #fNum, "Done."
    Close #fNum
End Sub
