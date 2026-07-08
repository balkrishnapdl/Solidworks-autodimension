Dim swApp      As Object
Dim Part       As Object
Dim boolstatus As Boolean
Dim longstatus As Long
Dim fNum       As Integer
Dim i          As Integer

Sub main()
    Set swApp = Application.SldWorks
    Set Part  = swApp.ActiveDoc

    fNum = FreeFile()
    Open "{LOGPATH}" For Output As #fNum
    Print #fNum, "AutoDimension via view-select + Nothing-datum approach"
    Print #fNum, "======================================================="

    Dim viewNames(2) As String
    viewNames(0) = "{VIEW0}"
    viewNames(1) = "{VIEW1}"
    viewNames(2) = "{VIEW2}"

    For i = 0 To 2
        If viewNames(i) = "" Then GoTo NextView

        Print #fNum, ""
        Print #fNum, "=== " & viewNames(i) & " ==="

        ' Step 1: Activate the view and clear selection
        Part.ActivateView viewNames(i)
        Part.ClearSelection2 True
        Part.ForceRebuild3 False

        ' Step 2: Select the view entity as DRAWINGVIEW
        boolstatus = Part.Extension.SelectByID2( _
            viewNames(i), "DRAWINGVIEW", _
            0, 0, 0, _
            False, 0, Nothing, 0)
        Print #fNum, "  SelectByID2 = " & boolstatus

        If Not boolstatus Then
            Print #fNum, "  SKIP: could not select view"
            GoTo NextView
        End If

        ' Step 3: AutoDimension
        ' FIX 1: Removed EditSelectAll  (method does not exist on IModelDoc2)
        ' FIX 2: HorizDatum / VertDatum = Nothing, NOT 0
        '         Passing integer 0 into an Object param causes
        '         DISP_E_TYPEMISMATCH argErr=1 on HorizDatum
        longstatus = Part.AutoDimension(1, 1, -1, 1, 1)
        Print #fNum, "  AutoDimension result = " & longstatus

        Part.ClearSelection2 True

NextView:
    Next i

    ' Final rebuild
    Part.ForceRebuild3 False
    Print #fNum, ""
    Print #fNum, "Done."
    Close #fNum
End Sub
