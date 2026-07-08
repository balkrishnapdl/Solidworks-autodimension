import os
import time
import tempfile
import subprocess

import pythoncom
import win32com.client


INPUT_PARTS      = r'D:\New folder\PairDrop_files_20260623_1531'
OUTPUT_PARTS     = r'D:\New folder\output_images'
OUTPUT_DRAWINGS  = r'D:\New folder\output_images'
DRAWING_TEMPLATE = r'D:\New folder\template.DRWDOT'
SW_PATH          = r'D:\SOLIDWORKS\2024\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe'


TEST_STEP_FILE   = r'D:\New folder\PairDrop_files_20260623_1531\flat_complex.step'

AUTO_DIMENSION_ENABLED       = True
AUTO_HOLE_DIMENSIONS_ENABLED = True
TRY_PREDEFINED_VIEWS_FIRST   = True
MAX_WAIT_VIEW_SECONDS        = 40
POLL_INTERVAL                = 1.0

VIEW_ENTITY_TYPE_EDGE   = 1
VIEW_ENTITY_TYPE_VERTEX = 3
SW_HORIZONTAL_ORDINATE  = 5
SW_VERTICAL_ORDINATE    = 6
ORDINATE_LABEL_OFFSET   = 0.025

SW_SAVE_ERRORS = {
    0: 'OK',
    1: 'swFileSaveError — generic failure',
    2: 'swFileSaveAsDoNotOverwrite — file exists',
    3: 'swFileSaveAsInvalidFileExtension',
    5: 'swFileSaveAsNameExceedsMaxPathLength',
    8: 'swFileSaveRequiresSavingReferences',
}


def check_configuration():
    """Cell 2 — Configuration (edit paths here).

    Verifies the template / SW exe / test STEP file exist, and creates
    the output directories.
    """
    for label, path in [('TEMPLATE', DRAWING_TEMPLATE),
                         ('SW_EXE',   SW_PATH),
                         ('TEST_STEP', TEST_STEP_FILE)]:
        exists = os.path.isfile(path)
        print(f'  {label}: {"EXISTS" if exists else "*** NOT FOUND ***"}  →  {path}')

    os.makedirs(OUTPUT_PARTS,    exist_ok=True)
    os.makedirs(OUTPUT_DRAWINGS, exist_ok=True)
    print('Directories OK')


def launch_or_attach_solidworks():
    """Cell 3 — Launch / attach SolidWorks.

    Tries to attach to an already-running SolidWorks instance; if none is
    running, launches it and polls until it becomes responsive.
    """
    pythoncom.CoInitialize()

    swApp = None
    try:
        swApp = win32com.client.GetActiveObject(SW_PATH)
        swApp.Visible = True
        print('Attached to running SolidWorks.')
    except Exception:
        print('No running instance — launching...')
        subprocess.Popen([SW_PATH])
        for attempt in range(60):
            time.sleep(30)
            try:
                candidate = win32com.client.Dispatch('SldWorks.Application')
                candidate.Visible = True
                _ = candidate.RevisionNumber
                swApp = candidate
                print(f'SolidWorks ready after ~{attempt + 1} s.')
                break
            except Exception:
                pass

    if swApp is None:
        raise RuntimeError('Could not connect to SolidWorks.')

    # Quick health probe
    print(f'  Revision  : {swApp.RevisionNumber}')
    print(f'  Doc count : {swApp.GetDocumentCount}')
    print(f'  Binding   : {type(swApp)}')

    return swApp


def import_step_to_sldprt(swApp, step_file_path=None):
    """Import a STEP file and save it as SLDPRT.

    step_file_path: path to the source .STEP/.STP file. Defaults to the
    module-level TEST_STEP_FILE for backwards compatibility with the
    original notebook-style script.
    """
    step_file_path = step_file_path or TEST_STEP_FILE

    base_name = os.path.splitext(os.path.basename(step_file_path))[0]
    part_path = os.path.join(OUTPUT_PARTS, base_name + '.SLDPRT')

    print(f'Importing STEP : {step_file_path}')
    print(f'STEP exists    : {os.path.isfile(step_file_path)}')
    print(f'Target SLDPRT  : {part_path}')

    load_errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    null_prefs  = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

    loaded_ok = swApp.LoadFile4(step_file_path, 'r', null_prefs, load_errors)

    print(f'  LoadFile4 ok      : {loaded_ok}')
    print(f'  LoadFile4 errors  : {load_errors.value}')

    part_doc = swApp.ActiveDoc if loaded_ok else None
    print(f'  ActiveDoc after LoadFile4 : {part_doc}')

    
    if part_doc is None:
        print('  ActiveDoc is None — trying GetOpenDocumentByName...')
        part_doc = swApp.GetDocuments(step_file_path)
        print(f'  GetOpenDocumentByName result: {part_doc}')

    if part_doc is None:
        raise RuntimeError(
            'Could not load STEP file via LoadFile4 — '
            f'loaded_ok={loaded_ok}, errors={load_errors.value}'
        )

    
    doc_type = part_doc.GetType
    print(f'  doc type: {doc_type}  (1=part ✓  2=assembly  3=drawing)')
    assert doc_type == 1, f'Expected a part doc (1) but got type {doc_type}'

    try:
        part_doc.ForceRebuild3(False)
    except Exception:
        pass

    save_errors   = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    save_warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    null_export   = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

    ok = part_doc.Extension.SaveAs(
        part_path,
        0,            # Version  — 0 = current SW version
        1,            # Options  — 1 = swSaveAsOptions_Silent (no blocking dialogs)
        null_export,  # ExportData — proper COM null, native SLDPRT needs none
        save_errors,
        save_warnings
    )
    print(f'  Extension.SaveAs -> SLDPRT : ok={ok}  errors={save_errors.value}  warnings={save_warnings.value}')
    print(f'  error meaning : {SW_SAVE_ERRORS.get(save_errors.value, f"code {save_errors.value}")}')

    if not ok or not os.path.isfile(part_path):
        raise RuntimeError(
            f'Failed to save imported STEP as SLDPRT at {part_path} '
            f'(ok={ok}, errors={save_errors.value}, warnings={save_warnings.value})'
        )

    print(f'\nPart path : {part_path}')

    return part_doc, part_path


def open_drawing_template(swApp):
    """Cell 5 — Open drawing template.

    Calls NewDocument and reads back the COM interface type, verifying
    it is a drawing (type 3) and probing for CreateDrawViewFromModelView3.
    """
    print(f'Template : {DRAWING_TEMPLATE}')
    print(f'Exists   : {os.path.isfile(DRAWING_TEMPLATE)}')

    # ── Step A: call NewDocument ───────────────────────────────────────────
    raw_doc = swApp.NewDocument(DRAWING_TEMPLATE, 0, 0.0, 0.0)
    time.sleep(1.0)   # give SW a moment to fully open it

    print('\nraw_doc (NewDocument return value):')
    print(f'  is None  : {raw_doc is None}')
    print(f'  type     : {type(raw_doc)}')

    # ── Step B: get ActiveDoc (IDrawingDoc-capable reference) ─────────────
    swDraw = swApp.ActiveDoc
    print('\nswDraw (swApp.ActiveDoc):')
    print(f'  is None  : {swDraw is None}')
    print(f'  type     : {type(swDraw)}')

    # ── Step C: verify it is a drawing (type 3) ────────────────────────────
    try:
        doc_type = swDraw.GetType
        print(f'  GetType  : {doc_type}  (3 = drawing ✓)')
        if doc_type != 3:
            print('  *** NOT A DRAWING — check your template file! ***')
    except Exception as e:
        print(f'  GetType failed: {e}')

    # ── Step D: probe whether CreateDrawViewFromModelView3 is visible ─────
    print('\nProbing CreateDrawViewFromModelView3 on raw_doc (NewDocument return):')
    try:
        fn = getattr(raw_doc, 'CreateDrawViewFromModelView3', 'MISSING')
        print(f'  raw_doc  : {fn}')
    except Exception as e:
        print(f'  raw_doc  error: {e}')

    print('Probing CreateDrawViewFromModelView3 on swDraw (ActiveDoc):')
    try:
        fn = getattr(swDraw, 'CreateDrawViewFromModelView3', 'MISSING')
        print(f'  swDraw   : {fn}')
    except Exception as e:
        print(f'  swDraw   error: {e}')

    return swDraw



def create_views(swDraw, part_path):

    # Read actual sheet dimensions first
    sheet = swDraw.GetCurrentSheet
    props = sheet.GetProperties
    # props[0]=scale_num, props[1]=scale_den, props[2]=type,
    # props[3]=template, props[4]=width, props[5]=height
    sheet_w = props[5]
    sheet_h = props[6]
    print(f'Sheet size: {sheet_w * 1000:.1f} mm × {sheet_h * 1000:.1f} mm')

    # Recompute layout to fit inside the sheet
    margin = 0.015
    col1_x = margin + sheet_w * 0.18
    col2_x = margin + sheet_w * 0.68
    row1_y = margin + sheet_h * 0.20
    row2_y = margin + sheet_h * 0.70

    view_layout_fitted = [
        ('*Front',     col1_x, row1_y),
        ('*Top',       col1_x, row2_y),
        ('*Right',     col2_x, row1_y),
        # ('*Isometric', col2_x, row2_y),
    ]
    print('Fitted layout:')
    for name, x, y in view_layout_fitted:
        print(f'  {name:12s}  x={x * 1000:.1f}mm  y={y * 1000:.1f}mm')

    # Now create views with fitted coordinates
    views = []
    for view_name, x, y in view_layout_fitted:
        print(f'\n  Trying "{view_name}" at ({x:.4f}, {y:.4f})...')
        try:
            v = swDraw.CreateDrawViewFromModelView3(part_path, view_name, x, y, 0.0)
            if v is not None:
                views.append((view_name, v))
                print('    ✓ Created')
            else:
                print('    ✗ Returned None')
        except Exception as e:
            print(f'    ✗ Exception: {e}')

    try:
        swDraw.ForceRebuild3(False)
    except Exception:
        pass

    print(f'\nViews created: {len(views)} / {len(view_layout_fitted)}')

    return views


def wait_for_view_geometry(swDraw, view, view_name):
    """Helper used by Cell 7. Polls a view until it reports visible edges,
    or until MAX_WAIT_VIEW_SECONDS elapses."""
    deadline = time.time() + MAX_WAIT_VIEW_SECONDS
    attempt = 0
    while time.time() < deadline:
        try:
            swDraw.ActivateView(view_name)
        except Exception:
            pass
        try:
            swDraw.ForceRebuild3(False)
        except Exception:
            pass
        try:
            comps = view.GetVisibleComponents
            comp_list = list(comps) if comps else [None]
            for comp in comp_list:
                edges = view.GetVisibleEntities2(comp, VIEW_ENTITY_TYPE_EDGE)
                if edges and len(list(edges)) > 0:
                    print(f'  ✓ "{view_name}" has edges after ~{attempt} s')
                    return True
        except Exception:
            pass
        attempt += 1
        print(f'  Waiting for "{view_name}" ({attempt} s)...')
        time.sleep(POLL_INTERVAL)
    print(f'  ✗ TIMEOUT: "{view_name}" never showed edges.')
    return False


def wait_for_all_views(swDraw, views):
    """Cell 7 — Wait for view geometry (edges) to appear."""
    ready_views = []
    for name, v in views:
        ok = wait_for_view_geometry(swDraw, v, name)
        if ok:
            ready_views.append((name, v))

    print(f'\nViews with geometry: {len(ready_views)} / {len(views)}')
    return ready_views



def diagnose_views(ready_views):
    """Cell 8 — Diagnose edges & vertices per view."""
    for name, v in ready_views:
        edge_count = vertex_count = 0
        try:
            comps = v.GetVisibleComponents
            comp_list = list(comps) if comps else [None]
            for comp in comp_list:
                try:
                    e = v.GetVisibleEntities2(comp, VIEW_ENTITY_TYPE_EDGE)
                    if e:
                        edge_count += len(list(e))
                except Exception:
                    pass
                try:
                    verts = v.GetVisibleEntities2(comp, VIEW_ENTITY_TYPE_VERTEX)
                    if verts:
                        vertex_count += len(list(verts))
                except Exception:
                    pass
        except Exception as e:
            print(f'  {name}: GetVisibleComponents error: {e}')
        print(f'  {name:12s} → edges: {edge_count:4d}   vertices: {vertex_count:4d}')



_VBA_AUTODIM = '''\
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
'''


def _get_view_name(view_obj):
    """Helper used by Cell 9 to resolve a view's display name."""
    for attr in ('GetName2', 'Name'):
        try:
            val = getattr(view_obj, attr)
            name = val() if callable(val) else val
            if isinstance(name, str) and name:
                return name
        except Exception:
            pass
    return None


def run_autodimension_macro(swApp, swDraw, ready_views):
    """Cell 9 — Autodimension.

    Builds and runs a VBA macro (via RunMacro2) that activates each ortho
    view, selects it, and calls AutoDimension.
    """
    # ── Collect ortho view names (exclude isometric) ───────────────────────
    ortho_views = [(n, v) for n, v in ready_views if 'sometric' not in n.lower()]
    view_names = ['', '', '']
    for idx, (_, view_obj) in enumerate(ortho_views[:3]):
        actual = _get_view_name(view_obj)
        if actual:
            view_names[idx] = actual

    print(f"View names resolved: {view_names}")

    # ── Write macro ──────────────────────────────────────────────────────
    log_path = os.path.join(tempfile.gettempdir(), 'sw_autodim_log.txt')
    macro_path = os.path.join(tempfile.gettempdir(), 'sw_autodim.swb')

    macro_code = (_VBA_AUTODIM
                  .replace('{LOGPATH}', log_path.replace('\\', '\\\\'))
                  .replace('{VIEW0}', view_names[0])
                  .replace('{VIEW1}', view_names[1])
                  .replace('{VIEW2}', view_names[2]))

    macro_code = macro_code.encode('ascii', errors='replace').decode('ascii')

    with open(macro_path, 'w', encoding='ascii') as f:
        f.write(macro_code)

    print(f"Macro written : {macro_path}")
    print(f"Log will be   : {log_path}")

    # ── Run macro ────────────────────────────────────────────────────────
    macro_errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    result = swApp.RunMacro2(macro_path, '', 'main', 0, macro_errors)
    print(f"RunMacro2  result   : {result}")
    print(f"RunMacro2 errors   : {macro_errors.value}")

    # ── Print log ────────────────────────────────────────────────────────
    time.sleep(0.8)
    if os.path.exists(log_path):
        with open(log_path, 'r') as f:
            print(f.read())
    else:
        print("Log file not found — macro may have crashed before writing it")

    swDraw.ForceRebuild3(False)



_VBA_FILLET_CHAMFER = r'''
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
'''


def run_fillet_chamfer_macro(swApp, swDraw):
    """Cell 9 (continued) — Fillet + Chamfer annotation."""
    log_path = os.path.join(tempfile.gettempdir(), 'sw_fillet_chamfer_log.txt')
    macro_path = os.path.join(tempfile.gettempdir(), 'sw_fillet_chamfer.swb')

    macro_code = _VBA_FILLET_CHAMFER.replace('{LOGPATH}', log_path.replace('\\', '\\\\'))
    macro_code = macro_code.encode('ascii', errors='replace').decode('ascii')

    with open(macro_path, 'w', encoding='ascii') as f:
        f.write(macro_code)

    print(f'Macro written : {macro_path}')

    macro_errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    result = swApp.RunMacro2(macro_path, '', 'main', 0, macro_errors)
    print(f'RunMacro2 result : {result}')
    print(f'RunMacro2 errors : {macro_errors.value}')

    time.sleep(1.0)
    if os.path.exists(log_path):
        with open(log_path) as f:
            print(f.read())
    else:
        print('Log not found — macro may have crashed before writing it')

    swDraw.ForceRebuild3(False)



_VBA_CLAMP = r'''
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
'''


def run_clamp_annotations_macro(swApp, swDraw):
    """Cell 9 (continued) — Clamp annotation positions to stay on-sheet."""
    log_path = os.path.join(tempfile.gettempdir(), 'sw_clamp_ann_log.txt')
    macro_path = os.path.join(tempfile.gettempdir(), 'sw_clamp_ann.swb')

    # NOTE: preserved exactly as in the original notebook cell, including the
    # (functionally inert, since LOGPATH has no backslashes after the first
    # .replace pass) double-escaped replace call.
    macro_code = _VBA_CLAMP.replace('{LOGPATH}', log_path.replace('\\\\', '\\\\\\\\'))
    macro_code = macro_code.encode('ascii', errors='replace').decode('ascii')

    with open(macro_path, 'w', encoding='ascii') as f:
        f.write(macro_code)

    print(f'Macro written : {macro_path}')

    macro_errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    result = swApp.RunMacro2(macro_path, '', 'main', 0, macro_errors)
    print(f'RunMacro2 result : {result}')
    print(f'RunMacro2 errors : {macro_errors.value}')

    time.sleep(0.5)
    if os.path.exists(log_path):
        with open(log_path) as f:
            print(f.read())
    else:
        print('Log not found — macro may have crashed')

    swDraw.ForceRebuild3(False)


_VBA_DIM_LAYOUT = r'''
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
'''


def run_dimension_layout_macro(swApp, swDraw):
    """Cell 9 (continued) — Orientation-aware, dual-axis dimension layout.

    Classifies each dimension as 'vertical-type' (height dims, originally
    placed to the side of their view) or 'horizontal-type' (width dims,
    originally placed above/below their view), based on AutoDimension's
    own initial placement. Horizontal-type dims are stacked above/below
    the view; vertical-type dims are stacked left/right of the view.
    The dimension count on each axis is split between its two sides
    proportional to how much room (capped at ZONE_PAD) each side has,
    so neither side gets overloaded the way a single-side stack did.

    Run BEFORE the collision checker, and re-run run_clamp_annotations_macro
    AFTER this so nothing that lands near a sheet edge gets pushed past it.
    """
    log_path = os.path.join(tempfile.gettempdir(), 'sw_dim_layout_log.txt')
    macro_path = os.path.join(tempfile.gettempdir(), 'sw_dim_layout.swb')

    macro_code = _VBA_DIM_LAYOUT.replace('{LOGPATH}', log_path.replace('\\', '\\\\'))
    macro_code = macro_code.encode('ascii', errors='replace').decode('ascii')

    with open(macro_path, 'w', encoding='ascii') as f:
        f.write(macro_code)

    print(f'Macro written : {macro_path}')

    macro_errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    result = swApp.RunMacro2(macro_path, '', 'main', 0, macro_errors)
    print(f'RunMacro2 result : {result}')
    print(f'RunMacro2 errors : {macro_errors.value}')

    time.sleep(0.8)
    if os.path.exists(log_path):
        with open(log_path, 'r') as f:
            print(f.read())
    else:
        print('Log not found — macro may have crashed before writing it')

    swDraw.ForceRebuild3(False)

_VBA_COLLISION_CHECK = r'''
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
'''
 

 
def run_collision_check_macro(swApp, swDraw):
    """Cell 9 (continued) — Position-based dimension collision check.
 
    Checks:
      1) Any two dimension POSITIONS (from AutoDimension, fillet/chamfer
         annotation, anywhere on the sheet) closer than COLLISION_RADIUS.
      2) Any dimension position that lands inside a neighboring view's
         outline instead of its own.
 
    On detection, nudges the offending dimension down by a fixed step
    (single pass — no re-check loop yet).
 
    No bounding box (GetBox2) and no tolerance data are read — only
    ann.GetPosition() — since GetBox2 was erroring 438 on every dimension.
    """
    log_path = os.path.join(tempfile.gettempdir(), 'sw_collision_check_log.txt')
    macro_path = os.path.join(tempfile.gettempdir(), 'sw_collision_check.swb')
 
    macro_code = _VBA_COLLISION_CHECK.replace('{LOGPATH}', log_path.replace('\\', '\\\\'))
    macro_code = macro_code.encode('ascii', errors='replace').decode('ascii')
 
    with open(macro_path, 'w', encoding='ascii') as f:
        f.write(macro_code)
 
    print(f'Macro written : {macro_path}')
    print(f'Log will be   : {log_path}')
 
    macro_errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    result = swApp.RunMacro2(macro_path, '', 'main', 0, macro_errors)
    print(f'RunMacro2 result : {result}')
    print(f'RunMacro2 errors : {macro_errors.value}')
 
    time.sleep(0.8)
    if os.path.exists(log_path):
        with open(log_path, 'r') as f:
            print(f.read())
    else:
        print('Log not found — macro may have crashed before writing it')
 
    swDraw.ForceRebuild3(False)

def save_drawing_and_export_jpeg(swApp, part_path):
    """Cell 10 (part A) — Save .slddrw and close.

    Re-fetches ActiveDoc (the macros in Cell 9 may have shifted SW focus),
    saves the drawing as .SLDDRW, then exports a JPEG via Extension.SaveAs.
    Returns (swDraw, jpg_path) for use by the close step.
    """
    # Re-fetch active doc — macro in Cell 9 may have shifted SW focus
    swDraw = swApp.ActiveDoc
    assert swDraw.GetType == 3, "Active doc is not a drawing!"

    base_name = os.path.splitext(os.path.basename(part_path))[0]

    # ── Step 1: Save as SLDDRW first (gives the doc a real path on disk) ───
    drw_path = os.path.join(OUTPUT_DRAWINGS, base_name + '.SLDDRW')
    try:
        ok_drw = swDraw.SaveAs3(drw_path, 0, 0)
        print(f'SaveAs SLDDRW : ok={bool(ok_drw)}  →  {drw_path}')
    except Exception as e:
        print(f'SaveAs SLDDRW failed: {e}')

    # ── Step 2: Export JPEG via Extension.SaveAs (6-param, avoids PicCallback)
    # Use VARIANT(VT_DISPATCH, None) for null COM object — Python None causes
    # "Type mismatch" when COM expects a dispatch pointer.
    jpg_path  = os.path.join(OUTPUT_PARTS, base_name + '.jpg')
    errors    = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings  = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    null_disp = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

    try:
        ok = swDraw.Extension.SaveAs(
            jpg_path,
            0,          # Version  — 0 = current SW version
            1,          # Options  — 1 = swSaveAsOptions_Silent
            null_disp,  # ExportData — proper COM null (JPEG needs no data object)
            errors,
            warnings
        )
        print(f'SaveAs JPEG : ok={ok}  errors={errors.value}  warnings={warnings.value}')
        print(f'  error meaning : {SW_SAVE_ERRORS.get(errors.value, f"code {errors.value}")}')

    except Exception as e:
        print(f'Extension.SaveAs failed: {e}')
        # Fallback — options=1 for silent, so SW does not block on a dialog
        try:
            ok = swDraw.SaveAs3(jpg_path, 0, 1)
            print(f'Fallback IModelDoc2.SaveAs3 : ok={bool(ok)}')
        except Exception as e2:
            print(f'Fallback also failed: {e2}')

    # ── Verify ───────────────────────────────────────────────────────────
    if os.path.isfile(jpg_path):
        print(f'✓ JPEG saved : {jpg_path}  ({os.path.getsize(jpg_path) / 1024:.1f} KB)')
    else:
        print(f'✗ JPEG not found : {jpg_path}')

    return swDraw, jpg_path


def close_documents(swApp, swDraw, part_doc):
    """Cell 10 (part B) — Close drawing first, then the part.

    Always close drawing before part — SW will complain if you close the
    part while a drawing that references it is still open.
    """
    try:
        drw_title = swDraw.GetTitle
        swApp.CloseDoc(drw_title)
        print(f'✓ Closed drawing : {drw_title}')
    except Exception as e:
        print(f'✗ CloseDoc drawing failed: {e}')

    try:
        part_title = part_doc.GetTitle
        swApp.CloseDoc(part_title)
        print(f'✓ Closed part    : {part_title}')
    except Exception as e:
        print(f'✗ CloseDoc part failed: {e}')



def process_single_step_file(swApp, step_file_path):
    """Run the full pipeline (import -> views -> dimension -> export JPG)
    for a single STEP file, using an already-running SolidWorks instance.

    Returns the path to the generated JPG drawing.
    """
    part_doc, part_path = import_step_to_sldprt(swApp, step_file_path)
    swDraw = open_drawing_template(swApp)
    views = create_views(swDraw, part_path)
    ready_views = wait_for_all_views(swDraw, views)
    diagnose_views(ready_views)

    run_autodimension_macro(swApp, swDraw, ready_views)
    run_fillet_chamfer_macro(swApp, swDraw)
    run_clamp_annotations_macro(swApp, swDraw)
    run_dimension_layout_macro(swApp, swDraw)
    run_clamp_annotations_macro(swApp, swDraw) 
    run_collision_check_macro(swApp, swDraw) 

    swDraw, jpg_path = save_drawing_and_export_jpeg(swApp, part_path)
    close_documents(swApp, swDraw, part_doc)

    if not os.path.isfile(jpg_path):
        raise RuntimeError(f'JPG export did not produce a file at {jpg_path}')

    return jpg_path


def main():
    """Original single-file entry point, preserved for standalone/manual use."""
    print('Imports OK')
    check_configuration()
    swApp = launch_or_attach_solidworks()
    jpg_path = process_single_step_file(swApp, TEST_STEP_FILE)
    print(f'\nDone. Drawing exported to: {jpg_path}')


if __name__ == '__main__':
    main()