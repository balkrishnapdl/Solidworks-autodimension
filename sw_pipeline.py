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


MACRO_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'macros')

SW_SAVE_ERRORS = {
    0: 'OK',
    1: 'swFileSaveError — generic failure',
    2: 'swFileSaveAsDoNotOverwrite — file exists',
    3: 'swFileSaveAsInvalidFileExtension',
    5: 'swFileSaveAsNameExceedsMaxPathLength',
    8: 'swFileSaveRequiresSavingReferences',
}


def load_macro(filename, **replacements):

    path = os.path.join(MACRO_DIR, filename)
    with open(path, 'r', encoding='utf-8') as f:
        code = f.read()

    for key, value in replacements.items():
        code = code.replace('{' + key + '}', value)

    return code.encode('ascii', errors='replace').decode('ascii')


def write_macro_file(filename, **replacements):

    log_path = os.path.join(
        tempfile.gettempdir(),
        os.path.splitext(filename)[0] + '_log.txt'
    )
    macro_path = os.path.join(
        tempfile.gettempdir(),
        os.path.splitext(filename)[0] + '.swb'
    )

    replacements.setdefault('LOGPATH', log_path.replace('\\', '\\\\'))
    macro_code = load_macro(filename, **replacements)

    with open(macro_path, 'w', encoding='ascii') as f:
        f.write(macro_code)

    return macro_path, log_path


def check_configuration():
    """Verifies the template / SW exe / test STEP file exist, and creates
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

    missing_macros = [
        name for name in (
            'autodim.bas', 'fillet_chamfer.bas', 'clamp_annotations.bas',
            'dim_layout.bas', 'collision_check.bas',
        )
        if not os.path.isfile(os.path.join(MACRO_DIR, name))
    ]
    if missing_macros:
        print(f'  *** MISSING MACRO FILES in {MACRO_DIR}: {missing_macros} ***')
    else:
        print(f'Macro files OK ({MACRO_DIR})')


def launch_or_attach_solidworks():
    """Tries to attach to an already-running SolidWorks instance; if none is
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
    """Calls NewDocument and reads back the COM interface type, verifying
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
    """ Polls a view until it reports visible edges,
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
    ready_views = []
    for name, v in views:
        ok = wait_for_view_geometry(swDraw, v, name)
        if ok:
            ready_views.append((name, v))

    print(f'\nViews with geometry: {len(ready_views)} / {len(views)}')
    return ready_views



def diagnose_views(ready_views):
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


def _get_view_name(view_obj):
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
    ortho_views = [(n, v) for n, v in ready_views if 'isometric' not in n.lower()]
    view_names = ['', '', '']
    for idx, (_, view_obj) in enumerate(ortho_views[:3]):
        actual = _get_view_name(view_obj)
        if actual:
            view_names[idx] = actual

    print(f"View names resolved: {view_names}")

    macro_path, log_path = write_macro_file(
        'autodim.bas',
        VIEW0=view_names[0],
        VIEW1=view_names[1],
        VIEW2=view_names[2],
    )

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


def run_fillet_chamfer_macro(swApp, swDraw):
    macro_path, log_path = write_macro_file('fillet_chamfer.bas')

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


def run_clamp_annotations_macro(swApp, swDraw):
    macro_path, log_path = write_macro_file('clamp_annotations.bas')

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


def run_dimension_layout_macro(swApp, swDraw):
    macro_path, log_path = write_macro_file('dim_layout.bas')

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


def run_collision_check_macro(swApp, swDraw):
    macro_path, log_path = write_macro_file('collision_check.bas')

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
    print('Imports OK')
    check_configuration()
    swApp = launch_or_attach_solidworks()
    jpg_path = process_single_step_file(swApp, TEST_STEP_FILE)
    print(f'\nDone. Drawing exported to: {jpg_path}')


if __name__ == '__main__':
    main()
