"""
config.py
---------
Central place for paths, constants and small tunables used across the app.
Edit this file when SolidWorks paths, the drawing template, or output
folders change on the host machine.
"""

import os

# ── Backend / SolidWorks paths ──────────────────────────────────────────────
# These are passed straight through to sw_pipeline.py — keep them in sync
# with that module's own constants if you edit both.
INPUT_PARTS      = r'C:\Users\PC-09\Desktop\New folder\PairDrop_files_20260623_1531'
OUTPUT_PARTS     = r'C:\Users\PC-09\Desktop\New folder\output_images'
OUTPUT_DRAWINGS  = r'C:\Users\PC-09\Desktop\New folder\output_images'
DRAWING_TEMPLATE = r'C:\Users\PC-09\Desktop\New folder\template.DRWDOT'
SW_PATH          = r'D:\SOLIDWORKS\2024\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe'

# ── App-local working directories ───────────────────────────────────────────
APP_DIR        = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR     = os.path.join(APP_DIR, '.tmp_uploads')
RESULTS_DIR    = os.path.join(APP_DIR, '.tmp_results')

ALLOWED_EXTENSIONS = ('.step', '.stp')

# ── UI text ──────────────────────────────────────────────────────────────────
APP_TITLE    = 'Automatic Drawing Generator'
APP_SUBTITLE = 'Converts STEP files into fully dimensioned engineering drawings — automatically.'

# ── Status labels (single source of truth, used by ui.py + services.py) ────
STATUS_WAITING   = 'Waiting'
STATUS_RUNNING   = 'Processing'
STATUS_COMPLETED = 'Completed'
STATUS_FAILED    = 'Failed'

# ── Theme tokens (referenced by ui.py when building the CSS block) ─────────
THEME = {
    'bg':            '#F4F6F8',
    'surface':       '#FFFFFF',
    'surface_alt':   '#FAFBFC',
    'border':        '#E1E6EB',
    'text':          '#1B2630',
    'text_muted':    '#67737E',
    'accent':        '#0F5DA8',
    'accent_dark':   '#0B4783',
    'accent_soft':   '#E7F0FA',
    'success':       '#1A8A53',
    'success_soft':  '#E6F6EE',
    'error':         '#C7402F',
    'error_soft':    '#FBEAE7',
    'waiting':       '#8893A0',
}
