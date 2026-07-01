"""
app.py
------
Entry point: `streamlit run app.py`

This file only orchestrates session state and calls into ui.py (rendering)
and services.py (the batch run, which in turn calls backend.py). It holds
no CAD logic and no CSS — see backend.py and ui.py respectively.
"""

import streamlit as st

import ui
from config import APP_TITLE, STATUS_WAITING, STATUS_COMPLETED, STATUS_FAILED
from services import FileJob, run_batch
from utils import ensure_clean_dirs, is_step_file, save_uploaded_file, build_zip_archive


st.set_page_config(page_title=APP_TITLE, page_icon='📐', layout='centered')
ui.inject_global_css()


# ── Session state ────────────────────────────────────────────────────────
def _init_state():
    defaults = {
        'jobs': [],            # list[FileJob]
        'batch_running': False,
        'batch_done': False,
        'completed_count': 0,
        'failed_count': 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _start_new_batch():
    ensure_clean_dirs()
    st.session_state.jobs = []
    st.session_state.batch_running = False
    st.session_state.batch_done = False
    st.session_state.completed_count = 0
    st.session_state.failed_count = 0
    st.session_state.pop('uploader', None)


_init_state()
if not st.session_state.jobs and not st.session_state.batch_done:
    ensure_clean_dirs()

ui.render_header()


# ── Pre-batch: upload + file table + generate button ───────────────────────
if not st.session_state.batch_running and not st.session_state.batch_done:

    uploaded_files = ui.render_upload_card(disabled=False)

    valid_files = [f for f in (uploaded_files or []) if is_step_file(f.name)]

    if valid_files:
        names = [f.name for f in valid_files]
        # Keep job list in sync with whatever is currently in the uploader.
        st.session_state.jobs = [
            FileJob(name=name, path='', status=STATUS_WAITING) for name in names
        ]
        ui.render_file_table(st.session_state.jobs)

        st.button(
            'Generate Drawings',
            type='primary',
            use_container_width=True,
            disabled=len(valid_files) == 0,
            key='generate_btn',
        )

        if st.session_state.get('generate_btn'):
            ensure_clean_dirs()
            jobs = []
            for f in valid_files:
                path = save_uploaded_file(f)
                jobs.append(FileJob(name=f.name, path=path, status=STATUS_WAITING))
            st.session_state.jobs = jobs
            st.session_state.batch_running = True
            st.rerun()


# ── Batch running: locked UI, live progress ─────────────────────────────────
elif st.session_state.batch_running:

    ui.render_upload_card(disabled=True)

    table_slot = st.empty()
    progress_slot = st.empty()

    jobs = st.session_state.jobs
    total = len(jobs)

    with table_slot.container():
        ui.render_file_table(jobs)
    with progress_slot.container():
        ui.render_progress_section(0, total)

    def on_update(index, job):
        done = sum(1 for j in jobs if j.status in (STATUS_COMPLETED, STATUS_FAILED))
        with table_slot.container():
            ui.render_file_table(jobs)
        with progress_slot.container():
            ui.render_progress_section(done, total)

    summary = run_batch(jobs, on_update)

    st.session_state.jobs = summary.jobs
    st.session_state.completed_count = summary.completed
    st.session_state.failed_count = summary.failed
    st.session_state.batch_running = False
    st.session_state.batch_done = True
    st.rerun()


# ── Batch done: summary + gallery + downloads ───────────────────────────────
else:
    jobs = st.session_state.jobs
    ui.render_summary_card(
        st.session_state.completed_count, st.session_state.failed_count, len(jobs)
    )

    jpg_paths = ui.render_gallery(jobs) or []
    ui.render_failed_list(jobs)

    col1, col2 = st.columns(2)
    with col1:
        if jpg_paths:
            st.download_button(
                'Download All (ZIP)',
                data=build_zip_archive(jpg_paths),
                file_name='drawings.zip',
                mime='application/zip',
                use_container_width=True,
            )
    with col2:
        if st.button('Start New Batch', use_container_width=True):
            _start_new_batch()
            st.rerun()
