"""
ui.py
-----
Every Streamlit rendering function lives here. app.py only calls into this
module and into services.py — it never builds HTML/CSS itself. Keeping
presentation isolated like this is what makes it easy to later swap this
whole layer for a React frontend without touching backend.py / services.py.
"""

import base64
import os

import streamlit as st

from config import THEME, APP_TITLE, APP_SUBTITLE, STATUS_COMPLETED, STATUS_FAILED, STATUS_RUNNING, STATUS_WAITING


# ──────────────────────────────────────────────────────────────────────────
# Global stylesheet
# ──────────────────────────────────────────────────────────────────────────

def inject_global_css():
    t = THEME
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');

        html, body, [class*="css"] {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }}

        .stApp {{
            background: {t['bg']};
        }}

        #MainMenu, header[data-testid="stHeader"], footer {{ visibility: hidden; height: 0; }}

        .block-container {{
            max-width: 880px;
            padding-top: 2.75rem;
            padding-bottom: 4rem;
        }}

        /* ── Header ────────────────────────────────────────────────────── */
        .adg-eyebrow {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-family: 'IBM Plex Mono', monospace;
            font-size: 0.72rem;
            letter-spacing: 0.08em;
            color: {t['accent']};
            background: {t['accent_soft']};
            border: 1px solid {t['accent']}22;
            padding: 4px 10px;
            border-radius: 100px;
            margin-bottom: 14px;
        }}
        .adg-title {{
            font-size: 2.05rem;
            font-weight: 700;
            color: {t['text']};
            letter-spacing: -0.02em;
            margin: 0 0 6px 0;
        }}
        .adg-subtitle {{
            font-size: 1rem;
            color: {t['text_muted']};
            margin: 0 0 28px 0;
            max-width: 560px;
            line-height: 1.5;
        }}

        /* ── Card shell ───────────────────────────────────────────────── */
        .adg-card {{
            background: {t['surface']};
            border: 1px solid {t['border']};
            border-radius: 16px;
            padding: 28px;
            box-shadow: 0 1px 2px rgba(16, 24, 32, 0.04), 0 8px 24px rgba(16, 24, 32, 0.03);
            margin-bottom: 20px;
        }}
        .adg-card-title {{
            font-size: 0.95rem;
            font-weight: 600;
            color: {t['text']};
            margin: 0 0 4px 0;
        }}
        .adg-card-hint {{
            font-size: 0.84rem;
            color: {t['text_muted']};
            margin: 0 0 18px 0;
        }}

        /* ── Uploader ─────────────────────────────────────────────────── */
        [data-testid="stFileUploaderDropzone"] {{
            background: {t['surface_alt']};
            border: 1.5px dashed {t['border']};
            border-radius: 12px;
            padding: 8px;
            transition: border-color 0.15s ease, background 0.15s ease;
        }}
        [data-testid="stFileUploaderDropzone"]:hover {{
            border-color: {t['accent']};
            background: {t['accent_soft']};
        }}
        [data-testid="stFileUploaderDropzoneInstructions"] svg {{
            display: none;
        }}

        /* ── Status badges ────────────────────────────────────────────── */
        .adg-badge {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-family: 'IBM Plex Mono', monospace;
            font-size: 0.72rem;
            font-weight: 600;
            letter-spacing: 0.02em;
            padding: 4px 10px;
            border-radius: 100px;
            white-space: nowrap;
        }}
        .adg-badge-dot {{ width: 6px; height: 6px; border-radius: 50%; }}
        .adg-badge-waiting   {{ background: {t['surface_alt']}; color: {t['waiting']}; border: 1px solid {t['border']}; }}
        .adg-badge-waiting .adg-badge-dot   {{ background: {t['waiting']}; }}
        .adg-badge-processing {{ background: {t['accent_soft']}; color: {t['accent_dark']}; border: 1px solid {t['accent']}33; }}
        .adg-badge-processing .adg-badge-dot {{ background: {t['accent']}; }}
        .adg-badge-completed {{ background: {t['success_soft']}; color: {t['success']}; border: 1px solid {t['success']}33; }}
        .adg-badge-completed .adg-badge-dot  {{ background: {t['success']}; }}
        .adg-badge-failed    {{ background: {t['error_soft']}; color: {t['error']}; border: 1px solid {t['error']}33; }}
        .adg-badge-failed .adg-badge-dot     {{ background: {t['error']}; }}

        /* ── File rows table ──────────────────────────────────────────── */
        .adg-row {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 11px 4px;
            border-bottom: 1px solid {t['border']};
        }}
        .adg-row:last-child {{ border-bottom: none; }}
        .adg-row-name {{
            font-size: 0.89rem;
            color: {t['text']};
            font-weight: 500;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            padding-right: 14px;
        }}

        /* ── Buttons ──────────────────────────────────────────────────── */
        div.stButton > button, div.stDownloadButton > button {{
            border-radius: 10px;
            font-weight: 600;
            font-size: 0.9rem;
            padding: 0.55rem 1.1rem;
            border: 1px solid {t['border']};
            transition: all 0.12s ease;
        }}
        div.stButton > button[kind="primary"] {{
            background: {t['accent']};
            border: 1px solid {t['accent']};
            color: white;
        }}
        div.stButton > button[kind="primary"]:hover {{
            background: {t['accent_dark']};
            border-color: {t['accent_dark']};
        }}
        div.stButton > button:disabled {{
            opacity: 0.5;
        }}

        /* ── Progress ─────────────────────────────────────────────────── */
        div[data-testid="stProgress"] div[role="progressbar"] > div {{
            background: {t['accent']};
        }}
        .adg-progress-label {{
            font-family: 'IBM Plex Mono', monospace;
            font-size: 0.8rem;
            color: {t['text_muted']};
            margin-bottom: 10px;
        }}

        /* ── Success / summary card ──────────────────────────────────── */
        .adg-success-card {{
            background: {t['success_soft']};
            border: 1px solid {t['success']}33;
            border-radius: 16px;
            padding: 24px 28px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 16px;
        }}
        .adg-success-icon {{
            width: 38px; height: 38px;
            border-radius: 50%;
            background: {t['success']};
            color: white;
            display: flex; align-items: center; justify-content: center;
            font-size: 1.1rem;
            flex-shrink: 0;
        }}
        .adg-success-title {{ font-size: 1rem; font-weight: 700; color: {t['text']}; margin: 0; }}
        .adg-success-sub {{ font-size: 0.86rem; color: {t['text_muted']}; margin: 2px 0 0 0; }}

        .adg-partial-card {{
            background: {t['error_soft']};
            border: 1px solid {t['error']}33;
        }}
        .adg-partial-card .adg-success-icon {{ background: {t['error']}; }}

        /* ── Gallery ──────────────────────────────────────────────────── */
        .adg-thumb-name {{
            font-size: 0.82rem;
            font-weight: 600;
            color: {t['text']};
            margin: 10px 0 8px 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        div[data-testid="stImage"] img {{
            border-radius: 10px;
            border: 1px solid {t['border']};
        }}

        section[data-testid="stSidebar"] {{ display: none; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────────────────────────────────
# Header
# ──────────────────────────────────────────────────────────────────────────

def render_header():
    st.markdown(
        f"""
        <div class="adg-eyebrow">● SOLIDWORKS AUTOMATION</div>
        <h1 class="adg-title">{APP_TITLE}</h1>
        <p class="adg-subtitle">{APP_SUBTITLE}</p>
        """,
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────────────────────────────────
# Upload card
# ──────────────────────────────────────────────────────────────────────────

def render_upload_card(disabled: bool):
    st.markdown('<div class="adg-card">', unsafe_allow_html=True)
    st.markdown('<p class="adg-card-title">Upload STEP files</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="adg-card-hint">Drag and drop one or more .step / .stp files, or browse from your computer.</p>',
        unsafe_allow_html=True,
    )
    files = st.file_uploader(
        'Upload STEP files',
        type=['step', 'stp'],
        accept_multiple_files=True,
        disabled=disabled,
        label_visibility='collapsed',
        key='uploader',
    )
    st.markdown('</div>', unsafe_allow_html=True)
    return files


# ──────────────────────────────────────────────────────────────────────────
# Status badge + file table
# ──────────────────────────────────────────────────────────────────────────

_BADGE_CLASS = {
    STATUS_WAITING:   'adg-badge-waiting',
    STATUS_RUNNING:   'adg-badge-processing',
    STATUS_COMPLETED: 'adg-badge-completed',
    STATUS_FAILED:    'adg-badge-failed',
}


def _badge_html(status: str) -> str:
    cls = _BADGE_CLASS.get(status, 'adg-badge-waiting')
    return f'<span class="adg-badge {cls}"><span class="adg-badge-dot"></span>{status}</span>'


def render_file_table(jobs):
    st.markdown('<div class="adg-card">', unsafe_allow_html=True)
    st.markdown(f'<p class="adg-card-title">Files ({len(jobs)})</p>', unsafe_allow_html=True)
    rows = ''.join(
        f'<div class="adg-row"><span class="adg-row-name">{job.name}</span>{_badge_html(job.status)}</div>'
        for job in jobs
    )
    st.markdown(rows, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────
# Progress section
# ──────────────────────────────────────────────────────────────────────────

def render_progress_section(done: int, total: int):
    st.markdown('<div class="adg-card">', unsafe_allow_html=True)
    st.markdown('<p class="adg-card-title">Generating drawings</p>', unsafe_allow_html=True)
    st.markdown(
        f'<p class="adg-progress-label">Processing {min(done + 1, total)} of {total} files</p>',
        unsafe_allow_html=True,
    )
    st.progress(done / total if total else 0)
    st.markdown('</div>', unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────
# Summary (success / partial) card
# ──────────────────────────────────────────────────────────────────────────

def render_summary_card(completed: int, failed: int, total: int):
    if failed == 0:
        st.markdown(
            f"""
            <div class="adg-success-card">
                <div class="adg-success-icon">✓</div>
                <div>
                    <p class="adg-success-title">Processing complete — {completed} of {total} drawings generated successfully.</p>
                    <p class="adg-success-sub">Every uploaded file produced a dimensioned drawing.</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""
            <div class="adg-success-card adg-partial-card">
                <div class="adg-success-icon">!</div>
                <div>
                    <p class="adg-success-title">Processing finished with errors — {completed} completed, {failed} failed.</p>
                    <p class="adg-success-sub">Re-upload the failed files to try again, or review them individually.</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ──────────────────────────────────────────────────────────────────────────
# Results gallery
# ──────────────────────────────────────────────────────────────────────────

def render_gallery(jobs):
    completed_jobs = [j for j in jobs if j.status == STATUS_COMPLETED and j.jpg_path and os.path.isfile(j.jpg_path)]
    if not completed_jobs:
        return

    st.markdown('<div class="adg-card">', unsafe_allow_html=True)
    st.markdown('<p class="adg-card-title">Generated drawings</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="adg-card-hint">Preview and download each drawing, or grab everything as a ZIP.</p>',
        unsafe_allow_html=True,
    )

    cols = st.columns(3)
    for i, job in enumerate(completed_jobs):
        with cols[i % 3]:
            st.image(job.jpg_path, use_container_width=True)
            st.markdown(f'<p class="adg-thumb-name">{job.name}</p>', unsafe_allow_html=True)
            with open(job.jpg_path, 'rb') as f:
                st.download_button(
                    'Download',
                    data=f.read(),
                    file_name=os.path.basename(job.jpg_path),
                    mime='image/jpeg',
                    key=f'dl_{i}_{job.name}',
                    use_container_width=True,
                )
    st.markdown('</div>', unsafe_allow_html=True)

    return [j.jpg_path for j in completed_jobs]


def render_failed_list(jobs):
    failed_jobs = [j for j in jobs if j.status == STATUS_FAILED]
    if not failed_jobs:
        return
    st.markdown('<div class="adg-card">', unsafe_allow_html=True)
    st.markdown('<p class="adg-card-title">Failed files</p>', unsafe_allow_html=True)
    for job in failed_jobs:
        st.markdown(
            f'<div class="adg-row"><span class="adg-row-name">{job.name}</span>{_badge_html(job.status)}</div>',
            unsafe_allow_html=True,
        )
    st.markdown('</div>', unsafe_allow_html=True)
