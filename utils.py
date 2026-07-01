"""
utils.py
--------
Small, dependency-free helper functions shared by services.py and ui.py.
Nothing in here talks to SolidWorks or Streamlit directly.
"""

import io
import os
import re
import shutil
import zipfile

from config import UPLOAD_DIR, RESULTS_DIR, ALLOWED_EXTENSIONS


def ensure_clean_dirs():
    """Create (or reset) the app's scratch directories for a fresh batch."""
    for d in (UPLOAD_DIR, RESULTS_DIR):
        if os.path.exists(d):
            shutil.rmtree(d, ignore_errors=True)
        os.makedirs(d, exist_ok=True)


def is_step_file(filename: str) -> bool:
    return filename.lower().endswith(ALLOWED_EXTENSIONS)


def safe_filename(filename: str) -> str:
    """Strip path separators and odd characters so the name is safe to use
    as a file on disk."""
    name = os.path.basename(filename)
    name = re.sub(r'[^A-Za-z0-9._\-]+', '_', name)
    return name


def save_uploaded_file(uploaded_file) -> str:
    """Persist a Streamlit UploadedFile to UPLOAD_DIR and return its path."""
    dest_name = safe_filename(uploaded_file.name)
    dest_path = os.path.join(UPLOAD_DIR, dest_name)
    with open(dest_path, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    return dest_path


def build_zip_archive(jpg_paths: list) -> bytes:
    """Bundle a list of JPG file paths into an in-memory ZIP and return its
    bytes, ready to hand to st.download_button."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        for path in jpg_paths:
            if os.path.isfile(path):
                zf.write(path, arcname=os.path.basename(path))
    buffer.seek(0)
    return buffer.getvalue()


def read_bytes(path: str) -> bytes:
    with open(path, 'rb') as f:
        return f.read()
