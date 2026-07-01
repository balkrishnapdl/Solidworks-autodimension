# Automatic Drawing Generator — Streamlit UI

A Streamlit front-end for the existing SolidWorks STEP-to-drawing automation
pipeline. This app does **not** reimplement any CAD logic — it only
orchestrates and monitors `sw_pipeline.py`, which performs all SolidWorks
work exactly as before (load STEP, save SLDPRT, open the drawing template,
insert views, auto-dimension, export JPG).

## Files

- `app.py` — Streamlit entry point (`streamlit run app.py`). Session state only.
- `ui.py` — All rendering: CSS, header, upload card, file table, progress, gallery.
- `services.py` — Batch orchestration loop, framework-agnostic (no Streamlit imports).
- `backend.py` — Thin adapter exposing `init_environment / start_solidworks / generate_drawing`.
- `sw_pipeline.py` — Your original automation code (one function parameterized: `import_step_to_sldprt` now accepts a STEP path instead of a hardcoded constant; `process_single_step_file` is a new wrapper that chains the existing per-file steps; `main()` is unchanged in behavior).
- `utils.py` — Filesystem helpers (save uploads, build ZIP, etc).
- `config.py` — All paths, status labels, and theme color tokens in one place.

## Running

Must run on the Windows machine with SolidWorks + pywin32 installed
(this is unchanged from your original script).

```
pip install -r requirements.txt
streamlit run app.py
```

Update the SolidWorks paths at the top of `config.py` (and `sw_pipeline.py`,
which keeps its own copies for standalone use) if they change.

## Notes on swapping the backend later

`services.py` and `backend.py` have zero Streamlit imports, so replacing
the local SolidWorks call with a FastAPI HTTP call later only means editing
`backend.generate_drawing()` — `services.py`, `ui.py`, and `app.py` stay as-is.
