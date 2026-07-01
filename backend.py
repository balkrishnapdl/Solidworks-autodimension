"""
backend.py
----------
Thin adapter around the existing SolidWorks automation pipeline
(sw_pipeline.py). This module intentionally does NOT reimplement any CAD
logic — it only exposes a small, stable interface that services.py can call:

    init_environment()           -> create output dirs, sanity-check paths
    start_solidworks()           -> launch/attach SolidWorks, return handle
    generate_drawing(swApp, p)   -> run the full per-file pipeline, return jpg path

Keeping this layer thin means the underlying sw_pipeline.py can be swapped
out (e.g. for a FastAPI microservice call) later without touching the UI.
"""

from config import OUTPUT_PARTS, OUTPUT_DRAWINGS


class BackendUnavailableError(RuntimeError):
    """Raised when the SolidWorks automation backend cannot be loaded —
    e.g. running this app on a non-Windows machine without SolidWorks/
    pywin32 installed."""
    pass


def _load_pipeline():
    try:
        import sw_pipeline
        return sw_pipeline
    except Exception as exc:  # pragma: no cover - environment dependent
        raise BackendUnavailableError(
            'The SolidWorks automation backend could not be loaded. '
            'This application must run on a Windows machine with SolidWorks '
            'and pywin32 installed.'
        ) from exc


def init_environment():
    """Verify configured paths and create output directories.
    Mirrors sw_pipeline.check_configuration()."""
    pipeline = _load_pipeline()
    pipeline.check_configuration()


def start_solidworks():
    """Launch or attach to a running SolidWorks instance. Returns the
    swApp COM handle, reused across every file in the batch so SolidWorks
    only has to start once."""
    pipeline = _load_pipeline()
    return pipeline.launch_or_attach_solidworks()


def generate_drawing(sw_app, step_file_path: str) -> str:
    """Run the full existing pipeline for a single STEP file:
    import -> SLDPRT -> open template -> views -> auto-dimension ->
    export JPG. Returns the path to the generated JPG.

    All intermediate CAD steps remain internal to sw_pipeline.py; this
    function only ever returns the final artifact path.
    """
    pipeline = _load_pipeline()
    return pipeline.process_single_step_file(sw_app, step_file_path)


def output_dirs():
    return OUTPUT_PARTS, OUTPUT_DRAWINGS
