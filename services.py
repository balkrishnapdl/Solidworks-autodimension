"""
services.py
-----------
Orchestrates a batch run: feeds each uploaded STEP file through backend.py
sequentially, reporting progress back to the UI via a callback. This module
has no Streamlit dependency, so it can be reused as-is behind a FastAPI
endpoint later.
"""

from dataclasses import dataclass, field
from typing import Callable, Optional

import backend
from config import STATUS_RUNNING, STATUS_COMPLETED, STATUS_FAILED


@dataclass
class FileJob:
    name: str
    path: str
    status: str
    jpg_path: Optional[str] = None
    error: Optional[str] = None


@dataclass
class BatchSummary:
    total: int
    completed: int = 0
    failed: int = 0
    jobs: list = field(default_factory=list)


def run_batch(jobs: list, on_update: Callable[[int, FileJob], None]) -> BatchSummary:
    """Process every job sequentially through the existing SolidWorks
    pipeline, calling on_update(index, job) after every status change so
    the UI can re-render the file table and progress bar.

    jobs: list[FileJob] (status starts as STATUS_WAITING)
    """
    total = len(jobs)
    summary = BatchSummary(total=total, jobs=jobs)

    # Start SolidWorks once and reuse it for the whole batch.
    try:
        backend.init_environment()
        sw_app = backend.start_solidworks()
    except backend.BackendUnavailableError as exc:
        for i, job in enumerate(jobs):
            job.status = STATUS_FAILED
            job.error = str(exc)
            summary.failed += 1
            on_update(i, job)
        return summary

    for i, job in enumerate(jobs):
        job.status = STATUS_RUNNING
        on_update(i, job)

        try:
            jpg_path = backend.generate_drawing(sw_app, job.path)
            job.status = STATUS_COMPLETED
            job.jpg_path = jpg_path
            summary.completed += 1
        except Exception as exc:
            job.status = STATUS_FAILED
            job.error = str(exc)
            summary.failed += 1

        on_update(i, job)

    return summary
