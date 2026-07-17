"""Load Scrum workbook tables used across slides."""

from __future__ import annotations

from pathlib import Path

from wsr.loaders import load_ddp_plan, load_tracker, load_visibility
from wsr.report.models import ScrumWorkbook
from wsr.run_log import RunLog
from wsr.tracker import tracker_lookup, tracker_rows_lookup


def load_scrum_workbook(scrum_path: Path, log: RunLog) -> ScrumWorkbook:
    log.info("Loading tracker, visibility, and DDP sheets…")
    tracker = load_tracker(str(scrum_path))
    visibility = load_visibility(str(scrum_path))
    ddp = load_ddp_plan(str(scrum_path))
    tracker_map = tracker_lookup(tracker)
    tracker_rows = tracker_rows_lookup(tracker)
    log.info(
        f"Tracker rows: {len(tracker)}; visibility rows: {len(visibility)}; "
        f"DCRs in tracker map: {len(tracker_map)}"
    )
    return ScrumWorkbook(
        path=scrum_path,
        tracker=tracker,
        visibility=visibility,
        ddp=ddp,
        tracker_map=tracker_map,
        tracker_rows=tracker_rows,
    )
