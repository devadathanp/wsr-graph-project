"""Load Scrum workbook tables used across slides."""

from __future__ import annotations

from pathlib import Path

from wsr.constants import ACTION_ITEM_SHEET, RISK_SHEET, TRACKER_SHEET
from wsr.loaders import load_action_items, load_ddp_plan, load_risks, load_tracker, load_visibility
from wsr.report.models import ScrumWorkbook
from wsr.rich_text import build_rich_run_index
from wsr.run_log import RunLog
from wsr.tracker import tracker_lookup, tracker_rows_lookup


def load_scrum_workbook(scrum_path: Path, log: RunLog) -> ScrumWorkbook:
    log.info("Loading tracker, visibility, DDP, risks, and action-item sheets…")
    tracker = load_tracker(str(scrum_path))
    visibility = load_visibility(str(scrum_path))
    ddp = load_ddp_plan(str(scrum_path))
    risks = load_risks(str(scrum_path))
    actions = load_action_items(str(scrum_path))
    tracker_map = tracker_lookup(tracker)
    tracker_rows = tracker_rows_lookup(tracker)

    log.info("Indexing Excel rich-text / strikethrough cells…")
    rich_runs: dict[str, dict[tuple[int, str], list[tuple[str, bool]]]] = {}
    for sheet_name in (TRACKER_SHEET, ACTION_ITEM_SHEET, RISK_SHEET):
        try:
            sheet_runs = build_rich_run_index(scrum_path, sheet_name)
            rich_runs[sheet_name] = sheet_runs
            log.info(f"  {sheet_name}: {len(sheet_runs)} strikethrough cell(s)")
        except Exception as exc:
            log.warning(f"Could not index rich text on '{sheet_name}': {exc}")

    log.info(
        f"Tracker rows: {len(tracker)}; visibility rows: {len(visibility)}; "
        f"risk rows: {len(risks)}; action-item rows: {len(actions)}; "
        f"DCRs in tracker map: {len(tracker_map)}"
    )
    return ScrumWorkbook(
        path=scrum_path,
        tracker=tracker,
        visibility=visibility,
        ddp=ddp,
        risks=risks,
        action_items=actions,
        tracker_map=tracker_map,
        tracker_rows=tracker_rows,
        rich_runs=rich_runs,
    )
