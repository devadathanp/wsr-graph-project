"""Pending DCR selection for slides 5 and 6."""

from __future__ import annotations

import pandas as pd

from wsr.constants import DEFAULT_DATA_FILE, PENDING_TABLE_ROW_CAP
from wsr.graph import get_evaluation_data, get_implementation_data
from wsr.tracker import (
    closure_date_from_row,
    coerce_tracker_date,
    eval_status_from_row,
    impl_status_from_row,
    latest_comment,
    support_required_from_row,
    tracker_row_for_mode,
    visibility_row,
)

PLAN_COMPLETION_COL = "Planned Completion Date\n<dd-mm-yyyy>"

_PENDING_PRCR_STATE = {
    "evaluation": "Evaluate",
    "implementation": "Implement",
}
_PENDING_AT_RISK = "On Track"


def build_pending_item(
    dcr_id: int,
    tracker_rows: dict[int, list[pd.Series]],
    visibility: pd.DataFrame,
    mode: str,
    tracker_row: pd.Series | None = None,
) -> dict | None:
    if tracker_row is None:
        tracker_row = tracker_row_for_mode(tracker_rows, dcr_id, mode)
    if tracker_row is None:
        return None
    vis_row = visibility_row(visibility, dcr_id)
    summary = tracker_row.get("Summary", vis_row.get("Subject", "-") if vis_row is not None else "-")

    if mode == "evaluation":
        return {
            "dcr_id": dcr_id,
            "summary": str(summary) if pd.notna(summary) else "-",
            "status": eval_status_from_row(tracker_row),
            "closure_date": closure_date_from_row(tracker_row, vis_row),
            "support": support_required_from_row(tracker_row),
            "remarks": latest_comment(tracker_row.get("Comments (Daily)"), max_len=None),
        }
    return {
        "dcr_id": dcr_id,
        "summary": str(summary) if pd.notna(summary) else "-",
        "status": impl_status_from_row(tracker_row),
        "closure_date": closure_date_from_row(tracker_row, vis_row),
        "support": support_required_from_row(tracker_row),
        "remarks": latest_comment(tracker_row.get("Comments (Daily)"), max_len=None),
    }


def _filtered_pending_items(
    tracker_rows: dict[int, list[pd.Series]],
    visibility: pd.DataFrame,
    *,
    mode: str,
    prcr_state: str,
    cutoff_date: pd.Timestamp,
) -> list[dict]:
    cutoff = pd.Timestamp(cutoff_date).normalize()
    selected: list[tuple[pd.Timestamp, int, pd.Series]] = []
    seen: set[int] = set()

    for dcr_id, rows in tracker_rows.items():
        for row in rows:
            if str(row.get("PRCRState", "")).strip() != prcr_state:
                continue
            if str(row.get("At Risk", "")).strip() != _PENDING_AT_RISK:
                continue
            planned = coerce_tracker_date(row.get(PLAN_COMPLETION_COL))
            if planned is None or planned.normalize() > cutoff:
                continue
            if dcr_id in seen:
                continue
            seen.add(dcr_id)
            selected.append((planned.normalize(), dcr_id, row))

    selected.sort(key=lambda item: (item[0], item[1]))

    items: list[dict] = []
    for _, dcr_id, row in selected:
        item = build_pending_item(dcr_id, tracker_rows, visibility, mode, tracker_row=row)
        if item is not None:
            items.append(item)
    return items


def pending_items(
    visibility: pd.DataFrame,
    tracker_rows: dict[int, list[pd.Series]],
    mode: str,
    pending_week: int | None = None,
    limit: int = PENDING_TABLE_ROW_CAP,
    cutoff_date: pd.Timestamp | str | None = None,
) -> list[dict]:
    del pending_week, limit
    mode = mode.lower()
    if mode not in _PENDING_PRCR_STATE:
        raise ValueError(f"Unknown pending mode: {mode}")
    if cutoff_date is None:
        raise ValueError(f"cutoff_date is required for {mode} pending selection")
    if not isinstance(cutoff_date, pd.Timestamp):
        cutoff_date = pd.to_datetime(cutoff_date, dayfirst=True)

    return _filtered_pending_items(
        tracker_rows,
        visibility,
        mode=mode,
        prcr_state=_PENDING_PRCR_STATE[mode],
        cutoff_date=cutoff_date,
    )


def pending_week_for_chart(chart_week: int) -> int:
    return chart_week


def graph_week_capacity(data_file: str, pending_week: int, mode: str) -> int:
    section = (
        get_evaluation_data(data_file=data_file)
        if mode == "evaluation"
        else get_implementation_data(data_file=data_file)
    )
    match = section[section["Week No"] == pending_week]
    if match.empty:
        return PENDING_TABLE_ROW_CAP
    value = match.iloc[0].get("Eval In Progress")
    if pd.isna(value):
        return PENDING_TABLE_ROW_CAP
    return max(1, min(int(value), PENDING_TABLE_ROW_CAP))
