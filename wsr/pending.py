"""Pending DCR selection for slides 5 and 6."""

from __future__ import annotations

import pandas as pd

from wsr.constants import DEFAULT_DATA_FILE, PENDING_TABLE_ROW_CAP
from wsr.graph import get_evaluation_data, get_implementation_data
from wsr.tracker import (
    closure_date_from_row,
    closure_sort_key,
    comment_activity_score,
    eval_status_from_row,
    impl_status_from_row,
    latest_comment,
    parse_dcr_id,
    support_required_from_row,
    tracker_row_for_mode,
    visibility_row,
)


def build_pending_item(
    dcr_id: int,
    tracker_rows: dict[int, list[pd.Series]],
    visibility: pd.DataFrame,
    mode: str,
) -> dict | None:
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
        "remarks": latest_comment(tracker_row.get("Comments (Daily)"), max_len=None),
    }


def pending_items(
    visibility: pd.DataFrame,
    tracker_rows: dict[int, list[pd.Series]],
    mode: str,
    pending_week: int | None = None,
    limit: int = PENDING_TABLE_ROW_CAP,
) -> list[dict]:
    del pending_week  # week label is applied on the slide; rows come from current workbook state.
    mode = mode.lower()

    if mode == "evaluation":
        type_mask = visibility["Evaluation/ Implementation"].astype(str).str.contains("Eval", case=False, na=False)
    else:
        type_mask = visibility["Evaluation/ Implementation"].astype(str).str.contains("Impl", case=False, na=False)

    excluded_status = visibility["Status"].astype(str).str.contains(
        r"Rejected|Cancelled|Deferred|Closed", case=False, na=False
    )

    candidates: list[tuple[int, int]] = []
    for _, vis_row in visibility[type_mask & ~excluded_status].iterrows():
        dcr_id = parse_dcr_id(vis_row.get("DCR Number"))
        if dcr_id is None:
            continue

        tracker_row = tracker_row_for_mode(tracker_rows, dcr_id, mode)
        if tracker_row is None:
            continue
        if pd.notna(tracker_row.get("Actual Cmpln Date\n<dd-mm-yyyy>")):
            continue

        status = str(vis_row.get("Status", ""))
        score = comment_activity_score(tracker_row.get("Comments (Daily)"))
        if status in ("On Track", "At Risk", "Yet to start", "On Hold"):
            score += 4
        if mode == "implementation":
            summary = str(tracker_row.get("Summary", "")).lower()
            comments = str(tracker_row.get("Comments (Daily)", "")).lower()
            if "l2" in comments:
                score += 5
            if "appguide" in summary or "app-guide" in summary:
                score += 4
            if "dummy" in summary:
                score -= 10
        else:
            if str(tracker_row.get("PRCR Stage", "")) in ("Verify", "Submit", "Assess", "In Progress"):
                score += 2

        candidates.append((score, dcr_id))

    candidates.sort(key=lambda item: (-item[0], item[1]))
    selected_ids: list[int] = []
    seen: set[int] = set()
    for _, dcr_id in candidates:
        if dcr_id in seen:
            continue
        seen.add(dcr_id)
        selected_ids.append(dcr_id)
        if len(selected_ids) >= limit:
            break

    items = []
    for dcr_id in selected_ids:
        item = build_pending_item(dcr_id, tracker_rows, visibility, mode)
        if item is not None:
            items.append(item)

    if mode == "evaluation":
        items.sort(key=lambda item: closure_sort_key(item["closure_date"]))

    return items


def pending_week_for_chart(chart_week: int) -> int:
    """PDF uses prior week for pending-closure tables while charts show current week."""
    return max(chart_week - 1, 1)


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
