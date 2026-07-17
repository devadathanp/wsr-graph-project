"""Slide 9 discussion points (available for manual/automated use)."""

from __future__ import annotations

import pandas as pd

from wsr.tracker import format_date, latest_comment


def _is_high_priority(priority) -> bool:
    text = str(priority).strip().lower()
    return text == "high"


def discussion_points(
    visibility: pd.DataFrame,
    tracker_lookup_map: dict[int, pd.Series],
    limit: int = 5,
) -> list[dict]:
    items = []
    at_risk = visibility[visibility["Status"].astype(str).str.contains("At Risk|On Hold", case=False, na=False)]
    for _, vis_row in at_risk.iterrows():
        dcr_raw = vis_row.get("DCR Number")
        if pd.isna(dcr_raw):
            continue
        try:
            dcr_id = int(dcr_raw)
        except (TypeError, ValueError):
            continue
        if any(item["dcr_id"] == dcr_id for item in items):
            continue

        tracker_row = tracker_lookup_map.get(dcr_id, vis_row)
        priority = str(tracker_row.get("CQ Priority", tracker_row.get("Internal Priority", "-")))
        if not _is_high_priority(priority):
            continue
        items.append(
            {
                "dcr_id": dcr_id,
                "description": str(tracker_row.get("Summary", vis_row.get("Subject", "-"))),
                "program": str(tracker_row.get("DCR Project", "CSAR"))
                if pd.notna(tracker_row.get("DCR Project"))
                else "CSAR",
                "priority": priority,
                "plan_date": format_date(tracker_row.get("Planned Completion Date\n<dd-mm-yyyy>")),
                "remarks": latest_comment(tracker_row.get("Comments (Daily)"), max_len=None),
            }
        )
        if len(items) >= limit:
            break
    return items
