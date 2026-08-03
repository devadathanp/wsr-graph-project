"""Slide 7 DDP MS4-5 rows from Non STLA tracker columns."""

from __future__ import annotations

import pandas as pd

from wsr.tracker import format_date, parse_dcr_id

COL_DCR = "DCR ID - PTC"
COL_SUMMARY = "Summary"
COL_DDP_NEEDED = "Is DDP Testing Needed"
COL_PLAN_DATE = "DDP Plan Date"
COL_APPEARED_DATE = "DDP Appeared Date"
COL_PROGRAM = "DDP- Program"
COL_DEPENDENCIES = "DDP Dependencies"
COL_REMARK = "DDP Remark"
COL_STATUS = "DDP Status"


def _cell(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    text = str(value).strip()
    if text.lower() in ("nan", "none", "nat"):
        return ""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _format_ddp_date(value) -> str:
    """Format plan/appeared dates; keep multi-line Excel text as slash-separated."""
    text = _cell(value)
    if not text:
        return "-"
    if "\n" in text:
        parts = [part.strip() for part in text.split("\n") if part.strip()]
        return " / ".join(parts) if parts else "-"
    formatted = format_date(value)
    return formatted if formatted else "-"


def _ddp_needed(value) -> bool:
    return _cell(value).lower() in {"yes", "y"}


def _status_active(value) -> bool:
    status = _cell(value).lower()
    # Empty status still shown if DDP testing is needed; only Closed is dropped.
    return status != "closed"


def ddp_ms45_items(tracker: pd.DataFrame) -> list[dict[str, str]]:
    """
    Active DDP MS4-5 rows for slide 7.

    Include rows where ``Is DDP Testing Needed`` is Yes and ``DDP Status``
    is not Closed.
    """
    if tracker is None or tracker.empty:
        return []
    if COL_DDP_NEEDED not in tracker.columns:
        return []

    items: list[dict[str, str]] = []
    seen: set[int] = set()
    for _, row in tracker.iterrows():
        if not _ddp_needed(row.get(COL_DDP_NEEDED)):
            continue
        if COL_STATUS in tracker.columns and not _status_active(row.get(COL_STATUS)):
            continue

        dcr_id = parse_dcr_id(row.get(COL_DCR))
        if dcr_id is not None:
            if dcr_id in seen:
                continue
            seen.add(dcr_id)
            dcr_text = str(dcr_id)
        else:
            dcr_text = _cell(row.get(COL_DCR)) or "-"

        items.append(
            {
                "dcr_id": dcr_text,
                "summary": _cell(row.get(COL_SUMMARY)) or "-",
                "plan_date": _format_ddp_date(row.get(COL_PLAN_DATE)),
                "appeared_date": _format_ddp_date(row.get(COL_APPEARED_DATE)),
                "program": _cell(row.get(COL_PROGRAM)) or "-",
                "dependencies": _cell(row.get(COL_DEPENDENCIES)) or "-",
                "remarks": _cell(row.get(COL_REMARK)) or "-",
                "status": _cell(row.get(COL_STATUS)) or "-",
            }
        )
    return items
