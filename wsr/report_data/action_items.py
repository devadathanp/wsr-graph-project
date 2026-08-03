"""Slide 3 Action Items from the ActionItem sheet."""

from __future__ import annotations

import pandas as pd

from wsr.tracker import format_date

# Column B — filter only; never shown on the slide.
COL_AUDIENCE = "Internal(KPIT)/External(Cummins)\n(Note : Internal will not be part of WSR ppt)\n"
COL_ACTION = "Action Items"
COL_PRIORITY = "Priority"
COL_STATUS = "Status"
COL_OWNER = "Owner"
COL_IDENTIFIED = "identified Date"
COL_TARGET = "Target closure date"
COL_REMARKS = "Remarks"

EXCLUDED_STATUSES = frozenset({"closed", "cancelled", "canceled", "rejected"})


def _normalize_header(name: str) -> str:
    return " ".join(str(name).replace("\n", " ").split()).strip().lower()


def _find_column(columns, *needles: str) -> str | None:
    normalized = {_normalize_header(col): col for col in columns}
    for needle in needles:
        key = _normalize_header(needle)
        if key in normalized:
            return normalized[key]
        for header, original in normalized.items():
            if header.startswith(key) or key in header:
                return original
    return None


def _cell(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    text = str(value).replace("_x000B_", "\n").strip()
    if text.lower() in ("nan", "none", "nat"):
        return ""
    return text


def _is_external(value) -> bool:
    return _cell(value).lower().startswith("external")


def _status_active(value) -> bool:
    return _cell(value).lower() not in EXCLUDED_STATUSES


def action_items(actions: pd.DataFrame) -> list[dict[str, str]]:
    """
    External action items for slide 3.

    - Column B must be External (Internal KPIT items are omitted).
    - Status Closed / Cancelled / Rejected are omitted.
    - Column B itself is never returned for the slide table.
    """
    if actions is None or actions.empty:
        return []

    audience_col = _find_column(actions.columns, COL_AUDIENCE, "Internal(KPIT)/External(Cummins)")
    action_col = _find_column(actions.columns, COL_ACTION)
    status_col = _find_column(actions.columns, COL_STATUS)
    if not audience_col or not action_col or not status_col:
        return []

    priority_col = _find_column(actions.columns, COL_PRIORITY)
    owner_col = _find_column(actions.columns, COL_OWNER)
    identified_col = _find_column(actions.columns, COL_IDENTIFIED, "identified Date")
    target_col = _find_column(actions.columns, COL_TARGET, "Target closure date")
    remarks_col = _find_column(actions.columns, COL_REMARKS)

    items: list[dict[str, str]] = []
    for _, row in actions.iterrows():
        if not _is_external(row.get(audience_col)):
            continue
        if not _status_active(row.get(status_col)):
            continue
        action_text = _cell(row.get(action_col))
        if not action_text:
            continue

        identified_raw = row.get(identified_col) if identified_col else None
        target_raw = row.get(target_col) if target_col else None
        identified = format_date(identified_raw) if _cell(identified_raw) else "-"
        target = format_date(target_raw) if _cell(target_raw) else "-"
        items.append(
            {
                "action": action_text,
                "priority": _cell(row.get(priority_col)) if priority_col else "",
                "status": _cell(row.get(status_col)),
                "owner": _cell(row.get(owner_col)) if owner_col else "",
                "identified_date": identified,
                "target_date": target,
                "remarks": _cell(row.get(remarks_col)) if remarks_col else "",
            }
        )
    return items
