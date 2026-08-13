"""PFS ECM testing rows from Non STLA (Is ECM Testing Needed = Yes)."""

from __future__ import annotations

import math

import pandas as pd

from wsr.rich_text import TextRun, full_cell_text, resolve_display
from wsr.tracker import parse_dcr_id

COL_DCR = "DCR ID - PTC"
COL_SUMMARY = "Summary"
COL_ECM_NEEDED = "Is ECM Testing Needed"
COL_STATUS = "ECM Status"
COL_BT_START = "BT Start"
COL_BT_END = "BT end"
COL_DEPENDENCIES = "ECM Dependencies"
COL_REMARK = "ECM Remark"

# Column DC — keep only these; drop cancelled / closed / complete / rejected / blank.
_ACTIVE_ECM_STATUSES = frozenset(
    {
        "at risk",
        "in progress",
        "on hold",
        "yet to start",
    }
)


def _cell(value) -> str:
    try:
        if isinstance(value, float) and math.isnan(value):
            return ""
        if isinstance(value, float) and value == int(value):
            return str(int(value))
    except (TypeError, ValueError, OverflowError):
        pass
    return full_cell_text(value)


def _excel_row(idx) -> int:
    if isinstance(idx, (int, float)) and idx == idx:
        return int(idx) + 2
    return -1


def _ecm_needed(value) -> bool:
    return _cell(value).lower() in {"yes", "y"}


def _status_active(value) -> bool:
    return " ".join(_cell(value).lower().split()) in _ACTIVE_ECM_STATUSES


def ecm_testing_items(
    tracker: pd.DataFrame,
    rich_runs: dict[tuple[int, str], list[TextRun]] | None = None,
) -> list[dict]:
    """
    Rows for the PFS ECM Testing Details slide.

    Include rows where ``Is ECM Testing Needed`` is Yes and ``ECM Status``
    is At Risk, In Progress, On Hold, or Yet to Start.
    Struck-through + replacement values are both kept.
    """
    if tracker is None or tracker.empty or COL_ECM_NEEDED not in tracker.columns:
        return []

    items: list[dict] = []
    seen: set[int] = set()
    for idx, row in tracker.iterrows():
        if not _ecm_needed(row.get(COL_ECM_NEEDED)):
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

        excel_row = _excel_row(idx)

        def field(column: str, *, as_date: bool = False):
            return resolve_display(
                row.get(column) if column in row.index else None,
                excel_row=excel_row,
                column=column,
                rich_runs=rich_runs,
                as_date=as_date,
            )

        items.append(
            {
                "dcr_id": dcr_text,
                "summary": field(COL_SUMMARY) or "-",
                "bt_start": field(COL_BT_START, as_date=True) or "-",
                "bt_end": field(COL_BT_END, as_date=True) or "-",
                "dependencies": field(COL_DEPENDENCIES) or "-",
                "remarks": field(COL_REMARK) or "-",
            }
        )
    return items
