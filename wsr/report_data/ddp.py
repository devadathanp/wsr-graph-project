"""Slide 7 DDP MS4-5 rows from Non STLA tracker columns."""

from __future__ import annotations

import math

import pandas as pd

from wsr.rich_text import TextRun, full_cell_text, resolve_display
from wsr.tracker import parse_dcr_id

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


def _ddp_needed(value) -> bool:
    return _cell(value).lower() in {"yes", "y"}


def _status_active(value) -> bool:
    return _cell(value).lower() != "closed"


def ddp_ms45_items(
    tracker: pd.DataFrame,
    rich_runs: dict[tuple[int, str], list[TextRun]] | None = None,
) -> list[dict]:
    """
    Active DDP MS4-5 rows for slide 7.

    Include rows where ``Is DDP Testing Needed`` is Yes and ``DDP Status``
    is not Closed. Struck-through + replacement values are both kept.
    """
    if tracker is None or tracker.empty:
        return []
    if COL_DDP_NEEDED not in tracker.columns:
        return []

    items: list[dict] = []
    seen: set[int] = set()
    for idx, row in tracker.iterrows():
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
                "plan_date": field(COL_PLAN_DATE, as_date=True) or "-",
                "appeared_date": field(COL_APPEARED_DATE, as_date=True) or "-",
                "program": field(COL_PROGRAM) or "-",
                "dependencies": field(COL_DEPENDENCIES) or "-",
                "remarks": field(COL_REMARK) or "-",
                "status": field(COL_STATUS) or "-",
            }
        )
    return items
