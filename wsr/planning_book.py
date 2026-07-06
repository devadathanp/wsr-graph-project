"""Quarterly planning metrics from Book2.xlsx (slide 11)."""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from wsr.constants import DEFAULT_PLANNING_BOOK


def _as_float(value) -> float | None:
    try:
        if value in (None, ""):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def load_quarterly_planning(planning_book: str | Path | None = None) -> dict[str, int] | None:
    workbook_path = Path(planning_book) if planning_book else DEFAULT_PLANNING_BOOK
    if not workbook_path.exists():
        return None

    wb = load_workbook(workbook_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    available_hours = _as_float(ws["H34"].value)
    planned_hours = _as_float(ws["H36"].value)
    planned_pct_raw = _as_float(ws["H37"].value)
    resources_raw = _as_float(ws["I44"].value)

    if None in (available_hours, planned_hours, planned_pct_raw, resources_raw):
        return None

    return {
        "available_hours": int(round(available_hours)),
        "planned_hours": int(round(planned_hours)),
        "planned_pct": int(round(planned_pct_raw * 100 if planned_pct_raw <= 1 else planned_pct_raw)),
        "resources": int(round(resources_raw)),
    }
