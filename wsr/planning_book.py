"""
Quarterly planning metrics from Book2.xlsx (slide 11).

Looks for a row labelled roughly:
  "Total work Hrs. Available for PFS team"

Then computes:
  available_hours  = that cell
  planned_hours     = planned_pct% of available  (from GUI / CLI, default 90)
  resources         = team size from column I when present
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from wsr.constants import DEFAULT_PLANNING_BOOK

# Default when the user does not override (GUI / CLI).
DEFAULT_PLANNED_BANDWIDTH_PCT = 90
# Backward-compatible alias
PLANNED_BANDWIDTH_PCT = DEFAULT_PLANNED_BANDWIDTH_PCT

_AVAILABLE_LABEL = "total work hrs available for pfs team"


def _as_float(value) -> float | None:
    try:
        if value in (None, ""):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _normalize_label(value) -> str:
    return " ".join(str(value).strip().lower().replace(".", " ").split())


def _lookup_available_hours(ws) -> tuple[float | None, int | None]:
    candidates: list[dict] = []
    for row_idx in range(1, (ws.max_row or 0) + 1):
        label = ws.cell(row_idx, 2).value
        if not label or _AVAILABLE_LABEL not in _normalize_label(label):
            continue
        hours = _as_float(ws.cell(row_idx, 4).value)
        if hours is None:
            continue
        members = _as_float(ws.cell(row_idx, 9).value)
        candidates.append(
            {
                "hours": hours,
                "members": int(round(members)) if members is not None else None,
            }
        )

    if not candidates:
        return None, None

    with_members = [c for c in candidates if c["members"] is not None]
    if len(with_members) >= 2:
        max_members = max(c["members"] for c in with_members)
        subsets = [c for c in with_members if c["members"] < max_members]
        if subsets:
            best = max(subsets, key=lambda c: c["hours"])
            return best["hours"], best["members"]

    best = with_members[-1] if with_members else candidates[-1]
    return best["hours"], best["members"]


def load_quarterly_planning(
    planning_book: str | Path | None = None,
    *,
    planned_pct: int = DEFAULT_PLANNED_BANDWIDTH_PCT,
) -> dict[str, int] | None:
    workbook_path = Path(planning_book) if planning_book else DEFAULT_PLANNING_BOOK
    if not workbook_path.exists():
        return None

    wb = load_workbook(workbook_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    available_hours, resources = _lookup_available_hours(ws)
    if available_hours is None:
        return None

    planned_hours = available_hours * (planned_pct / 100.0)

    return {
        "available_hours": int(round(available_hours)),
        "planned_hours": int(round(planned_hours)),
        "planned_pct": int(planned_pct),
        "resources": resources if resources is not None else 0,
    }
