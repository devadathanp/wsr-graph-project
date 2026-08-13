"""
Quarterly planning metrics for the planning slide — all from Non STLA.

Bar 1 — Actual Available hours from the ``Actual Available Estimate`` column
  (number may live in the header, e.g. Actual Available Estimate "7152").

Bar 2 — Sum of Scrum effort hours:
  - Prefer ``Total Revised Estimation`` (or legacy ``Revised Estimation``)
    when that cell has a value.
  - Otherwise use ``Estimated Hrs``.

Bar 3 — Burndown:
  sum(Actual Efforts (Hrs)) − sum(Competency Gap Efforts).
"""

from __future__ import annotations

import re

import pandas as pd

# Kept for CLI/API compatibility; no longer drives the second planning bar.
DEFAULT_PLANNED_BANDWIDTH_PCT = 90
PLANNED_BANDWIDTH_PCT = DEFAULT_PLANNED_BANDWIDTH_PCT

_AVAILABLE_HEADER_TOKEN = "actual available estimate"
_AVAILABLE_NUMBER = re.compile(r"(\d+(?:\.\d+)?)")

COL_ESTIMATED_HRS = "Estimated Hrs"
# Preferred / legacy names — actual header is resolved at runtime.
COL_REVISED_ESTIMATION = "Total Revised Estimation"
_REVISED_ESTIMATION_ALIASES = (
    "total revised estimation",
    "revised estimation",
)
COL_ACTUAL_EFFORTS = "Actual Efforts (Hrs)"
COL_COMPETENCY_GAP = "Competency Gap Efforts"
COL_PRCR_STATE = "PRCRState"


def _as_float(value) -> float | None:
    try:
        if value in (None, ""):
            return None
        if isinstance(value, float) and value != value:  # NaN
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _normalize_header(value) -> str:
    return " ".join(str(value).strip().lower().replace(".", " ").split())


def revised_estimation_column(columns) -> str | None:
    """Resolve Total Revised Estimation / Revised Estimation despite header typos."""
    normalized = {_normalize_header(col): col for col in columns}
    for alias in _REVISED_ESTIMATION_ALIASES:
        if alias in normalized:
            return normalized[alias]
    for key, original in normalized.items():
        if "revised" in key and "estimation" in key:
            return original
    return None


def available_estimate_column(columns) -> str | None:
    """Find the Actual Available Estimate column (CZ on current Scrum files)."""
    for col in columns:
        if _AVAILABLE_HEADER_TOKEN in _normalize_header(col):
            return col
    return None


def scrum_available_hours(tracker: pd.DataFrame | None) -> int | None:
    """
    Quarter available hours from Non STLA.

    Prefers a numeric cell in the Actual Available Estimate column; otherwise
    parses the number from the column header (e.g. ``… "7152"``).
    """
    if tracker is None or tracker.empty:
        return None
    column = available_estimate_column(tracker.columns)
    if not column:
        return None
    for value in tracker[column]:
        hours = _as_float(value)
        if hours is not None:
            return int(round(hours))
    match = _AVAILABLE_NUMBER.search(str(column))
    if match:
        return int(round(float(match.group(1))))
    return None


def _prcr_state_blank(value) -> bool:
    if value is None:
        return True
    try:
        if isinstance(value, float) and value != value:
            return True
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip().lower() in {"", "nan", "none", "nat"}


def row_estimated_hours(row: pd.Series, revised_col: str | None = None) -> float | None:
    """
    Hours for one tracker row.

    Total Revised Estimation wins when present; otherwise Estimated Hrs.
    """
    estimated = _as_float(row.get(COL_ESTIMATED_HRS)) if COL_ESTIMATED_HRS in row.index else None
    if revised_col is None:
        revised_col = revised_estimation_column(row.index)
    if not revised_col or revised_col not in row.index:
        return estimated
    revised = _as_float(row.get(revised_col))
    if revised is not None:
        return revised
    return estimated


def sum_scrum_estimated_hours(tracker: pd.DataFrame | None) -> int | None:
    """
    Sum estimated hours (revised when present) for rows with a PRCRState.

    Blank / empty PRCRState rows are excluded. Evaluate, Implement, Task,
    Task_NA, and any other non-blank state are kept.
    """
    if tracker is None or tracker.empty:
        return None
    revised_col = revised_estimation_column(tracker.columns)
    if COL_ESTIMATED_HRS not in tracker.columns and revised_col is None:
        return None

    total = 0.0
    saw_any = False
    for _, row in tracker.iterrows():
        if COL_PRCR_STATE in tracker.columns and _prcr_state_blank(row.get(COL_PRCR_STATE)):
            continue
        hours = row_estimated_hours(row, revised_col=revised_col)
        if hours is None:
            continue
        total += hours
        saw_any = True
    if not saw_any:
        return 0
    return int(round(total))


def _sum_column(tracker: pd.DataFrame, column: str) -> float:
    if column not in tracker.columns:
        return 0.0
    total = 0.0
    for value in tracker[column]:
        hours = _as_float(value)
        if hours is not None:
            total += hours
    return total


def sum_burndown_hours(tracker: pd.DataFrame | None) -> int | None:
    """Actual Efforts (Hrs) minus Competency Gap Efforts across the tracker."""
    if tracker is None or tracker.empty:
        return None
    if COL_ACTUAL_EFFORTS not in tracker.columns and COL_COMPETENCY_GAP not in tracker.columns:
        return None
    actual = _sum_column(tracker, COL_ACTUAL_EFFORTS)
    gap = _sum_column(tracker, COL_COMPETENCY_GAP)
    return int(round(actual - gap))


def load_quarterly_planning(
    planning_book=None,
    *,
    planned_pct: int = DEFAULT_PLANNED_BANDWIDTH_PCT,
    tracker: pd.DataFrame | None = None,
) -> dict[str, int] | None:
    """Build the three planning-bar metrics from the Scrum tracker only."""
    del planning_book, planned_pct

    available = scrum_available_hours(tracker)
    if available is None:
        return None

    estimated_hours = sum_scrum_estimated_hours(tracker)
    if estimated_hours is None:
        estimated_hours = 0

    burndown_hours = sum_burndown_hours(tracker)
    if burndown_hours is None:
        burndown_hours = 0

    planned_pct = int(round((estimated_hours / available) * 100)) if available else 0

    return {
        "available_hours": int(available),
        "planned_hours": int(estimated_hours),
        "estimated_hours": int(estimated_hours),
        "burndown_hours": int(burndown_hours),
        "planned_pct": planned_pct,
        "resources": 0,
    }
