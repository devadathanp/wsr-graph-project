"""Planning sheet helpers."""

from __future__ import annotations

import pandas as pd


def planning_dcr_column(planning: pd.DataFrame) -> str | None:
    for column in ("DCRID-PTC", planning.columns[1] if len(planning.columns) > 1 else None):
        if column and column in planning.columns:
            return column
    return None


def planning_dcr_ids(planning: pd.DataFrame) -> list[int]:
    from wsr.tracker import parse_dcr_id

    dcr_col = planning_dcr_column(planning)
    if dcr_col is None:
        return []
    ids = []
    for value in planning[dcr_col]:
        dcr_id = parse_dcr_id(value)
        if dcr_id is not None:
            ids.append(dcr_id)
    return ids


def planning_type_counts(planning: pd.DataFrame) -> dict[str, int]:
    type_col = "Type" if "Type" in planning.columns else "Eval+Impl"
    if type_col not in planning.columns:
        return {}
    counts = planning[type_col].astype(str).str.strip().value_counts()
    return {str(key): int(value) for key, value in counts.items() if str(key) not in ("nan", "")}


def core2_planned_count(planning: pd.DataFrame) -> int | None:
    if "Program" not in planning.columns:
        return None
    program = planning["Program"].astype(str).str.strip()
    count = program.str.contains(r"\bDAF\b|Core\s*2", case=False, na=False).sum()
    return int(count)
