"""Eval Handoff from onsite — Non STLA rows where Onsite Evaluator is Yes."""

from __future__ import annotations

import pandas as pd

from wsr.rich_text import TextRun, resolve_display
from wsr.tracker import parse_dcr_id

COL_ONSITE = "Onsite Evaluator"
COL_DCR = "DCR ID - PTC"
COL_SUMMARY = "Summary"
COL_OWNER = "DCR Owner"
COL_HANDOFF_DATE = "Eval Handoff Date"
COL_REMARK = "Eval Handoff Remark"


def _cell(value) -> str:
    text = resolve_display(value)
    return text if isinstance(text, str) else ""


def _is_yes(value) -> bool:
    return _cell(value).lower() in {"yes", "y"}


def _excel_row(idx) -> int:
    if isinstance(idx, (int, float)) and idx == idx:
        return int(idx) + 2
    return -1


def eval_handoff_items(
    tracker: pd.DataFrame,
    rich_runs: dict[tuple[int, str], list[TextRun]] | None = None,
) -> list[dict]:
    """
    Rows for the Eval Handoff slide.

    Include every Non STLA row where ``Onsite Evaluator`` is Yes.
    Struck-through + replacement values (e.g. handoff dates) are both kept.
    """
    if tracker is None or tracker.empty or COL_ONSITE not in tracker.columns:
        return []

    items: list[dict] = []
    seen: set[int] = set()
    for idx, row in tracker.iterrows():
        if not _is_yes(row.get(COL_ONSITE)):
            continue

        dcr_id = parse_dcr_id(row.get(COL_DCR)) if COL_DCR in tracker.columns else None
        if dcr_id is not None:
            if dcr_id in seen:
                continue
            seen.add(dcr_id)
            dcr_text = str(dcr_id)
        else:
            dcr_text = _cell(row.get(COL_DCR)) or "-"

        excel_row = _excel_row(idx)
        summary = resolve_display(
            row.get(COL_SUMMARY) if COL_SUMMARY in tracker.columns else None,
            excel_row=excel_row,
            column=COL_SUMMARY,
            rich_runs=rich_runs,
        )
        evaluator = resolve_display(
            row.get(COL_OWNER) if COL_OWNER in tracker.columns else None,
            excel_row=excel_row,
            column=COL_OWNER,
            rich_runs=rich_runs,
        )
        handoff_date = resolve_display(
            row.get(COL_HANDOFF_DATE) if COL_HANDOFF_DATE in tracker.columns else None,
            excel_row=excel_row,
            column=COL_HANDOFF_DATE,
            rich_runs=rich_runs,
            as_date=True,
        )
        remark = resolve_display(
            row.get(COL_REMARK) if COL_REMARK in tracker.columns else None,
            excel_row=excel_row,
            column=COL_REMARK,
            rich_runs=rich_runs,
        )

        items.append(
            {
                "dcr_id": dcr_text,
                "summary": summary or "-",
                "evaluator": evaluator or "-",
                "handoff_date": handoff_date if handoff_date not in ("", None) else "-",
                "remark": remark or "-",
            }
        )
    return items
