"""Slide 8 eval handoff table (available for manual/automated use)."""

from __future__ import annotations

import pandas as pd

from wsr.report_data.planning import planning_dcr_column
from wsr.tracker import format_date, latest_comment, parse_dcr_id


def _report_month_year(report_date: str) -> tuple[int, int]:
    parsed = pd.to_datetime(report_date, dayfirst=True)
    return parsed.month, parsed.year


def _first_evaluator_name(value) -> str:
    if pd.isna(value) or str(value).strip() in ("", "nan"):
        return "-"
    text = str(value).strip()
    return text.split("/")[0].split(",")[0].strip()


def _planning_handoff_date(row: pd.Series) -> tuple[pd.Timestamp | None, str]:
    for column in (
        "L2/DRB Eval Send dates\nDD-MM-YY",
        "Eval completion date, (if it is evaluation)\nDD-MM-YY",
    ):
        if column not in row.index:
            continue
        parsed = pd.to_datetime(row.get(column), errors="coerce")
        if pd.notna(parsed):
            return parsed, format_date(parsed)
    return None, "-"


def eval_handoff_items(
    planning: pd.DataFrame,
    tracker_lookup_map: dict[int, pd.Series],
    report_date: str,
    limit: int = 10,
) -> list[dict]:
    month, year = _report_month_year(report_date)
    dcr_col = planning_dcr_column(planning)
    if dcr_col is None:
        return []

    items = []
    for _, row in planning.iterrows():
        dcr_id = parse_dcr_id(row.get(dcr_col))
        if dcr_id is None:
            continue

        handoff_ts, handoff_date = _planning_handoff_date(row)
        if handoff_ts is None or handoff_ts.month != month or handoff_ts.year != year:
            continue

        tracker_row = tracker_lookup_map.get(dcr_id)
        summary = str(row.get("Summary", "-")) if pd.notna(row.get("Summary")) else "-"
        if tracker_row is not None and pd.notna(tracker_row.get("Summary")):
            summary = str(tracker_row.get("Summary"))

        evaluator = _first_evaluator_name(row.get("KPIT User"))
        if evaluator == "-":
            evaluator = _first_evaluator_name(row.get("Estimator"))

        remark = "-"
        if pd.notna(row.get("Remarks.1")) and str(row.get("Remarks.1")).strip() not in ("", "nan"):
            remark = str(row.get("Remarks.1")).strip()
        elif tracker_row is not None:
            remark = latest_comment(tracker_row.get("Comments (Daily)"), max_len=None)

        items.append(
            {
                "dcr_id": dcr_id,
                "evaluator": evaluator,
                "handoff_date": handoff_date,
                "remark": remark,
                "summary": summary,
            }
        )
        if len(items) >= limit:
            break
    return items
