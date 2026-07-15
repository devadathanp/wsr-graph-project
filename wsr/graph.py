"""Graph sheet loading and week-series helpers."""

from __future__ import annotations

import numpy as np
import pandas as pd

from wsr.constants import DEFAULT_DATA_FILE, GRAPH_SHEET

COL_WEEK = "Week No"
COL_DATE = "Date"
COL_CUMULATIVE_BASELINE = "Cumulative (Baseline Plan)"
COL_CUMULATIVE_REVISED = "Cumulative Revised basline plan"
COL_COMPLETED = "Cumulative (Completed)"
COL_REJECTED = "Cumulative Rejected / Transferred/ Moved to next quarter"
COL_IN_PROGRESS = "Eval In Progress"
COL_DRB = "DRB /l2 Reviews & Rework  in progress"
COL_PCT_CONFIDENCE = "% Completion Confidence - Overall"
COL_PCT_ACTUAL = "% Actual weekly completion wr.t  revised Baseline"
COL_REMARKS = "Remarks"


def to_percentage(series: pd.Series) -> pd.Series:
    def convert(value):
        if pd.isna(value):
            return np.nan
        if isinstance(value, (int, float)):
            val = float(value)
            return val * 100 if abs(val) <= 1.5 else val
        text = str(value).replace("%", "").strip()
        if text == "":
            return np.nan
        val = float(text)
        return val * 100 if abs(val) <= 1.5 else val

    return series.apply(convert)


def _safe_int(value) -> int | None:
    if pd.isna(value):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def load_graph_sheet(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name=GRAPH_SHEET, header=2)


def load_graph_summary(data_file: str = DEFAULT_DATA_FILE) -> dict:
    raw = pd.read_excel(data_file, sheet_name=GRAPH_SHEET, header=None)
    return {
        "eval_baseline": _safe_int(raw.iloc[16, 4]),
        "eval_revised": _safe_int(raw.iloc[16, 5]),
        "eval_completed": _safe_int(raw.iloc[16, 6]),
        "impl_baseline": _safe_int(raw.iloc[34, 4]),
        "impl_revised": _safe_int(raw.iloc[34, 5]),
        "impl_completed": _safe_int(raw.iloc[34, 6]),
    }


def get_evaluation_data(df: pd.DataFrame | None = None, data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    if df is None:
        df = load_graph_sheet(data_file)

    impl_header_idx = df.index[df["Tagged to Release"].astype(str).str.strip() == "Implementation"]
    eval_df = df.loc[: impl_header_idx[0] - 1]
    section = eval_df[eval_df[COL_WEEK].notna()].copy()
    section.reset_index(drop=True, inplace=True)
    return _coerce_graph_numeric(section)


def get_implementation_data(df: pd.DataFrame | None = None, data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    if df is None:
        df = load_graph_sheet(data_file)

    impl_start_idx = df.index[
        df["Tagged to Release"].astype(str).str.contains("Q3", na=False)
        & df["Tagged to Release"].astype(str).str.contains("Implementation", na=False)
    ]
    section = df.loc[impl_start_idx[0] :]
    section = section[section[COL_WEEK].notna()].copy()
    section.reset_index(drop=True, inplace=True)
    return _coerce_graph_numeric(section)


def _coerce_graph_numeric(section: pd.DataFrame) -> pd.DataFrame:
    section = section.copy()
    numeric_cols = [
        COL_CUMULATIVE_BASELINE,
        COL_CUMULATIVE_REVISED,
        COL_COMPLETED,
        COL_REJECTED,
        COL_IN_PROGRESS,
        COL_DRB,
        COL_PCT_CONFIDENCE,
        COL_PCT_ACTUAL,
        "% Actual weekly completion wr.t Baseline",
    ]
    for column in numeric_cols:
        if column in section.columns:
            section[column] = pd.to_numeric(section[column], errors="coerce")
    return section


def latest_reported_week(
    data_file: str = DEFAULT_DATA_FILE,
) -> tuple[int | None, str | None]:
    df = load_graph_sheet(data_file)
    week: int | None = None
    date_label: str | None = None
    for section in (get_evaluation_data(df), get_implementation_data(df)):
        reported = section[section[COL_PCT_ACTUAL].notna()]
        if reported.empty:
            continue
        row = reported.iloc[-1]
        candidate = int(row[COL_WEEK])
        if week is None or candidate > week:
            week = candidate
            date_label = pd.to_datetime(row[COL_DATE]).strftime("%d-%m-%Y")
    return week, date_label


def add_week_labels(section: pd.DataFrame) -> pd.DataFrame:
    section = section.copy()
    section["Week Label"] = (
        section[COL_WEEK].astype(str) + "\n" + pd.to_datetime(section[COL_DATE]).dt.strftime("%d-%m")
    )
    return section


def week_remarks(section: pd.DataFrame, week_no: int) -> str:
    match = section[section[COL_WEEK] == week_no]
    if match.empty:
        return ""
    remark = match.iloc[0].get("Remarks")
    drb = match.iloc[0].get("DRB /l2 Reviews & Rework  in progress")
    parts = []
    if pd.notna(remark) and str(remark).strip():
        parts.append(str(remark).strip())
    if pd.notna(drb) and str(drb).strip() not in ("", "nan"):
        parts.append(f"DRB/L2 in progress: {int(float(drb))}")
    return " | ".join(parts)


def graph_status_notes(
    eval_data: pd.DataFrame,
    impl_data: pd.DataFrame,
    weeks: list[int],
) -> list[str]:
    lines = []
    for week in weeks:
        eval_note = week_remarks(eval_data, week)
        impl_note = week_remarks(impl_data, week)
        if eval_note or impl_note:
            lines.append(f"WK:{week} — Eval: {eval_note or '-'} | Impl: {impl_note or '-'}")
    return lines
