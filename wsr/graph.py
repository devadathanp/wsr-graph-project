"""Graph sheet loading and week-series helpers."""

from __future__ import annotations

import numpy as np
import pandas as pd

from wsr.constants import DEFAULT_DATA_FILE, GRAPH_SHEET


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
    section = eval_df[eval_df["Week No"].notna()].copy()
    section.reset_index(drop=True, inplace=True)
    return section


def get_implementation_data(df: pd.DataFrame | None = None, data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    if df is None:
        df = load_graph_sheet(data_file)

    impl_start_idx = df.index[
        df["Tagged to Release"].astype(str).str.contains("Q3", na=False)
        & df["Tagged to Release"].astype(str).str.contains("Implementation", na=False)
    ]
    section = df.loc[impl_start_idx[0] :]
    section = section[section["Week No"].notna()].copy()
    section.reset_index(drop=True, inplace=True)
    return section


def add_week_labels(section: pd.DataFrame) -> pd.DataFrame:
    section = section.copy()
    section["Week Label"] = (
        section["Week No"].astype(str) + "\n" + pd.to_datetime(section["Date"]).dt.strftime("%d-%m")
    )
    return section


def week_remarks(section: pd.DataFrame, week_no: int) -> str:
    match = section[section["Week No"] == week_no]
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
