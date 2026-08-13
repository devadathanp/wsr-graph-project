"""Report generation result and intermediate data containers."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class ReportResult:
    output_path: Path
    log_path: Path
    warnings: list[str] = field(default_factory=list)


@dataclass
class ReportTiming:
    chart_week: int
    report_date: str
    pending_week: int
    quarter_long: str
    quarter_short: str
    fiscal_year: int


@dataclass
class ScrumWorkbook:
    path: Path
    tracker: pd.DataFrame
    visibility: pd.DataFrame
    ddp: pd.DataFrame
    risks: pd.DataFrame
    action_items: pd.DataFrame
    tracker_map: dict[int, pd.Series]
    tracker_rows: dict[int, list[pd.Series]]
    # sheet -> (excel_row, column_header) -> rich-text runs with strike flags
    rich_runs: dict[str, dict[tuple[int, str], list[tuple[str, bool]]]] = field(
        default_factory=dict
    )


@dataclass
class ChartAssets:
    impl_chart: Path
    eval_chart: Path
    planning_chart: Path | None
    quarterly_planning: dict[str, int | str] | None
