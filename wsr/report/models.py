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


@dataclass
class ScrumWorkbook:
    path: Path
    tracker: pd.DataFrame
    visibility: pd.DataFrame
    ddp: pd.DataFrame
    tracker_map: dict[int, pd.Series]
    tracker_rows: dict[int, list[pd.Series]]


@dataclass
class ChartAssets:
    impl_chart: Path
    eval_chart: Path
    planning_chart: Path | None
    quarterly_planning: dict[str, int] | None
