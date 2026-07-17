"""Resolve chart week and report date from workbook data."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from wsr.errors import WsrDataError
from wsr.graph import latest_reported_week
from wsr.pending import pending_week_for_chart
from wsr.report.models import ReportTiming
from wsr.run_log import RunLog


def resolve_report_timing(
    scrum_path: Path,
    *,
    chart_week: int | None,
    report_date: str | None,
    log: RunLog,
) -> ReportTiming:
    detected_week, detected_date = latest_reported_week(str(scrum_path))
    log.info(f"Detected week={detected_week!r}, date={detected_date!r}")

    if chart_week is None:
        if detected_week is None:
            raise WsrDataError(
                "Could not detect the reporting week from the graph sheet "
                f'("{scrum_path.name}"). Pass an explicit week number.'
            )
        chart_week = detected_week

    if report_date is None:
        report_date = detected_date or datetime.now().strftime("%d-%m-%Y")

    pending_week = pending_week_for_chart(chart_week)
    log.info(f"Using chart week={chart_week}, report date={report_date}")
    return ReportTiming(
        chart_week=chart_week,
        report_date=report_date,
        pending_week=pending_week,
    )
