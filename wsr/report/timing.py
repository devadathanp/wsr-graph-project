"""Resolve chart week and report date from workbook data."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from wsr.errors import WsrDataError
from wsr.fiscal import fiscal_year, iso_week_number, last_friday, quarter_label_long, quarter_label_short
from wsr.graph import latest_reported_week
from wsr.report.models import ReportTiming
from wsr.run_log import RunLog


def resolve_report_timing(
    scrum_path: Path,
    *,
    chart_week: int | None,
    report_date: str | None,
    log: RunLog,
) -> ReportTiming:
    if report_date is None:
        report_date = datetime.now().strftime("%d-%m-%Y")

    detected_week, _ = latest_reported_week(str(scrum_path), report_date=report_date)
    log.info(f"Detected week={detected_week!r}")

    if chart_week is None:
        if detected_week is None:
            raise WsrDataError(
                "Could not detect the reporting week from the graph sheet "
                f'("{scrum_path.name}"). Pass an explicit week number.'
            )
        chart_week = detected_week

    # Slide dates, quarter, and heading week use the day the report is generated.
    # Bar charts and the confidence line use the full quarter table; the actual
    # completion line stops at the last Friday on or before report_date.
    heading_week = iso_week_number(report_date)
    quarter_long = quarter_label_long(report_date)
    quarter_short = quarter_label_short(report_date)
    year = fiscal_year(report_date)
    log.info(
        f"Using chart week={chart_week}, report date={report_date}, "
        f"graph through {last_friday(report_date):%d-%m-%Y}, "
        f"heading={quarter_long} week {heading_week}"
    )
    return ReportTiming(
        chart_week=chart_week,
        report_date=report_date,
        pending_week=heading_week,
        quarter_long=quarter_long,
        quarter_short=quarter_short,
        fiscal_year=year,
    )
