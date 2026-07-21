"""WSR report orchestration entry point."""

from __future__ import annotations

import io
from pathlib import Path

from wsr.constants import DEFAULT_DATA_FILE
from wsr.errors import WsrDataError
from wsr.planning_book import DEFAULT_PLANNED_BANDWIDTH_PCT
from wsr.pptx_sanitize import sanitize_pptx
from wsr.report.assets import build_chart_assets
from wsr.report.deck import build_presentation
from wsr.report.models import ReportResult
from wsr.report.timing import resolve_report_timing
from wsr.report.workbook import load_scrum_workbook
from wsr.run_log import RunLog, default_log_path
from wsr.validate import validate_planning_book, validate_scrum_workbook
from wsr_style import DEFAULT_TEMPLATE


def generate_report(
    output_path: str | Path = "WSR_Report.pptx",
    data_file: str = DEFAULT_DATA_FILE,
    chart_week: int | None = None,
    report_date: str | None = None,
    assets_dir: str | Path = "report_assets",
    template_path: str | Path = DEFAULT_TEMPLATE,
    closing_image: str | Path | None = None,
    planning_book: str | Path | None = None,
    planned_pct: int = DEFAULT_PLANNED_BANDWIDTH_PCT,
    log_path: str | Path | None = None,
) -> ReportResult:
    output_path = Path(output_path)
    log_file = Path(log_path) if log_path else default_log_path(output_path)
    log = RunLog(log_file)

    try:
        return _generate_report(
            output_path=output_path,
            data_file=data_file,
            chart_week=chart_week,
            report_date=report_date,
            assets_dir=assets_dir,
            template_path=template_path,
            closing_image=closing_image,
            planning_book=planning_book,
            planned_pct=planned_pct,
            log=log,
        )
    except WsrDataError as exc:
        log.exception(exc)
        exc.log_path = log.path
        raise
    except Exception as exc:
        log.exception(exc)
        wrapped = WsrDataError(str(exc), log_path=log.path)
        raise wrapped from exc
    finally:
        log.close()


def _generate_report(
    *,
    output_path: Path,
    data_file: str,
    chart_week: int | None,
    report_date: str | None,
    assets_dir: str | Path,
    template_path: str | Path,
    closing_image: str | Path | None,
    planning_book: str | Path | None,
    planned_pct: int,
    log: RunLog,
) -> ReportResult:
    assets_dir = Path(assets_dir)
    assets_dir.mkdir(exist_ok=True)
    closing_image_path = Path(closing_image) if closing_image else None

    log.info(f"Output: {output_path}")
    log.info(f"Report date arg: {report_date!r}; chart week arg: {chart_week!r}")
    log.info(f"Planned quarter %: {planned_pct}")

    scrum_path = validate_scrum_workbook(data_file, log)
    planning_book_path = validate_planning_book(planning_book, log)

    timing = resolve_report_timing(
        scrum_path,
        chart_week=chart_week,
        report_date=report_date,
        log=log,
    )
    workbook = load_scrum_workbook(scrum_path, log)
    charts = build_chart_assets(
        scrum_path,
        assets_dir,
        planning_book_path,
        log,
        planned_pct=planned_pct,
    )
    prs = build_presentation(
        template_path=Path(template_path),
        timing=timing,
        workbook=workbook,
        charts=charts,
        assets_dir=assets_dir,
        closing_image=closing_image_path,
        log=log,
    )

    stream = io.BytesIO()
    prs.save(stream)
    output_path.write_bytes(stream.getvalue())
    sanitize_pptx(output_path)
    log.info(f"Report written: {output_path}")
    if log.warnings:
        log.info(f"Completed with {len(log.warnings)} warning(s).")
    else:
        log.info("Completed with no warnings.")

    return ReportResult(
        output_path=output_path,
        log_path=log.path,
        warnings=list(log.warnings),
    )
