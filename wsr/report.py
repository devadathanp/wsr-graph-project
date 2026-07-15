"""WSR report orchestration: load data, build slides, save deck."""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from pptx import Presentation

from wsr.charts import save_evaluation_chart, save_implementation_chart, save_planning_chart
from wsr.constants import DEFAULT_DATA_FILE
from wsr.errors import WsrDataError
from wsr.graph import latest_reported_week
from wsr.loaders import load_ddp_plan, load_tracker, load_visibility
from wsr.pending import pending_items, pending_week_for_chart
from wsr.planning_book import load_quarterly_planning
from wsr.pptx_sanitize import sanitize_pptx
from wsr.report_data import ddp_ms45_items, summary_table_rows
from wsr.run_log import RunLog, default_log_path
from wsr.slides import (
    add_agenda_slide,
    add_closing_slide,
    add_dcr_status_slide,
    add_ddp_slide,
    add_discussion_slide,
    add_handoff_slide,
    add_mom_slide,
    add_pending_slide,
    add_planning_slide,
    add_risks_slide,
    add_title_slide,
)
from wsr.tracker import tracker_lookup, tracker_rows_lookup
from wsr.validate import validate_planning_book, validate_scrum_workbook
from wsr_style import DEFAULT_TEMPLATE, delete_all_slides


@dataclass
class ReportResult:
    output_path: Path
    log_path: Path
    warnings: list[str] = field(default_factory=list)


def generate_report(
    output_path: str | Path = "WSR_Report.pptx",
    data_file: str = DEFAULT_DATA_FILE,
    chart_week: int | None = None,
    report_date: str | None = None,
    assets_dir: str | Path = "report_assets",
    template_path: str | Path = DEFAULT_TEMPLATE,
    closing_image: str | Path | None = None,
    planning_book: str | Path | None = None,
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
    log: RunLog,
) -> ReportResult:
    assets_dir = Path(assets_dir)
    assets_dir.mkdir(exist_ok=True)
    closing_image_path = Path(closing_image) if closing_image else None

    log.info(f"Output: {output_path}")
    log.info(f"Report date arg: {report_date!r}; chart week arg: {chart_week!r}")

    scrum_path = validate_scrum_workbook(data_file, log)
    planning_book_path = validate_planning_book(planning_book, log)

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
    log.info(f"Using chart week={chart_week}, report date={report_date}")

    pending_week = pending_week_for_chart(chart_week)

    log.info("Building charts…")
    impl_chart = save_implementation_chart(
        assets_dir / "implementation_chart.png", data_file=str(scrum_path)
    )
    eval_chart = save_evaluation_chart(
        assets_dir / "evaluation_chart.png", data_file=str(scrum_path)
    )

    summary_rows = summary_table_rows(str(scrum_path))
    tracker = load_tracker(str(scrum_path))
    visibility = load_visibility(str(scrum_path))
    ddp = load_ddp_plan(str(scrum_path))
    tracker_map = tracker_lookup(tracker)
    tracker_rows = tracker_rows_lookup(tracker)

    log.info("Selecting pending evaluation / implementation rows…")
    eval_pending = pending_items(
        visibility,
        tracker_rows,
        mode="evaluation",
        pending_week=pending_week,
        cutoff_date=report_date,
    )
    impl_pending = pending_items(
        visibility,
        tracker_rows,
        mode="implementation",
        pending_week=pending_week,
        cutoff_date=report_date,
    )
    log.info(f"Pending eval rows: {len(eval_pending)}; impl rows: {len(impl_pending)}")

    ddp_items = ddp_ms45_items(ddp, tracker_map)
    log.info(f"DDP MS4-5 rows: {len(ddp_items)}")

    quarterly_planning = load_quarterly_planning(planning_book_path)
    planning_chart = None
    if quarterly_planning is None:
        if planning_book_path is not None:
            log.warning(
                f'Could not find "Total work Hrs. Available for PFS team" in '
                f"{planning_book_path.name}; slide 11 will show a placeholder."
            )
    else:
        planning_chart = save_planning_chart(
            quarterly_planning,
            assets_dir / "planning_chart.png",
        )
        log.info(
            f"Planning chart: available={quarterly_planning['available_hours']}, "
            f"planned={quarterly_planning['planned_hours']}, "
            f"resources={quarterly_planning['resources']}"
        )

    log.info("Assembling PowerPoint…")
    prs = Presentation(str(template_path))
    delete_all_slides(prs)

    add_title_slide(prs, report_date)
    add_agenda_slide(prs, report_date)
    add_mom_slide(prs, report_date)
    add_dcr_status_slide(prs, report_date, impl_chart, eval_chart, summary_rows)
    add_pending_slide(
        prs,
        f"Q3-2026 – Evaluations pending for closure for week {pending_week}",
        report_date,
        5,
        eval_pending,
        mode="evaluation",
    )
    add_pending_slide(
        prs,
        f"Q3-2026 – Implementation pending for closure for week {pending_week}",
        report_date,
        6,
        impl_pending,
        mode="implementation",
    )
    add_ddp_slide(prs, report_date, ddp_items)
    add_handoff_slide(prs, report_date)
    add_discussion_slide(prs, report_date, [])
    add_risks_slide(prs, report_date)
    add_planning_slide(prs, report_date, quarterly_planning, chart_image=planning_chart)
    add_closing_slide(
        prs,
        report_date,
        assets_dir=assets_dir,
        backdrop_path=closing_image_path,
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
