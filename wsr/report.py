"""WSR report orchestration: load data, build slides, save deck."""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

from pptx import Presentation

from wsr.charts import save_evaluation_chart, save_implementation_chart, save_planning_chart
from wsr.constants import DEFAULT_DATA_FILE
from wsr.graph import latest_reported_week
from wsr.loaders import load_ddp_plan, load_tracker, load_visibility
from wsr.pending import pending_items, pending_week_for_chart
from wsr.planning_book import load_quarterly_planning
from wsr.pptx_sanitize import sanitize_pptx
from wsr.report_data import (
    ddp_ms45_items,
    discussion_points,
    summary_table_rows,
)
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
from wsr_style import DEFAULT_TEMPLATE, delete_all_slides


def generate_report(
    output_path: str | Path = "WSR_Report.pptx",
    data_file: str = DEFAULT_DATA_FILE,
    chart_week: int | None = None,
    report_date: str | None = None,
    assets_dir: str | Path = "report_assets",
    template_path: str | Path = DEFAULT_TEMPLATE,
    closing_image: str | Path | None = None,
    planning_book: str | Path | None = None,
) -> Path:
    assets_dir = Path(assets_dir)
    assets_dir.mkdir(exist_ok=True)
    closing_image_path = Path(closing_image) if closing_image else None
    planning_book_path = Path(planning_book) if planning_book else None

    detected_week, detected_date = latest_reported_week(data_file)
    if chart_week is None:
        if detected_week is None:
            raise ValueError(
                "Could not detect the reporting week from the graph sheet. "
                "Pass an explicit week number."
            )
        chart_week = detected_week
    if report_date is None:
        report_date = detected_date or datetime.now().strftime("%d-%m-%Y")

    pending_week = pending_week_for_chart(chart_week)

    impl_chart = save_implementation_chart(assets_dir / "implementation_chart.png", data_file=data_file)
    eval_chart = save_evaluation_chart(assets_dir / "evaluation_chart.png", data_file=data_file)

    summary_rows = summary_table_rows(data_file)
    tracker = load_tracker(data_file)
    visibility = load_visibility(data_file)
    ddp = load_ddp_plan(data_file)
    tracker_map = tracker_lookup(tracker)
    tracker_rows = tracker_rows_lookup(tracker)

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
    ddp_items = ddp_ms45_items(ddp, tracker_map)
    quarterly_planning = load_quarterly_planning(planning_book_path)
    planning_chart = None
    if quarterly_planning is not None:
        planning_chart = save_planning_chart(
            quarterly_planning,
            assets_dir / "planning_chart.png",
        )

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

    output_path = Path(output_path)
    stream = io.BytesIO()
    prs.save(stream)
    output_path.write_bytes(stream.getvalue())
    sanitize_pptx(output_path)
    return output_path
