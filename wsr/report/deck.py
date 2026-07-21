"""
Assemble all WSR slides into one PowerPoint.

This is where slide order is defined. Each add_*_slide() function lives under
wsr/slides/ and is responsible for ONE slide's layout.

Automation notes for stakeholders:
  Automated:     1, 2, 4, 5, 6, 11, 12
  Manual body:   3 (MOM), 7 (DDP), 8 (handoff), 9 (discussion), 10 (risks body)
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from wsr.pending import pending_items
from wsr.report.models import ChartAssets, ReportTiming, ScrumWorkbook
from wsr.report_data import summary_table_rows
from wsr.run_log import RunLog
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
from wsr_style import delete_all_slides


def build_presentation(
    *,
    template_path: Path,
    timing: ReportTiming,
    workbook: ScrumWorkbook,
    charts: ChartAssets,
    assets_dir: Path,
    closing_image: Path | None,
    log: RunLog,
) -> Presentation:
    log.info("Selecting pending evaluation / implementation rows…")
    eval_pending = pending_items(
        workbook.visibility,
        workbook.tracker_rows,
        mode="evaluation",
        pending_week=timing.pending_week,
        cutoff_date=timing.report_date,
    )
    impl_pending = pending_items(
        workbook.visibility,
        workbook.tracker_rows,
        mode="implementation",
        pending_week=timing.pending_week,
        cutoff_date=timing.report_date,
    )
    log.info(f"Pending eval rows: {len(eval_pending)}; impl rows: {len(impl_pending)}")

    summary_rows = summary_table_rows(str(workbook.path))

    log.info("Assembling PowerPoint…")
    prs = Presentation(str(template_path))
    delete_all_slides(prs)

    report_date = timing.report_date
    pending_week = timing.pending_week

    add_title_slide(prs, report_date)
    add_agenda_slide(prs, report_date)
    add_mom_slide(prs, report_date)
    add_dcr_status_slide(prs, report_date, charts.impl_chart, charts.eval_chart, summary_rows)
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
    add_ddp_slide(prs, report_date)
    add_handoff_slide(prs, report_date)
    add_discussion_slide(prs, report_date)
    add_risks_slide(prs, report_date)
    add_planning_slide(
        prs,
        report_date,
        charts.quarterly_planning,
        chart_image=charts.planning_chart,
    )
    add_closing_slide(
        prs,
        report_date,
        assets_dir=assets_dir,
        backdrop_path=closing_image,
    )
    return prs
