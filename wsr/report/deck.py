"""
Assemble all WSR slides into one PowerPoint.

This is where slide order is defined. Each add_*_slide() function lives under
wsr/slides/ and is responsible for ONE slide's layout.

Automation notes for stakeholders:
  Automated:     1–11 (handoff from Onsite Evaluator = Yes)
  Manual body:   (none currently)
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from wsr.constants import TRACKER_SHEET
from wsr.pending import pending_items
from wsr.report.models import ChartAssets, ReportTiming, ScrumWorkbook
from wsr.report_data import summary_table_rows
from wsr.report_data.action_items import action_items
from wsr.report_data.ddp import ddp_ms45_items
from wsr.report_data.ecm import ecm_testing_items
from wsr.report_data.handoff import eval_handoff_items
from wsr.report_data.risks import active_risk_items
from wsr.run_log import RunLog
from wsr.slides import (
    add_agenda_slide,
    add_closing_slide,
    add_dcr_status_slide,
    add_ddp_slide,
    add_ecm_slide,
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

    risk_items = active_risk_items(workbook.risks)
    log.info(f"Active risk rows (open/pending/in progress): {len(risk_items)}")

    tracker_rich = workbook.rich_runs.get(TRACKER_SHEET, {})

    ddp_items = ddp_ms45_items(workbook.tracker, rich_runs=tracker_rich)
    log.info(f"Active DDP MS4-5 rows (testing needed, not closed): {len(ddp_items)}")

    ecm_items = ecm_testing_items(workbook.tracker, rich_runs=tracker_rich)
    log.info(
        "ECM testing rows (needed = Yes; status At Risk / In Progress / "
        f"On Hold / Yet to Start): {len(ecm_items)}"
    )

    handoff_items = eval_handoff_items(workbook.tracker, rich_runs=tracker_rich)
    log.info(f"Eval handoff rows (Onsite Evaluator = Yes): {len(handoff_items)}")

    action_item_rows = action_items(workbook.action_items)
    log.info(
        f"Active action items (external; not closed/cancelled/rejected): {len(action_item_rows)}"
    )

    summary_rows = summary_table_rows(str(workbook.path), tracker=workbook.tracker)
    log.info(
        "Slide 4 Rejected/Deferred from Non STLA "
        f"(Rejected={summary_rows[-2][1]}; Deferred={summary_rows[-1][1]})"
    )

    log.info("Assembling PowerPoint…")
    prs = Presentation(str(template_path))
    delete_all_slides(prs)

    report_date = timing.report_date
    pending_week = timing.pending_week
    quarter = timing.quarter_long

    add_title_slide(prs, report_date)
    add_agenda_slide(prs, report_date)
    add_mom_slide(prs, report_date, action_item_rows)
    add_dcr_status_slide(prs, report_date, charts.impl_chart, charts.eval_chart, summary_rows)

    slide_no = 5
    eval_pages = add_pending_slide(
        prs,
        f"{quarter} – Evaluations planned for closure for week {pending_week}",
        report_date,
        slide_no,
        eval_pending,
        mode="evaluation",
    )
    log.info(f"Eval pending split across {eval_pages} slide(s)")
    slide_no += eval_pages

    impl_pages = add_pending_slide(
        prs,
        f"{quarter} – Implementation planned for closure for week {pending_week}",
        report_date,
        slide_no,
        impl_pending,
        mode="implementation",
    )
    log.info(f"Impl pending split across {impl_pages} slide(s)")
    slide_no += impl_pages

    add_ddp_slide(prs, report_date, ddp_items, slide_number=slide_no)
    slide_no += 1
    ecm_pages = add_ecm_slide(prs, report_date, ecm_items, slide_number=slide_no)
    log.info(f"ECM testing split across {ecm_pages} slide(s)")
    slide_no += ecm_pages
    add_handoff_slide(
        prs,
        report_date,
        handoff_items,
        slide_number=slide_no,
        quarter_label=quarter,
    )
    slide_no += 1
    add_risks_slide(prs, report_date, risk_items, slide_number=slide_no)
    slide_no += 1
    add_planning_slide(
        prs,
        report_date,
        charts.quarterly_planning,
        chart_image=charts.planning_chart,
        slide_number=slide_no,
        quarter_label=quarter,
        fiscal_year=timing.fiscal_year,
    )
    slide_no += 1
    add_closing_slide(
        prs,
        report_date,
        assets_dir=assets_dir,
        backdrop_path=closing_image,
        slide_number=slide_no,
    )
    return prs
