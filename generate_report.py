#!/usr/bin/env python3
"""
Generate CES PFS CSAR (Non-STLA) Weekly Status Report (PowerPoint).

Uses templates/CES_CSAR_WSR_Template.pptx for brand fonts (Work Sans), colours, and
table styles extracted from the reference deck.

================================================================================
TODO — HARDCODED SECTIONS (not sourced from data.xlsm yet)
================================================================================
1. Presenter name on title slide
2. Agenda slide bullets
3. MOM & Action Items (meeting-specific narrative)
4. DCR status summary callout boxes (CSAR/Core2/ECM/DDP split, rejected/deferred)
5. DCR status weekly narrative paragraphs (WK22–WK24 editorial comments)
6. Eval handoff evaluator names and handoff dates
7. Risks & Mitigation content
8. Issues slide content
9. Quarterly planning chart values
10. Closing branding slide
11. Reference DDP DCR ID list (when row missing from DDP_Plan)
12. Reference pending-eval / pending-impl DCR snapshots per week
================================================================================
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from wsr_charts import save_evaluation_chart, save_implementation_chart
from wsr_common import (
    DEFAULT_DATA_FILE,
    discussion_points,
    ddp_ms45_items,
    get_evaluation_data,
    get_implementation_data,
    graph_week_capacity,
    load_ddp_plan,
    load_graph_summary,
    load_tracker,
    load_visibility,
    pending_items,
    pending_week_for_chart,
    tracker_lookup,
    week_remarks,
)
from wsr_style import (
    DEFAULT_TEMPLATE,
    FONT_BODY,
    TEXT_DARK,
    TEXT_MUTED,
    add_number_badge,
    delete_all_slides,
    find_placeholder,
    set_run_font,
    set_slide_footer,
    set_slide_title,
    style_agenda_run,
    style_body_run,
    style_table_cells,
    style_title_run,
)

LAYOUT_OPENING = 13
LAYOUT_CONTENT = 3

HARDCODED_PRESENTER = "Chitharenjan Nair"
HARDCODED_AGENDA = [
    "MOM & Action Items",
    "DCR Status",
    "Discussion Points",
    "Issues and Risks",
]
HARDCODED_MOM_ROWS = [
    {
        "discussion": "DCR 18922156 Phase 35 OBD REAL Regen Trigger Issue (Core 2 DCR) — "
        "Implementation deferred to Q4'26. Eval completed.",
        "action": "PFS Team to work on new Defect DCR for AftSootFiltFrqntRegenDiag.",
        "status": "Eval completed, Implementation deferred",
        "owner": "PFS Team",
        "closure": "18-06-2026",
        "remarks": "Yue and Dhanraj confirmed OBD REAL Regen Trigger issue is not urgent in Q3'26.",
    }
]
HARDCODED_SUMMARY_CALLOUTS = {
    "total_planned": "95 >> 97 (Non STLA + Core 2) + ECM Testing",
    "csar": "CSAR 74 >> 76",
    "core2": "Core2 9",
    "ecm_testing": "ECM Testing 12",
    "ddp_testing": "DDP Testing 7",
    "eval_planned": "DCR's Planned for Evaluation 46 >> 48",
    "impl_planned": "DCR's Planned for Implementation 77 >> 79",
    "rejected": "DCR's Rejected — Eval: 02, Impl: 03",
    "deferred": "DCR's Deferred — Eval: 01, Impl: 01",
}
HARDCODED_WEEKLY_NARRATIVE = {
    22: "WK:22 — 1 DCR: L2 Comments fixing",
    23: "WK:23 — Same as WK22 | 1 DCR: Awaiting CDS Review | 2 DCR's Awaiting L2 Review",
    24: "WK:24 — 12 DCR Closure pending | 5 DCR for L2 Reverification | 1 DCR WIP | 6 DCR Eval Closed",
}
HARDCODED_HANDOFF_ROWS = [
    {"dcr_id": 19138223, "evaluator": "Yue", "handoff_date": "19 Jun 2026", "remark": "Eval handoff from onsite"},
    {"dcr_id": 19578849, "evaluator": "Duygu", "handoff_date": "19 Jun 2026", "remark": "Eval handoff from onsite"},
    {
        "dcr_id": 18922156,
        "evaluator": "Duygu",
        "handoff_date": "03 Jun 2026 >> 16 Jun 2026",
        "remark": "Eval artifacts received",
    },
]
HARDCODED_QUARTERLY_PLANNING = {
    "available_hours": 7514,
    "planned_pct": 90,
    "planned_hours": 6754,
    "resources": 17,
}


def _new_content_slide(prs: Presentation, title: str, report_date: str, slide_number: int):
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_CONTENT])
    set_slide_title(slide, title)
    set_slide_footer(slide, report_date, slide_number)
    return slide


def _add_title_slide(prs: Presentation, report_date: str):
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_OPENING])

    title_ph = find_placeholder(slide, idx=0)
    if title_ph is not None:
        title_ph.text = "CES PFS CSAR (Non- STLA) and CORE 2\nWeekly Status Report"
        for paragraph in title_ph.text_frame.paragraphs:
            for run in paragraph.runs:
                style_title_run(run)

    date_ph = find_placeholder(slide, idx=16)
    if date_ph is not None:
        date_ph.text = report_date
        for paragraph in date_ph.text_frame.paragraphs:
            for run in paragraph.runs:
                style_body_run(run)

    presenter_ph = find_placeholder(slide, idx=23)
    if presenter_ph is not None:
        presenter_ph.text = HARDCODED_PRESENTER
        for paragraph in presenter_ph.text_frame.paragraphs:
            for run in paragraph.runs:
                style_body_run(run, bold=True)

    set_slide_footer(slide, report_date, 1)


def _add_agenda_slide(prs: Presentation, report_date: str):
    slide = _new_content_slide(prs, "Agenda", report_date, 2)
    top_positions = [1.55, 2.65, 3.57, 4.59]
    for idx, (item, top) in enumerate(zip(HARDCODED_AGENDA, top_positions), start=1):
        add_number_badge(slide, str(idx), left=0.67, top=top - 0.33, accent=idx >= 3)
        box = slide.shapes.add_textbox(Inches(1.66), Inches(top), Inches(9.5), Inches(0.45))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = item
        style_agenda_run(run)


def _add_mom_slide(prs: Presentation, report_date: str):
    slide = _new_content_slide(prs, "MOM of 11/06,18/06", report_date, 3)
    headers = ["Sr. No.", "Discussion points", "Action Items", "Status", "Ownership", "Closure", "Remarks"]
    rows = [
        [
            str(idx),
            row["discussion"],
            row["action"],
            row["status"],
            row["owner"],
            row["closure"],
            row["remarks"],
        ]
        for idx, row in enumerate(HARDCODED_MOM_ROWS, start=1)
    ]
    _add_table(slide, headers, rows, top=1.0, col_widths=[0.55, 2.3, 1.9, 1.2, 0.9, 0.8, 2.0])


def _add_dcr_status_slide(
    prs: Presentation,
    report_date: str,
    impl_chart: Path,
    eval_chart: Path,
    summary: dict,
    chart_week: int,
    pending_week: int,
    eval_data,
    impl_data,
):
    slide = _new_content_slide(
        prs,
        "DCR status Q3'26 – CSAR (Non-STLA)  & Core 2 program - PFS (Evaluation and Implementation)",
        report_date,
        4,
    )

    slide.shapes.add_picture(str(impl_chart), Inches(0.12), Inches(1.05), width=Inches(6.45))
    slide.shapes.add_picture(str(eval_chart), Inches(6.62), Inches(1.05), width=Inches(6.45))

    callout = slide.shapes.add_textbox(Inches(0.2), Inches(4.7), Inches(6.2), Inches(2.1))
    tf = callout.text_frame
    tf.word_wrap = True
    lines = [
        f"Total DCR's planned: {HARDCODED_SUMMARY_CALLOUTS['total_planned']}",
        f"{HARDCODED_SUMMARY_CALLOUTS['csar']} | {HARDCODED_SUMMARY_CALLOUTS['core2']} | "
        f"{HARDCODED_SUMMARY_CALLOUTS['ecm_testing']} | {HARDCODED_SUMMARY_CALLOUTS['ddp_testing']}",
        f"{HARDCODED_SUMMARY_CALLOUTS['eval_planned']} | {HARDCODED_SUMMARY_CALLOUTS['impl_planned']}",
        f"{HARDCODED_SUMMARY_CALLOUTS['rejected']} | {HARDCODED_SUMMARY_CALLOUTS['deferred']}",
        f"Data totals — Eval: {summary.get('eval_baseline')}/{summary.get('eval_revised')}/"
        f"{summary.get('eval_completed')} | Impl: {summary.get('impl_baseline')}/"
        f"{summary.get('impl_revised')}/{summary.get('impl_completed')}",
    ]
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = para.add_run()
        run.text = line
        set_run_font(run, size=Pt(8), color=TEXT_DARK, name=FONT_BODY)

    notes = slide.shapes.add_textbox(Inches(6.7), Inches(4.7), Inches(6.1), Inches(2.1))
    ntf = notes.text_frame
    ntf.word_wrap = True
    note_lines = []
    for week in (pending_week - 1, pending_week, chart_week):
        hardcoded = HARDCODED_WEEKLY_NARRATIVE.get(week)
        from_data_eval = week_remarks(eval_data, week)
        from_data_impl = week_remarks(impl_data, week)
        if hardcoded:
            note_lines.append(hardcoded)
        elif from_data_eval or from_data_impl:
            note_lines.append(f"WK:{week} — Eval: {from_data_eval or '-'} | Impl: {from_data_impl or '-'}")
    for i, line in enumerate(note_lines):
        para = ntf.paragraphs[0] if i == 0 else ntf.add_paragraph()
        run = para.add_run()
        run.text = line
        set_run_font(run, size=Pt(8), color=TEXT_MUTED, name=FONT_BODY)


def _add_pending_slide(
    prs: Presentation,
    title: str,
    report_date: str,
    slide_number: int,
    items: list[dict],
    mode: str,
):
    slide = _new_content_slide(prs, title, report_date, slide_number)

    if mode == "evaluation":
        headers = ["Sr No", "DCR ID", "Summary", "Current Status", "Eval Closure date", "Support Required"]
        rows = [
            [
                str(i + 1),
                str(item["dcr_id"]),
                item["summary"],
                item["status"],
                item["closure_date"],
                item["support"],
            ]
            for i, item in enumerate(items)
        ]
        widths = [0.5, 0.75, 2.7, 1.6, 1.0, 2.6]
    else:
        headers = ["Sr No", "DCR ID", "Summary", "Current Status", "Remarks"]
        rows = [
            [
                str(i + 1),
                str(item["dcr_id"]),
                item["summary"],
                item["status"],
                item["remarks"],
            ]
            for i, item in enumerate(items)
        ]
        widths = [0.5, 0.75, 3.2, 1.8, 2.9]

    if not rows:
        rows = [["-", "-", "No open items found with current filters", "-", "-"]]
        if mode == "evaluation":
            rows[0].append("-")

    _add_table(slide, headers, rows, top=0.95, col_widths=widths)


def _add_ddp_slide(prs: Presentation, report_date: str, items: list[dict]):
    slide = _new_content_slide(prs, "PFS DDP Details MS 4-5", report_date, 7)
    headers = ["Sr.No", "DCR ID", "DCR Summary", "Plan Dates", "Appeared Dates", "Program", "Remarks"]
    rows = [
        [
            str(item["sr_no"]),
            item["dcr_id"],
            item["summary"],
            item["plan_date"],
            item["appeared_date"],
            item["program"],
            item["remarks"],
        ]
        for item in items
    ]
    if not rows:
        rows = [["-", "-", "No DDP MS4-5 rows matched filters", "-", "-", "-", "-"]]
    _add_table(slide, headers, rows, top=0.95, col_widths=[0.5, 0.75, 2.5, 1.0, 1.0, 0.8, 2.5])


def _add_handoff_slide(prs: Presentation, report_date: str, tracker_map: dict):
    slide = _new_content_slide(prs, "Q3-2026 – Eval Handoff from onsite", report_date, 8)
    headers = ["Sr. No.", "DCR ID", "Summary", "Evaluator", "Eval Handoff Date", "Remark"]
    rows = []
    for idx, item in enumerate(HARDCODED_HANDOFF_ROWS, start=1):
        tracker_row = tracker_map.get(item["dcr_id"])
        summary = tracker_row.get("Summary", "-") if tracker_row is not None else "-"
        rows.append(
            [
                str(idx),
                str(item["dcr_id"]),
                str(summary),
                item["evaluator"],
                item["handoff_date"],
                item["remark"],
            ]
        )
    _add_table(slide, headers, rows, top=0.95, col_widths=[0.55, 0.75, 2.9, 0.9, 1.2, 1.9])


def _add_discussion_slide(prs: Presentation, report_date: str, items: list[dict]):
    slide = _new_content_slide(prs, "Discussion Points", report_date, 9)
    headers = ["#", "Descriptions", "Program", "Priority", "Plan completion dates", "Remarks"]
    rows = [
        [
            str(i + 1),
            item["description"],
            item["program"],
            str(item["priority"]),
            item["plan_date"],
            item["remarks"],
        ]
        for i, item in enumerate(items)
    ]
    if not rows:
        rows = [["-", "No at-risk / on-hold DCRs found", "-", "-", "-", "-"]]
    _add_table(slide, headers, rows, top=0.95, col_widths=[0.35, 3.0, 0.75, 0.7, 1.2, 2.5])


def _add_risks_slide(prs: Presentation, report_date: str):
    slide = _new_content_slide(prs, "Risks & Mitigation Plan", report_date, 10)
    headers = ["#", "Risk", "Impact", "Risk Mitigation / Contingency"]
    rows = [["NA", "NA", "NA", "NA"]]
    _add_table(slide, headers, rows, top=1.1, col_widths=[0.5, 3.5, 1.2, 3.5])

    issue_headers = ["#", "Issues", "Impact", "Contingency action"]
    issue_rows = [["-", "-", "-", "-"]]
    _add_table(slide, issue_headers, issue_rows, top=2.9, col_widths=[0.5, 3.5, 1.2, 3.5])


def _add_planning_slide(prs: Presentation, report_date: str):
    slide = _new_content_slide(prs, "Quarterly Planning 2026-Non STLA", report_date, 11)
    body = slide.shapes.add_textbox(Inches(0.7), Inches(1.4), Inches(11.5), Inches(4.5))
    tf = body.text_frame
    qp = HARDCODED_QUARTERLY_PLANNING
    lines = [
        "Note: The Estimations for Q3 Planning are in progress, these are high level tentative estimations",
        f"**Bandwidth is planned with {qp['resources']} Resources",
        str(qp["available_hours"]),
        f"{qp['planned_pct']}%, {qp['planned_hours']}",
        "Q3 Actual Available    90% of Q3 is Planned",
    ]
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = para.add_run()
        run.text = line
        style_body_run(run)


def _add_closing_slide(prs: Presentation, report_date: str):
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_OPENING])
    title_ph = find_placeholder(slide, idx=0)
    if title_ph is not None:
        title_ph.text = "We are reimagining mobility with you\nLets collaborate"
        for paragraph in title_ph.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                style_title_run(run)
    set_slide_footer(slide, report_date, 12)


def _add_table(slide, headers, rows, top=1.0, col_widths=None):
    row_height = min(0.34 * (len(rows) + 1), 5.8)
    table_shape = slide.shapes.add_table(
        len(rows) + 1,
        len(headers),
        Inches(0.25),
        Inches(top),
        Inches(12.85),
        Inches(row_height),
    )
    table = table_shape.table
    if col_widths:
        for idx, width in enumerate(col_widths):
            table.columns[idx].width = Inches(width)

    for col_idx, header in enumerate(headers):
        table.cell(0, col_idx).text = header

    for row_idx, row in enumerate(rows, start=1):
        for col_idx, value in enumerate(row):
            table.cell(row_idx, col_idx).text = str(value)

    style_table_cells(table)


def generate_report(
    output_path: str | Path = "WSR_Report.pptx",
    data_file: str = DEFAULT_DATA_FILE,
    chart_week: int = 25,
    report_date: str | None = None,
    assets_dir: str | Path = "report_assets",
    template_path: str | Path = DEFAULT_TEMPLATE,
) -> Path:
    assets_dir = Path(assets_dir)
    assets_dir.mkdir(exist_ok=True)
    pending_week = pending_week_for_chart(chart_week)

    if report_date is None:
        report_date = datetime.now().strftime("%d-%m-%Y")

    impl_chart = save_implementation_chart(assets_dir / "implementation_chart.png", data_file=data_file)
    eval_chart = save_evaluation_chart(assets_dir / "evaluation_chart.png", data_file=data_file)

    summary = load_graph_summary(data_file)
    eval_data = get_evaluation_data(data_file=data_file)
    impl_data = get_implementation_data(data_file=data_file)
    tracker = load_tracker(data_file)
    visibility = load_visibility(data_file)
    ddp = load_ddp_plan(data_file)
    tracker_map = tracker_lookup(tracker)

    eval_limit = graph_week_capacity(data_file, pending_week, "evaluation")
    eval_pending = pending_items(
        visibility, tracker_map, mode="evaluation", pending_week=pending_week, limit=eval_limit
    )
    impl_pending = pending_items(
        visibility, tracker_map, mode="implementation", pending_week=pending_week, limit=2
    )
    ddp_items = ddp_ms45_items(ddp, tracker_map)
    discussion = discussion_points(visibility, tracker_map)

    prs = Presentation(str(template_path))
    delete_all_slides(prs)

    _add_title_slide(prs, report_date)
    _add_agenda_slide(prs, report_date)
    _add_mom_slide(prs, report_date)
    _add_dcr_status_slide(
        prs, report_date, impl_chart, eval_chart, summary, chart_week, pending_week, eval_data, impl_data
    )
    _add_pending_slide(
        prs,
        f"Q3-2026 – Evaluations pending for closure for week {pending_week}",
        report_date,
        5,
        eval_pending,
        mode="evaluation",
    )
    _add_pending_slide(
        prs,
        f"Q3-2026 – Implementation pending for closure for week {pending_week}",
        report_date,
        6,
        impl_pending,
        mode="implementation",
    )
    _add_ddp_slide(prs, report_date, ddp_items)
    _add_handoff_slide(prs, report_date, tracker_map)
    _add_discussion_slide(prs, report_date, discussion)
    _add_risks_slide(prs, report_date)
    _add_planning_slide(prs, report_date)
    _add_closing_slide(prs, report_date)

    output_path = Path(output_path)
    prs.save(output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Generate WSR PowerPoint report from data.xlsm")
    parser.add_argument("--data", default=DEFAULT_DATA_FILE, help="Path to Excel workbook")
    parser.add_argument("--output", default="WSR_Report.pptx", help="Output PowerPoint path")
    parser.add_argument(
        "--week",
        type=int,
        default=25,
        help="Chart week number (pending tables use week-1, matching reference PDF)",
    )
    parser.add_argument("--date", default=None, help="Report date label (dd-mm-yyyy)")
    parser.add_argument("--assets-dir", default="report_assets", help="Directory for chart images")
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE),
        help="Branded PowerPoint template (default: templates/CES_CSAR_WSR_Template.pptx)",
    )
    args = parser.parse_args()

    output = generate_report(
        output_path=args.output,
        data_file=args.data,
        chart_week=args.week,
        report_date=args.date,
        assets_dir=args.assets_dir,
        template_path=args.template,
    )
    print(f"Report generated: {output}")
    print(f"Template: {args.template}")
    print(f"Chart week: {args.week} | Pending tables week: {pending_week_for_chart(args.week)}")


if __name__ == "__main__":
    main()
