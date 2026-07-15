"""Data table slides: pending DCRs, DDP, handoff, and discussion."""

from __future__ import annotations

from pptx import Presentation

from wsr.slides.base import add_table, empty_row, new_content_slide


def add_pending_slide(
    prs: Presentation,
    title: str,
    report_date: str,
    slide_number: int,
    items: list[dict],
    mode: str,
) -> None:
    slide = new_content_slide(prs, title, report_date, slide_number)

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
        widths = [0.58, 0.9, 3.4, 1.85, 1.45, 3.0]
    else:
        headers = [
            "Sr No",
            "DCR ID",
            "Summary",
            "Current Status",
            "Impl Closure Date",
            "Support Required",
        ]
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
        widths = [0.58, 0.9, 3.4, 1.85, 1.45, 3.0]

    if not rows:
        rows = [empty_row(len(headers))]

    add_table(slide, headers, rows, col_widths=widths)


def add_ddp_slide(prs: Presentation, report_date: str, items: list[dict]) -> None:
    slide = new_content_slide(prs, "PFS DDP Details MS 4-5", report_date, 7)
    headers = [
        "Sr.No",
        "DCR ID",
        "DCR Summary",
        "Plan Dates",
        "Appeared Dates",
        "Program",
        "Dependencies",
        "Remarks",
    ]
    rows = [
        [
            str(item["sr_no"]),
            item["dcr_id"],
            item["summary"],
            item["plan_date"],
            item["appeared_date"],
            item["program"],
            item.get("dependencies", "-"),
            item["remarks"],
        ]
        for item in items
    ]
    if not rows:
        rows = [empty_row(len(headers))]
    add_table(
        slide,
        headers,
        rows,
        col_widths=[0.55, 0.9, 3.2, 1.2, 1.25, 0.9, 2.2, 2.5],
    )


def add_handoff_slide(prs: Presentation, report_date: str, items: list[dict] | None = None) -> None:
    del items
    slide = new_content_slide(prs, "Q3-2026 – Eval Handoff from onsite", report_date, 8)
    headers = ["Sr. No.", "DCR ID", "Summary", "Evaluator", "Eval Handoff Date", "Remark"]
    rows = [[""] * len(headers)]
    add_table(slide, headers, rows, col_widths=[0.62, 0.9, 3.6, 1.05, 1.55, 2.6])


def add_discussion_slide(prs: Presentation, report_date: str, items: list[dict]) -> None:
    del items
    slide = new_content_slide(prs, "Discussion Points", report_date, 9)
    headers = ["#", "Descriptions", "Program", "Priority", "Plan completion dates", "Remarks"]
    rows = [[""] * len(headers)]
    add_table(slide, headers, rows, col_widths=[0.42, 3.8, 0.95, 0.9, 1.55, 3.0])
