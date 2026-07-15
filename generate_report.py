#!/usr/bin/env python3
"""Generate CES PFS CSAR (Non-STLA) Weekly Status Report (PowerPoint)."""

from __future__ import annotations

import argparse

from wsr.constants import DEFAULT_DATA_FILE
from wsr.graph import latest_reported_week
from wsr.pending import pending_week_for_chart
from wsr.report import generate_report
from wsr_style import DEFAULT_TEMPLATE


def main():
    parser = argparse.ArgumentParser(description="Generate WSR PowerPoint report from data.xlsm")
    parser.add_argument("--data", default=DEFAULT_DATA_FILE, help="Path to Excel workbook")
    parser.add_argument("--output", default="WSR_Report.pptx", help="Output PowerPoint path")
    parser.add_argument(
        "--week",
        type=int,
        default=None,
        help="Chart week number (auto-detected from the graph sheet if omitted; "
        "pending table titles use the same week)",
    )
    parser.add_argument(
        "--date",
        default=None,
        help="Report date dd-mm-yyyy (shown on slides; also used as the Planned "
        "Completion cutoff for evaluation/implementation pending on slides 5–6). "
        "Auto-detected from the graph sheet if omitted.",
    )
    parser.add_argument("--assets-dir", default="report_assets", help="Directory for chart images")
    parser.add_argument(
        "--closing-image",
        default=None,
        help="Closing slide backdrop image (default: report_assets/closing_backdrop.png)",
    )
    parser.add_argument(
        "--planning-book",
        default=None,
        help="Quarterly planning workbook for slide 11 (default: ./Book2.xlsx)",
    )
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE),
        help="Branded PowerPoint template (default: templates/CES_CSAR_WSR_Template.pptx)",
    )
    args = parser.parse_args()

    week = args.week
    if week is None:
        detected_week, _ = latest_reported_week(args.data)
        week = detected_week

    output = generate_report(
        output_path=args.output,
        data_file=args.data,
        chart_week=args.week,
        report_date=args.date,
        assets_dir=args.assets_dir,
        template_path=args.template,
        closing_image=args.closing_image,
        planning_book=args.planning_book,
    )
    print(f"Report generated: {output}")
    print(f"Template: {args.template}")
    if week is not None:
        print(f"Chart week: {week} | Pending tables week: {pending_week_for_chart(week)}")


if __name__ == "__main__":
    main()
