#!/usr/bin/env python3
"""
Generate CES PFS CSAR (Non-STLA) Weekly Status Report (PowerPoint).

Uses templates/CES_CSAR_WSR_Template.pptx for brand fonts (Work Sans), colours, and
table styles extracted from the reference deck.

================================================================================
TODO — STRUCTURAL / NARRATIVE SECTIONS (not tabular data from Excel)
================================================================================
1. Agenda slide bullets
2. Closing slide backdrop image path (optional --closing-image)
================================================================================
SOURCED FROM data.xlsm and Book2.xlsx
================================================================================
- Pending eval/impl DCR tables
- DCR status summary callout boxes
- DDP MS4-5 slide
- Eval handoff slide (Non_STLA Planning dates)
- Discussion points (at-risk DCRs)
- Charts and graph totals
- Quarterly planning chart (Book2.xlsx)
================================================================================
"""

from __future__ import annotations

import argparse

from wsr.constants import DEFAULT_DATA_FILE
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
        default=25,
        help="Chart week number (pending tables use week-1, matching reference PDF)",
    )
    parser.add_argument("--date", default=None, help="Report date label (dd-mm-yyyy)")
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
    print(f"Chart week: {args.week} | Pending tables week: {pending_week_for_chart(args.week)}")


if __name__ == "__main__":
    main()
