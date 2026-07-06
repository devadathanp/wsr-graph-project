"""Slide 10 — Risks, issues, and impact legend."""

from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches

from wsr.slides.base import add_table, empty_row, new_content_slide
from wsr_style import TABLE_LEFT_IN, style_body_run


def _add_impact_legend(slide) -> None:
    """Risk/issues impact legend with spacing so labels do not overlap."""
    legend_items = [
        (RGBColor(0xFF, 0x00, 0x00), "High Impact / High Possibility"),
        (RGBColor(0xFF, 0xC0, 0x00), "Medium Impact / Medium Possibility"),
        (RGBColor(0x92, 0xD0, 0x50), "Low Impact / Low Possibility"),
    ]
    square_top = 6.36
    text_top = 6.33
    item_width = 3.95
    gap = 0.35
    start_left = TABLE_LEFT_IN

    for index, (color, label) in enumerate(legend_items):
        left = start_left + index * (item_width + gap)
        square = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(left),
            Inches(square_top),
            Inches(0.22),
            Inches(0.22),
        )
        square.fill.solid()
        square.fill.fore_color.rgb = color
        square.line.fill.background()

        box = slide.shapes.add_textbox(
            Inches(left + 0.3),
            Inches(text_top),
            Inches(item_width - 0.3),
            Inches(0.34),
        )
        text_frame = box.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        run = text_frame.paragraphs[0].add_run()
        run.text = label
        style_body_run(run)


def add_risks_slide(prs: Presentation, report_date: str) -> None:
    slide = new_content_slide(prs, "Risks & Mitigation Plan", report_date, 10)
    headers = ["#", "Risk", "Impact", "Risk Mitigation / Contingency"]
    rows = [empty_row(len(headers))]
    risks_top = add_table(slide, headers, rows, col_widths=[0.45, 4.4, 1.15, 4.6])

    issue_headers = ["#", "Issues", "Impact", "Contingency action"]
    issue_rows = [empty_row(len(issue_headers))]
    issue_top = risks_top + min(0.34 * (len(rows) + 1), 5.8) + 0.35
    add_table(slide, issue_headers, issue_rows, top=issue_top, col_widths=[0.45, 4.4, 1.15, 4.6])

    _add_impact_legend(slide)
