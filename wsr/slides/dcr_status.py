"""Slide 4 — DCR status charts and summary panel."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt

from wsr.constants import (
    DCR_CHART_HEIGHT,
    DCR_CHART_LEFT,
    DCR_CHART_WIDTH,
    DCR_EVAL_TOP,
    DCR_IMPL_TOP,
    DCR_NOTES_GAP,
    DCR_PANEL_LEFT,
    DCR_PANEL_WIDTH,
    DCR_STATUS_NOTE_LINES,
    DCR_SUMMARY_TOP,
    DCR_TITLE_TOP,
)
from wsr.slides.base import add_summary_key_value_table, new_content_slide
from wsr.tracker import format_ordinal_day_month, format_quarter_label
from wsr_style import FONT_BODY, TEXT_DARK, raise_slide_title, set_run_font


def dcr_status_slide_title(report_date: str) -> str:
    quarter = format_quarter_label(report_date)
    till = format_ordinal_day_month(report_date)
    return (
        f"DCR Status {quarter} - CSAR (Non-STLA) & Core 2 program - PFS (till {till})"
    )


def add_dcr_status_slide(
    prs: Presentation,
    report_date: str,
    impl_chart: Path,
    eval_chart: Path,
    summary_rows: list[tuple[str, str]],
) -> None:
    slide = new_content_slide(
        prs,
        dcr_status_slide_title(report_date),
        report_date,
        4,
        title_size=Pt(18),
    )
    raise_slide_title(slide, top_in=DCR_TITLE_TOP)

    slide.shapes.add_picture(
        str(eval_chart),
        Inches(DCR_CHART_LEFT),
        Inches(DCR_EVAL_TOP),
        width=Inches(DCR_CHART_WIDTH),
        height=Inches(DCR_CHART_HEIGHT),
    )
    slide.shapes.add_picture(
        str(impl_chart),
        Inches(DCR_CHART_LEFT),
        Inches(DCR_IMPL_TOP),
        width=Inches(DCR_CHART_WIDTH),
        height=Inches(DCR_CHART_HEIGHT),
    )

    table_bottom = add_summary_key_value_table(
        slide,
        summary_rows,
        left=DCR_PANEL_LEFT,
        top=DCR_SUMMARY_TOP,
        width=DCR_PANEL_WIDTH,
    )

    notes_top = table_bottom + DCR_NOTES_GAP
    notes = slide.shapes.add_textbox(
        Inches(DCR_PANEL_LEFT),
        Inches(notes_top),
        Inches(DCR_PANEL_WIDTH),
        Inches(7.45 - notes_top),
    )
    notes_tf = notes.text_frame
    notes_tf.word_wrap = True
    notes_tf.clear()

    header = notes_tf.paragraphs[0]
    header_run = header.add_run()
    header_run.text = "# Note :"
    set_run_font(header_run, size=Pt(10), bold=True, color=TEXT_DARK, name=FONT_BODY)

    for index, line in enumerate(DCR_STATUS_NOTE_LINES, start=1):
        paragraph = notes_tf.add_paragraph()
        run = paragraph.add_run()
        run.text = f"{index}. {line}"
        set_run_font(run, size=Pt(7.5), color=TEXT_DARK, name=FONT_BODY)
