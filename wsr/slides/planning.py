"""Slide 11 — Quarterly planning chart."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from wsr.slides.base import new_content_slide
from wsr_style import style_body_run


def add_planning_slide(
    prs: Presentation,
    report_date: str,
    qp: dict[str, int] | None,
    *,
    chart_image: Path | None = None,
) -> None:
    slide = new_content_slide(prs, "Quarterly Planning 2026-Non STLA", report_date, 11)

    if qp is None:
        note = slide.shapes.add_textbox(Inches(0.59), Inches(2.5), Inches(12.18), Inches(0.8))
        note_run = note.text_frame.paragraphs[0].add_run()
        note_run.text = "No quarterly planning data found in Book2.xlsx."
        style_body_run(note_run)
        return

    if chart_image is not None and chart_image.exists():
        slide.shapes.add_picture(
            str(chart_image),
            Inches(1.29),
            Inches(0.93),
            width=Inches(11.20),
            height=Inches(4.91),
        )

    bandwidth = slide.shapes.add_textbox(Inches(0.59), Inches(5.85), Inches(10.85), Inches(0.40))
    bw_run = bandwidth.text_frame.paragraphs[0].add_run()
    bw_run.text = f"Bandwidth is planned with {qp['resources']} Resources"
    style_body_run(bw_run, bold=False)

    note = slide.shapes.add_textbox(Inches(0.59), Inches(6.25), Inches(12.18), Inches(0.40))
    note_run = note.text_frame.paragraphs[0].add_run()
    note_run.text = (
        "Note: The Estimations for Q3 Planning are in progress, "
        "these are high level tentative estimations"
    )
    style_body_run(note_run)
