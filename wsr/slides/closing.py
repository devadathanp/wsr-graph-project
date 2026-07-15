"""Slide 12 — Closing slide with backdrop image."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from wsr.constants import LAYOUT_OPENING
from wsr_style import (
    CLOSING_HEADLINE_SIZE,
    CLOSING_SUBLINE_SIZE,
    DEFAULT_CLOSING_BACKDROP,
    FONT_BODY,
    FONT_MAJOR,
    FOOTER_DATE_HEIGHT,
    FOOTER_DATE_LEFT,
    FOOTER_DATE_TOP,
    FOOTER_DATE_WIDTH,
    FOOTER_NUMBER_HEIGHT,
    FOOTER_NUMBER_LEFT,
    FOOTER_NUMBER_TOP,
    FOOTER_NUMBER_WIDTH,
    FOOTER_SIZE,
    WHITE,
    find_placeholder,
    set_run_font,
)


def _send_shape_to_back(shape) -> None:
    sp_tree = shape._element.getparent()
    if sp_tree is None:
        return
    sp_tree.remove(shape._element)
    insert_at = 1 if len(sp_tree) > 0 else 0
    sp_tree.insert(insert_at, shape._element)


def _clear_placeholder_text(slide, idx: int) -> None:
    placeholder = find_placeholder(slide, idx=idx)
    if placeholder is not None:
        placeholder.text = ""


def resolve_closing_backdrop(assets_dir: Path, override: Path | None = None) -> Path | None:
    if override is not None and override.exists():
        return override
    candidates = [
        assets_dir / "closing_backdrop.png",
        assets_dir / "closing_backdrop.jpg",
        Path("closing_backdrop.png"),
        Path("closing_backdrop.jpg"),
        DEFAULT_CLOSING_BACKDROP,
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _set_closing_footer(slide, report_date: str, slide_number: int) -> None:
    for left, top, width, height, text in (
        (FOOTER_DATE_LEFT, FOOTER_DATE_TOP, FOOTER_DATE_WIDTH, FOOTER_DATE_HEIGHT, report_date),
        (FOOTER_NUMBER_LEFT, FOOTER_NUMBER_TOP, FOOTER_NUMBER_WIDTH, FOOTER_NUMBER_HEIGHT, str(slide_number)),
    ):
        box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        text_frame = box.text_frame
        text_frame.clear()
        text_frame.margin_left = text_frame.margin_right = text_frame.margin_top = text_frame.margin_bottom = 0
        paragraph = text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.RIGHT
        run = paragraph.add_run()
        run.text = text
        set_run_font(run, size=FOOTER_SIZE, color=WHITE, name=FONT_BODY)


def add_closing_slide(
    prs: Presentation,
    report_date: str,
    *,
    assets_dir: Path,
    backdrop_path: Path | None = None,
) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_OPENING])

    backdrop = resolve_closing_backdrop(assets_dir, backdrop_path)
    if backdrop is not None:
        picture = slide.shapes.add_picture(
            str(backdrop),
            Inches(0),
            Inches(0),
            width=prs.slide_width,
            height=prs.slide_height,
        )
        _send_shape_to_back(picture)

    for placeholder_idx in (0, 16, 23):
        _clear_placeholder_text(slide, placeholder_idx)

    headline_box = slide.shapes.add_textbox(Inches(0.7), Inches(4.05), Inches(11.9), Inches(1.5))
    headline_frame = headline_box.text_frame
    headline_frame.clear()
    headline_frame.word_wrap = True
    headline_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    headline_paragraph = headline_frame.paragraphs[0]
    headline_paragraph.alignment = PP_ALIGN.CENTER
    headline_run = headline_paragraph.add_run()
    headline_run.text = "We are imagining mobility with you"
    set_run_font(
        headline_run,
        size=CLOSING_HEADLINE_SIZE,
        bold=True,
        color=WHITE,
        name=FONT_MAJOR,
    )

    subline_box = slide.shapes.add_textbox(Inches(1.2), Inches(5.45), Inches(10.9), Inches(0.55))
    subline_frame = subline_box.text_frame
    subline_frame.clear()
    subline_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    subline_paragraph = subline_frame.paragraphs[0]
    subline_paragraph.alignment = PP_ALIGN.CENTER
    subline_run = subline_paragraph.add_run()
    subline_run.text = "Lets collaborate"
    set_run_font(
        subline_run,
        size=CLOSING_SUBLINE_SIZE,
        bold=False,
        color=WHITE,
        name=FONT_BODY,
    )

    _set_closing_footer(slide, report_date, 12)
