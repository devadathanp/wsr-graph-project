"""Visual theme for WSR PowerPoint reports (CES CSAR Cummins template)."""

from __future__ import annotations

from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

DEFAULT_TEMPLATE = Path(__file__).parent / "templates" / "CES_CSAR_WSR_Template.pptx"

# Cummins / CES deck palette (from template theme1.xml)
ACCENT_LIME = RGBColor(0xB0, 0xFF, 0x45)      # accent1
TEXT_DARK = RGBColor(0x16, 0x17, 0x18)       # dk1
TEXT_MUTED = RGBColor(0x4D, 0x51, 0x54)      # dk2
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT_MAJOR = "Work Sans Medium"
FONT_BODY = "Work Sans"

TABLE_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

TITLE_SIZE = Pt(28)
SECTION_TITLE_SIZE = Pt(16)
BODY_SIZE = Pt(12)
AGENDA_ITEM_SIZE = Pt(24)
TABLE_HEADER_SIZE = Pt(10.5)
TABLE_BODY_SIZE = Pt(10)
FOOTER_SIZE = Pt(9)


def set_run_font(
    run,
    *,
    size=BODY_SIZE,
    bold=False,
    color=TEXT_DARK,
    name=FONT_BODY,
):
    run.font.name = name
    run.font.size = size
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def style_title_run(run):
    set_run_font(run, size=TITLE_SIZE, bold=True, color=TEXT_DARK, name=FONT_MAJOR)


def style_section_title_run(run):
    set_run_font(run, size=SECTION_TITLE_SIZE, bold=True, color=TEXT_DARK, name=FONT_MAJOR)


def style_body_run(run, *, bold=False, color=TEXT_DARK):
    set_run_font(run, size=BODY_SIZE, bold=bold, color=color, name=FONT_BODY)


def style_agenda_run(run):
    set_run_font(run, size=AGENDA_ITEM_SIZE, color=TEXT_DARK, name=FONT_BODY)


def style_footer_run(run):
    set_run_font(run, size=FOOTER_SIZE, color=TEXT_MUTED, name=FONT_BODY)


def delete_all_slides(prs) -> None:
    for index in range(len(prs.slides) - 1, -1, -1):
        slide_id = prs.slides._sldIdLst[index]
        r_id = slide_id.rId
        prs.part.drop_rel(r_id)
        del prs.slides._sldIdLst[index]


def find_placeholder(slide, idx: int | None = None, name_contains: str | None = None):
    for shape in slide.placeholders:
        if idx is not None and shape.placeholder_format.idx == idx:
            return shape
        if name_contains and name_contains.lower() in shape.name.lower():
            return shape
    return None


def set_slide_footer(slide, report_date: str, slide_number: int | None = None):
    date_ph = find_placeholder(slide, idx=13) or find_placeholder(slide, name_contains="Date")
    if date_ph is not None:
        date_ph.text = report_date
        for paragraph in date_ph.text_frame.paragraphs:
            for run in paragraph.runs:
                style_footer_run(run)

    if slide_number is not None:
        num_ph = find_placeholder(slide, idx=10) or find_placeholder(slide, name_contains="Slide Number")
        if num_ph is not None:
            num_ph.text = str(slide_number)
            for paragraph in num_ph.text_frame.paragraphs:
                for run in paragraph.runs:
                    style_footer_run(run)


def set_slide_title(slide, title: str):
    title_ph = find_placeholder(slide, idx=0) or find_placeholder(slide, name_contains="Title")
    if title_ph is None:
        return
    title_ph.text = title
    for paragraph in title_ph.text_frame.paragraphs:
        for run in paragraph.runs:
            style_section_title_run(run)


def apply_table_style(table) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    for child in list(tbl_pr):
        if child.tag == qn("a:tableStyleId"):
            tbl_pr.remove(child)
    style = OxmlElement("a:tableStyleId")
    style.text = TABLE_STYLE_ID
    tbl_pr.append(style)
    if tbl_pr.get("firstRow") is None:
        tbl_pr.set("firstRow", "1")
    if tbl_pr.get("bandRow") is None:
        tbl_pr.set("bandRow", "1")


def style_table_cells(table, *, header_rows: int = 1):
    apply_table_style(table)
    for row_idx in range(len(table.rows)):
        for col_idx in range(len(table.columns)):
            cell = table.cell(row_idx, col_idx)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.04)
            cell.margin_right = Inches(0.04)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)

            is_header = row_idx < header_rows
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT
                if not paragraph.runs:
                    paragraph.text = paragraph.text or ""
                for run in paragraph.runs:
                    set_run_font(
                        run,
                        size=TABLE_HEADER_SIZE if is_header else TABLE_BODY_SIZE,
                        bold=is_header,
                        color=TEXT_DARK,
                        name=FONT_BODY,
                    )


def add_number_badge(slide, number: str, left: float, top: float, accent: bool = False):
    size = Inches(0.42)
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), size, size)
    shape.fill.solid()
    shape.fill.fore_color.rgb = ACCENT_LIME if accent else TEXT_DARK
    shape.line.fill.background()
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = number
    set_run_font(run, size=Pt(14), bold=True, color=WHITE if not accent else TEXT_DARK, name=FONT_MAJOR)
    return shape
