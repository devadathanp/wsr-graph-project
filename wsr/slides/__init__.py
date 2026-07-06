"""Slide builders for the WSR PowerPoint report."""

from wsr.slides.closing import add_closing_slide
from wsr.slides.dcr_status import add_dcr_status_slide
from wsr.slides.opening import add_agenda_slide, add_mom_slide, add_title_slide
from wsr.slides.planning import add_planning_slide
from wsr.slides.risks import add_risks_slide
from wsr.slides.tables import (
    add_ddp_slide,
    add_discussion_slide,
    add_handoff_slide,
    add_pending_slide,
)

__all__ = [
    "add_agenda_slide",
    "add_closing_slide",
    "add_dcr_status_slide",
    "add_ddp_slide",
    "add_discussion_slide",
    "add_handoff_slide",
    "add_mom_slide",
    "add_pending_slide",
    "add_planning_slide",
    "add_risks_slide",
    "add_title_slide",
]
