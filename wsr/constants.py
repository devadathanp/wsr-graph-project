"""Shared constants for WSR data sources and slide layout."""

from __future__ import annotations

from pathlib import Path

DEFAULT_DATA_FILE = "data.xlsm"
GRAPH_SHEET = "CSAR_WSR_Graph (Non-STLA)"
TRACKER_SHEET = "Non STLA"
VISIBILITY_SHEET = "Visibility Sheet."
DDP_SHEET = "DDP_Plan"
PLANNING_SHEET = "Non_STLA (Planning)"

DEFAULT_PLANNING_BOOK = Path(__file__).resolve().parent.parent / "Book2.xlsx"

# PowerPoint slide layout indices (CES CSAR template).
LAYOUT_OPENING = 13
LAYOUT_CONTENT = 3

# Slide 4 — DCR status charts and right-hand summary panel.
DCR_CHART_LEFT = 0.08
DCR_CHART_WIDTH = 9.72
DCR_EVAL_TOP = 0.95
DCR_CHART_HEIGHT = 3.2
DCR_IMPL_TOP = 4.22
DCR_PANEL_LEFT = 9.85
DCR_PANEL_WIDTH = 3.2
DCR_SUMMARY_TOP = 0.95
DCR_NOTES_GAP = 0.12
DCR_STATUS_NOTE_LINES = [
    "The initial plan is based on high level estimations and DCRs in KPIT Pune team's bucket as on date.",
    "Revised baseline plan is updated based on reshuffling done during execution in the quarter "
    "based on DCR Rejections / moved to next quarter/ dependencies.",
]

# Slide 2 — Agenda.
AGENDA_ITEMS = [
    "MOM & Action Items",
    "DCR Status",
    "Discussion Points",
    "Issues and Risks",
]
AGENDA_BADGE_SIZE = 1.04
AGENDA_LAYOUT = [
    {"badge_top": 1.22, "text_left": 1.72, "text_top": 1.55},
    {"badge_top": 2.26, "text_left": 1.66, "text_top": 2.65},
    {"badge_top": 3.29, "text_left": 1.66, "text_top": 3.57},
    {"badge_top": 4.27, "text_left": 1.72, "text_top": 4.59},
]

PENDING_TABLE_ROW_CAP = 12
