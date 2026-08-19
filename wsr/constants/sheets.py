"""Excel workbook sheet names and default file paths."""

from __future__ import annotations

from pathlib import Path

DEFAULT_DATA_FILE = "SCRUM_PFS_Aug'26-Non STLA.xlsm"
GRAPH_SHEET = "CSAR_WSR_Graph (Non-STLA)"
TRACKER_SHEET = "Non STLA"
VISIBILITY_SHEET = "Visibility Sheet."
DDP_SHEET = "DDP_Plan"
PLANNING_SHEET = "Non_STLA (Planning)"
RISK_SHEET = "Risk and Mitigation Plan"
ACTION_ITEM_SHEET = "ActionItem"

DEFAULT_PLANNING_BOOK = Path(__file__).resolve().parent.parent.parent / "Book2.xlsx"
