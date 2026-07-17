"""Excel workbook sheet names and default file paths."""

from __future__ import annotations

from pathlib import Path

DEFAULT_DATA_FILE = "SCRUM_PFS_Jul'26.xlsm"
GRAPH_SHEET = "CSAR_WSR_Graph (Non-STLA)"
TRACKER_SHEET = "Non STLA"
VISIBILITY_SHEET = "Visibility Sheet."
DDP_SHEET = "DDP_Plan"
PLANNING_SHEET = "Non_STLA (Planning)"

DEFAULT_PLANNING_BOOK = Path(__file__).resolve().parent.parent.parent / "Book2.xlsx"
