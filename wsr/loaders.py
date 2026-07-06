"""Excel workbook loaders."""

from __future__ import annotations

import pandas as pd

from wsr.constants import (
    DDP_SHEET,
    DEFAULT_DATA_FILE,
    PLANNING_SHEET,
    TRACKER_SHEET,
    VISIBILITY_SHEET,
)


def load_tracker(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name=TRACKER_SHEET)


def load_visibility(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name=VISIBILITY_SHEET)


def load_ddp_plan(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name=DDP_SHEET)


def load_non_stla_planning(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name=PLANNING_SHEET, header=1)
