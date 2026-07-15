"""Excel workbook loaders."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from wsr.constants import (
    DDP_SHEET,
    DEFAULT_DATA_FILE,
    PLANNING_SHEET,
    TRACKER_SHEET,
    VISIBILITY_SHEET,
)
from wsr.errors import WsrDataError


def _read_sheet(
    data_file: str | Path,
    sheet_name: str,
    *,
    header: int = 0,
) -> pd.DataFrame:
    path = Path(data_file)
    try:
        return pd.read_excel(path, sheet_name=sheet_name, header=header)
    except ValueError as exc:
        raise WsrDataError(
            f'Sheet "{sheet_name}" not found or unreadable in {path.name}:\n{exc}'
        ) from exc
    except Exception as exc:
        raise WsrDataError(
            f'Could not read sheet "{sheet_name}" in {path.name}:\n{exc}'
        ) from exc


def load_tracker(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return _read_sheet(data_file, TRACKER_SHEET)


def load_visibility(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return _read_sheet(data_file, VISIBILITY_SHEET)


def load_ddp_plan(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return _read_sheet(data_file, DDP_SHEET)


def load_non_stla_planning(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return _read_sheet(data_file, PLANNING_SHEET, header=1)
