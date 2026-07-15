"""Pre-flight checks for Scrum / Planning Excel inputs."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from wsr.constants import (
    DDP_SHEET,
    DEFAULT_PLANNING_BOOK,
    GRAPH_SHEET,
    TRACKER_SHEET,
    VISIBILITY_SHEET,
)
from wsr.errors import WsrDataError
from wsr.graph import (
    COL_COMPLETED,
    COL_CUMULATIVE_BASELINE,
    COL_CUMULATIVE_REVISED,
    COL_DATE,
    COL_DRB,
    COL_IN_PROGRESS,
    COL_PCT_ACTUAL,
    COL_PCT_CONFIDENCE,
    COL_REJECTED,
    COL_WEEK,
    get_evaluation_data,
    get_implementation_data,
    load_graph_sheet,
)
from wsr.run_log import RunLog

REQUIRED_SCRUM_SHEETS = (
    GRAPH_SHEET,
    TRACKER_SHEET,
    VISIBILITY_SHEET,
    DDP_SHEET,
)

TRACKER_REQUIRED_COLUMNS = (
    "DCR ID - PTC",
    "PRCRState",
    "At Risk",
    "Planned Completion Date\n<dd-mm-yyyy>",
    "Summary",
)

GRAPH_REQUIRED_COLUMNS = (
    "Tagged to Release",
    COL_WEEK,
    COL_DATE,
    COL_CUMULATIVE_BASELINE,
    COL_CUMULATIVE_REVISED,
    COL_COMPLETED,
    COL_REJECTED,
    COL_IN_PROGRESS,
    COL_DRB,
    COL_PCT_CONFIDENCE,
    COL_PCT_ACTUAL,
)


def _sheet_names(path: Path) -> list[str]:
    try:
        return list(pd.ExcelFile(path).sheet_names)
    except Exception as exc:
        raise WsrDataError(
            f"Could not open workbook (file locked or corrupt): {path.name}\n{exc}"
        ) from exc


def require_file(path: Path, *, label: str) -> None:
    if not path.exists():
        raise WsrDataError(f"{label} not found:\n{path}")
    if not path.is_file():
        raise WsrDataError(f"{label} is not a file:\n{path}")


def validate_scrum_workbook(data_file: str | Path, log: RunLog | None = None) -> Path:
    path = Path(data_file)
    require_file(path, label="Scrum workbook")
    if log:
        log.info(f"Scrum workbook: {path}")

    names = _sheet_names(path)
    if log:
        log.info(f"Sheets found: {', '.join(names)}")

    missing = [name for name in REQUIRED_SCRUM_SHEETS if name not in names]
    if missing:
        raise WsrDataError(
            "Required sheet(s) missing from the Scrum workbook:\n"
            + "\n".join(f"  - {name}" for name in missing)
            + f"\n\nWorkbook: {path.name}"
        )

    try:
        tracker = pd.read_excel(path, sheet_name=TRACKER_SHEET)
    except Exception as exc:
        raise WsrDataError(
            f'Could not read sheet "{TRACKER_SHEET}" in {path.name}:\n{exc}'
        ) from exc

    missing_cols = [col for col in TRACKER_REQUIRED_COLUMNS if col not in tracker.columns]
    if missing_cols:
        raise WsrDataError(
            f'Required column(s) missing on sheet "{TRACKER_SHEET}":\n'
            + "\n".join(f"  - {col!r}" for col in missing_cols)
            + f"\n\nWorkbook: {path.name}"
        )
    if tracker.empty:
        raise WsrDataError(f'Sheet "{TRACKER_SHEET}" has no data rows in {path.name}.')

    try:
        graph = load_graph_sheet(str(path))
    except Exception as exc:
        raise WsrDataError(
            f'Could not read sheet "{GRAPH_SHEET}" in {path.name}:\n{exc}'
        ) from exc

    missing_graph = [col for col in GRAPH_REQUIRED_COLUMNS if col not in graph.columns]
    if missing_graph:
        raise WsrDataError(
            f'Required column(s) missing on sheet "{GRAPH_SHEET}":\n'
            + "\n".join(f"  - {col!r}" for col in missing_graph)
            + f"\n\nWorkbook: {path.name}"
        )

    try:
        eval_section = get_evaluation_data(graph)
        impl_section = get_implementation_data(graph)
    except Exception as exc:
        raise WsrDataError(
            f'Graph sheet "{GRAPH_SHEET}" looks corrupt or incomplete in {path.name}:\n{exc}'
        ) from exc

    if eval_section.empty:
        raise WsrDataError(
            f'No evaluation week rows found on "{GRAPH_SHEET}" in {path.name}.'
        )
    if impl_section.empty:
        raise WsrDataError(
            f'No implementation week rows found on "{GRAPH_SHEET}" in {path.name}.'
        )

    if log:
        log.info(
            f"Graph OK — eval weeks: {len(eval_section)}, impl weeks: {len(impl_section)}; "
            f"tracker rows: {len(tracker)}"
        )
    return path


def validate_planning_book(
    planning_book: str | Path | None,
    log: RunLog | None = None,
) -> Path | None:
    path = Path(planning_book) if planning_book else DEFAULT_PLANNING_BOOK
    if not path.exists():
        if log:
            label = "selected planning workbook" if planning_book else "default Book2.xlsx"
            log.warning(f"No {label} found ({path}); slide 11 will show a placeholder.")
        return None

    if log:
        log.info(f"Planning workbook: {path}")
    return path
