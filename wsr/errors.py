"""Domain errors for WSR generation."""

from __future__ import annotations

from pathlib import Path


class WsrError(Exception):
    """Base error for the WSR generator."""


class WsrDataError(WsrError):
    """Workbook / sheet / column data is missing or unusable."""

    def __init__(self, message: str, *, log_path: Path | None = None):
        super().__init__(message)
        self.log_path = log_path
