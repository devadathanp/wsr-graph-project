"""Append-only run log for GUI and CLI report generation."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import traceback


class RunLog:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.warnings: list[str] = []
        self._fh = self.path.open("w", encoding="utf-8")
        self.info(f"WSR run log started at {datetime.now():%Y-%m-%d %H:%M:%S}")

    def _write(self, level: str, message: str) -> None:
        line = f"{datetime.now():%H:%M:%S} [{level}] {message}"
        self._fh.write(line + "\n")
        self._fh.flush()

    def info(self, message: str) -> None:
        self._write("INFO", message)

    def warning(self, message: str) -> None:
        self.warnings.append(message)
        self._write("WARN", message)

    def error(self, message: str) -> None:
        self._write("ERROR", message)

    def exception(self, exc: BaseException) -> None:
        self.error(f"{type(exc).__name__}: {exc}")
        self._fh.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
        self._fh.flush()

    def close(self) -> None:
        if self._fh.closed:
            return
        self.info(f"WSR run log finished at {datetime.now():%Y-%m-%d %H:%M:%S}")
        self._fh.close()


def default_log_path(output_path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path(output_path).with_name(f"{Path(output_path).stem}_{stamp}.log")
