"""WSR report orchestration: load data, build slides, save deck."""

from wsr.report.generate import generate_report
from wsr.report.models import ReportResult

__all__ = ["ReportResult", "generate_report"]
