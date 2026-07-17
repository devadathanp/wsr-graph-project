"""Slide table data: DDP, handoff, discussion, and summary metrics."""

from wsr.report_data.ddp import ddp_ms45_items
from wsr.report_data.discussion import discussion_points
from wsr.report_data.handoff import eval_handoff_items
from wsr.report_data.planning import (
    core2_planned_count,
    planning_dcr_column,
    planning_dcr_ids,
    planning_type_counts,
)
from wsr.report_data.summary import summary_callouts, summary_table_rows
from wsr.report_data.visibility import visibility_status_counts

__all__ = [
    "core2_planned_count",
    "ddp_ms45_items",
    "discussion_points",
    "eval_handoff_items",
    "planning_dcr_column",
    "planning_dcr_ids",
    "planning_type_counts",
    "summary_callouts",
    "summary_table_rows",
    "visibility_status_counts",
]
