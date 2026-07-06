"""
Backward-compatible re-exports for WSR data helpers.

New code should import from the ``wsr`` package directly, e.g.::

    from wsr.loaders import load_tracker
    from wsr.pending import pending_items
"""

from wsr.constants import (
    DDP_SHEET,
    DEFAULT_DATA_FILE,
    GRAPH_SHEET,
    PLANNING_SHEET,
    TRACKER_SHEET,
    VISIBILITY_SHEET,
)
from wsr.graph import (
    add_week_labels,
    get_evaluation_data,
    get_implementation_data,
    graph_status_notes,
    load_graph_sheet,
    load_graph_summary,
    to_percentage,
    week_remarks,
)
from wsr.loaders import load_ddp_plan, load_non_stla_planning, load_tracker, load_visibility
from wsr.pending import (
    build_pending_item,
    graph_week_capacity,
    pending_items,
    pending_week_for_chart,
)
from wsr.report_data import (
    core2_planned_count,
    ddp_ms45_items,
    discussion_points,
    eval_handoff_items,
    planning_dcr_column,
    planning_dcr_ids,
    planning_type_counts,
    summary_callouts,
    summary_table_rows,
    visibility_status_counts,
)
from wsr.tracker import (
    closure_date_from_row,
    closure_sort_key,
    comment_activity_score,
    eval_status_from_row,
    format_date,
    impl_status_from_row,
    latest_comment,
    parse_dcr_id,
    support_required_from_row,
    tracker_lookup,
    tracker_row_for_mode,
    tracker_rows_lookup,
    visibility_row,
)

# Legacy private-name aliases.
_tracker_row_for_mode = tracker_row_for_mode
_visibility_row = visibility_row
_build_pending_item = build_pending_item
_comment_activity_score = comment_activity_score
_planning_dcr_column = planning_dcr_column
_planning_dcr_ids = planning_dcr_ids
_core2_planned_count = core2_planned_count

__all__ = [
    "DEFAULT_DATA_FILE",
    "GRAPH_SHEET",
    "TRACKER_SHEET",
    "VISIBILITY_SHEET",
    "DDP_SHEET",
    "add_week_labels",
    "closure_date_from_row",
    "closure_sort_key",
    "ddp_ms45_items",
    "discussion_points",
    "eval_handoff_items",
    "eval_status_from_row",
    "format_date",
    "get_evaluation_data",
    "get_implementation_data",
    "graph_status_notes",
    "graph_week_capacity",
    "impl_status_from_row",
    "latest_comment",
    "load_ddp_plan",
    "load_graph_sheet",
    "load_graph_summary",
    "load_non_stla_planning",
    "load_tracker",
    "load_visibility",
    "parse_dcr_id",
    "pending_items",
    "pending_week_for_chart",
    "summary_callouts",
    "summary_table_rows",
    "support_required_from_row",
    "to_percentage",
    "tracker_lookup",
    "tracker_rows_lookup",
    "week_remarks",
]
