"""Slide 4 summary table and legacy callout text."""

from __future__ import annotations

import pandas as pd

from wsr.constants import DEFAULT_DATA_FILE
from wsr.graph import load_graph_summary
from wsr.loaders import load_non_stla_planning, load_tracker
from wsr.report_data.planning import (
    planning_dcr_ids,
    planning_type_counts,
)
from wsr.tracker import parse_dcr_id

COL_PRCR_STATE = "PRCRState"
COL_AT_RISK = "At Risk"
COL_DCR = "DCR ID - PTC"
COL_DDP_NEEDED = "Is DDP Testing Needed"
COL_ECM_NEEDED = "Is ECM Testing Needed"
COL_DCR_PHASE = "DCR Phase"

# At Risk values excluded from CSAR / Core2 program counts.
_EXCLUDED_AT_RISK = frozenset({"cancelled", "canceled", "rejected"})


def _unique_dcr_count(
    tracker: pd.DataFrame,
    *,
    row_matches,
) -> int:
    """Count matching rows once per DCR ID (rows without a DCR ID count individually)."""
    seen: set[int] = set()
    orphan = 0
    for _, row in tracker.iterrows():
        if not row_matches(row):
            continue
        dcr_id = parse_dcr_id(row.get(COL_DCR)) if COL_DCR in tracker.columns else None
        if dcr_id is None:
            orphan += 1
        else:
            seen.add(dcr_id)
    return len(seen) + orphan


def _at_risk_excluded(row) -> bool:
    risk = str(row.get(COL_AT_RISK, "")).strip().lower()
    return risk in _EXCLUDED_AT_RISK


def phase_program_count(tracker: pd.DataFrame | None, phase_token: str) -> int:
    """
    Unique DCR count whose ``DCR Phase`` contains ``phase_token`` (case-insensitive).

    Skips DCRs whose At Risk is Cancelled or Rejected.
    """
    if tracker is None or tracker.empty or COL_DCR_PHASE not in tracker.columns:
        return 0
    token = phase_token.strip().lower()

    def _matches(row) -> bool:
        phase = str(row.get(COL_DCR_PHASE, "") or "").strip().lower()
        if token not in phase:
            return False
        return not _at_risk_excluded(row)

    return _unique_dcr_count(tracker, row_matches=_matches)


def csar_count(tracker: pd.DataFrame | None) -> int:
    """CES CSAR phases (all versions), excluding Cancelled/Rejected."""
    return phase_program_count(tracker, "CES CSAR")


def core2_count(tracker: pd.DataFrame | None) -> int:
    """CES ATS phases (all versions) for Core2, excluding Cancelled/Rejected."""
    return phase_program_count(tracker, "CES ATS")


def at_risk_eval_impl_counts(
    tracker: pd.DataFrame | None,
    at_risk_value: str,
) -> dict[str, int]:
    """
    Count unique DCRs for Eval / Impl with a given At Risk value.

    Eval  → PRCRState == Evaluate and At Risk == ``at_risk_value``
    Impl  → PRCRState == Implement and At Risk == ``at_risk_value``
    Duplicate DCR IDs are counted once.
    """
    empty = {"eval": 0, "impl": 0}
    if tracker is None or tracker.empty:
        return empty
    if COL_PRCR_STATE not in tracker.columns or COL_AT_RISK not in tracker.columns:
        return empty

    target = at_risk_value.strip().lower()

    def _count(prcr_state: str) -> int:
        return _unique_dcr_count(
            tracker,
            row_matches=lambda row: (
                str(row.get(COL_PRCR_STATE, "")).strip() == prcr_state
                and str(row.get(COL_AT_RISK, "")).strip().lower() == target
            ),
        )

    return {"eval": _count("Evaluate"), "impl": _count("Implement")}


def ddp_testing_count(tracker: pd.DataFrame | None) -> int:
    """Unique DCR count where ``Is DDP Testing Needed`` is Yes."""
    if tracker is None or tracker.empty or COL_DDP_NEEDED not in tracker.columns:
        return 0

    def _is_yes(row) -> bool:
        return str(row.get(COL_DDP_NEEDED, "")).strip().lower() in {"yes", "y"}

    return _unique_dcr_count(tracker, row_matches=_is_yes)


def ecm_testing_count(tracker: pd.DataFrame | None) -> int:
    """Unique DCR count where ``Is ECM Testing Needed`` is Yes."""
    if tracker is None or tracker.empty or COL_ECM_NEEDED not in tracker.columns:
        return 0

    def _is_yes(row) -> bool:
        return str(row.get(COL_ECM_NEEDED, "")).strip().lower() in {"yes", "y"}

    return _unique_dcr_count(tracker, row_matches=_is_yes)


def _format_eval_impl(counts: dict[str, int]) -> str:
    return f"Eval: {counts['eval']}, Impl: {counts['impl']}"


def summary_callouts(data_file: str = DEFAULT_DATA_FILE) -> dict[str, str]:
    planning = load_non_stla_planning(data_file)
    graph = load_graph_summary(data_file)
    tracker = load_tracker(data_file)
    type_counts = planning_type_counts(planning)

    planned_ids = planning_dcr_ids(planning)
    total = len(set(planned_ids))
    eval_impl = type_counts.get("Eval+Impl", 0)
    impl_only = type_counts.get("Impl", 0)

    eval_baseline = graph.get("eval_baseline")
    eval_revised = graph.get("eval_revised")
    impl_baseline = graph.get("impl_baseline")
    impl_revised = graph.get("impl_revised")

    rejected = at_risk_eval_impl_counts(tracker, "Rejected")
    deferred = at_risk_eval_impl_counts(tracker, "Deferred")

    return {
        "total_planned": f"{total} (Non STLA + Core 2) + ECM Testing",
        "csar": f"CSAR {csar_count(tracker)}",
        "core2": f"Core2 {core2_count(tracker)}",
        "ecm_testing": f"ECM Testing {ecm_testing_count(tracker)}",
        "ddp_testing": f"DDP Testing {ddp_testing_count(tracker)}",
        "eval_planned": (
            f"DCR's Planned for Evaluation {eval_baseline} >> {eval_revised}"
            if eval_baseline is not None and eval_revised is not None
            else f"DCR's Planned for Evaluation {eval_impl + type_counts.get('Eval', 0)}"
        ),
        "impl_planned": (
            f"DCR's Planned for Implementation {impl_baseline} >> {impl_revised}"
            if impl_baseline is not None and impl_revised is not None
            else f"DCR's Planned for Implementation {impl_only + eval_impl}"
        ),
        "rejected": f"DCR's Rejected — {_format_eval_impl(rejected)}",
        "deferred": f"DCR's Deferred — {_format_eval_impl(deferred)}",
    }


def summary_table_rows(
    data_file: str = DEFAULT_DATA_FILE,
    tracker: pd.DataFrame | None = None,
) -> list[tuple[str, str]]:
    graph = load_graph_summary(data_file)

    eval_baseline = graph.get("eval_baseline")
    impl_baseline = graph.get("impl_baseline")

    if eval_baseline is not None and impl_baseline is not None:
        total_value = str(eval_baseline + impl_baseline)
        eval_planned = str(eval_baseline)
        impl_planned = str(impl_baseline)
    else:
        planning = load_non_stla_planning(data_file)
        type_counts = planning_type_counts(planning)
        eval_planned = str(type_counts.get("Eval+Impl", 0) + type_counts.get("Eval", 0))
        impl_planned = str(type_counts.get("Impl", 0))
        total_value = str(int(eval_planned) + int(impl_planned))

    if tracker is None:
        tracker = load_tracker(data_file)
    rejected = at_risk_eval_impl_counts(tracker, "Rejected")
    deferred = at_risk_eval_impl_counts(tracker, "Deferred")
    ddp_count = ddp_testing_count(tracker)
    ecm_count = ecm_testing_count(tracker)

    return [
        ("Total DCR's planned", total_value),
        ("CSAR", str(csar_count(tracker))),
        ("Core2", str(core2_count(tracker))),
        ("ECM Testing", str(ecm_count)),
        ("DDP Testing", str(ddp_count)),
        ("DCR's Planned for Evaluation", eval_planned),
        ("DCR's Planned for Implementation", impl_planned),
        ("DCR's Rejected", _format_eval_impl(rejected)),
        ("DCR's Deferred", _format_eval_impl(deferred)),
    ]
