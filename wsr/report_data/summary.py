"""Slide 4 summary table and legacy callout text."""

from __future__ import annotations

from wsr.constants import DEFAULT_DATA_FILE
from wsr.graph import load_graph_summary
from wsr.loaders import load_non_stla_planning, load_visibility
from wsr.report_data.planning import (
    core2_planned_count,
    planning_dcr_ids,
    planning_type_counts,
)
from wsr.report_data.visibility import visibility_status_counts


def summary_callouts(data_file: str = DEFAULT_DATA_FILE) -> dict[str, str]:
    planning = load_non_stla_planning(data_file)
    visibility = load_visibility(data_file)
    graph = load_graph_summary(data_file)
    type_counts = planning_type_counts(planning)
    status_counts = visibility_status_counts(visibility)

    planned_ids = planning_dcr_ids(planning)
    total = len(set(planned_ids))
    eval_impl = type_counts.get("Eval+Impl", 0)
    impl_only = type_counts.get("Impl", 0)

    eval_baseline = graph.get("eval_baseline")
    eval_revised = graph.get("eval_revised")
    impl_baseline = graph.get("impl_baseline")
    impl_revised = graph.get("impl_revised")

    return {
        "total_planned": f"{total} (Non STLA + Core 2) + ECM Testing",
        "csar": (
            f"CSAR {impl_baseline} >> {impl_revised}"
            if impl_baseline is not None and impl_revised is not None
            else f"CSAR planned {total}"
        ),
        "core2": (
            f"Core2 {core2_count}"
            if (core2_count := core2_planned_count(planning)) is not None
            else "-"
        ),
        "ecm_testing": f"ECM Testing {type_counts.get('ECM_Testing', 0)}",
        "ddp_testing": f"DDP Testing {type_counts.get('DDP', 0)}",
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
        "rejected": (
            "DCR's Rejected — Eval: {eval:02d}, Impl: {impl:02d}".format(**status_counts["rejected"])
        ),
        "deferred": (
            "DCR's Deferred — Eval: {eval:02d}, Impl: {impl:02d}".format(**status_counts["deferred"])
        ),
    }


def summary_table_rows(data_file: str = DEFAULT_DATA_FILE) -> list[tuple[str, str]]:
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

    return [
        ("Total DCR's planned", total_value),
        ("CSAR", ""),
        ("Core2", ""),
        ("ECM Testing", ""),
        ("DDP Testing", ""),
        ("DCR's Planned for Evaluation", eval_planned),
        ("DCR's Planned for Implementation", impl_planned),
        ("DCR's Rejected", "Eval: , Impl: "),
        ("DCR's Deferred", "Eval: , Impl: "),
    ]
