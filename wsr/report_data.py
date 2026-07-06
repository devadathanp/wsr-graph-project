"""Slide table data: DDP, handoff, discussion, and summary metrics."""

from __future__ import annotations

import pandas as pd

from wsr.constants import DEFAULT_DATA_FILE
from wsr.graph import load_graph_summary
from wsr.loaders import load_non_stla_planning, load_visibility
from wsr.tracker import format_date, latest_comment, parse_dcr_id


def planning_dcr_column(planning: pd.DataFrame) -> str | None:
    for column in ("DCRID-PTC", planning.columns[1] if len(planning.columns) > 1 else None):
        if column and column in planning.columns:
            return column
    return None


def planning_dcr_ids(planning: pd.DataFrame) -> list[int]:
    dcr_col = planning_dcr_column(planning)
    if dcr_col is None:
        return []
    ids = []
    for value in planning[dcr_col]:
        dcr_id = parse_dcr_id(value)
        if dcr_id is not None:
            ids.append(dcr_id)
    return ids


def planning_type_counts(planning: pd.DataFrame) -> dict[str, int]:
    type_col = "Type" if "Type" in planning.columns else "Eval+Impl"
    if type_col not in planning.columns:
        return {}
    counts = planning[type_col].astype(str).str.strip().value_counts()
    return {str(key): int(value) for key, value in counts.items() if str(key) not in ("nan", "")}


def core2_planned_count(planning: pd.DataFrame) -> int | None:
    if "Program" not in planning.columns:
        return None
    program = planning["Program"].astype(str).str.strip()
    count = program.str.contains(r"\bDAF\b|Core\s*2", case=False, na=False).sum()
    return int(count)


def visibility_status_counts(visibility: pd.DataFrame) -> dict[str, dict[str, int]]:
    rejected = visibility[visibility["Status"].astype(str).str.contains("Reject", case=False, na=False)]
    deferred = visibility[visibility["Status"].astype(str).str.contains("Defer", case=False, na=False)]

    def _split_eval_impl(frame: pd.DataFrame) -> dict[str, int]:
        eval_count = 0
        impl_count = 0
        for _, row in frame.iterrows():
            work_type = str(row.get("Evaluation/ Implementation", ""))
            if "Eval" in work_type:
                eval_count += 1
            elif "Impl" in work_type:
                impl_count += 1
        return {"eval": eval_count, "impl": impl_count}

    return {
        "rejected": _split_eval_impl(rejected),
        "deferred": _split_eval_impl(deferred),
    }


def summary_callouts(data_file: str = DEFAULT_DATA_FILE) -> dict[str, str]:
    """Build DCR status summary callout text from planning, graph, and visibility sheets."""
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
    """Two-column DCR status summary for slide 4 (values sourced from workbooks)."""
    planning = load_non_stla_planning(data_file)
    type_counts = planning_type_counts(planning)
    status_counts = visibility_status_counts(load_visibility(data_file))
    graph = load_graph_summary(data_file)

    eval_baseline = graph.get("eval_baseline")
    eval_revised = graph.get("eval_revised")
    impl_baseline = graph.get("impl_baseline")
    impl_revised = graph.get("impl_revised")

    if None not in (eval_baseline, eval_revised, impl_baseline, impl_revised):
        total_value = (
            f"{eval_baseline + impl_baseline} >> {eval_revised + impl_revised} "
            "(Non STLA + Core 2) + ECM Testing"
        )
    else:
        total_value = f"{len(set(planning_dcr_ids(planning)))} (Non STLA + Core 2) + ECM Testing"

    csar_value = (
        f"{impl_baseline} >> {impl_revised}"
        if impl_baseline is not None and impl_revised is not None
        else "-"
    )
    core2_count = core2_planned_count(planning)
    eval_planned = (
        f"{eval_baseline} >> {eval_revised}"
        if eval_baseline is not None and eval_revised is not None
        else str(type_counts.get("Eval+Impl", 0) + type_counts.get("Eval", 0))
    )
    impl_planned = (
        f"{impl_baseline} >> {impl_revised}"
        if impl_baseline is not None and impl_revised is not None
        else str(type_counts.get("Impl", 0) + type_counts.get("Eval+Impl", 0))
    )

    return [
        ("Total DCR's planned", total_value),
        ("CSAR", csar_value),
        ("Core2", str(core2_count) if core2_count is not None else "-"),
        ("ECM Testing", str(type_counts.get("ECM_Testing", 0))),
        ("DDP Testing", str(type_counts.get("DDP", 0))),
        ("DCR's Planned for Evaluation", eval_planned),
        ("DCR's Planned for Implementation", impl_planned),
        (
            "DCR's Rejected",
            f"Eval: {status_counts['rejected']['eval']:02d}, "
            f"Impl: {status_counts['rejected']['impl']:02d}",
        ),
        (
            "DCR's Deferred",
            f"Eval: {status_counts['deferred']['eval']:02d}, "
            f"Impl: {status_counts['deferred']['impl']:02d}",
        ),
    ]


def ddp_row_item(
    ddp_row: pd.Series,
    tracker_lookup_map: dict[int, pd.Series],
    sr_no: int,
) -> dict:
    dcr_no = ddp_row.get("DCR No")
    dcr_text = "-" if pd.isna(dcr_no) else str(dcr_no).replace(".0", "").strip()
    dcr_id = parse_dcr_id(dcr_no)
    tracker_row = tracker_lookup_map.get(dcr_id) if dcr_id is not None else None

    if pd.notna(ddp_row.get("Diagnostics Name")):
        summary = str(ddp_row.get("Diagnostics Name"))
    elif tracker_row is not None:
        summary = str(tracker_row.get("Summary", "-"))
    else:
        summary = "-"

    remarks = ddp_row.get("Current status", ddp_row.get("Status", "-"))
    if pd.isna(remarks) or str(remarks).strip() in ("", "nan"):
        remarks = (
            latest_comment(tracker_row.get("Comments (Daily)"), max_len=200)
            if tracker_row is not None
            else "-"
        )

    dependencies = "-"
    if tracker_row is not None:
        for field in (
            "Support Required from team",
            "Reasons for delay",
            "Mitigation Plan",
        ):
            value = tracker_row.get(field)
            if pd.notna(value) and str(value).strip() not in ("", "nan", "0"):
                dependencies = str(value).strip().replace("\n", " ")
                break

    return {
        "sr_no": sr_no,
        "dcr_id": dcr_text,
        "summary": summary,
        "plan_date": format_date(ddp_row.get("Revised planned dates", ddp_row.get("Appeared Plan date"))),
        "appeared_date": format_date(ddp_row.get("Appeared Plan date")),
        "program": str(ddp_row.get("Bench Type", "-")) if pd.notna(ddp_row.get("Bench Type")) else "-",
        "dependencies": dependencies,
        "remarks": str(remarks),
    }


def ddp_ms45_items(
    ddp: pd.DataFrame,
    tracker_lookup_map: dict[int, pd.Series],
    limit: int = 7,
) -> list[dict]:
    rows = ddp[ddp["DCR No"].notna()].copy()
    rows = rows[~rows["DCR No"].astype(str).str.strip().str.upper().eq("TBD")]
    rows["_ms45"] = rows["Status"].astype(str).str.contains(r"MS\s*4|4-5|4_5", case=False, na=False)
    rows["_has_diag"] = rows["Diagnostics Name"].notna()
    rows = rows.sort_values(by=["_ms45", "_has_diag"], ascending=[False, False])

    items = []
    for _, row in rows.iterrows():
        items.append(ddp_row_item(row, tracker_lookup_map, len(items) + 1))
        if len(items) >= limit:
            return items

    rows = ddp[ddp["Diagnostics Name"].notna()].copy()
    for _, row in rows.iterrows():
        items.append(ddp_row_item(row, tracker_lookup_map, len(items) + 1))
        if len(items) >= limit:
            break
    return items


def _is_high_priority(priority) -> bool:
    text = str(priority).strip().lower()
    return text == "high"


def _report_month_year(report_date: str) -> tuple[int, int]:
    parsed = pd.to_datetime(report_date, dayfirst=True)
    return parsed.month, parsed.year


def _first_evaluator_name(value) -> str:
    if pd.isna(value) or str(value).strip() in ("", "nan"):
        return "-"
    text = str(value).strip()
    return text.split("/")[0].split(",")[0].strip()


def _planning_handoff_date(row: pd.Series) -> tuple[pd.Timestamp | None, str]:
    for column in (
        "L2/DRB Eval Send dates\nDD-MM-YY",
        "Eval completion date, (if it is evaluation)\nDD-MM-YY",
    ):
        if column not in row.index:
            continue
        parsed = pd.to_datetime(row.get(column), errors="coerce")
        if pd.notna(parsed):
            return parsed, format_date(parsed)
    return None, "-"


def eval_handoff_items(
    planning: pd.DataFrame,
    tracker_lookup_map: dict[int, pd.Series],
    report_date: str,
    limit: int = 10,
) -> list[dict]:
    """Eval handoffs whose planning-sheet eval send/completion date falls in the report month."""
    month, year = _report_month_year(report_date)
    dcr_col = planning_dcr_column(planning)
    if dcr_col is None:
        return []

    items = []
    for _, row in planning.iterrows():
        dcr_id = parse_dcr_id(row.get(dcr_col))
        if dcr_id is None:
            continue

        handoff_ts, handoff_date = _planning_handoff_date(row)
        if handoff_ts is None or handoff_ts.month != month or handoff_ts.year != year:
            continue

        tracker_row = tracker_lookup_map.get(dcr_id)
        summary = str(row.get("Summary", "-")) if pd.notna(row.get("Summary")) else "-"
        if tracker_row is not None and pd.notna(tracker_row.get("Summary")):
            summary = str(tracker_row.get("Summary"))

        evaluator = _first_evaluator_name(row.get("KPIT User"))
        if evaluator == "-":
            evaluator = _first_evaluator_name(row.get("Estimator"))

        remark = "-"
        if pd.notna(row.get("Remarks.1")) and str(row.get("Remarks.1")).strip() not in ("", "nan"):
            remark = str(row.get("Remarks.1")).strip()
        elif tracker_row is not None:
            remark = latest_comment(tracker_row.get("Comments (Daily)"), max_len=None)

        items.append(
            {
                "dcr_id": dcr_id,
                "evaluator": evaluator,
                "handoff_date": handoff_date,
                "remark": remark,
                "summary": summary,
            }
        )
        if len(items) >= limit:
            break
    return items


def discussion_points(
    visibility: pd.DataFrame,
    tracker_lookup_map: dict[int, pd.Series],
    limit: int = 5,
) -> list[dict]:
    items = []
    at_risk = visibility[visibility["Status"].astype(str).str.contains("At Risk|On Hold", case=False, na=False)]
    for _, vis_row in at_risk.iterrows():
        dcr_raw = vis_row.get("DCR Number")
        if pd.isna(dcr_raw):
            continue
        try:
            dcr_id = int(dcr_raw)
        except (TypeError, ValueError):
            continue
        if any(item["dcr_id"] == dcr_id for item in items):
            continue

        tracker_row = tracker_lookup_map.get(dcr_id, vis_row)
        priority = str(tracker_row.get("CQ Priority", tracker_row.get("Internal Priority", "-")))
        if not _is_high_priority(priority):
            continue
        items.append(
            {
                "dcr_id": dcr_id,
                "description": str(tracker_row.get("Summary", vis_row.get("Subject", "-"))),
                "program": str(tracker_row.get("DCR Project", "CSAR"))
                if pd.notna(tracker_row.get("DCR Project"))
                else "CSAR",
                "priority": priority,
                "plan_date": format_date(tracker_row.get("Planned Completion Date\n<dd-mm-yyyy>")),
                "remarks": latest_comment(tracker_row.get("Comments (Daily)"), max_len=None),
            }
        )
        if len(items) >= limit:
            break
    return items
