"""Shared data loading and helpers for WSR graph/report generation."""

from __future__ import annotations

import re

import numpy as np
import pandas as pd

DEFAULT_DATA_FILE = "data.xlsm"
GRAPH_SHEET = "CSAR_WSR_Graph (Non-STLA)"
TRACKER_SHEET = "Non STLA"
VISIBILITY_SHEET = "Visibility Sheet."
DDP_SHEET = "DDP_Plan"


def to_percentage(series: pd.Series) -> pd.Series:
    def convert(value):
        if pd.isna(value):
            return np.nan
        if isinstance(value, (int, float)):
            val = float(value)
            return val * 100 if abs(val) <= 1.5 else val
        text = str(value).replace("%", "").strip()
        if text == "":
            return np.nan
        val = float(text)
        return val * 100 if abs(val) <= 1.5 else val

    return series.apply(convert)


def load_graph_sheet(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name=GRAPH_SHEET, header=2)


def load_graph_summary(data_file: str = DEFAULT_DATA_FILE) -> dict:
    raw = pd.read_excel(data_file, sheet_name=GRAPH_SHEET, header=None)
    return {
        "eval_baseline": _safe_int(raw.iloc[16, 4]),
        "eval_revised": _safe_int(raw.iloc[16, 5]),
        "eval_completed": _safe_int(raw.iloc[16, 6]),
        "impl_baseline": _safe_int(raw.iloc[34, 4]),
        "impl_revised": _safe_int(raw.iloc[34, 5]),
        "impl_completed": _safe_int(raw.iloc[34, 6]),
    }


def _safe_int(value) -> int | None:
    if pd.isna(value):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def get_evaluation_data(df: pd.DataFrame | None = None, data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    if df is None:
        df = load_graph_sheet(data_file)

    impl_header_idx = df.index[df["Tagged to Release"].astype(str).str.strip() == "Implementation"]
    eval_df = df.loc[: impl_header_idx[0] - 1]
    section = eval_df[eval_df["Week No"].notna()].copy()
    section.reset_index(drop=True, inplace=True)
    return section


def get_implementation_data(df: pd.DataFrame | None = None, data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    if df is None:
        df = load_graph_sheet(data_file)

    impl_start_idx = df.index[
        df["Tagged to Release"].astype(str).str.contains("Q3", na=False)
        & df["Tagged to Release"].astype(str).str.contains("Implementation", na=False)
    ]
    section = df.loc[impl_start_idx[0] :]
    section = section[section["Week No"].notna()].copy()
    section.reset_index(drop=True, inplace=True)
    return section


def add_week_labels(section: pd.DataFrame) -> pd.DataFrame:
    section = section.copy()
    section["Week Label"] = (
        section["Week No"].astype(str) + "\n" + pd.to_datetime(section["Date"]).dt.strftime("%d-%m")
    )
    return section


def load_tracker(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name=TRACKER_SHEET)


def load_visibility(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name=VISIBILITY_SHEET)


def load_ddp_plan(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name=DDP_SHEET)


def parse_dcr_id(value) -> int | None:
    if pd.isna(value):
        return None
    try:
        return int(float(str(value).strip().split("\n")[0].strip()))
    except (TypeError, ValueError):
        return None


def tracker_lookup(tracker: pd.DataFrame) -> dict[int, pd.Series]:
    """Most recently seen Non STLA row per DCR (backward compatible)."""
    return {dcr_id: rows[-1] for dcr_id, rows in tracker_rows_lookup(tracker).items()}


def tracker_rows_lookup(tracker: pd.DataFrame) -> dict[int, list[pd.Series]]:
    lookup: dict[int, list[pd.Series]] = {}
    id_col = "DCR ID - PTC"
    for _, row in tracker.iterrows():
        dcr_id = parse_dcr_id(row.get(id_col))
        if dcr_id is None:
            continue
        lookup.setdefault(dcr_id, []).append(row)
    return lookup


def _tracker_row_for_mode(
    tracker_rows: dict[int, list[pd.Series]],
    dcr_id: int,
    mode: str,
) -> pd.Series | None:
    rows = tracker_rows.get(dcr_id, [])
    if not rows:
        return None
    if mode == "evaluation":
        for row in rows:
            if "Eval" in str(row.get("PRCRState", "")):
                return row
    else:
        for row in rows:
            state = str(row.get("PRCRState", ""))
            if "Impl" in state and "Eval" not in state:
                return row
    return rows[-1]


def format_date(value) -> str:
    if pd.isna(value) or value in (0, "0"):
        return "-"
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d-%b-%Y")
    try:
        parsed = pd.to_datetime(value, dayfirst=True)
        return parsed.strftime("%d-%b-%Y")
    except (TypeError, ValueError):
        return str(value)


def latest_comment(comments: str | float, max_len: int | None = 120) -> str:
    if pd.isna(comments):
        return "-"
    text = str(comments).strip()
    if not text:
        return "-"
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    latest = lines[0]
    if max_len is not None and len(latest) > max_len:
        return latest[: max_len - 3] + "..."
    return latest


def _comment_activity_score(comments: str | float) -> int:
    if pd.isna(comments):
        return 0
    text = str(comments)
    lower = text.lower()
    score = min(len(text) // 150, 5)
    if any(token in lower for token in ("l2", "drb", "ccb", "rrb", "evaluation", "peer review")):
        score += 3
    if any(token in text for token in ("18-06", "17-06", "16-06", "15-06", "14-06", "13-06", "12-06")):
        score += 2
    return score


def eval_status_from_row(row: pd.Series) -> str:
    comments = str(row.get("Comments (Daily)", "")).lower()
    if "working on l2" in comments or "l2 comments" in comments or "l2 review" in comments:
        return "Working on L2 Review Comments"
    if "evaluation submitted" in comments or "sent for l2" in comments or "send for l2" in comments:
        return "Evaluation Submitted"
    if "ccb" in comments and ("progress" in comments or "review" in comments):
        return "CCB in progress"
    if "working on the comments" in comments or "comments received" in comments:
        return "Working on the Comments received"
    if "working on the inputs" in comments:
        return "Working on the inputs received"
    for col in ("Current Activity", "DCR State", "PRCRState"):
        value = row.get(col)
        if pd.notna(value) and str(value).strip():
            text = str(value).strip().replace("\n", " ")
            if len(text) > 90:
                return text[:87] + "..."
            return text
    return "Evaluation Submitted"


def impl_status_from_row(row: pd.Series) -> str:
    comments = str(row.get("Comments (Daily)", "")).lower()
    if "l2" in comments:
        return "Working on the L2 Review Comment"
    for col in ("Current Activity", "DCR State", "PRCRState"):
        value = row.get(col)
        if pd.notna(value) and str(value).strip():
            return str(value).strip().replace("\n", " ")[:90]
    return "In Progress"


def _coerce_tracker_date(value, default_year: int = 2026) -> pd.Timestamp | None:
    if pd.isna(value) or value in (0, "0"):
        return None
    if isinstance(value, pd.Timestamp):
        if value.year < 1990:
            return None
        return value
    text = str(value).strip()
    if not text or text.lower() in ("nan", "nat", "00:00:00"):
        return None
    try:
        parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    except (TypeError, ValueError):
        return None
    if pd.isna(parsed) or parsed.year < 1990:
        return None
    if parsed.year == 1900 and default_year:
        parsed = parsed.replace(year=default_year)
    return parsed


def _dates_from_comments(comments, default_year: int = 2026) -> list[pd.Timestamp]:
    if pd.isna(comments):
        return []
    text = str(comments)
    found: list[pd.Timestamp] = []

    month_map = {
        "jan": 1,
        "feb": 2,
        "mar": 3,
        "apr": 4,
        "may": 5,
        "jun": 6,
        "jul": 7,
        "aug": 8,
        "sep": 9,
        "oct": 10,
        "nov": 11,
        "dec": 12,
    }

    for match in re.finditer(
        r"(\d{1,2})(?:st|nd|rd|th)?\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*(?:\s+(\d{4}))?",
        text,
        flags=re.IGNORECASE,
    ):
        day = int(match.group(1))
        month = month_map[match.group(2).lower()[:3]]
        year = int(match.group(3)) if match.group(3) else default_year
        found.append(pd.Timestamp(year=year, month=month, day=day))

    for match in re.finditer(r"(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?", text):
        day = int(match.group(1))
        month = int(match.group(2))
        year_raw = match.group(3)
        if year_raw:
            year = int(year_raw)
            if year < 100:
                year += 2000
        else:
            year = default_year
        try:
            found.append(pd.Timestamp(year=year, month=month, day=day))
        except ValueError:
            continue

    return found


def closure_date_from_row(
    row: pd.Series,
    vis_row: pd.Series | None = None,
    fallback: str = "-",
    default_year: int = 2026,
) -> str:
    candidates: list[pd.Timestamp] = []
    date_columns = [
        "Planned Completion Date\n<dd-mm-yyyy>",
        "Deadline Date as per tagged phase/ Commit Date",
        "DRB/L2\nPlanned date3",
        "PEER \nPlanned date",
        "CCB\nPlanned date2",
        "OBD FORUM\nPLANNED DATE",
        "Support Required by date",
        "Expected completion Dt? (if Delayed)",
    ]
    for column in date_columns:
        parsed = _coerce_tracker_date(row.get(column), default_year=default_year)
        if parsed is not None:
            candidates.append(parsed)

    if vis_row is not None:
        for column in ("Planned End Date", "Planned Start Date"):
            parsed = _coerce_tracker_date(vis_row.get(column), default_year=default_year)
            if parsed is not None and parsed.year >= 1990:
                candidates.append(parsed)

    latest_line = latest_comment(row.get("Comments (Daily)"), max_len=None)
    comment_dates = _dates_from_comments(latest_line, default_year=default_year)
    if comment_dates:
        return format_date(max(comment_dates))

    if candidates:
        return format_date(min(candidates))

    return fallback


def closure_sort_key(closure_date: str) -> tuple[int, pd.Timestamp]:
    text = str(closure_date).strip()
    if not text or text in ("-", "nan"):
        return (1, pd.Timestamp.max)
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        return (1, pd.Timestamp.max)
    return (0, parsed)


def support_required_from_row(row: pd.Series) -> str:
    support = row.get("Support Required from team")
    if pd.notna(support) and str(support).strip() not in ("", "nan"):
        return str(support).strip()
    comments = str(row.get("Comments (Daily)", "")).lower()
    if "yogesh" in comments and "rrb" in comments:
        return "Support Required from Yogesh to complete RRB Tags."
    if "l2 review delayed" in comments:
        return "L2 Review delayed from Component owner (beyond SLA)."
    return "-"


def load_non_stla_planning(data_file: str = DEFAULT_DATA_FILE) -> pd.DataFrame:
    return pd.read_excel(data_file, sheet_name="Non_STLA (Planning)", header=1)


def _planning_dcr_column(planning: pd.DataFrame) -> str | None:
    for column in ("DCRID-PTC", planning.columns[1] if len(planning.columns) > 1 else None):
        if column and column in planning.columns:
            return column
    return None


def _planning_dcr_ids(planning: pd.DataFrame) -> list[int]:
    dcr_col = _planning_dcr_column(planning)
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


def _core2_planned_count(planning: pd.DataFrame) -> int | None:
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

    planned_ids = _planning_dcr_ids(planning)
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
            if (core2_count := _core2_planned_count(planning)) is not None
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


def _visibility_row(visibility: pd.DataFrame, dcr_id: int) -> pd.Series | None:
    matches = visibility[visibility["DCR Number"] == dcr_id]
    if matches.empty:
        return None
    return matches.iloc[0]


def _build_pending_item(
    dcr_id: int,
    tracker_rows: dict[int, list[pd.Series]],
    visibility: pd.DataFrame,
    mode: str,
) -> dict | None:
    tracker_row = _tracker_row_for_mode(tracker_rows, dcr_id, mode)
    if tracker_row is None:
        return None
    vis_row = _visibility_row(visibility, dcr_id)
    summary = tracker_row.get("Summary", vis_row.get("Subject", "-") if vis_row is not None else "-")

    if mode == "evaluation":
        return {
            "dcr_id": dcr_id,
            "summary": str(summary) if pd.notna(summary) else "-",
            "status": eval_status_from_row(tracker_row),
            "closure_date": closure_date_from_row(tracker_row, vis_row),
            "support": support_required_from_row(tracker_row),
            "remarks": latest_comment(tracker_row.get("Comments (Daily)"), max_len=None),
        }
    return {
        "dcr_id": dcr_id,
        "summary": str(summary) if pd.notna(summary) else "-",
        "status": impl_status_from_row(tracker_row),
        "remarks": latest_comment(tracker_row.get("Comments (Daily)"), max_len=None),
    }


def pending_items(
    visibility: pd.DataFrame,
    tracker_rows: dict[int, list[pd.Series]],
    mode: str,
    pending_week: int | None = None,
    limit: int = 12,
) -> list[dict]:
    del pending_week  # week label is applied on the slide; rows come from current workbook state.
    mode = mode.lower()

    if mode == "evaluation":
        type_mask = visibility["Evaluation/ Implementation"].astype(str).str.contains("Eval", case=False, na=False)
    else:
        type_mask = visibility["Evaluation/ Implementation"].astype(str).str.contains("Impl", case=False, na=False)

    excluded_status = visibility["Status"].astype(str).str.contains(
        r"Rejected|Cancelled|Deferred|Closed", case=False, na=False
    )

    candidates: list[tuple[int, int]] = []
    for _, vis_row in visibility[type_mask & ~excluded_status].iterrows():
        dcr_id = parse_dcr_id(vis_row.get("DCR Number"))
        if dcr_id is None:
            continue

        tracker_row = _tracker_row_for_mode(tracker_rows, dcr_id, mode)
        if tracker_row is None:
            continue
        if pd.notna(tracker_row.get("Actual Cmpln Date\n<dd-mm-yyyy>")):
            continue

        status = str(vis_row.get("Status", ""))
        score = _comment_activity_score(tracker_row.get("Comments (Daily)"))
        if status in ("On Track", "At Risk", "Yet to start", "On Hold"):
            score += 4
        if mode == "implementation":
            summary = str(tracker_row.get("Summary", "")).lower()
            comments = str(tracker_row.get("Comments (Daily)", "")).lower()
            if "l2" in comments:
                score += 5
            if "appguide" in summary or "app-guide" in summary:
                score += 4
            if "dummy" in summary:
                score -= 10
        else:
            if str(tracker_row.get("PRCR Stage", "")) in ("Verify", "Submit", "Assess", "In Progress"):
                score += 2

        candidates.append((score, dcr_id))

    candidates.sort(key=lambda item: (-item[0], item[1]))
    selected_ids: list[int] = []
    seen: set[int] = set()
    for _, dcr_id in candidates:
        if dcr_id in seen:
            continue
        seen.add(dcr_id)
        selected_ids.append(dcr_id)
        if len(selected_ids) >= limit:
            break

    items = []
    for dcr_id in selected_ids:
        item = _build_pending_item(dcr_id, tracker_rows, visibility, mode)
        if item is not None:
            items.append(item)

    if mode == "evaluation":
        items.sort(key=lambda item: closure_sort_key(item["closure_date"]))

    return items


def pending_week_for_chart(chart_week: int) -> int:
    """PDF uses prior week for pending-closure tables while charts show current week."""
    return max(chart_week - 1, 1)


def graph_week_capacity(data_file: str, pending_week: int, mode: str) -> int:
    section = (
        get_evaluation_data(data_file=data_file)
        if mode == "evaluation"
        else get_implementation_data(data_file=data_file)
    )
    match = section[section["Week No"] == pending_week]
    if match.empty:
        return 12
    value = match.iloc[0].get("Eval In Progress")
    if pd.isna(value):
        return 12
    return max(1, min(int(value), 12))


def _ddp_row_item(
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
        items.append(_ddp_row_item(row, tracker_lookup_map, len(items) + 1))
        if len(items) >= limit:
            return items

    rows = ddp[ddp["Diagnostics Name"].notna()].copy()
    for _, row in rows.iterrows():
        items.append(_ddp_row_item(row, tracker_lookup_map, len(items) + 1))
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
    dcr_col = _planning_dcr_column(planning)
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


def week_remarks(section: pd.DataFrame, week_no: int) -> str:
    match = section[section["Week No"] == week_no]
    if match.empty:
        return ""
    remark = match.iloc[0].get("Remarks")
    drb = match.iloc[0].get("DRB /l2 Reviews & Rework  in progress")
    parts = []
    if pd.notna(remark) and str(remark).strip():
        parts.append(str(remark).strip())
    if pd.notna(drb) and str(drb).strip() not in ("", "nan"):
        parts.append(f"DRB/L2 in progress: {int(float(drb))}")
    return " | ".join(parts)
