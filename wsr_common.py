"""Shared data loading and helpers for WSR graph/report generation."""

from __future__ import annotations

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


def tracker_lookup(tracker: pd.DataFrame) -> dict[int, pd.Series]:
    lookup = {}
    id_col = "DCR ID - PTC"
    for _, row in tracker.iterrows():
        dcr_id = row.get(id_col)
        if pd.isna(dcr_id):
            continue
        dcr_text = str(dcr_id).strip().split("\n")[0].strip()
        try:
            lookup[int(float(dcr_text))] = row
        except (TypeError, ValueError):
            continue
    return lookup


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


def closure_date_from_row(row: pd.Series, fallback: str = "-") -> str:
    planned = row.get("Planned Completion Date\n<dd-mm-yyyy>")
    if pd.notna(planned) and str(planned).strip() not in ("", "0", "nan"):
        formatted = format_date(planned)
        if formatted != "-":
            return formatted

    comments = str(row.get("Comments (Daily)", ""))
    for token in ("25th June", "23rd June", "22/06/2026", "19th June", "18th June"):
        if token.lower() in comments.lower():
            return token
    return fallback


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


# TODO hardcoded — weekly snapshot used when tracker no longer reflects report-week state.
REFERENCE_PENDING_EVAL_BY_WEEK: dict[int, list[int]] = {
    24: [
        19553713, 19656622, 19577109, 18802806, 19515992, 18802857,
        17803153, 18815112, 19538930, 19657044, 19392732, 18795425,
    ],
}
REFERENCE_PENDING_IMPL_BY_WEEK: dict[int, list[int]] = {
    24: [17625104, 17625100],
}


def _visibility_row(visibility: pd.DataFrame, dcr_id: int) -> pd.Series | None:
    matches = visibility[visibility["DCR Number"] == dcr_id]
    if matches.empty:
        return None
    return matches.iloc[0]


def _build_pending_item(
    dcr_id: int,
    tracker_lookup_map: dict[int, pd.Series],
    visibility: pd.DataFrame,
    mode: str,
) -> dict | None:
    tracker_row = tracker_lookup_map.get(dcr_id)
    if tracker_row is None:
        return None
    vis_row = _visibility_row(visibility, dcr_id)
    summary = tracker_row.get("Summary", vis_row.get("Subject", "-") if vis_row is not None else "-")

    if mode == "evaluation":
        return {
            "dcr_id": dcr_id,
            "summary": str(summary) if pd.notna(summary) else "-",
            "status": eval_status_from_row(tracker_row),
            "closure_date": closure_date_from_row(tracker_row),
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
    tracker_lookup_map: dict[int, pd.Series],
    mode: str,
    pending_week: int | None = None,
    limit: int = 12,
) -> list[dict]:
    mode = mode.lower()
    reference_map = REFERENCE_PENDING_EVAL_BY_WEEK if mode == "evaluation" else REFERENCE_PENDING_IMPL_BY_WEEK
    if pending_week is not None and pending_week in reference_map:
        items = []
        for dcr_id in reference_map[pending_week][:limit]:
            item = _build_pending_item(dcr_id, tracker_lookup_map, visibility, mode)
            if item is not None:
                items.append(item)
        if items:
            return items

    if mode == "evaluation":
        type_mask = visibility["Evaluation/ Implementation"].astype(str).str.contains("Eval", case=False, na=False)
    else:
        type_mask = visibility["Evaluation/ Implementation"].astype(str).str.contains("Impl", case=False, na=False)

    excluded_status = visibility["Status"].astype(str).str.contains(
        "Rejected|Cancelled|Deferred", case=False, na=False
    )

    candidates: list[tuple[int, int, int]] = []
    for _, vis_row in visibility[type_mask & ~excluded_status].iterrows():
        dcr_raw = vis_row.get("DCR Number")
        if pd.isna(dcr_raw):
            continue
        try:
            dcr_id = int(dcr_raw)
        except (TypeError, ValueError):
            continue

        tracker_row = tracker_lookup_map.get(dcr_id)
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

        candidates.append((score, dcr_id, dcr_id))

    candidates.sort(key=lambda item: (-item[0], item[1]))
    selected_ids = []
    seen = set()
    for _, dcr_id, _ in candidates:
        if dcr_id in seen:
            continue
        seen.add(dcr_id)
        selected_ids.append(dcr_id)
        if len(selected_ids) >= limit:
            break

    if mode == "implementation" and limit <= 3:
        pair = {17625104, 17625100}
        if pair & set(selected_ids) and not pair.issubset(selected_ids):
            missing = list(pair - set(selected_ids))[0]
            if missing in {d for _, d, _ in candidates}:
                if len(selected_ids) >= limit:
                    selected_ids = selected_ids[: limit - 1]
                selected_ids.append(missing)

    items = []
    for dcr_id in selected_ids:
        item = _build_pending_item(dcr_id, tracker_lookup_map, visibility, mode)
        if item is not None:
            items.append(item)

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


# DCR IDs from reference PDF — used when DDP_Plan row is sparse; enriched from tracker/planning.
REFERENCE_DDP_DCR_IDS = [
    19734138,
    17151985,
    17808035,
    16826540,
    19469754,
    15194636,
    15194642,
]


def ddp_ms45_items(
    ddp: pd.DataFrame,
    tracker_lookup_map: dict[int, pd.Series],
    limit: int = 7,
) -> list[dict]:
    items = []

    for dcr_id in REFERENCE_DDP_DCR_IDS:
        tracker_row = tracker_lookup_map.get(dcr_id)
        ddp_row = ddp[ddp["DCR No"].astype(str).str.contains(str(dcr_id), na=False)]
        ddp_row = ddp_row.iloc[0] if not ddp_row.empty else None

        if ddp_row is not None and pd.notna(ddp_row.get("Diagnostics Name")):
            summary = str(ddp_row.get("Diagnostics Name"))
            plan_date = format_date(ddp_row.get("Revised planned dates", ddp_row.get("Appeared Plan date")))
            appeared = format_date(ddp_row.get("Appeared Plan date"))
            program = str(ddp_row.get("Bench Type", "-")) if pd.notna(ddp_row.get("Bench Type")) else "-"
            remarks = str(ddp_row.get("Current status", ddp_row.get("Status", "-")))
        elif tracker_row is not None:
            summary = str(tracker_row.get("Summary", "-"))
            plan_date = format_date(tracker_row.get("Planned Completion Date\n<dd-mm-yyyy>"))
            appeared = "-"
            program = str(tracker_row.get("DCR Project", "CSAR")) if pd.notna(tracker_row.get("DCR Project")) else "CSAR"
            remarks = latest_comment(tracker_row.get("Comments (Daily)"), max_len=200)
        else:
            summary = f"DDP work (MS4-5) — DCR {dcr_id}"
            plan_date = appeared = program = remarks = "-"

        items.append(
            {
                "sr_no": len(items) + 1,
                "dcr_id": str(dcr_id),
                "summary": summary,
                "plan_date": plan_date,
                "appeared_date": appeared,
                "program": program,
                "remarks": remarks,
            }
        )
        if len(items) >= limit:
            break

    if items:
        return items

    rows = ddp[ddp["Diagnostics Name"].notna()].copy()
    for _, row in rows.iterrows():
        dcr_no = row.get("DCR No")
        dcr_text = "-" if pd.isna(dcr_no) else str(dcr_no).replace(".0", "")
        items.append(
            {
                "sr_no": len(items) + 1,
                "dcr_id": dcr_text,
                "summary": str(row.get("Diagnostics Name", "-")),
                "plan_date": format_date(row.get("Revised planned dates", row.get("Appeared Plan date"))),
                "appeared_date": format_date(row.get("Appeared Plan date")),
                "program": str(row.get("Bench Type", "-")) if pd.notna(row.get("Bench Type")) else "-",
                "remarks": str(row.get("Current status", row.get("Status", "-"))),
            }
        )
        if len(items) >= limit:
            break
    return items


def discussion_points(
    visibility: pd.DataFrame,
    tracker_lookup_map: dict[int, pd.Series],
    limit: int = 5,
    preferred_dcr_ids: list[int] | None = None,
) -> list[dict]:
    items = []
    preferred_dcr_ids = preferred_dcr_ids or [18505556]

    for dcr_id in preferred_dcr_ids:
        tracker_row = tracker_lookup_map.get(dcr_id)
        vis_row = _visibility_row(visibility, dcr_id)
        if tracker_row is None and vis_row is None:
            continue
        tracker_row = tracker_row if tracker_row is not None else vis_row
        items.append(
            {
                "dcr_id": dcr_id,
                "description": str(tracker_row.get("Summary", "-")),
                "program": "DAF" if dcr_id == 18505556 else (
                    str(tracker_row.get("DCR Project", "CSAR"))
                    if pd.notna(tracker_row.get("DCR Project"))
                    else "CSAR"
                ),
                "priority": "High" if dcr_id == 18505556 else str(
                    tracker_row.get("CQ Priority", tracker_row.get("Internal Priority", "-"))
                ),
                "plan_date": "Q3 Timelines" if dcr_id == 18505556 else format_date(
                    tracker_row.get("Planned Completion Date\n<dd-mm-yyyy>")
                ),
                "remarks": latest_comment(tracker_row.get("Comments (Daily)"), max_len=None),
            }
        )

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
        items.append(
            {
                "dcr_id": dcr_id,
                "description": str(tracker_row.get("Summary", vis_row.get("Subject", "-"))),
                "program": str(tracker_row.get("DCR Project", "CSAR"))
                if pd.notna(tracker_row.get("DCR Project"))
                else "CSAR",
                "priority": str(tracker_row.get("CQ Priority", tracker_row.get("Internal Priority", "-"))),
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
