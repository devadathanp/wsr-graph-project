"""Tracker row parsing, dates, and status derivation."""

from __future__ import annotations

import re

import pandas as pd

from wsr.loaders import load_non_stla_planning, load_visibility


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


def tracker_row_for_mode(
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


def comment_activity_score(comments: str | float) -> int:
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


def coerce_tracker_date(value, default_year: int = 2026) -> pd.Timestamp | None:
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


def dates_from_comments(comments, default_year: int = 2026) -> list[pd.Timestamp]:
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
        parsed = coerce_tracker_date(row.get(column), default_year=default_year)
        if parsed is not None:
            candidates.append(parsed)

    if vis_row is not None:
        for column in ("Planned End Date", "Planned Start Date"):
            parsed = coerce_tracker_date(vis_row.get(column), default_year=default_year)
            if parsed is not None and parsed.year >= 1990:
                candidates.append(parsed)

    latest_line = latest_comment(row.get("Comments (Daily)"), max_len=None)
    comment_dates = dates_from_comments(latest_line, default_year=default_year)
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


def visibility_row(visibility: pd.DataFrame, dcr_id: int) -> pd.Series | None:
    matches = visibility[visibility["DCR Number"] == dcr_id]
    if matches.empty:
        return None
    return matches.iloc[0]


# Backward-compatible alias used internally.
_tracker_row_for_mode = tracker_row_for_mode
_visibility_row = visibility_row
