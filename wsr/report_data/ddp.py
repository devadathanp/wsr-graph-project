"""Slide 7 DDP MS4-5 table rows."""

from __future__ import annotations

import pandas as pd

from wsr.tracker import format_date, latest_comment, parse_dcr_id


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
