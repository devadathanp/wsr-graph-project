"""Active risks for slide 10 from the Risk and Mitigation Plan sheet."""

from __future__ import annotations

import pandas as pd

# Status values that still belong on the WSR (closed risks are omitted).
ACTIVE_RISK_STATUSES = frozenset({"open", "pending", "in progress"})

COL_SNO = "S_no"
COL_DCR = "DCR(If any)"
COL_RISK = "What is the Risk/Issue"
COL_IMPACT = "Impact"
COL_SUPPORT = "Support Required"
COL_STATUS = "Status"


def _cell(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in ("nan", "none", "nat"):
        return ""
    # Excel often stores DCR IDs as floats (19381827.0).
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return text


def active_risk_items(risks: pd.DataFrame) -> list[dict[str, str]]:
    """Return open / pending / in-progress risks for the WSR risks slide."""
    if risks is None or risks.empty:
        return []

    required = (COL_RISK, COL_STATUS)
    missing = [col for col in required if col not in risks.columns]
    if missing:
        return []

    items: list[dict[str, str]] = []
    for _, row in risks.iterrows():
        status = _cell(row.get(COL_STATUS))
        if status.lower() not in ACTIVE_RISK_STATUSES:
            continue
        risk_text = _cell(row.get(COL_RISK))
        if not risk_text:
            continue
        items.append(
            {
                "dcr": _cell(row.get(COL_DCR)) if COL_DCR in risks.columns else "",
                "risk": risk_text,
                "impact": _cell(row.get(COL_IMPACT)) if COL_IMPACT in risks.columns else "",
                "support": _cell(row.get(COL_SUPPORT)) if COL_SUPPORT in risks.columns else "",
                "status": status,
            }
        )
    return items
