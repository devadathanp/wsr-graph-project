"""Visibility sheet aggregations."""

from __future__ import annotations

import pandas as pd


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
