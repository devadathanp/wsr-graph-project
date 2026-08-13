"""Chart PNG generation for the report."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from wsr.charts import save_evaluation_chart, save_implementation_chart, save_planning_chart
from wsr.planning_book import load_quarterly_planning
from wsr.report.models import ChartAssets
from wsr.run_log import RunLog


def build_chart_assets(
    scrum_path: Path,
    assets_dir: Path,
    log: RunLog,
    *,
    tracker: pd.DataFrame | None = None,
    quarter_short: str = "Q3'26",
    fiscal_year: int = 2026,
) -> ChartAssets:
    log.info("Building charts…")
    impl_chart = save_implementation_chart(
        assets_dir / "implementation_chart.png",
        data_file=str(scrum_path),
        quarter_short=quarter_short,
    )
    eval_chart = save_evaluation_chart(
        assets_dir / "evaluation_chart.png",
        data_file=str(scrum_path),
        quarter_short=quarter_short,
    )

    quarterly_planning = load_quarterly_planning(tracker=tracker)
    planning_chart = None
    if quarterly_planning is None:
        log.warning(
            'Could not find Actual Available Estimate on Non STLA; '
            "the quarterly planning slide will show a placeholder."
        )
    else:
        quarterly_planning["quarter_token"] = quarter_short.split("'")[0]
        quarterly_planning["fiscal_year"] = fiscal_year
        planning_chart = save_planning_chart(
            quarterly_planning,
            assets_dir / "planning_chart.png",
        )
        log.info(
            f"Planning chart: available={quarterly_planning['available_hours']}, "
            f"estimated_hrs={quarterly_planning['estimated_hours']} "
            f"({quarterly_planning['planned_pct']}% of available), "
            f"burndown={quarterly_planning['burndown_hours']}"
        )

    return ChartAssets(
        impl_chart=impl_chart,
        eval_chart=eval_chart,
        planning_chart=planning_chart,
        quarterly_planning=quarterly_planning,
    )
