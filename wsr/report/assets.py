"""Chart PNG generation for the report."""

from __future__ import annotations

from pathlib import Path

from wsr.charts import save_evaluation_chart, save_implementation_chart, save_planning_chart
from wsr.planning_book import load_quarterly_planning
from wsr.report.models import ChartAssets
from wsr.run_log import RunLog


def build_chart_assets(
    scrum_path: Path,
    assets_dir: Path,
    planning_book_path: Path | None,
    log: RunLog,
) -> ChartAssets:
    log.info("Building charts…")
    impl_chart = save_implementation_chart(
        assets_dir / "implementation_chart.png",
        data_file=str(scrum_path),
    )
    eval_chart = save_evaluation_chart(
        assets_dir / "evaluation_chart.png",
        data_file=str(scrum_path),
    )

    quarterly_planning = load_quarterly_planning(planning_book_path)
    planning_chart = None
    if quarterly_planning is None:
        if planning_book_path is not None:
            log.warning(
                f'Could not find "Total work Hrs. Available for PFS team" in '
                f"{planning_book_path.name}; slide 11 will show a placeholder."
            )
    else:
        planning_chart = save_planning_chart(
            quarterly_planning,
            assets_dir / "planning_chart.png",
        )
        log.info(
            f"Planning chart: available={quarterly_planning['available_hours']}, "
            f"planned={quarterly_planning['planned_hours']}, "
            f"resources={quarterly_planning['resources']}"
        )

    return ChartAssets(
        impl_chart=impl_chart,
        eval_chart=eval_chart,
        planning_chart=planning_chart,
        quarterly_planning=quarterly_planning,
    )
