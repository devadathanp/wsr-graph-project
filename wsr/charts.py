"""Chart generation for WSR reports (data from CSAR_WSR_Graph Non-STLA sheet)."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

from wsr.constants import DEFAULT_DATA_FILE
from wsr.graph import (
    COL_COMPLETED,
    COL_CUMULATIVE_BASELINE,
    COL_CUMULATIVE_REVISED,
    COL_DRB,
    COL_IN_PROGRESS,
    COL_PCT_ACTUAL,
    COL_PCT_CONFIDENCE,
    COL_REJECTED,
    add_week_labels,
    get_evaluation_data,
    get_implementation_data,
    to_percentage,
)

# Reference WSR graph palette (slide 4).
CHART_BASELINE = "#7ED321"
CHART_REVISED = "#B8E986"
CHART_COMPLETED = "#00B050"
CHART_REJECTED = "#FFC000"
CHART_IN_PROGRESS = "#FFFF00"
CHART_DRB = "#00B0F0"
CHART_CONFIDENCE = "#8064A2"
CHART_ACTUAL = "#7030A0"

# Backward-compatible alias (sheet column name).
DRB_COLUMN = COL_DRB

# Bar groups that receive numeric labels (index into bar_specs below).
_LABEL_COMPLETED = 2
_LABEL_IN_PROGRESS = 4


def _bar_value_labels(ax, bars, *, min_value: float = 1, y_pad: float = 0.8) -> None:
    """Place a single value above each bar with a small white backing for readability."""
    for bar in bars:
        height = bar.get_height()
        if np.isnan(height) or height < min_value:
            continue
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            height + y_pad,
            f"{int(height)}",
            ha="center",
            va="bottom",
            fontsize=7.5,
            color="#161718",
            bbox={"boxstyle": "round,pad=0.12", "facecolor": "white", "edgecolor": "none", "alpha": 0.9},
            zorder=5,
            clip_on=False,
        )


def _changed_value_labels(ax, bars, *, min_value: float = 1) -> None:
    """Label bars only when the value changes from the previous week."""
    previous: float | None = None
    for bar in bars:
        height = bar.get_height()
        if np.isnan(height) or height < min_value:
            previous = height
            continue
        if previous is None or height != previous:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                height + 0.8,
                f"{int(height)}",
                ha="center",
                va="bottom",
                fontsize=7.5,
                color="#161718",
                bbox={"boxstyle": "round,pad=0.12", "facecolor": "white", "edgecolor": "none", "alpha": 0.9},
                zorder=5,
                clip_on=False,
            )
        previous = height


def _plot_section(
    section,
    *,
    title: str,
    progress_col: str,
    progress_label: str,
    output_path: str | Path,
    figsize=(11.8, 4.35),
) -> Path:
    section = add_week_labels(section)
    x = np.arange(len(section))
    bar_width = 0.1
    offsets = (-2.5, -1.5, -0.5, 0.5, 1.5, 2.5)

    fig, ax1 = plt.subplots(figsize=figsize)

    bar_specs = [
        (offsets[0], COL_CUMULATIVE_BASELINE, CHART_BASELINE, "Cumulative (Baseline Plan)"),
        (offsets[1], COL_CUMULATIVE_REVISED, CHART_REVISED, "Cumulative Revised baseline plan"),
        (offsets[2], COL_COMPLETED, CHART_COMPLETED, "Cumulative (Completed)"),
        (offsets[3], COL_REJECTED, CHART_REJECTED, "Rejected / Transferred / Moved to next quarter"),
        (offsets[4], progress_col, CHART_IN_PROGRESS, progress_label),
        (offsets[5], COL_DRB, CHART_DRB, "DRB / L2 Reviews & Rework in progress"),
    ]

    bars_groups = []
    bar_max = 0.0
    for offset, column, color, label in bar_specs:
        values = section[column] if column in section.columns else [0] * len(section)
        bars = ax1.bar(
            x + offset * bar_width,
            values,
            width=bar_width,
            color=color,
            label=label,
            zorder=2,
        )
        bars_groups.append(bars)
        for bar in bars:
            height = bar.get_height()
            if not np.isnan(height):
                bar_max = max(bar_max, height)

    ax1.set_ylim(0, max(bar_max * 1.12, bar_max + 6))

    # In-progress is the key weekly metric; completed only when it moves.
    _bar_value_labels(ax1, bars_groups[_LABEL_IN_PROGRESS], min_value=1)
    _changed_value_labels(ax1, bars_groups[_LABEL_COMPLETED], min_value=1)

    ax2 = ax1.twinx()
    confidence = to_percentage(section[COL_PCT_CONFIDENCE])
    actual = to_percentage(section[COL_PCT_ACTUAL])

    ax2.plot(x, confidence, color=CHART_CONFIDENCE, linewidth=2.2, label="% Completion Confidence - Overall", zorder=3)
    ax2.plot(x, actual, color=CHART_ACTUAL, linewidth=2.2, label="% Actual weekly completion w.r.t revised Baseline", zorder=3)
    ax2.set_ylim(0, 120)

    ax1.set_title(title, fontsize=15, fontweight="bold", pad=10)
    ax1.set_xticks(x)
    ax1.set_xticklabels(section["Week Label"], fontsize=8.5)
    ax1.set_ylabel("DCR Count", fontsize=10)
    ax2.set_ylabel("Percentage", fontsize=10)
    ax1.grid(axis="y", linestyle="--", alpha=0.3, zorder=0)

    handles1, labels1 = ax1.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(
        handles1 + handles2,
        labels1 + labels2,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.2),
        ncol=2,
        fontsize=6.5,
        frameon=False,
    )

    plt.tight_layout(rect=(0, 0.08, 1, 0.98))
    output_path = Path(output_path)
    plt.savefig(output_path, dpi=220, bbox_inches="tight")
    plt.close(fig)
    return output_path


def save_implementation_chart(output_path: str | Path, data_file: str = DEFAULT_DATA_FILE) -> Path:
    section = get_implementation_data(data_file=data_file)
    return _plot_section(
        section,
        title="Q3'26 Implementation",
        progress_col=COL_IN_PROGRESS,
        progress_label="Impl In Progress",
        output_path=output_path,
    )


def save_evaluation_chart(output_path: str | Path, data_file: str = DEFAULT_DATA_FILE) -> Path:
    section = get_evaluation_data(data_file=data_file)
    return _plot_section(
        section,
        title="Q3'26 Evaluation",
        progress_col=COL_IN_PROGRESS,
        progress_label="Eval In Progress",
        output_path=output_path,
    )


def save_planning_chart(
    planning: dict[str, int],
    output_path: str | Path,
) -> Path:
    """Quarterly planning bar chart as PNG (avoids native PPT chart repair issues)."""
    categories = ["Q3 Actual Available", f"{planning['planned_pct']}% of Q3 is Planned"]
    values = [planning["available_hours"], planning["planned_hours"]]

    fig, ax = plt.subplots(figsize=(11.2, 4.6))
    bars = ax.bar(categories, values, color="#92D050", width=0.45)
    ax.set_title("PFS Quarterly Planning 2026", fontsize=14, fontweight="bold", color="#161718", pad=12)
    ax.set_ylabel("")
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", linestyle="--", alpha=0.3)

    for bar in bars:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            height,
            f"{int(height)}",
            ha="center",
            va="bottom",
            fontsize=11,
            color="#161718",
        )

    plt.tight_layout()
    output_path = Path(output_path)
    plt.savefig(output_path, dpi=220, bbox_inches="tight")
    plt.close(fig)
    return output_path
