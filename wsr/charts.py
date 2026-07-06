"""Chart generation for WSR reports."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

from wsr.constants import DEFAULT_DATA_FILE
from wsr.graph import add_week_labels, get_evaluation_data, get_implementation_data, to_percentage

# Reference WSR graph palette (slide 4).
CHART_BASELINE = "#7ED321"
CHART_REVISED = "#B8E986"
CHART_COMPLETED = "#00B050"
CHART_REJECTED = "#FFC000"
CHART_IN_PROGRESS = "#FFFF00"
CHART_DRB = "#00B0F0"
CHART_CONFIDENCE = "#8064A2"
CHART_ACTUAL = "#7030A0"
DRB_COLUMN = "DRB /l2 Reviews & Rework  in progress"


def _plot_section(
    section,
    *,
    title: str,
    progress_col: str,
    progress_label: str,
    output_path: str | Path,
    figsize=(11.5, 4.05),
) -> Path:
    section = add_week_labels(section)
    x = np.arange(len(section))
    bar_width = 0.11
    offsets = (-2.5, -1.5, -0.5, 0.5, 1.5, 2.5)

    fig, ax1 = plt.subplots(figsize=figsize)

    bar_specs = [
        (offsets[0], "Cumulative (Baseline Plan)", CHART_BASELINE, "Cumulative (Baseline Plan)"),
        (offsets[1], "Cumulative Revised basline plan", CHART_REVISED, "Cumulative Revised baseline plan"),
        (offsets[2], "Cumulative (Completed)", CHART_COMPLETED, "Cumulative (Completed)"),
        (
            offsets[3],
            "Cumulative Rejected / Transferred/ Moved to next quarter",
            CHART_REJECTED,
            "Rejected / Transferred / Moved to next quarter",
        ),
        (offsets[4], progress_col, CHART_IN_PROGRESS, progress_label),
        (offsets[5], DRB_COLUMN, CHART_DRB, "DRB / L2 Reviews & Rework in progress"),
    ]

    bars_groups = []
    for offset, column, color, label in bar_specs:
        values = section[column] if column in section.columns else [0] * len(section)
        bars = ax1.bar(
            x + offset * bar_width,
            values,
            width=bar_width,
            color=color,
            label=label,
        )
        bars_groups.append(bars)

    ax2 = ax1.twinx()
    confidence = to_percentage(section["% Completion Confidence - Overall"])
    actual = to_percentage(section["% Actual weekly completion wr.t  revised Baseline"])

    ax2.plot(
        x,
        confidence,
        color=CHART_CONFIDENCE,
        linewidth=2.5,
        label="% Completion Confidence - Overall",
    )
    ax2.plot(
        x,
        actual,
        color=CHART_ACTUAL,
        linewidth=2.5,
        label="% Actual weekly completion w.r.t revised Baseline",
    )

    bar_max = 0.0
    for bars in bars_groups:
        for bar in bars:
            height = bar.get_height()
            if not np.isnan(height):
                bar_max = max(bar_max, height)

    lane_spacing = 3.4
    lane_base = bar_max + 3.0
    label_ceiling = lane_base + lane_spacing * (len(bars_groups) - 1) + 2.0
    ax1.set_ylim(0, label_ceiling)

    for group_idx, bars in enumerate(bars_groups):
        lane_y = lane_base + group_idx * lane_spacing
        for bar in bars:
            height = bar.get_height()
            if not np.isnan(height) and height > 0:
                bar_x = bar.get_x() + bar.get_width() / 2
                if height < lane_y - 1.5:
                    ax1.plot(
                        [bar_x, bar_x],
                        [height + 0.4, lane_y - 0.35],
                        color="#666666",
                        linewidth=0.45,
                        alpha=0.55,
                        zorder=1,
                    )
                ax1.text(
                    bar_x,
                    lane_y,
                    f"{int(height)}",
                    ha="center",
                    va="bottom",
                    fontsize=6.5,
                    zorder=2,
                    clip_on=False,
                )

    for index, value in enumerate(confidence):
        if not np.isnan(value):
            ax2.text(
                index + 0.38,
                min(value + 6, 118),
                f"{int(value)}%",
                ha="left",
                va="bottom",
                fontsize=7,
                color=CHART_CONFIDENCE,
                clip_on=False,
            )

    for index, value in enumerate(actual):
        if not np.isnan(value):
            ax2.text(
                index + 0.38,
                max(value - 8, 2),
                f"{int(value)}%",
                ha="left",
                va="top",
                fontsize=7,
                color=CHART_ACTUAL,
                clip_on=False,
            )

    ax1.set_title(title, fontsize=15, fontweight="bold", pad=8)
    ax1.set_xticks(x)
    ax1.set_xticklabels(section["Week Label"], fontsize=9)
    ax1.set_ylabel("DCR Count", fontsize=10)
    ax2.set_ylabel("Percentage", fontsize=10)
    ax2.set_ylim(0, 120)
    ax1.grid(axis="y", linestyle="--", alpha=0.35)

    handles1, labels1 = ax1.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(
        handles1 + handles2,
        labels1 + labels2,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.22),
        ncol=2,
        fontsize=7,
        frameon=False,
    )

    plt.tight_layout(rect=(0, 0.06, 1, 0.98))
    output_path = Path(output_path)
    plt.savefig(output_path, dpi=220, bbox_inches="tight")
    plt.close(fig)
    return output_path


def save_implementation_chart(output_path: str | Path, data_file: str = DEFAULT_DATA_FILE) -> Path:
    section = get_implementation_data(data_file=data_file)
    return _plot_section(
        section,
        title="Q3'26 Implementation",
        progress_col="Eval In Progress",
        progress_label="Impl In Progress",
        output_path=output_path,
    )


def save_evaluation_chart(output_path: str | Path, data_file: str = DEFAULT_DATA_FILE) -> Path:
    section = get_evaluation_data(data_file=data_file)
    return _plot_section(
        section,
        title="Q3'26 Evaluation",
        progress_col="Eval In Progress",
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
