"""Chart generation for WSR reports."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

from wsr_common import add_week_labels, get_evaluation_data, get_implementation_data, to_percentage


def _plot_section(
    section,
    *,
    title: str,
    progress_col: str,
    progress_label: str,
    output_path: str | Path,
    figsize=(16, 8),
) -> Path:
    section = add_week_labels(section)
    x = np.arange(len(section))
    bar_width = 0.12

    fig, ax1 = plt.subplots(figsize=figsize)

    bars1 = ax1.bar(
        x - 2 * bar_width,
        section["Cumulative (Baseline Plan)"],
        width=bar_width,
        color="#7ED321",
        label="Cumulative (Baseline Plan)",
    )
    bars2 = ax1.bar(
        x - bar_width,
        section["Cumulative Revised basline plan"],
        width=bar_width,
        color="#B8E986",
        label="Cumulative Revised baseline plan",
    )
    bars3 = ax1.bar(
        x,
        section["Cumulative (Completed)"],
        width=bar_width,
        color="#00B050",
        label="Cumulative (Completed)",
    )
    bars4 = ax1.bar(
        x + bar_width,
        section["Cumulative Rejected / Transferred/ Moved to next quarter"],
        width=bar_width,
        color="#F4B400",
        label="Rejected/Transferred",
    )
    bars5 = ax1.bar(
        x + 2 * bar_width,
        section[progress_col],
        width=bar_width,
        color="#FFF000",
        label=progress_label,
    )

    ax2 = ax1.twinx()
    confidence = to_percentage(section["% Completion Confidence - Overall"])
    actual = to_percentage(section["% Actual weekly completion wr.t  revised Baseline"])

    ax2.plot(x, confidence, color="black", linewidth=3, label="% Completion Confidence")
    ax2.plot(x, actual, color="#7030A0", linewidth=3, label="% Actual weekly completion")

    for bars in [bars1, bars2, bars3, bars4, bars5]:
        for bar in bars:
            height = bar.get_height()
            if not np.isnan(height) and height > 0:
                ax1.text(
                    bar.get_x() + bar.get_width() / 2,
                    height,
                    f"{int(height)}",
                    ha="center",
                    va="bottom",
                    fontsize=9,
                )

    for i, value in enumerate(confidence):
        if not np.isnan(value):
            ax2.text(i, value + 2, f"{int(value)}%", ha="center", fontsize=10, color="black")

    for i, value in enumerate(actual):
        if not np.isnan(value):
            ax2.text(i, value - 8, f"{int(value)}%", ha="center", fontsize=10, color="#7030A0")

    ax1.set_title(title, fontsize=20, fontweight="bold")
    ax1.set_xticks(x)
    ax1.set_xticklabels(section["Week Label"])
    ax1.set_ylabel("DCR Count")
    ax2.set_ylabel("Percentage")
    ax2.set_ylim(0, 120)
    ax1.grid(axis="y", linestyle="--", alpha=0.4)

    handles1, labels1 = ax1.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(
        handles1 + handles2,
        labels1 + labels2,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.12),
        ncol=3,
    )

    plt.tight_layout()
    output_path = Path(output_path)
    plt.savefig(output_path, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return output_path


def save_implementation_chart(output_path: str | Path, data_file: str = "data.xlsm") -> Path:
    section = get_implementation_data(data_file=data_file)
    return _plot_section(
        section,
        title="Q3'26 Implementation",
        progress_col="Eval In Progress",
        progress_label="Impl In Progress",
        output_path=output_path,
    )


def save_evaluation_chart(output_path: str | Path, data_file: str = "data.xlsm") -> Path:
    section = get_evaluation_data(data_file=data_file)
    return _plot_section(
        section,
        title="Q3'26 Evaluation",
        progress_col="Eval In Progress",
        progress_label="Eval In Progress",
        output_path=output_path,
    )
