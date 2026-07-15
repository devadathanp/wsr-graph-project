"""Chart generation for WSR reports (data from CSAR_WSR_Graph Non-STLA sheet)."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patheffects as pe
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

CHART_BASELINE = "#7ED321"
CHART_REVISED = "#B8E986"
CHART_COMPLETED = "#00B050"
CHART_REJECTED = "#FFC000"
CHART_IN_PROGRESS = "#FFFF00"
CHART_DRB = "#00B0F0"
CHART_CONFIDENCE = "#8064A2"
CHART_ACTUAL = "#7030A0"

DRB_COLUMN = COL_DRB


_LABEL_STROKE = [pe.withStroke(linewidth=2.2, foreground="white")]


def _label_bars(ax, bars_groups, bar_max: float, ylim_top: float) -> None:
    inside_min = max(bar_max * 0.14, 9)
    # Keep above-bar numbers out of the top band reserved for % line labels.
    max_above_y = ylim_top * 0.72
    n_weeks = len(bars_groups[0]) if bars_groups else 0

    for week_idx in range(n_weeks):
        above_slot = 0
        for series_idx, bars in enumerate(bars_groups):
            bar = bars[week_idx]
            height = bar.get_height()
            if np.isnan(height) or height < 0.5:
                continue
            x = bar.get_x() + bar.get_width() / 2
            value = f"{int(round(height))}"

            # Prefer on-bar when tall; otherwise above — but never into the % label band.
            place_inside = height >= inside_min and (
                series_idx in (0, 2) or height + max(bar_max * 0.05, 2) > max_above_y
            )
            if place_inside:
                ax.text(
                    x,
                    height * 0.48,
                    value,
                    ha="center",
                    va="center",
                    fontsize=6.5,
                    color="#161718",
                    zorder=5,
                    clip_on=False,
                    path_effects=_LABEL_STROKE,
                )
            else:
                y_pad = max(bar_max * 0.014, 0.75) + above_slot * max(bar_max * 0.028, 1.2)
                above_slot += 1
                y = min(height + y_pad, max_above_y)
                if y <= height + 0.2:
                    ax.text(
                        x,
                        height * 0.48,
                        value,
                        ha="center",
                        va="center",
                        fontsize=6.5,
                        color="#161718",
                        zorder=5,
                        clip_on=False,
                        path_effects=_LABEL_STROKE,
                    )
                else:
                    ax.text(
                        x,
                        y,
                        value,
                        ha="center",
                        va="bottom",
                        fontsize=6.5,
                        color="#161718",
                        zorder=5,
                        clip_on=False,
                        path_effects=_LABEL_STROKE,
                    )


def _label_percentages(
    ax,
    x_positions: np.ndarray,
    confidence: np.ndarray,
    actual: np.ndarray,
    *,
    y_pad: float = 1.8,
) -> None:
    """Anchor each % label to its line point (tiny offset so the glyph sits on the line)."""
    y_lo, y_hi = ax.get_ylim()

    def _draw(xpos: float, value: float, color: str, above: bool) -> None:
        if above:
            y = min(value + y_pad, y_hi - 0.5)
            va = "bottom"
        else:
            y = max(value - y_pad, y_lo + 0.5)
            va = "top"
        ax.text(
            xpos,
            y,
            f"{int(round(value))}%",
            ha="center",
            va=va,
            fontsize=7,
            color=color,
            zorder=6,
            clip_on=False,
            path_effects=_LABEL_STROKE,
        )

    for index, xpos in enumerate(x_positions):
        conf = confidence[index] if index < len(confidence) else np.nan
        act = actual[index] if index < len(actual) else np.nan
        has_conf = not np.isnan(conf)
        has_act = not np.isnan(act)

        if has_conf and has_act:
            conf_above = float(conf) >= float(act)
            _draw(xpos, float(conf), CHART_CONFIDENCE, above=conf_above)
            _draw(xpos, float(act), CHART_ACTUAL, above=not conf_above)
        else:
            if has_conf:
                _draw(xpos, float(conf), CHART_CONFIDENCE, above=True)
            if has_act:
                _draw(xpos, float(act), CHART_ACTUAL, above=False)


def _plot_section(
    section,
    *,
    title: str,
    progress_col: str,
    progress_label: str,
    output_path: str | Path,
    figsize=(13.4, 4.65),
) -> Path:
    section = add_week_labels(section)
    x = np.arange(len(section))
    bar_width = 0.105
    offsets = (-2.6, -1.55, -0.5, 0.55, 1.6, 2.65)

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

    ylim_top = max(bar_max * 1.22, bar_max + 10)
    ax1.set_ylim(0, ylim_top)
    _label_bars(ax1, bars_groups, bar_max, ylim_top)

    ax2 = ax1.twinx()
    confidence = to_percentage(section[COL_PCT_CONFIDENCE]).to_numpy(dtype=float)
    actual = to_percentage(section[COL_PCT_ACTUAL]).to_numpy(dtype=float)

    ax2.plot(x, confidence, color=CHART_CONFIDENCE, linewidth=2.2, label="% Completion Confidence - Overall", zorder=3)
    ax2.plot(x, actual, color=CHART_ACTUAL, linewidth=2.2, label="% Actual weekly completion w.r.t revised Baseline", zorder=3)
    pct_max = np.nanmax([*confidence, *actual, 100.0])
    ax2.set_ylim(0, max(120.0, float(pct_max) + 12.0))
    _label_percentages(ax2, x, confidence, actual)

    ax1.set_title(title, fontdict={"fontsize": 14, "fontweight": "normal"}, pad=6)
    ax1.set_xticks(x)
    ax1.set_xticklabels(section["Week Label"], fontsize=8)
    ax1.set_ylabel("DCR Count", fontsize=9)
    ax2.set_ylabel("Percentage", fontsize=9)
    ax1.grid(axis="y", linestyle="--", alpha=0.28, zorder=0)

    for spine in ax1.spines.values():
        spine.set_visible(False)
    for spine in ax2.spines.values():
        spine.set_visible(False)
    ax1.spines["bottom"].set_visible(True)
    ax1.spines["bottom"].set_color("#9aa0a6")
    ax1.tick_params(axis="both", length=0)
    ax2.tick_params(axis="both", length=0)

    handles1, labels1 = ax1.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(
        handles1 + handles2,
        labels1 + labels2,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.18),
        ncol=2,
        fontsize=6,
        frameon=False,
    )

    plt.tight_layout(rect=(0, 0.07, 1, 0.98))
    output_path = Path(output_path)
    plt.savefig(output_path, dpi=220, bbox_inches="tight", pad_inches=0.08)
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
    categories = ["Q3 Actual Available", f"{planning['planned_pct']}% of Q3 is Planned"]
    values = [planning["available_hours"], planning["planned_hours"]]

    fig, ax = plt.subplots(figsize=(11.2, 4.6))
    bars = ax.bar(categories, values, color="#92D050", width=0.45)
    ax.set_title(
        "PFS Quarterly Planning 2026",
        fontdict={"fontsize": 14, "fontweight": "normal"},
        color="#161718",
        pad=12,
    )
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
