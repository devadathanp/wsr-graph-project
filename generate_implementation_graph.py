import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# =====================================================
# READ EXCEL
# =====================================================

file_path = "data.xlsm"

sheet_name = "CSAR_WSR_Graph (Non-STLA)"

df = pd.read_excel(
    file_path,
    sheet_name=sheet_name,
    header=2,
)

# -----------------------------------------------------
# FILTER IMPLEMENTATION SECTION
# -----------------------------------------------------

impl_start_idx = df.index[
    df["Tagged to Release"].astype(str).str.contains("Q3", na=False)
    & df["Tagged to Release"].astype(str).str.contains("Implementation", na=False)
]

impl = df.loc[impl_start_idx[0]:]
impl = impl[impl["Week No"].notna()].copy()
impl.reset_index(drop=True, inplace=True)

progress_col = "Eval In Progress"


def to_percentage(series):
    def convert(value):
        if pd.isna(value):
            return np.nan
        if isinstance(value, (int, float)):
            val = float(value)
            return val * 100 if abs(val) <= 1.5 else val
        text = str(value).replace("%", "").strip()
        if text == "":
            return np.nan
        val = float(text)
        return val * 100 if abs(val) <= 1.5 else val

    return series.apply(convert)

# -----------------------------------------------------
# X AXIS LABELS
# -----------------------------------------------------

impl["Week Label"] = (
    impl["Week No"].astype(str)
    + "\n"
    + pd.to_datetime(impl["Date"]).dt.strftime("%d-%m")
)

x = np.arange(len(impl))

# =====================================================
# FIGURE
# =====================================================

fig, ax1 = plt.subplots(figsize=(16,8))

bar_width = 0.12

# =====================================================
# BARS
# =====================================================

bars1 = ax1.bar(
    x - 2*bar_width,
    impl["Cumulative (Baseline Plan)"],
    width=bar_width,
    color="#7ED321",
    label="Cumulative (Baseline Plan)"
)

bars2 = ax1.bar(
    x - bar_width,
    impl["Cumulative Revised basline plan"],
    width=bar_width,
    color="#B8E986",
    label="Cumulative Revised baseline plan"
)

bars3 = ax1.bar(
    x,
    impl["Cumulative (Completed)"],
    width=bar_width,
    color="#00B050",
    label="Cumulative (Completed)"
)

bars4 = ax1.bar(
    x + bar_width,
    impl["Cumulative Rejected / Transferred/ Moved to next quarter"],
    width=bar_width,
    color="#F4B400",
    label="Rejected/Transferred"
)

bars5 = ax1.bar(
    x + 2*bar_width,
    impl[progress_col],
    width=bar_width,
    color="#FFF000",
    label="Impl In Progress"
)

# =====================================================
# SECONDARY AXIS
# =====================================================

ax2 = ax1.twinx()

confidence = to_percentage(impl["% Completion Confidence - Overall"])
actual = to_percentage(impl["% Actual weekly completion wr.t  revised Baseline"])

line1 = ax2.plot(
    x,
    confidence,
    color="black",
    linewidth=3,
    label="% Completion Confidence"
)

line2 = ax2.plot(
    x,
    actual,
    color="#7030A0",
    linewidth=3,
    label="% Actual weekly completion"
)

# =====================================================
# DATA LABELS
# =====================================================

for bars in [bars1, bars2, bars3, bars4, bars5]:
    for bar in bars:
        height = bar.get_height()

        if pd.notna(height) and height > 0:
            ax1.text(
                bar.get_x() + bar.get_width()/2,
                height,
                f"{int(height)}",
                ha="center",
                va="bottom",
                fontsize=9
            )

for i, value in enumerate(confidence):
    ax2.text(
        i,
        value + 2,
        f"{int(value)}%",
        ha="center",
        fontsize=10,
        color="black"
    )

for i, value in enumerate(actual):
    if pd.notna(value):
        ax2.text(
            i,
            value - 8,
            f"{int(value)}%",
            ha="center",
            fontsize=10,
            color="#7030A0"
        )

# =====================================================
# FORMATTING
# =====================================================

ax1.set_title(
    "Q3'26 Implementation",
    fontsize=20,
    fontweight="bold"
)

ax1.set_xticks(x)
ax1.set_xticklabels(impl["Week Label"])

ax1.set_ylabel("DCR Count")
ax2.set_ylabel("Percentage")

ax2.set_ylim(0,120)

ax1.grid(
    axis="y",
    linestyle="--",
    alpha=0.4
)

# =====================================================
# COMBINED LEGEND
# =====================================================

handles1, labels1 = ax1.get_legend_handles_labels()
handles2, labels2 = ax2.get_legend_handles_labels()

ax1.legend(
    handles1 + handles2,
    labels1 + labels2,
    loc="upper center",
    bbox_to_anchor=(0.5,-0.12),
    ncol=3
)

plt.tight_layout()

plt.savefig(
    "implementation_graph.png",
    dpi=300,
    bbox_inches="tight"
)

plt.show()