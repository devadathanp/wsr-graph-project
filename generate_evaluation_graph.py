#!/usr/bin/env python3
"""Generate the Q3'26 Evaluation chart PNG from data.xlsm."""

from __future__ import annotations

import argparse

from wsr_charts import save_evaluation_chart
from wsr_common import DEFAULT_DATA_FILE


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate evaluation status chart")
    parser.add_argument("--data", default=DEFAULT_DATA_FILE, help="Path to Excel workbook")
    parser.add_argument("--output", default="Evaluation_graph.png", help="Output PNG path")
    args = parser.parse_args()
    path = save_evaluation_chart(args.output, data_file=args.data)
    print(f"Saved: {path}")


if __name__ == "__main__":
    main()
