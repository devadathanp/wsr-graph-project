#!/usr/bin/env python3
"""Generate the Q3'26 Implementation chart PNG from data.xlsm."""

from __future__ import annotations

import argparse

from wsr_charts import save_implementation_chart
from wsr_common import DEFAULT_DATA_FILE


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate implementation status chart")
    parser.add_argument("--data", default=DEFAULT_DATA_FILE, help="Path to Excel workbook")
    parser.add_argument("--output", default="implementation_graph.png", help="Output PNG path")
    args = parser.parse_args()
    path = save_implementation_chart(args.output, data_file=args.data)
    print(f"Saved: {path}")


if __name__ == "__main__":
    main()
