"""Fiscal calendar used on WSR headings.

The fiscal year starts on 1 November (not 1 January):

  Q1  Nov, Dec, Jan
  Q2  Feb, Mar, Apr
  Q3  May, Jun, Jul
  Q4  Aug, Sep, Oct

The fiscal year number is the calendar year in which that year ends (31 Oct).
Week number is the ISO week of the report / system date.
"""

from __future__ import annotations

from datetime import datetime

import pandas as pd


def parse_report_date(value) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    parsed = pd.to_datetime(value, dayfirst=True)
    if pd.isna(parsed):
        raise ValueError(f"Could not parse report date: {value!r}")
    return parsed.to_pydatetime()


def fiscal_quarter_and_year(value) -> tuple[int, int]:
    dt = parse_report_date(value)
    month = int(dt.month)
    if month >= 11:
        return 1, int(dt.year) + 1
    if month == 1:
        return 1, int(dt.year)
    if month <= 4:
        return 2, int(dt.year)
    if month <= 7:
        return 3, int(dt.year)
    return 4, int(dt.year)


def iso_week_number(value) -> int:
    dt = parse_report_date(value)
    return int(dt.isocalendar()[1])


def quarter_label_long(value) -> str:
    quarter, year = fiscal_quarter_and_year(value)
    return f"Q{quarter}-{year}"


def quarter_label_short(value) -> str:
    quarter, year = fiscal_quarter_and_year(value)
    return f"Q{quarter}'{str(year)[2:]}"


def fiscal_year(value) -> int:
    return fiscal_quarter_and_year(value)[1]
