"""core/reader.py — Reads and validates attendance Excel files."""

import pandas as pd
from datetime import datetime, time
from pathlib import Path
from typing import Tuple


def _try_parse_time(val, formats: list) -> time | None:
    if pd.isna(val):
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, datetime):
        return val.time()
    s = str(val).strip()
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    return None


def _try_parse_date(val, formats: list):
    if pd.isna(val):
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return pd.Timestamp(val).date()
    s = str(val).strip()
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def load_file(path: str | Path, settings: dict) -> Tuple[pd.DataFrame, list]:
    """
    Load an attendance Excel file.
    Returns (dataframe, list_of_warnings).
    """
    path = Path(path)
    warnings = []

    try:
        df = pd.read_excel(path, dtype=str)
    except Exception as e:
        raise ValueError(f"Cannot read file: {e}")

    # Strip whitespace from column names
    df.columns = [str(c).strip() for c in df.columns]

    col_map = settings["excel"]["expected_columns"]
    date_fmts = settings["excel"]["date_formats"]
    time_fmts = settings["excel"]["time_formats"]

    # Flexible column matching (case-insensitive)
    rename = {}
    for internal, expected in col_map.items():
        for col in df.columns:
            if col.lower() == expected.lower():
                rename[col] = internal
                break
        else:
            warnings.append(f"Column not found: '{expected}' — using defaults where possible.")

    df = df.rename(columns=rename)

    # Parse date column
    if "date" in df.columns:
        df["date"] = df["date"].apply(lambda v: _try_parse_date(v, date_fmts))

    # Parse time columns
    for tcol in ("time_in", "time_out"):
        if tcol in df.columns:
            df[tcol] = df[tcol].apply(lambda v: _try_parse_time(v, time_fmts))

    # Normalize employee_name
    if "employee_name" in df.columns:
        df["employee_name"] = df["employee_name"].str.strip()

    df = df.dropna(subset=["date"])
    df = df.reset_index(drop=True)

    return df, warnings


def validate_dataframe(df: pd.DataFrame) -> list:
    """Return list of data quality issues."""
    issues = []
    required = ["employee_name", "date"]
    for col in required:
        if col not in df.columns:
            issues.append(f"Missing required column: {col}")
    if "time_in" in df.columns and df["time_in"].isna().any():
        n = df["time_in"].isna().sum()
        issues.append(f"{n} row(s) have missing Time In values.")
    if "time_out" in df.columns and df["time_out"].isna().any():
        n = df["time_out"].isna().sum()
        issues.append(f"{n} row(s) have missing Time Out values.")
    return issues