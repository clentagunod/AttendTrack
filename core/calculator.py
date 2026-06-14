"""core/calculator.py — Computes attendance metrics from raw data."""

import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta


def _time_to_minutes(t: time | None) -> float | None:
    if t is None:
        return None
    return t.hour * 60 + t.minute + t.second / 60


def _minutes_to_hours(m: float | None) -> float:
    if m is None:
        return 0.0
    return round(m / 60, 2)


def _normalize_status(val, status_labels: dict) -> str:
    if pd.isna(val):
        return "unknown"
    v = str(val).strip().lower()
    for canonical, variants in status_labels.items():
        if v in [x.lower() for x in variants]:
            return canonical
    return v


def compute_records(df: pd.DataFrame, settings: dict) -> pd.DataFrame:
    """
    Enrich the raw dataframe with computed columns:
      - hours_worked, is_late, late_minutes, is_absent,
        is_half_day, overtime_hours, status_normalized
    """
    att = settings["attendance"]
    sl  = settings["status_labels"]

    work_start   = datetime.strptime(att["work_start_time"], "%H:%M").time()
    late_thresh  = att["late_threshold_minutes"]
    break_mins   = att["break_duration_minutes"]
    ot_threshold = att["overtime_threshold_hours"] * 60  # in minutes

    start_min = _time_to_minutes(work_start)

    rows = []
    for _, row in df.iterrows():
        r = row.to_dict()

        ti = r.get("time_in")
        to = r.get("time_out")
        status_raw = r.get("status", None)

        status = _normalize_status(status_raw, sl)
        r["status_normalized"] = status

        # Hours worked
        if ti is not None and to is not None:
            in_min  = _time_to_minutes(ti)
            out_min = _time_to_minutes(to)
            if out_min > in_min:
                worked = out_min - in_min - break_mins
                r["hours_worked"] = max(0.0, _minutes_to_hours(worked))
            else:
                r["hours_worked"] = 0.0
        else:
            r["hours_worked"] = 0.0

        # Late arrival
        if ti is not None and start_min is not None:
            in_min = _time_to_minutes(ti)
            diff   = in_min - start_min
            r["late_minutes"] = max(0.0, round(diff, 1))
            r["is_late"]      = diff > late_thresh
        else:
            r["late_minutes"] = 0.0
            r["is_late"]      = False

        # Absence
        r["is_absent"]   = status in ("absent",)
        r["is_leave"]    = status in ("leave",)
        r["is_holiday"]  = status in ("holiday",)
        r["is_half_day"] = status in ("half_day",) or (
            0 < r["hours_worked"] < att["half_day_hours"]
        )

        # Overtime
        worked_min = r["hours_worked"] * 60
        r["overtime_hours"] = max(0.0, _minutes_to_hours(worked_min - ot_threshold))

        rows.append(r)

    return pd.DataFrame(rows)


def summarize_by_employee(df: pd.DataFrame) -> pd.DataFrame:
    """Return one-row-per-employee summary."""
    if df.empty:
        return pd.DataFrame()

    grp = df.groupby("employee_name", sort=True)

    summary = pd.DataFrame({
        "Employee":        grp["employee_name"].first(),
        "Department":      grp["department"].first() if "department" in df.columns else "—",
        "Days Present":    grp.apply(lambda g: (~g["is_absent"] & ~g["is_leave"] & ~g["is_holiday"]).sum()),
        "Days Absent":     grp["is_absent"].sum(),
        "Days Leave":      grp["is_leave"].sum(),
        "Days Late":       grp["is_late"].sum(),
        "Total Hours":     grp["hours_worked"].sum().round(1),
        "Avg Hours/Day":   grp["hours_worked"].mean().round(2),
        "Overtime Hours":  grp["overtime_hours"].sum().round(1),
        "Avg Late (min)":  grp.apply(lambda g: g.loc[g["is_late"], "late_minutes"].mean()).round(1).fillna(0),
    }).reset_index(drop=True)

    return summary


def summarize_by_department(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "department" not in df.columns:
        return pd.DataFrame()

    grp = df.groupby("department", sort=True)

    summary = pd.DataFrame({
        "Department":      grp["department"].first(),
        "Headcount":       grp["employee_name"].nunique(),
        "Avg Hours/Day":   grp["hours_worked"].mean().round(2),
        "Total Absences":  grp["is_absent"].sum(),
        "Total Late":      grp["is_late"].sum(),
        "Total Overtime":  grp["overtime_hours"].sum().round(1),
    }).reset_index(drop=True)

    return summary


def overall_kpis(df: pd.DataFrame) -> dict:
    """Return headline KPI numbers for the dashboard."""
    if df.empty:
        return {}

    total_records  = len(df)
    total_employees = df["employee_name"].nunique() if "employee_name" in df.columns else 0
    total_present  = int((~df["is_absent"] & ~df["is_leave"] & ~df["is_holiday"]).sum())
    total_absent   = int(df["is_absent"].sum())
    total_late     = int(df["is_late"].sum())
    avg_hours      = round(float(df["hours_worked"].mean()), 2)
    total_overtime = round(float(df["overtime_hours"].sum()), 1)
    attendance_rate = round(total_present / total_records * 100, 1) if total_records else 0.0
    punctuality_rate = round((total_records - total_late) / total_records * 100, 1) if total_records else 0.0

    return {
        "total_employees":   total_employees,
        "total_records":     total_records,
        "total_present":     total_present,
        "total_absent":      total_absent,
        "total_late":        total_late,
        "avg_hours":         avg_hours,
        "total_overtime":    total_overtime,
        "attendance_rate":   attendance_rate,
        "punctuality_rate":  punctuality_rate,
    }