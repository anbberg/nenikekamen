"""
Shared logic for sync and analyse jobs.
"""
from __future__ import annotations

from datetime import datetime
from typing import Optional

from .graph_excel import get_range_values


def parse_training_start_date(value: str) -> datetime:
    try:
        return datetime.fromisoformat(value)
    except Exception as exc:
        raise RuntimeError(f"Invalid TRAINING_START_DATE: {value!r}") from exc


def build_current_week_plan_agg_summary(
    access_token: str,
    excel_path: str,
    sheet_name: str,
) -> Optional[str]:
    """
    Find the row for current ISO week in Plan+Agg (column B = Vecka),
    read that row's Utfall columns (L–T), return a short status line.
    """
    if not sheet_name or not sheet_name.strip():
        return None
    sheet_name = sheet_name.strip()
    y, w, _ = datetime.now().isocalendar()
    current_week = y * 100 + w  # YYYYWW
    try:
        # Column B = Vecka (week numbers). Read B2:B60 to find current week row.
        week_values = get_range_values(
            access_token, excel_path, sheet_name, "B2:B60"
        )
        if not week_values:
            return None
        excel_row = None
        for i, row in enumerate(week_values):
            if not row:
                continue
            val = row[0]
            if val is None:
                continue
            # Compare as number or string
            if val == current_week or str(val).strip() == str(current_week):
                excel_row = 2 + i  # 1-based Excel row (first data row = 2)
                break
        if excel_row is None:
            return None
        # Read Utfall for this row: L (count), M (volume_h), N (volume_km), O (long_run_km),
        # P (remaining_count), Q (remaining_hours), R (long_run_offset), S (long_run_status), T (Kommentar)
        range_addr = f"L{excel_row}:T{excel_row}"
        row_values = get_range_values(
            access_token, excel_path, sheet_name, range_addr
        )
        if not row_values or not row_values[0]:
            return None
        v = row_values[0]
        count = v[0] if len(v) > 0 else None
        volume_h = v[1] if len(v) > 1 else None
        volume_km = v[2] if len(v) > 2 else None
        long_run_km = v[3] if len(v) > 3 else None
        remaining_count = v[4] if len(v) > 4 else None
        remaining_hours = v[5] if len(v) > 5 else None
        long_run_offset = v[6] if len(v) > 6 else None
        long_run_status = v[7] if len(v) > 7 else None
        kommentar = v[8] if len(v) > 8 and v[8] else None
        parts = [f"Vecka {current_week}:"]
        if count is not None:
            parts.append(f" {count} pass")
        if volume_km is not None:
            parts.append(f", {volume_km} km")
        if volume_h is not None:
            parts.append(f" ({volume_h} h)")
        if remaining_count is not None or remaining_hours is not None:
            parts.append(". Återstående:")
            if remaining_count is not None:
                parts.append(f" {remaining_count} pass")
            if remaining_hours is not None:
                parts.append(f", {remaining_hours} h")
        if long_run_status is not None and str(long_run_status).strip():
            parts.append(f". Långpass: {long_run_status}")
        elif long_run_offset is not None:
            parts.append(f". Långpass: {long_run_offset}")
        if kommentar:
            parts.append(f" ({kommentar})")
        return "".join(parts).strip()
    except Exception as e:
        print(f"Kunde inte läsa Plan+Agg för nuvarande vecka: {e}")
        return None
