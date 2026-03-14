"""
Shared logic for sync and analyse jobs.
"""

from __future__ import annotations

from datetime import datetime
from typing import Optional

from .graph_excel import get_range_values


def format_hours(decimal_hours: Optional[float]) -> str:
    """
    Convert decimal hours (e.g. 1.2802777) to a readable string (e.g. "1 h 17 min").
    If 0 or None, returns "0 min".
    """
    if decimal_hours is None:
        return "0 min"
    h = float(decimal_hours)
    if h <= 0:
        return "0 min"
    hours = int(h)
    minutes = round((h - hours) * 60)
    if hours == 0:
        return f"{minutes} min"
    return f"{hours} h {minutes} min"


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
        week_values = get_range_values(access_token, excel_path, sheet_name, "B2:B60")
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

        # Read Fokus from column D
        fokus_cell = get_range_values(
            access_token, excel_path, sheet_name, f"D{excel_row}:D{excel_row}"
        )
        fokus = ""
        if fokus_cell and fokus_cell[0] and fokus_cell[0][0] is not None:
            fokus = str(fokus_cell[0][0]).strip()

        # Read Utfall for this row: L (count), M (volume_h), N (volume_km), O (long_run_km),
        # P (remaining_count), Q (remaining_hours), R (long_run_offset), S (long_run_status), T (Kommentar)
        range_addr = f"L{excel_row}:T{excel_row}"
        row_values = get_range_values(access_token, excel_path, sheet_name, range_addr)
        if not row_values or not row_values[0]:
            return None
        v = row_values[0]
        volume_h = v[1] if len(v) > 1 else None
        volume_km = v[2] if len(v) > 2 else None
        long_run_km = v[3] if len(v) > 3 else None
        remaining_hours = v[5] if len(v) > 5 else None
        long_run_status = v[7] if len(v) > 7 else None

        # Format values for display
        try:
            km_val = round(float(volume_km), 1) if volume_km is not None else 0
        except (TypeError, ValueError):
            km_val = 0
        time_done = format_hours(volume_h)
        time_left = format_hours(remaining_hours)

        # Build message: Vecka | Fokus, Avklarat, Kvarstår, optional Långpass kvar
        lines = [
            f"📅 **Vecka {w} | Fokus: {fokus or '–'}**",
            f"✅ **Avklarat:** {km_val} km ({time_done})",
            f"⏳ **Kvarstår:** {time_left} volym",
        ]
        # Show long-run line only when status indicates not completed (e.g. NOT_COMPLETED); hide when OK
        status_str = (
            str(long_run_status).strip().upper() if long_run_status is not None else ""
        )
        if status_str != "OK" and long_run_km is not None:
            try:
                lr_km = round(float(long_run_km), 1)
                lines.append(f"🏃‍♂️ **Långpass kvar:** {lr_km} km")
            except (TypeError, ValueError):
                lines.append("🏃‍♂️ **Långpass kvar:** – km")
        return "\n".join(lines)
    except Exception as e:
        print(f"Kunde inte läsa Plan+Agg för nuvarande vecka: {e}")
        return None
