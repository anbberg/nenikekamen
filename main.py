from __future__ import annotations

import re
import time
import traceback
from datetime import datetime
from typing import Any, List, Optional, Tuple

from config_loader import load_config
from excel_log import LOG_HEADERS, STRAVA_ID_COLUMN_INDEX, run_to_log_row
from graph_auth import get_graph_access_token
from graph_excel import (
    append_table_rows,
    get_log_table_name,
    get_range_values,
    get_table_values,
    workbook_calculate,
)
from strava_client import StravaClient
from telegram_notify import send_telegram_message


def _parse_training_start_date(value: str) -> datetime:
    try:
        return datetime.fromisoformat(value)
    except Exception as exc:
        raise RuntimeError(f"Invalid TRAINING_START_DATE: {value!r}") from exc


def _activity_week(activity: dict) -> Optional[int]:
    """Return ISO week as YYYYWW for the activity's start, or None."""
    raw = activity.get("start_date_local") or activity.get("start_date")
    if not raw:
        return None
    try:
        if isinstance(raw, datetime):
            dt = raw
        elif isinstance(raw, str):
            s = raw.replace("Z", "+00:00") if raw.endswith("Z") else raw
            dt = datetime.fromisoformat(s)
        else:
            return None
        y, w, _ = dt.isocalendar()
        return y * 100 + w
    except Exception:
        return None


def _week_summary_from_activities(activities: List[dict]) -> str:
    """Build a short 'Denna vecka: N pass, X km' from activities in the current ISO week."""
    if not activities:
        return ""
    y, w, _ = datetime.now().isocalendar()
    current_week = y * 100 + w
    this_week = [a for a in activities if _activity_week(a) == current_week]
    if not this_week:
        return ""
    count = len(this_week)
    distance_km = sum((a.get("distance_m") or 0) / 1000.0 for a in this_week)
    return f"Denna vecka: {count} pass, {round(distance_km, 1)} km"


def _cell_to_row_col(cell: str) -> Tuple[int, int]:
    """Parse A1-style cell to (row, col) 1-based. Supports A-Z and AA, AB, ..."""
    m = re.match(r"^([A-Za-z]+)(\d+)$", cell.strip())
    if not m:
        raise ValueError(f"Invalid cell address: {cell!r}")
    letters, row_s = m.group(1).upper(), m.group(2)
    row = int(row_s)
    col = 0
    for c in letters:
        col = col * 26 + (ord(c) - ord("A") + 1)
    return (row, col)


def _col_to_letter(n: int) -> str:
    """1-based column number to Excel letter(s). 1->A, 26->Z, 27->AA."""
    letters = ""
    while n:
        n, r = divmod(n - 1, 26)
        letters = chr(65 + r) + letters
    return letters


def _build_current_week_plan_agg_summary(
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


def _build_plan_summary(
    access_token: str,
    excel_path: str,
    sheet_name: str,
    summary_cells: str,
) -> Optional[str]:
    """
    Parse summary_cells (e.g. "B2:Vecka,B3:Återstående volym (min)"),
    GET the range values from the sheet, return a single summary line or None on error.
    """
    if not sheet_name or not summary_cells:
        return None
    parts = [p.strip() for p in summary_cells.split(",") if p.strip()]
    if not parts:
        return None
    cells_and_labels: List[Tuple[str, str]] = []
    for p in parts:
        idx = p.find(":")
        if idx <= 0:
            continue
        cell, label = p[:idx].strip(), p[idx + 1 :].strip()
        if cell and label:
            cells_and_labels.append((cell, label))
    if not cells_and_labels:
        return None
    try:
        rows_cols = [_cell_to_row_col(c) for c, _ in cells_and_labels]
        min_row = min(r for r, _ in rows_cols)
        max_row = max(r for r, _ in rows_cols)
        min_col = min(c for _, c in rows_cols)
        max_col = max(c for _, c in rows_cols)
        range_address = f"{_col_to_letter(min_col)}{min_row}:{_col_to_letter(max_col)}{max_row}"
    except (ValueError, KeyError):
        return None
    try:
        values = get_range_values(access_token, excel_path, sheet_name, range_address)
    except Exception as e:
        print(f"Kunde inte läsa Plan+Agg: {e}")
        return None
    if not values:
        return None
    result_parts = []
    for (cell, label), (row, col) in zip(cells_and_labels, rows_cols):
        r_off = row - min_row
        c_off = col - min_col
        if r_off < len(values) and c_off < len(values[r_off]):
            val = values[r_off][c_off]
            if val is not None and str(val).strip() != "":
                result_parts.append(f"{label}: {val}")
    if not result_parts:
        return None
    return ". ".join(result_parts)


def main() -> None:
    config = load_config()

    graph_cfg = config["graph"]
    strava_cfg = config["strava"]
    runtime_cfg = config["runtime"]
    telegram_cfg = config.get("telegram", {})
    plan_cfg = config.get("plan_summary", {})
    bot_token = (telegram_cfg.get("bot_token") or "").strip()
    chat_id = (telegram_cfg.get("chat_id") or "").strip()
    telegram_configured = bool(bot_token and chat_id)

    try:
        token_cache_path = graph_cfg.get("token_cache_path", "graph_token_cache.bin")
        training_start_str = runtime_cfg["training_start_date"]
        training_start = _parse_training_start_date(training_start_str)

        strava_client = StravaClient(
            client_id=strava_cfg["client_id"],
            client_secret=strava_cfg["client_secret"],
            refresh_token=strava_cfg["refresh_token"],
        )

        print(
            f"Fetching all activities from Strava since {training_start.isoformat()}..."
        )
        activities = strava_client.fetch_activities(since=training_start)
        print(f"Found {len(activities)} activities")

        print("Acquiring Microsoft Graph access token...")
        token = get_graph_access_token(
            client_id=graph_cfg["client_id"],
            tenant_id=graph_cfg["tenant_id"],
            scopes=graph_cfg["scopes"],
            token_cache_path=token_cache_path,
        )

        excel_path = graph_cfg["excel_path"]
        log_sheet_name = graph_cfg.get("log_sheet_name", "Logg")

        table_name = get_log_table_name(token, excel_path, log_sheet_name)
        table_values = get_table_values(token, excel_path, table_name)
        if table_values and len(table_values) > 0 and len(table_values[0]) != len(LOG_HEADERS):
            raise RuntimeError(
                f"Tabellen '{table_name}' har {len(table_values[0])} kolumner men scriptet "
                f"förväntar sig {len(LOG_HEADERS)}. Kolumnerna ska heta (i ordning): "
                + ", ".join(LOG_HEADERS)
            )
        # First row is header; strava_id is at STRAVA_ID_COLUMN_INDEX
        existing_ids = set()
        for row in table_values[1:]:
            if len(row) > STRAVA_ID_COLUMN_INDEX and row[STRAVA_ID_COLUMN_INDEX] is not None:
                existing_ids.add(row[STRAVA_ID_COLUMN_INDEX])

        new_runs = [r for r in activities if r.get("id") not in existing_ids]
        if new_runs:
            values_2d = [run_to_log_row(r) for r in new_runs]
            print(
                f"Appending {len(new_runs)} activities to table '{table_name}'..."
            )
            # Append in batches to reduce risk of InsertDeleteConflict
            batch_size = 15
            for i in range(0, len(values_2d), batch_size):
                chunk = values_2d[i : i + batch_size]
                append_table_rows(token, excel_path, table_name, chunk)
                if i + batch_size < len(values_2d):
                    time.sleep(0.5)
            try:
                workbook_calculate(token, excel_path)
            except Exception as e:
                print(f"Calculate (valfritt) misslyckades: {e}")
            wait_sec = plan_cfg.get("wait_seconds", 3)
            time.sleep(wait_sec)
        else:
            print("No new activities to append.")

        # Sammanfattning från Plan+Agg (nuvarande veckas rad) om arket är konfigurerat
        summary_str = None
        if plan_cfg.get("sheet_name"):
            summary_str = _build_current_week_plan_agg_summary(
                token,
                excel_path,
                plan_cfg["sheet_name"].strip(),
            )
        if summary_str is None:
            summary_str = _week_summary_from_activities(activities)

        print("Done.")
        if telegram_configured:
            msg = "Nenikekamen: Körningen lyckades."
            if summary_str:
                msg += "\n\n" + summary_str
            send_telegram_message(bot_token, chat_id, msg)
        else:
            print("Telegram not configured. No message is sent.")

    except Exception as exc:
        if telegram_configured:
            err_msg = f"Nenikekamen: Körningen kraschade.\n\n{type(exc).__name__}: {exc}"
            tb = traceback.format_exc()
            if tb:
                err_msg += "\n\n" + (tb[-500:] if len(tb) > 500 else tb)
            send_telegram_message(bot_token, chat_id, err_msg)
        raise


if __name__ == "__main__":
    main()
