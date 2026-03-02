from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet


LOG_HEADERS = [
    "Datum",
    "Distans (km)",
    "Tid (min:sek)",
    "Pace (min/km)",
    "Höjd (m)",
    "Puls snitt",
    "Puls max",
    "Typ",
    "Strava ID",
    "Strava länk",
    "Namn",
]


def _get_or_create_log_sheet(wb_path: Path, sheet_name: str) -> Worksheet:
    if wb_path.exists():
        wb = load_workbook(wb_path)
    else:
        wb = Workbook()

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(title=sheet_name)

    # Ensure headers on first row
    if ws.max_row == 0 or all(cell.value is None for cell in ws[1]):
        for idx, header in enumerate(LOG_HEADERS, start=1):
            ws.cell(row=1, column=idx, value=header)

    wb.save(wb_path)
    return ws


def _parse_iso_datetime(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        # Strava returns ISO timestamps, e.g. "2023-01-01T10:00:00Z"
        if value.endswith("Z"):
            value = value.replace("Z", "+00:00")
        return datetime.fromisoformat(value)
    except Exception:
        return None


def _get_existing_strava_ids(path: Path, sheet_name: str) -> Set[Any]:
    """
    Return a set of Strava IDs that are already present in the log sheet.
    """
    if not path.exists():
        return set()

    wb = load_workbook(path, read_only=True)
    if sheet_name not in wb.sheetnames:
        return set()

    ws = wb[sheet_name]
    ids: Set[Any] = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        # Some rows may have fewer columns if the sheet was edited manually.
        # Guard against short tuples before accessing index 8 (9th column).
        if len(row) <= 8:
            continue
        strava_id = row[8]  # index 8 -> 9th column
        if strava_id is not None:
            ids.add(strava_id)
    return ids


def _format_duration(seconds: int) -> str:
    minutes, sec = divmod(int(seconds), 60)
    return f"{minutes:d}:{sec:02d}"


def _format_pace(distance_km: float, moving_time_s: int) -> Optional[str]:
    if distance_km <= 0 or moving_time_s <= 0:
        return None
    pace_s_per_km = moving_time_s / distance_km
    minutes, sec = divmod(int(pace_s_per_km), 60)
    return f"{minutes:d}:{sec:02d}"


def append_runs_to_log(
    path: Path,
    sheet_name: str,
    runs: Iterable[Dict[str, Any]],
) -> None:
    """
    Append Strava run data as rows to the log sheet, skipping runs
    whose Strava ID already exists in the log.
    """
    # Ensure workbook and sheet exist
    _get_or_create_log_sheet(path, sheet_name)
    existing_ids = _get_existing_strava_ids(path, sheet_name)

    wb = load_workbook(path)
    ws = wb[sheet_name]

    for run in runs:
        strava_id = run.get("id")
        if strava_id in existing_ids:
            continue

        start_dt_local = _parse_iso_datetime(run.get("start_date_local") or run.get("start_date"))
        distance_km = (run.get("distance_m") or 0.0) / 1000.0
        moving_time_s = int(run.get("moving_time_s") or 0)
        total_elev = run.get("total_elevation_gain")
        avg_hr = run.get("average_heartrate")
        max_hr = run.get("max_heartrate")
        run_type = run.get("sport_type") or run.get("type") or "Run"
        strava_url = run.get("strava_url")
        name = run.get("name")

        duration_str = _format_duration(moving_time_s) if moving_time_s else None
        pace_str = _format_pace(distance_km, moving_time_s)

        row = [
            start_dt_local.date().isoformat() if start_dt_local else None,
            round(distance_km, 2) if distance_km else None,
            duration_str,
            pace_str,
            total_elev,
            avg_hr,
            max_hr,
            run_type,
            strava_id,
            strava_url,
            name,
        ]

        ws.append(row)
        existing_ids.add(strava_id)

    wb.save(path)
