"""
Log row format for Strava runs. Used when appending to the Logg sheet via Graph API.
"""
from datetime import datetime
from typing import Any, Dict, List, Optional

LOG_HEADERS = [
    "start_datetime",
    "week_number",
    "distance_km",
    "duration_s",
    "speed_km_h",
    "elevation_m",
    "avg_heart_rate",
    "max_heart_rate",
    "sport_type",
    "strava_id",
    "strava_url",
    "name",
]

# 0-based index of strava_id in a log row (for extracting existing IDs from usedRange values).
STRAVA_ID_COLUMN_INDEX = 9


def _parse_iso_datetime(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        if value.endswith("Z"):
            value = value.replace("Z", "+00:00")
        return datetime.fromisoformat(value)
    except Exception:
        return None


def run_to_log_row(run: Dict[str, Any]) -> List[Any]:
    """
    Build a single log row (list of values in LOG_HEADERS order) from a Strava run dict.
    Suitable for PATCH range via Graph API. Datetimes are returned as ISO strings.
    """
    start_dt_local = _parse_iso_datetime(
        run.get("start_date_local") or run.get("start_date")
    )
    distance_km = (run.get("distance_m") or 0.0) / 1000.0
    moving_time_s = int(run.get("moving_time_s") or 0)
    total_elev = run.get("total_elevation_gain")
    avg_hr = run.get("average_heartrate")
    max_hr = run.get("max_heartrate")
    run_type = run.get("sport_type") or run.get("type") or "Run"
    strava_url = run.get("strava_url")
    name = run.get("name")
    strava_id = run.get("id")

    duration_s: Optional[int] = moving_time_s or None
    speed_km_h: Optional[float] = None
    if distance_km > 0 and moving_time_s > 0:
        speed_km_h = round(distance_km / (moving_time_s / 3600.0), 2)

    # ISO string for Graph API (Excel accepts ISO for date/time cells).
    start_dt_excel: Optional[str] = None
    if start_dt_local:
        start_dt_excel = start_dt_local.replace(tzinfo=None).isoformat()

    week_number: Optional[int] = None
    if start_dt_local:
        y, w, _ = start_dt_local.isocalendar()
        week_number = y * 100 + w

    return [
        start_dt_excel,
        week_number,
        round(distance_km, 2) if distance_km else None,
        duration_s,
        speed_km_h,
        total_elev,
        avg_hr,
        max_hr,
        run_type,
        strava_id,
        strava_url,
        name,
    ]
