from excel_log import (
    LOG_HEADERS,
    STRAVA_ID_COLUMN_INDEX,
    _parse_iso_datetime,
    run_to_log_row,
)


def test_parse_iso_datetime_parses_z_and_offset():
    z_ts = "2026-01-01T10:00:00Z"
    offset_ts = "2026-01-01T11:00:00+01:00"
    invalid = "not-a-timestamp"

    parsed_z = _parse_iso_datetime(z_ts)
    parsed_offset = _parse_iso_datetime(offset_ts)
    parsed_invalid = _parse_iso_datetime(invalid)

    assert parsed_z is not None
    assert parsed_offset is not None
    assert parsed_z == parsed_offset
    assert parsed_invalid is None


def test_run_to_log_row_produces_log_headers_order():
    run = {
        "id": 123,
        "name": "Test Run",
        "start_date_local": "2026-01-02T09:10:21Z",
        "distance_m": 5000.0,
        "moving_time_s": 1500,
        "total_elevation_gain": 42.0,
        "average_heartrate": 140.0,
        "max_heartrate": 170.0,
        "sport_type": "Run",
        "strava_url": "https://www.strava.com/activities/123",
    }
    row = run_to_log_row(run)
    assert len(row) == len(LOG_HEADERS)
    # start_datetime as ISO string
    assert row[0].startswith("2026-01-02")
    assert row[1] == 202601  # week_number
    assert row[2] == 5.0  # distance_km
    assert row[3] == 1500  # duration_s
    assert row[4] == 12.0  # speed_km_h
    assert row[5] == 42.0  # elevation_m
    assert row[6] == 140.0  # avg_heart_rate
    assert row[7] == 170.0  # max_heart_rate
    assert row[8] == "Run"  # sport_type
    assert row[9] == 123  # strava_id
    assert row[10] == "https://www.strava.com/activities/123"
    assert row[11] == "Test Run"


def test_strava_id_column_index_matches_headers():
    assert LOG_HEADERS[STRAVA_ID_COLUMN_INDEX] == "strava_id"
