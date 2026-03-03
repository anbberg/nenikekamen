from datetime import datetime, timezone

import pytest

from strava_client import STRAVA_ACTIVITIES_URL, STRAVA_TOKEN_URL, StravaClient


class DummyResponse:
    def __init__(self, status_code=200, json_data=None, ok=True):
        self.status_code = status_code
        self._json_data = json_data or {}
        self.ok = ok
        self.text = str(self._json_data)

    def json(self):
        return self._json_data


def test_refresh_access_token_success(monkeypatch):
    calls = {}

    def fake_post(url, data):
        calls["url"] = url
        calls["data"] = data
        return DummyResponse(
            status_code=200,
            json_data={"access_token": "ACCESS", "refresh_token": "NEW_REFRESH"},
            ok=True,
        )

    monkeypatch.setattr("strava_client.requests.post", fake_post)

    client = StravaClient("cid", "secret", "old_refresh")
    token = client._refresh_access_token()

    assert token == "ACCESS"
    assert client._access_token == "ACCESS"
    assert client.refresh_token == "NEW_REFRESH"
    assert calls["url"] == STRAVA_TOKEN_URL
    assert calls["data"]["grant_type"] == "refresh_token"


def test_fetch_activities_uses_after_and_maps_fields(monkeypatch):
    # Arrange fake token refresh and GET
    def fake_post(url, data):
        return DummyResponse(
            status_code=200,
            json_data={"access_token": "TOKEN"},
            ok=True,
        )

    captured = {}

    def fake_get(url, headers, params):
        captured["url"] = url
        captured["headers"] = headers
        captured["params"] = params
        activities = [
            {
                "id": 1,
                "name": "Run 1",
                "sport_type": "Run",
                "type": "Run",
                "start_date": "2026-01-01T10:00:00Z",
                "start_date_local": "2026-01-01T11:00:00Z",
                "distance": 1000.0,
                "moving_time": 400,
                "elapsed_time": 450,
                "total_elevation_gain": 10.0,
                "average_heartrate": 150.0,
                "max_heartrate": 180.0,
            }
        ]
        return DummyResponse(status_code=200, json_data=activities, ok=True)

    monkeypatch.setattr("strava_client.requests.post", fake_post)
    monkeypatch.setattr("strava_client.requests.get", fake_get)

    client = StravaClient("cid", "secret", "refresh")

    since = datetime(2026, 1, 1, tzinfo=timezone.utc)
    result = client.fetch_activities(since=since, per_page=50)

    # Request details
    assert captured["url"] == STRAVA_ACTIVITIES_URL
    assert captured["headers"]["Authorization"] == "Bearer TOKEN"
    assert captured["params"]["per_page"] == 50
    assert captured["params"]["after"] == int(since.timestamp())

    # Mapped activity
    assert len(result) == 1
    act = result[0]
    assert act["id"] == 1
    assert act["name"] == "Run 1"
    assert act["distance_m"] == 1000.0
    assert act["moving_time_s"] == 400
    assert act["elapsed_time_s"] == 450
    assert act["total_elevation_gain"] == 10.0
    assert act["average_heartrate"] == 150.0
    assert act["max_heartrate"] == 180.0
    assert act["sport_type"] == "Run"
    assert act["strava_url"].endswith("/1")


def test_fetch_activities_retries_on_401(monkeypatch):
    # First GET returns 401, second GET succeeds
    def fake_post(url, data):
        return DummyResponse(
            status_code=200,
            json_data={"access_token": "NEW_TOKEN"},
            ok=True,
        )

    calls = {"get": 0}

    def fake_get(url, headers, params):
        calls["get"] += 1
        if calls["get"] == 1:
            return DummyResponse(status_code=401, json_data={"message": "unauthorized"}, ok=False)
        return DummyResponse(status_code=200, json_data=[], ok=True)

    monkeypatch.setattr("strava_client.requests.post", fake_post)
    monkeypatch.setattr("strava_client.requests.get", fake_get)

    client = StravaClient("cid", "secret", "refresh")
    result = client.fetch_activities()

    assert result == []
    # Should have called GET twice (initial + retry)
    assert calls["get"] == 2

