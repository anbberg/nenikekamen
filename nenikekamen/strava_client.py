from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

import requests

STRAVA_TOKEN_URL = "https://www.strava.com/oauth/token"
STRAVA_ACTIVITIES_URL = "https://www.strava.com/api/v3/athlete/activities"


class StravaClient:
    def __init__(self, client_id: str, client_secret: str, refresh_token: str) -> None:
        self.client_id = client_id
        self.client_secret = client_secret
        self.refresh_token = refresh_token
        self._access_token: Optional[str] = None

    def _refresh_access_token(self) -> str:
        payload = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "refresh_token",
            "refresh_token": self.refresh_token,
        }
        resp = requests.post(STRAVA_TOKEN_URL, data=payload)
        if not resp.ok:
            try:
                err = resp.json()
                msg = err.get("message", resp.text)
            except Exception:
                msg = resp.text
            raise RuntimeError(
                f"Strava token refresh failed ({resp.status_code}): {msg}. "
                "Check STRAVA_CLIENT_ID, STRAVA_CLIENT_SECRET, and STRAVA_REFRESH_TOKEN. "
                "You may need to re-authorize and get a new refresh token (see README)."
            ) from None
        data = resp.json()
        self._access_token = data["access_token"]
        # Strava can rotate refresh_token, so update in memory if provided.
        if "refresh_token" in data:
            self.refresh_token = data["refresh_token"]
        return self._access_token

    def _get_access_token(self) -> str:
        if not self._access_token:
            return self._refresh_access_token()
        return self._access_token

    def fetch_activities(
        self,
        since: Optional[datetime] = None,
        per_page: int = 200,
    ) -> List[Dict[str, Any]]:
        """
        Fetch all activities from Strava since the given datetime (UTC).
        No sport-type filter; filter in Excel if you only want runs etc.
        """
        token = self._get_access_token()

        params: Dict[str, Any] = {"per_page": per_page}
        if since is not None:
            if since.tzinfo is None:
                since = since.replace(tzinfo=timezone.utc)
            params["after"] = int(since.timestamp())

        headers = {"Authorization": f"Bearer {token}"}
        resp = requests.get(STRAVA_ACTIVITIES_URL, headers=headers, params=params)

        if resp.status_code == 401:
            # Token may have been rejected (wrong scope or expired). Try refresh once and retry.
            self._access_token = None
            token = self._get_access_token()
            headers = {"Authorization": f"Bearer {token}"}
            resp = requests.get(STRAVA_ACTIVITIES_URL, headers=headers, params=params)

        if not resp.ok:
            try:
                err = resp.json()
                msg = err.get("message", resp.text)
            except Exception:
                msg = resp.text
            raise RuntimeError(
                f"Strava API error ({resp.status_code}): {msg}. "
                "If 401: re-authorize the app and request scope 'activity:read_all', "
                "then put the new refresh_token in your .env (see README)."
            ) from None

        activities = resp.json()

        result: List[Dict[str, Any]] = []
        for act in activities:
            sport_type = act.get("sport_type") or act.get("type")
            start_date = act.get("start_date")
            start_date_local = act.get("start_date_local")
            distance_m = act.get("distance", 0.0)
            moving_time_s = act.get("moving_time", 0)
            elapsed_time_s = act.get("elapsed_time", 0)
            total_elevation_gain = act.get("total_elevation_gain", 0.0)
            average_heartrate = act.get("average_heartrate")
            max_heartrate = act.get("max_heartrate")

            result.append(
                {
                    "id": act.get("id"),
                    "name": act.get("name"),
                    "start_date": start_date,
                    "start_date_local": start_date_local,
                    "distance_m": distance_m,
                    "moving_time_s": moving_time_s,
                    "elapsed_time_s": elapsed_time_s,
                    "total_elevation_gain": total_elevation_gain,
                    "average_heartrate": average_heartrate,
                    "max_heartrate": max_heartrate,
                    "type": act.get("type"),
                    "sport_type": sport_type,
                    "strava_url": f"https://www.strava.com/activities/{act.get('id')}",
                }
            )

        return result
