from __future__ import annotations

import traceback
from datetime import datetime
from pathlib import Path

from config_loader import load_config
from telegram_notify import send_telegram_message
from excel_log import append_runs_to_log
from graph_auth import get_graph_access_token
from graph_excel import download_excel_file, upload_excel_file
from strava_client import StravaClient


def _parse_training_start_date(value: str) -> datetime:
    """
    Parse the TRAINING_START_DATE from config (ISO-format).
    Treats naive datetimes as local time; for Strava's 'after' filter,
    that is sufficient for the fixed training window.
    """
    try:
        return datetime.fromisoformat(value)
    except Exception as exc:
        raise RuntimeError(f"Invalid TRAINING_START_DATE: {value!r}") from exc


def main() -> None:
    config = load_config()

    graph_cfg = config["graph"]
    strava_cfg = config["strava"]
    runtime_cfg = config["runtime"]
    telegram_cfg = config.get("telegram", {})
    bot_token = (telegram_cfg.get("bot_token") or "").strip()
    chat_id = (telegram_cfg.get("chat_id") or "").strip()
    telegram_configured = bool(bot_token and chat_id)

    try:
        # Resolve paths
        temp_excel_path = Path(
            runtime_cfg.get("temp_excel_path", "local_training.xlsx")
        )
        token_cache_path = graph_cfg.get("token_cache_path", "graph_token_cache.bin")

        # Fixed training window start (same every run)
        training_start_str = runtime_cfg["training_start_date"]
        training_start = _parse_training_start_date(training_start_str)

        # Prepare Strava client
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
        if not activities:
            print("No activities found in the configured training window.")
            if telegram_configured:
                print("Sending Telegram message")
                send_telegram_message(
                    bot_token, chat_id, "Nenikekamen: Körningen lyckades."
                )
            else:
                print("Telegram not configured. No message is sent.")
            return

        # Get Graph access token
        print("Acquiring Microsoft Graph access token...")
        token = get_graph_access_token(
            client_id=graph_cfg["client_id"],
            tenant_id=graph_cfg["tenant_id"],
            scopes=graph_cfg["scopes"],
            token_cache_path=token_cache_path,
        )

        excel_path = graph_cfg["excel_path"]
        log_sheet_name = graph_cfg.get("log_sheet_name", "Logg")

        print("Downloading Excel file from OneDrive...")
        download_excel_file(token, excel_path, str(temp_excel_path))

        print(
            f"Appending activities to Excel log sheet '{log_sheet_name}' (skipping already logged IDs)..."
        )
        append_runs_to_log(temp_excel_path, log_sheet_name, activities)

        print("Uploading updated Excel file to OneDrive...")
        upload_excel_file(token, excel_path, str(temp_excel_path))

        print("Done.")
        if telegram_configured:
            print("Sending telegram message")
            send_telegram_message(
                bot_token, chat_id, "Nenikekamen: Körningen lyckades."
            )
        else:
            print("Telegram not configured. No message is sent.")

    except Exception as exc:
        if telegram_configured:
            msg = f"Nenikekamen: Körningen kraschade.\n\n{type(exc).__name__}: {exc}"
            tb = traceback.format_exc()
            if tb:
                msg += "\n\n" + (tb[-500:] if len(tb) > 500 else tb)
            send_telegram_message(bot_token, chat_id, msg)
        raise


if __name__ == "__main__":
    main()
