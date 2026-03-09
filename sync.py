"""
Sync job: Strava → Excel Logg.
Fetches activities from Strava and appends new ones to the Excel log table.
Sends Telegram notification on success or failure.
"""

from __future__ import annotations

import time
import traceback

from nenikekamen.config_loader import load_config
from nenikekamen.excel_log import LOG_HEADERS, STRAVA_ID_COLUMN_INDEX, run_to_log_row
from nenikekamen.graph_auth import get_graph_access_token
from nenikekamen.graph_excel import (
    append_table_rows,
    get_log_table_name,
    get_table_values,
    workbook_calculate,
)
from nenikekamen.strava_client import StravaClient
from nenikekamen.sync_analyse import parse_training_start_date
from nenikekamen.telegram_notify import send_telegram_message


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
        training_start = parse_training_start_date(training_start_str)

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
        if (
            table_values
            and len(table_values) > 0
            and len(table_values[0]) != len(LOG_HEADERS)
        ):
            raise RuntimeError(
                f"Tabellen '{table_name}' har {len(table_values[0])} kolumner men scriptet "
                f"förväntar sig {len(LOG_HEADERS)}. Kolumnerna ska heta (i ordning): "
                + ", ".join(LOG_HEADERS)
            )
        # First row is header; strava_id is at STRAVA_ID_COLUMN_INDEX
        existing_ids = set()
        for row in table_values[1:]:
            if (
                len(row) > STRAVA_ID_COLUMN_INDEX
                and row[STRAVA_ID_COLUMN_INDEX] is not None
            ):
                existing_ids.add(row[STRAVA_ID_COLUMN_INDEX])

        new_runs = [r for r in activities if r.get("id") not in existing_ids]
        if new_runs:
            values_2d = [run_to_log_row(r) for r in new_runs]
            print(f"Appending {len(new_runs)} activities to table '{table_name}'...")
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

        print("Done.")
        if telegram_configured:
            if new_runs:
                msg = f"✅ Sync lyckades. Hittade {len(new_runs)} nya pass."
            else:
                msg = "✅ Sync lyckades. Inga nya pass."
            send_telegram_message(bot_token, chat_id, msg)
        else:
            print("Telegram not configured. No message is sent.")

    except Exception as exc:
        if telegram_configured:
            err_msg = f"Sync kraschade.\n\n{type(exc).__name__}: {exc}"
            tb = traceback.format_exc()
            if tb:
                # Skicka med lite mer av tracebacken för felsökning.
                snippet = tb[-2000:] if len(tb) > 2000 else tb
                err_msg += "\n\n" + snippet
            send_telegram_message(bot_token, chat_id, err_msg)
        raise


if __name__ == "__main__":
    main()
