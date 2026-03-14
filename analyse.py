"""
Analyse job: Excel → Telegram.
Reads Plan+Agg sheet, builds status summary, sends to Telegram.
"""

from __future__ import annotations

import traceback

from nenikekamen.config_loader import load_config
from nenikekamen.graph_auth import get_graph_access_token
from nenikekamen.graph_excel import workbook_calculate
from nenikekamen.sync_analyse import build_current_week_plan_agg_summary
from nenikekamen.telegram_notify import send_telegram_message


def main() -> None:
    config = load_config()

    graph_cfg = config["graph"]
    telegram_cfg = config.get("telegram", {})
    plan_cfg = config.get("plan_summary", {})
    bot_token = (telegram_cfg.get("bot_token") or "").strip()
    chat_id = (telegram_cfg.get("chat_id") or "").strip()
    telegram_configured = bool(bot_token and chat_id)

    try:
        token_cache_path = graph_cfg.get("token_cache_path", "graph_token_cache.bin")

        print("Acquiring Microsoft Graph access token...")
        token = get_graph_access_token(
            client_id=graph_cfg["client_id"],
            tenant_id=graph_cfg["tenant_id"],
            scopes=graph_cfg["scopes"],
            token_cache_path=token_cache_path,
        )

        excel_path = graph_cfg["excel_path"]

        # Ensure formulas are recalculated before reading
        try:
            workbook_calculate(token, excel_path)
        except Exception as e:
            print(f"Calculate (valfritt) misslyckades: {e}")

        summary_str = None
        sheet_name = (plan_cfg.get("sheet_name") or "").strip()
        if sheet_name:
            summary_str = build_current_week_plan_agg_summary(
                token,
                excel_path,
                sheet_name,
            )

        if summary_str is None:
            summary_str = "Ingen Plan+Agg konfigurerad."

        print("Done.")
        if telegram_configured:
            msg = summary_str
            print(msg)
            send_telegram_message(bot_token, chat_id, msg)
        else:
            print("Telegram not configured. No message is sent.")
            print(f"Status: {summary_str}")

    except Exception as exc:
        if telegram_configured:
            err_msg = f"Nenikekamen: Analys kraschade.\n\n{type(exc).__name__}: {exc}"
            tb = traceback.format_exc()
            if tb:
                snippet = tb[-2000:] if len(tb) > 2000 else tb
                err_msg += "\n\n" + snippet
            send_telegram_message(bot_token, chat_id, err_msg)
        raise


if __name__ == "__main__":
    main()
