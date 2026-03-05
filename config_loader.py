import os
from pathlib import Path
from typing import Any, Dict

from dotenv import load_dotenv


def load_config(dotenv_path: str | None = ".env") -> Dict[str, Any]:
    """
    Load configuration from environment variables (optionally via a .env file).

    This lets you keep secrets out of source control and still have a readable,
    comment-friendly config file.
    """
    if dotenv_path is not None:
        # Load variables from .env into the process environment if the file exists.
        env_file = Path(dotenv_path)
        if env_file.exists():
            load_dotenv(dotenv_path)

    graph_scopes_raw = os.getenv("GRAPH_SCOPES", "Files.ReadWrite.All offline_access")
    graph_scopes = [s for s in graph_scopes_raw.split() if s]

    config: Dict[str, Any] = {
        "graph": {
            "client_id": os.environ.get("GRAPH_CLIENT_ID", ""),
            "tenant_id": os.environ.get("GRAPH_TENANT_ID", ""),
            "scopes": graph_scopes,
            "excel_path": os.environ.get("GRAPH_EXCEL_PATH", ""),
            "log_sheet_name": os.environ.get("GRAPH_LOG_SHEET_NAME", "Logg"),
            "token_cache_path": os.environ.get("GRAPH_TOKEN_CACHE_PATH", "graph_token_cache.bin"),
        },
        "strava": {
            "client_id": os.environ.get("STRAVA_CLIENT_ID", ""),
            "client_secret": os.environ.get("STRAVA_CLIENT_SECRET", ""),
            "refresh_token": os.environ.get("STRAVA_REFRESH_TOKEN", ""),
        },
        "runtime": {
            "temp_excel_path": os.environ.get("TEMP_EXCEL_PATH", "local_training.xlsx"),
            "training_start_date": os.environ.get("TRAINING_START_DATE", ""),
        },
        "telegram": {
            "bot_token": os.environ.get("TELEGRAM_BOT_TOKEN", ""),
            "chat_id": os.environ.get("TELEGRAM_CHAT_ID", ""),
        },
    }

    # Basic sanity check to fail fast if key values are missing.
    missing = []
    if not config["graph"]["client_id"]:
        missing.append("GRAPH_CLIENT_ID")
    if not config["graph"]["tenant_id"]:
        missing.append("GRAPH_TENANT_ID")
    if not config["graph"]["excel_path"]:
        missing.append("GRAPH_EXCEL_PATH")
    if not config["strava"]["client_id"]:
        missing.append("STRAVA_CLIENT_ID")
    if not config["strava"]["client_secret"]:
        missing.append("STRAVA_CLIENT_SECRET")
    if not config["strava"]["refresh_token"]:
        missing.append("STRAVA_REFRESH_TOKEN")
    if not config["runtime"]["training_start_date"]:
        missing.append("TRAINING_START_DATE")

    if missing:
        raise RuntimeError(
            "Missing required environment variables: " + ", ".join(missing)
        )

    return config


