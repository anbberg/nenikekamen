import os
from pathlib import Path

import pytest

from config_loader import load_config


def test_load_config_reads_required_values_from_env(tmp_path, monkeypatch):
    # Prepare a minimal .env file
    env_file = tmp_path / ".env"
    env_file.write_text(
        "\n".join(
            [
                "GRAPH_CLIENT_ID=gid",
                "GRAPH_TENANT_ID=tenant",
                "GRAPH_EXCEL_PATH=/path/to/file.xlsx",
                "GRAPH_LOG_SHEET_NAME=Logg",
                "GRAPH_TOKEN_CACHE_PATH=cache.bin",
                "STRAVA_CLIENT_ID=scid",
                "STRAVA_CLIENT_SECRET=secret",
                "STRAVA_REFRESH_TOKEN=refresh",
                "TRAINING_START_DATE=2026-01-01",
                "TEMP_EXCEL_PATH=local.xlsx",
                "GRAPH_SCOPES=Files.ReadWrite.All offline_access",
            ]
        ),
        encoding="utf-8",
    )

    # Ensure environment is clean for these keys
    for key in list(os.environ.keys()):
        if key.startswith("GRAPH_") or key.startswith("STRAVA_") or key in {
            "TRAINING_START_DATE",
            "TEMP_EXCEL_PATH",
        }:
            monkeypatch.delenv(key, raising=False)

    cfg = load_config(dotenv_path=str(env_file))

    assert cfg["graph"]["client_id"] == "gid"
    assert cfg["graph"]["tenant_id"] == "tenant"
    assert cfg["graph"]["excel_path"] == "/path/to/file.xlsx"
    assert cfg["graph"]["log_sheet_name"] == "Logg"
    assert cfg["graph"]["token_cache_path"] == "cache.bin"

    assert cfg["strava"]["client_id"] == "scid"
    assert cfg["strava"]["client_secret"] == "secret"
    assert cfg["strava"]["refresh_token"] == "refresh"

    assert cfg["runtime"]["temp_excel_path"] == "local.xlsx"
    assert cfg["runtime"]["training_start_date"] == "2026-01-01"

    # GRAPH_SCOPES is split on spaces
    assert cfg["graph"]["scopes"] == ["Files.ReadWrite.All", "offline_access"]


def test_load_config_raises_on_missing_required_keys(tmp_path, monkeypatch):
    # Empty env file -> missing everything
    env_file = tmp_path / ".env"
    env_file.write_text("", encoding="utf-8")

    for key in list(os.environ.keys()):
        if key.startswith("GRAPH_") or key.startswith("STRAVA_") or key in {
            "TRAINING_START_DATE",
            "TEMP_EXCEL_PATH",
        }:
            monkeypatch.delenv(key, raising=False)

    with pytest.raises(RuntimeError) as excinfo:
        load_config(dotenv_path=str(env_file))

    msg = str(excinfo.value)
    # Should mention several missing keys
    assert "GRAPH_CLIENT_ID" in msg
    assert "STRAVA_CLIENT_ID" in msg
    assert "TRAINING_START_DATE" in msg

