import time
from urllib.parse import quote
from typing import Any, List

import requests

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def _workbook_url(excel_path: str, *segments: str) -> str:
    """Build URL for workbook resource. Path is quoted so spaces/special chars work."""
    path_encoded = quote(excel_path, safe="/")
    base = f"{GRAPH_BASE_URL}/me/drive/root:{path_encoded}:/workbook"
    if segments:
        return base + "/" + "/".join(segments)
    return base


def _headers(access_token: str) -> dict:
    return {"Authorization": f"Bearer {access_token}"}


def _sheet_segment(sheet_name: str) -> str:
    """OData segment for worksheet by name (handles '+' etc. in name)."""
    encoded = quote(sheet_name, safe="")
    return f"worksheets('{encoded}')"


def _table_segment(table_name: str) -> str:
    """OData segment for table by name."""
    encoded = quote(table_name, safe="")
    return f"tables('{encoded}')"


def get_range_values(
    access_token: str,
    excel_path: str,
    sheet_name: str,
    address: str,
) -> List[List[Any]]:
    """
    GET a range and return its values (calculated for formula cells).
    address e.g. "A1:B2" or "J2:J100".
    """
    url = _workbook_url(
        excel_path,
        _sheet_segment(sheet_name),
        "range(address='" + address + "')",
    )
    params = {"$select": "values"}
    resp = requests.get(url, headers=_headers(access_token), params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data.get("values") or []


def get_log_table_name(
    access_token: str,
    excel_path: str,
    log_sheet_name: str,
) -> str:
    """
    GET the first table on the worksheet (e.g. Logg). Returns its name.
    Raises if the sheet has no tables.
    """
    url = _workbook_url(
        excel_path,
        _sheet_segment(log_sheet_name),
        "tables",
    )
    params = {"$select": "name"}
    resp = requests.get(url, headers=_headers(access_token), params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    tables = data.get("value") or []
    if not tables:
        raise RuntimeError(
            f"Arket '{log_sheet_name}' har inga tabeller. "
            "Gör om loggdatat till en tabell (Infoga → Tabell) och kör igen."
        )
    return tables[0]["name"]


def get_table_values(
    access_token: str,
    excel_path: str,
    table_name: str,
) -> List[List[Any]]:
    """GET a table's range values (includes header row as first row)."""
    url = _workbook_url(
        excel_path,
        _table_segment(table_name),
        "range",
    )
    params = {"$select": "values"}
    resp = requests.get(url, headers=_headers(access_token), params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data.get("values") or []


def append_table_rows(
    access_token: str,
    excel_path: str,
    table_name: str,
    values_2d: List[List[Any]],
    *,
    retry_on_lock: int = 3,
    retry_delay_sec: int = 60,
) -> None:
    """
    Append rows to the end of the table. values_2d: one row per inner list.

    Vid EditModeCannotAcquireLock (OneDrive släpper inte låset direkt): väntar och
    försöker igen upp till retry_on_lock gånger med retry_delay_sec sekunder mellan.
    """
    url = _workbook_url(
        excel_path,
        _table_segment(table_name),
        "rows",
    )
    payload = {"values": values_2d, "index": None}

    last_error = None
    for attempt in range(retry_on_lock):
        resp = requests.post(
            url, headers=_headers(access_token), json=payload, timeout=30
        )
        if resp.status_code != 409:
            resp.raise_for_status()
            return

        # 409 Conflict – läs felinfo
        try:
            full = resp.json()
        except Exception:
            full = None
        err = (full or {}).get("error", {}) if isinstance(full, dict) else {}
        code = err.get("code") or (err.get("innerError") or {}).get("code")
        message = err.get("message") or ""
        code_l = str(code or "").lower()
        msg_l = message.lower()

        # EditModeCannotAcquireLock – OneDrive släpper inte låset direkt. Retry.
        if "editmodecannotacquirelock" in code_l or "editing this workbook" in msg_l:
            last_error = RuntimeError(
                f"Excel-filen verkar låst (OneDrive släpper inte alltid direkt). "
                f"Försök {attempt + 1}/{retry_on_lock} misslyckades."
            )
            if attempt < retry_on_lock - 1:
                print(
                    f"  Låst – väntar {retry_delay_sec}s innan försök {attempt + 2}/{retry_on_lock}..."
                )
                time.sleep(retry_delay_sec)
                continue
            raise RuntimeError(
                "Excel-filen är (eller verkar) öppen och låst. "
                "Stäng alla instanser i Excel, Excel Online och OneDrive-webbläsaren. "
                "OneDrive kan ta flera minuter att släppa låset. Kör sync igen om en stund."
            ) from last_error

        # InsertDeleteConflict – ingen retry
        if "insertdeleteconflict" in code_l or "move cells" in msg_l:
            raise RuntimeError(
                "Excel tillåter inte att fler rader läggs till i tabellen: något under eller "
                "bredvid tabellen blockerar. Ta bort innehåll direkt under tabellen på arket Logg, "
                "eller flytta tabellen så att den har utrymme att växa nedåt."
            ) from None

        # Annan 409
        detail_str = full if full is not None else (resp.text or "(ingen body)")
        raise RuntimeError(
            f"409 Conflict vid append till tabellen '{table_name}'. "
            f"API-svar: {detail_str!r}"
        ) from None

    if last_error:
        raise last_error


def workbook_calculate(access_token: str, excel_path: str) -> None:
    """POST workbook/application/calculate to force recalculation of formulas."""
    url = _workbook_url(excel_path, "application", "calculate")
    resp = requests.post(url, headers=_headers(access_token), json={}, timeout=30)
    resp.raise_for_status()
