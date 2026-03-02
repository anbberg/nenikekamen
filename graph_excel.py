from pathlib import Path
from typing import Optional

import requests

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def _build_file_content_url(excel_path: str) -> str:
    # excel_path is the path as it appears in OneDrive, e.g. "/Documents/Training/Marathon.xlsx"
    # Graph expects a URL-encoded path, but simple paths without special chars work as-is.
    return f"{GRAPH_BASE_URL}/me/drive/root:{excel_path}:/content"


def download_excel_file(
    access_token: str,
    excel_path: str,
    local_path: str,
) -> Path:
    """
    Download the Excel file from OneDrive to a local temp path.
    """
    url = _build_file_content_url(excel_path)
    headers = {"Authorization": f"Bearer {access_token}"}

    response = requests.get(url, headers=headers)
    response.raise_for_status()

    target = Path(local_path)
    target.write_bytes(response.content)
    return target


def upload_excel_file(
    access_token: str,
    excel_path: str,
    local_path: str,
) -> None:
    """
    Upload the local Excel file back to OneDrive, replacing the existing file.
    """
    url = _build_file_content_url(excel_path)
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }

    data = Path(local_path).read_bytes()
    response = requests.put(url, headers=headers, data=data)
    response.raise_for_status()

