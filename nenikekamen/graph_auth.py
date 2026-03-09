from pathlib import Path
from typing import Dict, List

import msal

# MSAL rejects these if passed explicitly; Azure app still needs offline_access in the portal for refresh tokens.
MSAL_RESERVED_SCOPES = {"offline_access", "profile", "openid"}


def _build_authority(tenant_id: str) -> str:
    """
    Build the MSAL authority URL.

    For personal accounts, using the tenant ID provided by the portal is fine.
    """
    return f"https://login.microsoftonline.com/{tenant_id}"


def _get_token_cache(cache_path: Path) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))
    return cache


def _save_token_cache(cache: msal.SerializableTokenCache, cache_path: Path) -> None:
    if cache.has_state_changed:
        cache_path.write_text(cache.serialize(), encoding="utf-8")


def get_graph_access_token(
    client_id: str,
    tenant_id: str,
    scopes: List[str],
    token_cache_path: str,
) -> str:
    """
    Acquire an access token for Microsoft Graph.

    Flow (same on Windows and Pi):
    - First try to acquire a token silently from the local cache.
    - If that fails, use the device code flow where you open a URL
      on any device and enter the code shown in the terminal.
    """
    # MSAL does not accept reserved scopes (offline_access, profile, openid) in the list.
    scopes_for_msal = [s for s in scopes if s not in MSAL_RESERVED_SCOPES]
    if not scopes_for_msal:
        scopes_for_msal = ["Files.ReadWrite.All"]  # fallback

    cache_file = Path(token_cache_path)
    cache = _get_token_cache(cache_file)

    authority = _build_authority(tenant_id)
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=cache,
    )

    # Try silent acquisition first using any cached account
    accounts = app.get_accounts()
    result: Dict | None = None
    if accounts:
        result = app.acquire_token_silent(scopes=scopes_for_msal, account=accounts[0])

    if not result:
        # Device code flow: shows a code + URL in the terminal.
        flow: Dict = app.initiate_device_flow(scopes=scopes_for_msal)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to initiate device flow: {flow}")

        print("\nTo sign in to Microsoft Graph (device code flow):")
        print(flow["message"])
        print()

        result = app.acquire_token_by_device_flow(flow)  # blocks until completed

    if "access_token" not in result:
        raise RuntimeError(f"Failed to obtain access token: {result.get('error_description')}")

    _save_token_cache(cache, cache_file)
    return result["access_token"]
