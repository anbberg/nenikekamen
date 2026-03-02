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
    Acquire an access token for Microsoft Graph using device code flow.

    On first run, this will print a device code and URL. You complete
    authentication on another device, and the token will then be cached.
    """
    # MSAL does not accept reserved scopes (offline_access, profile, openid) in the list.
    scopes_for_msal = [s for s in scopes if s not in MSAL_RESERVED_SCOPES]
    if not scopes_for_msal:
        scopes_for_msal = ["Files.ReadWrite.All"]  # fallback

    cache_file = Path(token_cache_path)
    cache = _get_token_cache(cache_file)

    authority = _build_authority(tenant_id)
    app = msal.PublicClientApplication(client_id=client_id, authority=authority, token_cache=cache)

    # Try silent acquisition first
    result: Dict = app.acquire_token_silent(scopes=scopes_for_msal, account=None)

    if not result:
        # Interactive flow: opens your browser automatically for normal OAuth login.
        # Uses fixed port so redirect_uri matches Azure. Add http://localhost:8400
        # to your app's "Mobile and desktop applications" redirect URIs.
        print("Opening browser for Microsoft sign-in...")
        result = app.acquire_token_interactive(scopes=scopes_for_msal, port=8400)

    if "access_token" not in result:
        raise RuntimeError(f"Failed to obtain access token: {result.get('error_description')}")

    _save_token_cache(cache, cache_file)
    return result["access_token"]

