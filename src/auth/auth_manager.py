"""
src/auth/auth_manager.py

Handles Azure AD authentication (MSAL) and Dropbox token management.
Supports:
  - Client credentials flow (app-only, no user)
  - Token caching to avoid repeated round-trips
"""

from __future__ import annotations

import logging
from functools import lru_cache

import dropbox
import msal
import requests

from config.settings import settings

logger = logging.getLogger(__name__)


class AzureAuthManager:
    """
    Obtain and cache Azure AD access tokens using the client credentials flow.

    Usage:
        auth = AzureAuthManager()
        token = auth.get_token()
        headers = auth.get_auth_headers()
    """

    _GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]

    def __init__(self) -> None:
        self._app = msal.ConfidentialClientApplication(
            client_id=settings.azure_client_id,
            client_credential=settings.azure_client_secret,
            authority=settings.graph_authority,
        )
        self._token: dict | None = None

    def get_token(self) -> str:
        """Return a valid access token, refreshing if necessary."""
        # Try the cache first
        result = self._app.acquire_token_silent(self._GRAPH_SCOPES, account=None)
        if not result:
            logger.debug("Token cache miss — fetching new token from Azure AD")
            result = self._app.acquire_token_for_client(scopes=self._GRAPH_SCOPES)

        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "unknown"))
            raise RuntimeError(f"Failed to acquire Azure AD token: {error}")

        return result["access_token"]

    def get_auth_headers(self) -> dict[str, str]:
        """Return HTTP headers with a valid Bearer token."""
        return {
            "Authorization": f"Bearer {self.get_token()}",
            "Content-Type": "application/json",
        }

    def get_requests_session(self) -> requests.Session:
        """Return a requests.Session pre-configured with auth headers."""
        session = requests.Session()
        session.headers.update(self.get_auth_headers())
        return session


class DropboxAuthManager:
    """
    Obtain a Dropbox client using the refresh-token (offline access) flow.
    Falls back to a static access token when DROPBOX_ACCESS_TOKEN is set.
    """

    def __init__(self) -> None:
        self._client: dropbox.Dropbox | None = None

    def get_client(self) -> dropbox.Dropbox:
        """Return a cached, authenticated Dropbox client."""
        if self._client is not None:
            return self._client

        if settings.dropbox_access_token:
            logger.debug("Using static Dropbox access token")
            self._client = dropbox.Dropbox(settings.dropbox_access_token)
        else:
            logger.debug("Using Dropbox refresh token flow")
            self._client = dropbox.Dropbox(
                oauth2_refresh_token=settings.dropbox_refresh_token,
                app_key=settings.dropbox_app_key,
                app_secret=settings.dropbox_app_secret,
            )

        # Sanity-check the credentials
        self._client.users_get_current_account()
        logger.info("Dropbox client authenticated successfully")
        return self._client


# ── Module-level singletons ───────────────────────────────────────────────────

@lru_cache(maxsize=1)
def get_azure_auth() -> AzureAuthManager:
    return AzureAuthManager()


@lru_cache(maxsize=1)
def get_dropbox_auth() -> DropboxAuthManager:
    return DropboxAuthManager()

