"""
src/sharepoint/site_resolver.py

Converts a human-readable SharePoint site URL into the Graph site ID
format required by most Graph API endpoints.

Graph site IDs look like:
  contoso.sharepoint.com,abc123...,def456...
"""

from __future__ import annotations

import logging
from urllib.parse import urlparse

import requests

from src.auth.auth_manager import get_azure_auth

logger = logging.getLogger(__name__)
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class SiteResolver:
    """
    Resolve a SharePoint site URL to its Graph site ID.

    The result is cached so only one API call is made per instance lifetime.
    """

    def __init__(self, site_url: str) -> None:
        self._site_url = site_url.rstrip("/")
        self._site_id: str | None = None
        self._auth = get_azure_auth()

    def get_site_id(self) -> str:
        if self._site_id:
            return self._site_id

        parsed = urlparse(self._site_url)
        hostname = parsed.hostname  # e.g. contoso.sharepoint.com
        # Strip leading slash from path, e.g. "/sites/mysite" → "sites/mysite"
        site_path = parsed.path.lstrip("/")

        url = f"{GRAPH_BASE}/sites/{hostname}:/{site_path}"
        session: requests.Session = self._auth.get_requests_session()
        resp = session.get(url)
        resp.raise_for_status()

        self._site_id = resp.json()["id"]
        logger.debug("Resolved site ID: %s → %s", self._site_url, self._site_id)
        return self._site_id
