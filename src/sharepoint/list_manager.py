"""
src/sharepoint/list_manager.py

Full CRUD operations on SharePoint lists via Microsoft Graph API.

Graph endpoint pattern:
  /sites/{site-id}/lists/{list-id}/items
"""

from __future__ import annotations

import logging
from typing import Any, Iterator

import requests
from tenacity import retry, stop_after_attempt, wait_exponential

from src.auth.auth_manager import get_azure_auth
from src.sharepoint.site_resolver import SiteResolver

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class SharePointListManager:
    """
    Read and write SharePoint list items using the Graph API.

    Example:
        mgr = SharePointListManager(site_url="https://contoso.sharepoint.com/sites/proj")
        items = mgr.get_all_items("Tasks")
        mgr.create_item("Tasks", {"Title": "New task", "Status": "Open"})
    """

    def __init__(self, site_url: str) -> None:
        self._auth = get_azure_auth()
        self._resolver = SiteResolver(site_url)
        self._site_id: str | None = None

    # ── Internals ─────────────────────────────────────────────────────────────

    @property
    def site_id(self) -> str:
        if self._site_id is None:
            self._site_id = self._resolver.get_site_id()
        return self._site_id

    def _session(self) -> requests.Session:
        return self._auth.get_requests_session()

    def _list_url(self, list_name: str) -> str:
        return f"{GRAPH_BASE}/sites/{self.site_id}/lists/{list_name}/items"

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
    def _get(self, url: str, params: dict | None = None) -> dict:
        resp = self._session().get(url, params=params)
        resp.raise_for_status()
        return resp.json()

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
    def _post(self, url: str, payload: dict) -> dict:
        resp = self._session().post(url, json=payload)
        resp.raise_for_status()
        return resp.json()

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
    def _patch(self, url: str, payload: dict) -> dict:
        resp = self._session().patch(url, json=payload)
        resp.raise_for_status()
        return resp.json()

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
    def _delete(self, url: str) -> None:
        resp = self._session().delete(url)
        resp.raise_for_status()

    # ── Public API ────────────────────────────────────────────────────────────

    def get_all_items(
        self,
        list_name: str,
        expand_fields: bool = True,
        filter_query: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        Return all items in a SharePoint list, handling Graph pagination automatically.

        Args:
            list_name:     Display name or GUID of the list.
            expand_fields: Whether to include `fields` in the response.
            filter_query:  Optional OData $filter string, e.g. "fields/Status eq 'Open'".
        """
        params: dict[str, str] = {}
        if expand_fields:
            params["expand"] = "fields(select=*)"
        if filter_query:
            params["$filter"] = filter_query

        url: str | None = self._list_url(list_name)
        items: list[dict] = []

        while url:
            data = self._get(url, params=params)
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")  # Follow pagination
            params = {}  # nextLink already contains params

        logger.info("Fetched %d items from list '%s'", len(items), list_name)
        return items

    def get_item(self, list_name: str, item_id: int) -> dict[str, Any]:
        """Fetch a single list item by its integer ID."""
        url = f"{self._list_url(list_name)}/{item_id}?expand=fields(select=*)"
        return self._get(url)

    def create_item(self, list_name: str, fields: dict[str, Any]) -> dict[str, Any]:
        """
        Create a new list item.

        Args:
            list_name: Display name or GUID of the list.
            fields:    Column name → value mapping.

        Returns:
            The created item as returned by Graph.
        """
        payload = {"fields": fields}
        result = self._post(self._list_url(list_name), payload)
        logger.info("Created item %s in list '%s'", result.get("id"), list_name)
        return result

    def update_item(
        self, list_name: str, item_id: int, fields: dict[str, Any]
    ) -> dict[str, Any]:
        """Update specific fields on an existing item."""
        url = f"{self._list_url(list_name)}/{item_id}/fields"
        result = self._patch(url, fields)
        logger.info("Updated item %d in list '%s'", item_id, list_name)
        return result

    def delete_item(self, list_name: str, item_id: int) -> None:
        """Permanently delete a list item."""
        url = f"{self._list_url(list_name)}/{item_id}"
        self._delete(url)
        logger.info("Deleted item %d from list '%s'", item_id, list_name)

    def upsert_item(
        self,
        list_name: str,
        match_field: str,
        match_value: Any,
        fields: dict[str, Any],
    ) -> dict[str, Any]:
        """
        Create or update an item based on a matching field value.

        Useful for idempotent imports — avoids duplicates.
        """
        existing = self.get_all_items(
            list_name,
            filter_query=f"fields/{match_field} eq '{match_value}'",
        )
        if existing:
            item_id = existing[0]["id"]
            return self.update_item(list_name, int(item_id), fields)
        return self.create_item(list_name, fields)

    def iter_items(
        self, list_name: str, batch_size: int = 500
    ) -> Iterator[dict[str, Any]]:
        """Memory-efficient generator that yields items one at a time."""
        params = {
            "expand": "fields(select=*)",
            "$top": str(batch_size),
        }
        url: str | None = self._list_url(list_name)
        while url:
            data = self._get(url, params=params)
            for item in data.get("value", []):
                yield item
            url = data.get("@odata.nextLink")
            params = {}

