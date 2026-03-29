"""
src/provisioning/provisioner.py

Automate SharePoint provisioning:
  - Create SharePoint lists with custom columns
  - Set list permissions
  - Create document library folder structures
  - Provision from a YAML/JSON template

Uses Microsoft Graph API for modern provisioning.
"""

from __future__ import annotations

import logging
from typing import Any

import requests

from src.auth.auth_manager import get_azure_auth
from src.sharepoint.site_resolver import SiteResolver

logger = logging.getLogger(__name__)
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# ── Column type mapping ───────────────────────────────────────────────────────
# Maps friendly names to Graph column definition dicts
COLUMN_TYPES: dict[str, dict] = {
    "text":     {"text": {}},
    "number":   {"number": {}},
    "boolean":  {"boolean": {}},
    "choice":   {"choice": {"allowTextEntry": False, "choices": []}},
    "datetime": {"dateTime": {"format": "dateTime"}},
    "person":   {"personOrGroup": {"allowMultipleSelection": False}},
    "url":      {"hyperlinkOrPicture": {"isPicture": False}},
    "lookup":   {"lookup": {}},
}


class SharePointProvisioner:
    """
    Provision SharePoint resources programmatically.

    Example:
        p = SharePointProvisioner("https://contoso.sharepoint.com/sites/proj")

        # Create a list with custom columns
        p.create_list("Project Tasks", columns=[
            {"name": "Status",    "type": "choice", "choices": ["Open", "Closed"]},
            {"name": "DueDate",   "type": "datetime"},
            {"name": "Assignee",  "type": "person"},
            {"name": "Priority",  "type": "number"},
        ])

        # Create a folder structure in a library
        p.create_folder_structure("Documents", [
            "2024/Q1", "2024/Q2", "2024/Q3", "2024/Q4",
            "Archive",
        ])
    """

    def __init__(self, site_url: str) -> None:
        self._auth = get_azure_auth()
        self._resolver = SiteResolver(site_url)

    @property
    def _site_id(self) -> str:
        return self._resolver.get_site_id()

    def _session(self) -> requests.Session:
        return self._auth.get_requests_session()

    # ── List provisioning ─────────────────────────────────────────────────────

    def create_list(
        self,
        display_name: str,
        columns: list[dict[str, Any]] | None = None,
        list_template: str = "genericList",
    ) -> dict[str, Any]:
        """
        Create a SharePoint list and optionally add custom columns.

        Args:
            display_name:  Display name for the list.
            columns:       List of column specs, each with at minimum
                           {"name": str, "type": str}.
            list_template: Graph list template: "genericList" | "documentLibrary".

        Returns:
            Created list resource from Graph.
        """
        url = f"{GRAPH_BASE}/sites/{self._site_id}/lists"
        payload: dict[str, Any] = {
            "displayName": display_name,
            "list": {"template": list_template},
        }

        # Embed column definitions in the creation request when possible
        if columns:
            payload["columns"] = [self._build_column_def(c) for c in columns]

        resp = self._session().post(url, json=payload)
        resp.raise_for_status()
        created = resp.json()
        logger.info(
            "Created list '%s' (id: %s)", display_name, created.get("id")
        )
        return created

    def add_column_to_list(
        self, list_id: str, column_spec: dict[str, Any]
    ) -> dict[str, Any]:
        """Add a single column to an existing list."""
        url = f"{GRAPH_BASE}/sites/{self._site_id}/lists/{list_id}/columns"
        col_def = self._build_column_def(column_spec)
        resp = self._session().post(url, json=col_def)
        resp.raise_for_status()
        return resp.json()

    def _build_column_def(self, spec: dict[str, Any]) -> dict[str, Any]:
        """Convert a friendly column spec into a Graph column definition."""
        col_type = spec.get("type", "text").lower()
        type_def = dict(COLUMN_TYPES.get(col_type, COLUMN_TYPES["text"]))

        # Handle choice columns
        if col_type == "choice" and "choices" in spec:
            type_def["choice"] = {
                "allowTextEntry": spec.get("allowTextEntry", False),
                "choices": spec["choices"],
            }

        col_def: dict[str, Any] = {
            "name": spec["name"],
            "displayName": spec.get("displayName", spec["name"]),
            "required": spec.get("required", False),
            "enforceUniqueValues": spec.get("unique", False),
            "hidden": spec.get("hidden", False),
            **type_def,
        }
        return col_def

    def get_or_create_list(
        self, display_name: str, columns: list[dict] | None = None
    ) -> dict[str, Any]:
        """Return an existing list by name, or create it if absent."""
        url = f"{GRAPH_BASE}/sites/{self._site_id}/lists"
        resp = self._session().get(url)
        resp.raise_for_status()
        for lst in resp.json().get("value", []):
            if lst["displayName"].lower() == display_name.lower():
                logger.info("List '%s' already exists — skipping creation", display_name)
                return lst
        return self.create_list(display_name, columns=columns)

    def delete_list(self, list_id: str) -> None:
        """Permanently delete a list by its GUID."""
        url = f"{GRAPH_BASE}/sites/{self._site_id}/lists/{list_id}"
        resp = self._session().delete(url)
        resp.raise_for_status()
        logger.info("Deleted list %s", list_id)

    # ── Folder structure provisioning ─────────────────────────────────────────

    def create_folder_structure(
        self, library_name: str, folders: list[str]
    ) -> list[dict[str, Any]]:
        """
        Create a set of folders inside a document library.

        Accepts nested paths like "2024/Q1/Invoices" — intermediate folders
        are created automatically.

        Args:
            library_name: Target document library name.
            folders:      List of folder paths relative to the library root.

        Returns:
            List of created DriveItem dicts.
        """
        # Get the drive ID for the library
        drives_url = f"{GRAPH_BASE}/sites/{self._site_id}/drives"
        resp = self._session().get(drives_url)
        resp.raise_for_status()
        drive_id: str | None = None
        for drive in resp.json().get("value", []):
            if drive["name"].lower() == library_name.lower():
                drive_id = drive["id"]
                break

        if not drive_id:
            raise ValueError(f"Library '{library_name}' not found")

        created_folders: list[dict] = []
        for folder_path in folders:
            parts = folder_path.strip("/").split("/")
            current_path = ""
            for part in parts:
                parent_ref = (
                    f"{GRAPH_BASE}/drives/{drive_id}/root/children"
                    if not current_path
                    else f"{GRAPH_BASE}/drives/{drive_id}/root:/{current_path}:/children"
                )
                current_path = f"{current_path}/{part}".lstrip("/")
                try:
                    folder_resp = self._session().post(
                        parent_ref,
                        json={
                            "name": part,
                            "folder": {},
                            "@microsoft.graph.conflictBehavior": "fail",
                        },
                    )
                    if folder_resp.status_code == 409:
                        logger.debug("Folder already exists: %s", current_path)
                        continue
                    folder_resp.raise_for_status()
                    created_folders.append(folder_resp.json())
                    logger.info("Created folder: %s/%s", library_name, current_path)
                except requests.HTTPError as exc:
                    if exc.response is not None and exc.response.status_code == 409:
                        logger.debug("Folder already exists: %s", current_path)
                    else:
                        raise

        return created_folders

    # ── Permission management ─────────────────────────────────────────────────

    def set_list_permissions(
        self,
        list_id: str,
        role: str,
        user_emails: list[str],
    ) -> list[dict]:
        """
        Grant permissions on a list to a set of users.

        Args:
            list_id:     The list GUID.
            role:        Permission role: "read" | "write" | "owner".
            user_emails: List of user email addresses.

        Returns:
            List of permission objects created.
        """
        url = f"{GRAPH_BASE}/sites/{self._site_id}/lists/{list_id}/permissions"
        results = []
        for email in user_emails:
            payload = {
                "roles": [role],
                "grantedToIdentities": [
                    {"user": {"email": email}}
                ],
            }
            resp = self._session().post(url, json=payload)
            resp.raise_for_status()
            results.append(resp.json())
            logger.info("Granted '%s' role to %s on list %s", role, email, list_id)
        return results

