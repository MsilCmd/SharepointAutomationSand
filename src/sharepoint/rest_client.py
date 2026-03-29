"""
src/sharepoint/rest_client.py

Alternative SharePoint access via Office365-REST-Python-Client.
This is useful for operations that Graph doesn't support well, such as:
  - CAML queries for complex filtered list reads
  - Classic SharePoint features (site columns, content types)
  - Legacy on-prem SharePoint compatibility

Complements graph-based managers — use whichever suits the operation.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.lists.list import List as SPList
from office365.sharepoint.files.file import File as SPFile

from config.settings import settings

logger = logging.getLogger(__name__)


class SharePointRestClient:
    """
    Thin wrapper around Office365-REST-Python-Client (CSOM-style API).

    Supports two auth modes:
      - App-only: client ID + secret (recommended for automation)
      - User:     username + password (legacy, avoid in production)

    Example:
        client = SharePointRestClient()
        items = client.get_list_items("Tasks", caml_query="<Query>...</Query>")
        client.upload_file("Documents", Path("./report.pdf"))
    """

    def __init__(self, auth_mode: str = "app") -> None:
        """
        Args:
            auth_mode: "app" for client credential, "user" for username/password.
        """
        self._site_url = settings.sharepoint_site_url
        self._ctx = self._build_context(auth_mode)

    def _build_context(self, auth_mode: str) -> ClientContext:
        if auth_mode == "app":
            cred = ClientCredential(
                settings.azure_client_id,
                settings.azure_client_secret,
            )
        elif auth_mode == "user":
            if not settings.sharepoint_username or not settings.sharepoint_password:
                raise ValueError(
                    "SHAREPOINT_USERNAME and SHAREPOINT_PASSWORD must be set for user auth"
                )
            cred = UserCredential(
                settings.sharepoint_username,
                settings.sharepoint_password,
            )
        else:
            raise ValueError(f"Unknown auth_mode: {auth_mode!r}")

        ctx = ClientContext(self._site_url).with_credentials(cred)
        logger.debug("Built SharePoint context (%s auth) for %s", auth_mode, self._site_url)
        return ctx

    # ── List operations ───────────────────────────────────────────────────────

    def get_list(self, list_name: str) -> SPList:
        """Return the SharePoint List object for a given name."""
        lst = self._ctx.web.lists.get_by_title(list_name)
        self._ctx.load(lst)
        self._ctx.execute_query()
        return lst

    def get_list_items(
        self,
        list_name: str,
        caml_query: str | None = None,
        fields: list[str] | None = None,
    ) -> list[dict[str, Any]]:
        """
        Retrieve items from a SharePoint list.

        Args:
            list_name:  Display name of the list.
            caml_query: Optional CAML XML query string for server-side filtering.
            fields:     Column names to include. None = all.

        Returns:
            List of dicts with field name → value.

        Example CAML query to filter by status:
            caml_query = '''
            <View><Query><Where>
              <Eq><FieldRef Name="Status"/><Value Type="Text">Open</Value></Eq>
            </Where></Query></View>
            '''
        """
        from office365.sharepoint.caml.query import CamlQuery

        lst = self._ctx.web.lists.get_by_title(list_name)

        if caml_query:
            query = CamlQuery()
            query.ViewXml = caml_query
            items = lst.get_items(query)
        else:
            items = lst.items

        self._ctx.load(items)
        self._ctx.execute_query()

        result = []
        for item in items:
            row = dict(item.properties)
            if fields:
                row = {k: v for k, v in row.items() if k in fields}
            result.append(row)

        logger.info("CSOM: fetched %d items from '%s'", len(result), list_name)
        return result

    def create_list_item(
        self, list_name: str, fields: dict[str, Any]
    ) -> dict[str, Any]:
        """Create a new list item using CSOM."""
        lst = self._ctx.web.lists.get_by_title(list_name)
        item = lst.add_item(fields)
        self._ctx.execute_query()
        logger.info("CSOM: created item in '%s'", list_name)
        return dict(item.properties)

    def update_list_item(
        self, list_name: str, item_id: int, fields: dict[str, Any]
    ) -> None:
        """Update fields on an existing list item using CSOM."""
        lst = self._ctx.web.lists.get_by_title(list_name)
        item = lst.get_item_by_id(item_id)
        for key, value in fields.items():
            item.set_property(key, value)
        item.update()
        self._ctx.execute_query()
        logger.info("CSOM: updated item %d in '%s'", item_id, list_name)

    def delete_list_item(self, list_name: str, item_id: int) -> None:
        """Delete a list item by ID using CSOM."""
        lst = self._ctx.web.lists.get_by_title(list_name)
        item = lst.get_item_by_id(item_id)
        item.delete_object()
        self._ctx.execute_query()
        logger.info("CSOM: deleted item %d from '%s'", item_id, list_name)

    # ── File operations ───────────────────────────────────────────────────────

    def upload_file(
        self,
        library_name: str,
        local_path: Path,
        remote_folder: str = "/",
        overwrite: bool = True,
    ) -> dict[str, Any]:
        """
        Upload a file to a document library using the CSOM API.

        For files under ~250 MB. Use DocumentManager (Graph) for larger files.
        """
        local_path = Path(local_path)
        folder_url = f"{self._site_url}/{library_name}/{remote_folder.strip('/')}"
        folder = self._ctx.web.get_folder_by_server_relative_url(
            f"/{library_name}/{remote_folder.strip('/')}"
        )

        with open(local_path, "rb") as f:
            content = f.read()

        uploaded = folder.upload_file(local_path.name, content)
        self._ctx.execute_query()
        logger.info("CSOM: uploaded %s to %s", local_path.name, library_name)
        return dict(uploaded.properties)

    def download_file(
        self, library_name: str, remote_path: str, local_dest: Path
    ) -> Path:
        """Download a file from SharePoint using CSOM."""
        server_relative_url = f"/{library_name}/{remote_path.lstrip('/')}"
        local_dest = Path(local_dest)
        local_dest.parent.mkdir(parents=True, exist_ok=True)

        with open(local_dest, "wb") as f:
            (
                self._ctx.web.get_file_by_server_relative_url(server_relative_url)
                .download(f)
                .execute_query()
            )

        logger.info("CSOM: downloaded %s → %s", remote_path, local_dest)
        return local_dest

    def list_files_in_folder(
        self, library_name: str, folder_path: str = "/"
    ) -> list[dict[str, Any]]:
        """Return metadata for all files in a library folder using CSOM."""
        server_rel = f"/{library_name}/{folder_path.lstrip('/')}"
        folder = self._ctx.web.get_folder_by_server_relative_url(server_rel)
        files = folder.files
        self._ctx.load(files)
        self._ctx.execute_query()
        return [dict(f.properties) for f in files]

    # ── Site & web operations ─────────────────────────────────────────────────

    def get_web_properties(self) -> dict[str, Any]:
        """Return properties of the current SharePoint web (site)."""
        web = self._ctx.web
        self._ctx.load(web)
        self._ctx.execute_query()
        return dict(web.properties)

    def get_all_lists(self) -> list[dict[str, Any]]:
        """Return metadata for all lists/libraries in the site."""
        lists = self._ctx.web.lists
        self._ctx.load(lists)
        self._ctx.execute_query()
        return [dict(lst.properties) for lst in lists]

    def get_site_users(self) -> list[dict[str, Any]]:
        """Return all users with access to the current site."""
        users = self._ctx.web.site_users
        self._ctx.load(users)
        self._ctx.execute_query()
        return [dict(u.properties) for u in users]
