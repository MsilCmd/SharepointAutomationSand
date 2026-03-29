"""
src/sharepoint/document_manager.py

Upload and download documents to/from SharePoint document libraries
using the Microsoft Graph large-file upload session API.

Handles files of any size:
  - Small files (< 4 MB): direct PUT
  - Large files (≥ 4 MB): resumable upload session
"""

from __future__ import annotations

import logging
import math
from pathlib import Path
from typing import BinaryIO

import requests
from tenacity import retry, stop_after_attempt, wait_exponential

from src.auth.auth_manager import get_azure_auth
from src.sharepoint.site_resolver import SiteResolver

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SMALL_FILE_THRESHOLD = 4 * 1024 * 1024  # 4 MB
CHUNK_SIZE = 10 * 1024 * 1024  # 10 MB per chunk


class DocumentManager:
    """
    Upload and download SharePoint document library files via Graph.

    Example:
        dm = DocumentManager(site_url="https://contoso.sharepoint.com/sites/proj")
        dm.upload("Documents", Path("./report.pdf"))
        dm.download("Documents", "report.pdf", Path("./local/report.pdf"))
    """

    def __init__(self, site_url: str) -> None:
        self._auth = get_azure_auth()
        self._resolver = SiteResolver(site_url)

    @property
    def _site_id(self) -> str:
        return self._resolver.get_site_id()

    def _session(self) -> requests.Session:
        return self._auth.get_requests_session()

    def _drive_url(self, library_name: str) -> str:
        return f"{GRAPH_BASE}/sites/{self._site_id}/drives"

    def _get_drive_id(self, library_name: str) -> str:
        """Resolve a document library display name to its Graph drive ID."""
        resp = self._session().get(self._drive_url(library_name))
        resp.raise_for_status()
        drives = resp.json().get("value", [])
        for drive in drives:
            if drive["name"].lower() == library_name.lower():
                return drive["id"]
        raise ValueError(
            f"Document library '{library_name}' not found. "
            f"Available: {[d['name'] for d in drives]}"
        )

    # ── Upload ────────────────────────────────────────────────────────────────

    def upload(
        self,
        library_name: str,
        local_path: Path,
        remote_folder: str = "/",
        overwrite: bool = True,
    ) -> dict:
        """
        Upload a file to a SharePoint document library.

        Automatically chooses direct upload or resumable session based on size.

        Args:
            library_name:  Name of the document library (e.g. "Documents").
            local_path:    Path to the local file.
            remote_folder: Target folder within the library (default: root).
            overwrite:     Replace the file if it already exists.

        Returns:
            Graph DriveItem response dict.
        """
        local_path = Path(local_path)
        if not local_path.exists():
            raise FileNotFoundError(f"File not found: {local_path}")

        file_size = local_path.stat().st_size
        drive_id = self._get_drive_id(library_name)
        folder = remote_folder.strip("/")
        remote_path = f"{folder}/{local_path.name}" if folder else local_path.name

        item_url = (
            f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}:/content"
            if file_size < SMALL_FILE_THRESHOLD
            else f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}:/createUploadSession"
        )

        if file_size < SMALL_FILE_THRESHOLD:
            return self._small_upload(item_url, local_path, overwrite)
        return self._large_upload(item_url, local_path, file_size, overwrite)

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=30))
    def _small_upload(self, url: str, path: Path, overwrite: bool) -> dict:
        headers = {
            **self._auth.get_auth_headers(),
            "Content-Type": "application/octet-stream",
        }
        if overwrite:
            headers["@microsoft.graph.conflictBehavior"] = "replace"

        with open(path, "rb") as f:
            resp = requests.put(url, headers=headers, data=f)
        resp.raise_for_status()
        logger.info("Uploaded (direct) %s → SharePoint", path.name)
        return resp.json()

    def _large_upload(
        self, session_url: str, path: Path, file_size: int, overwrite: bool
    ) -> dict:
        """Use Graph's resumable upload session for large files."""
        # 1. Create upload session
        conflict = "replace" if overwrite else "fail"
        payload = {
            "item": {
                "@microsoft.graph.conflictBehavior": conflict,
                "name": path.name,
            }
        }
        sess_resp = self._session().post(session_url, json=payload)
        sess_resp.raise_for_status()
        upload_url = sess_resp.json()["uploadUrl"]

        # 2. Upload in chunks
        total_chunks = math.ceil(file_size / CHUNK_SIZE)
        logger.info(
            "Starting large-file upload: %s (%d MB, %d chunks)",
            path.name,
            file_size // (1024 * 1024),
            total_chunks,
        )

        with open(path, "rb") as f:
            chunk_index = 0
            while True:
                chunk = f.read(CHUNK_SIZE)
                if not chunk:
                    break
                start = chunk_index * CHUNK_SIZE
                end = start + len(chunk) - 1
                headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{file_size}",
                }
                chunk_resp = requests.put(upload_url, headers=headers, data=chunk)
                chunk_resp.raise_for_status()
                chunk_index += 1
                logger.debug("Uploaded chunk %d/%d", chunk_index, total_chunks)

                # Final chunk returns the DriveItem
                if chunk_resp.status_code in (200, 201):
                    logger.info("Large-file upload complete: %s", path.name)
                    return chunk_resp.json()

        raise RuntimeError("Upload session ended without a completion response")

    # ── Download ──────────────────────────────────────────────────────────────

    def download(
        self,
        library_name: str,
        remote_file_path: str,
        local_dest: Path,
        chunk_size: int = 8192,
    ) -> Path:
        """
        Download a file from a SharePoint document library.

        Args:
            library_name:     Document library name.
            remote_file_path: Path within library, e.g. "Reports/Q1.xlsx".
            local_dest:       Local path to write the file to.
            chunk_size:       Streaming chunk size in bytes.

        Returns:
            Path to the downloaded file.
        """
        drive_id = self._get_drive_id(library_name)
        content_url = (
            f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_file_path}:/content"
        )

        local_dest = Path(local_dest)
        local_dest.parent.mkdir(parents=True, exist_ok=True)

        with self._session().get(content_url, stream=True) as resp:
            resp.raise_for_status()
            with open(local_dest, "wb") as f:
                for chunk in resp.iter_content(chunk_size=chunk_size):
                    f.write(chunk)

        logger.info("Downloaded %s → %s", remote_file_path, local_dest)
        return local_dest

    def list_files(self, library_name: str, folder_path: str = "/") -> list[dict]:
        """List all files in a library folder."""
        drive_id = self._get_drive_id(library_name)
        folder = folder_path.strip("/")
        url = (
            f"{GRAPH_BASE}/drives/{drive_id}/root/children"
            if not folder
            else f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder}:/children"
        )
        resp = self._session().get(url)
        resp.raise_for_status()
        return resp.json().get("value", [])

    def delete_file(self, library_name: str, remote_file_path: str) -> None:
        """Delete a file from the library."""
        drive_id = self._get_drive_id(library_name)
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_file_path}"
        resp = self._session().delete(url)
        resp.raise_for_status()
        logger.info("Deleted %s from library '%s'", remote_file_path, library_name)
