"""
src/dropbox/dropbox_manager.py

Browse and download files from Dropbox using the official Python SDK.
Used primarily as the source side of the Dropbox → SharePoint import pipeline.
"""

from __future__ import annotations

import logging
import tempfile
from pathlib import Path

import dropbox
from dropbox.files import FileMetadata, FolderMetadata

from src.auth.auth_manager import get_dropbox_auth

logger = logging.getLogger(__name__)


class DropboxManager:
    """
    Interact with a Dropbox account.

    Example:
        dbx = DropboxManager()
        files = dbx.list_files("/Reports")
        local_path = dbx.download_file("/Reports/Q1.xlsx")
    """

    def __init__(self) -> None:
        self._auth = get_dropbox_auth()

    @property
    def _client(self) -> dropbox.Dropbox:
        return self._auth.get_client()

    def list_files(
        self,
        folder_path: str = "",
        recursive: bool = False,
        extensions: list[str] | None = None,
    ) -> list[FileMetadata]:
        """
        Return FileMetadata objects for all files in a Dropbox folder.

        Args:
            folder_path: Dropbox path, e.g. "/Reports". Empty string = root.
            recursive:   Whether to recurse into sub-folders.
            extensions:  Optional whitelist of file extensions, e.g. [".xlsx", ".pdf"].

        Returns:
            List of FileMetadata objects.
        """
        result = self._client.files_list_folder(
            folder_path, recursive=recursive
        )
        entries: list[FileMetadata] = []

        while True:
            for entry in result.entries:
                if isinstance(entry, FileMetadata):
                    if extensions is None or Path(entry.name).suffix.lower() in extensions:
                        entries.append(entry)

            if not result.has_more:
                break
            result = self._client.files_list_folder_continue(result.cursor)

        logger.info(
            "Found %d files in Dropbox path '%s'", len(entries), folder_path or "/"
        )
        return entries

    def download_file(self, dropbox_path: str, local_dest: Path | None = None) -> Path:
        """
        Download a single file from Dropbox.

        Args:
            dropbox_path: Full Dropbox path, e.g. "/Reports/Q1.xlsx".
            local_dest:   Optional local path. Defaults to a temp file.

        Returns:
            Path to the downloaded local file.
        """
        if local_dest is None:
            suffix = Path(dropbox_path).suffix
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            local_dest = Path(tmp.name)
            tmp.close()

        local_dest = Path(local_dest)
        local_dest.parent.mkdir(parents=True, exist_ok=True)

        metadata, response = self._client.files_download(dropbox_path)
        with open(local_dest, "wb") as f:
            f.write(response.content)

        logger.info(
            "Downloaded Dropbox file %s → %s (%d bytes)",
            dropbox_path,
            local_dest,
            len(response.content),
        )
        return local_dest

    def get_metadata(self, dropbox_path: str) -> FileMetadata | FolderMetadata:
        """Return metadata for a Dropbox path."""
        return self._client.files_get_metadata(dropbox_path)

    def folder_exists(self, dropbox_path: str) -> bool:
        """Check whether a Dropbox folder exists."""
        try:
            meta = self.get_metadata(dropbox_path)
            return isinstance(meta, FolderMetadata)
        except dropbox.exceptions.ApiError:
            return False
