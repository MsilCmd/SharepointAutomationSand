"""
src/dropbox/import_pipeline.py

Orchestrates the Dropbox → SharePoint import workflow:

  1. List files in a Dropbox folder
  2. Download each file locally (temp dir)
  3. Upload each file to a SharePoint document library
  4. Optionally write an import record to a SharePoint list for auditing
  5. Clean up temp files

Designed to be idempotent: re-running the pipeline won't duplicate files
if `overwrite=True` is set (which calls Graph's replace conflict behaviour).
"""

from __future__ import annotations

import logging
import tempfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from src.dropbox.dropbox_manager import DropboxManager
from src.sharepoint.document_manager import DocumentManager
from src.sharepoint.list_manager import SharePointListManager

logger = logging.getLogger(__name__)


@dataclass
class ImportResult:
    """Outcome of a single file import."""

    dropbox_path: str
    sharepoint_path: str
    success: bool
    error: str | None = None
    bytes_transferred: int = 0
    timestamp: datetime = field(default_factory=lambda: datetime.now(timezone.utc))


@dataclass
class ImportSummary:
    """Aggregate results of a pipeline run."""

    total: int = 0
    succeeded: int = 0
    failed: int = 0
    results: list[ImportResult] = field(default_factory=list)

    @property
    def success_rate(self) -> float:
        return (self.succeeded / self.total * 100) if self.total else 0.0


class DropboxToSharePointPipeline:
    """
    Import files from a Dropbox folder into a SharePoint document library.

    Example:
        pipeline = DropboxToSharePointPipeline(
            site_url="https://contoso.sharepoint.com/sites/proj",
            sp_library="ImportedFiles",
            audit_list="Import Log",          # optional SP list for audit trail
        )
        summary = pipeline.run(
            dropbox_folder="/Client Reports",
            sp_folder="Dropbox Imports",
            extensions=[".pdf", ".xlsx"],
        )
        print(f"Imported {summary.succeeded}/{summary.total} files")
    """

    def __init__(
        self,
        site_url: str,
        sp_library: str,
        audit_list: str | None = None,
        overwrite: bool = True,
        on_progress: Callable[[ImportResult], None] | None = None,
    ) -> None:
        self._dbx = DropboxManager()
        self._docs = DocumentManager(site_url)
        self._lists = SharePointListManager(site_url) if audit_list else None
        self._sp_library = sp_library
        self._audit_list = audit_list
        self._overwrite = overwrite
        self._on_progress = on_progress

    def run(
        self,
        dropbox_folder: str = "",
        sp_folder: str = "/",
        extensions: list[str] | None = None,
        recursive: bool = False,
        dry_run: bool = False,
    ) -> ImportSummary:
        """
        Execute the import pipeline.

        Args:
            dropbox_folder: Source Dropbox folder path.
            sp_folder:      Destination folder within the SP library.
            extensions:     File extension whitelist, e.g. [".pdf", ".xlsx"].
            recursive:      Recurse into sub-folders in Dropbox.
            dry_run:        Log what would be imported without actually doing it.

        Returns:
            ImportSummary with per-file results.
        """
        summary = ImportSummary()
        files = self._dbx.list_files(
            dropbox_folder, recursive=recursive, extensions=extensions
        )
        summary.total = len(files)

        logger.info(
            "Starting Dropbox → SharePoint import: %d files from '%s'",
            summary.total,
            dropbox_folder or "/",
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            for file_meta in files:
                result = self._import_one(
                    file_meta.path_display,
                    sp_folder,
                    Path(tmpdir),
                    dry_run,
                )
                summary.results.append(result)
                if result.success:
                    summary.succeeded += 1
                else:
                    summary.failed += 1

                if self._on_progress:
                    self._on_progress(result)

        logger.info(
            "Import complete: %d/%d succeeded (%.1f%%)",
            summary.succeeded,
            summary.total,
            summary.success_rate,
        )
        return summary

    def _import_one(
        self,
        dropbox_path: str,
        sp_folder: str,
        tmp_dir: Path,
        dry_run: bool,
    ) -> ImportResult:
        filename = Path(dropbox_path).name
        sp_path = f"{sp_folder}/{filename}".lstrip("/")

        if dry_run:
            logger.info("[DRY RUN] Would import: %s → %s", dropbox_path, sp_path)
            return ImportResult(
                dropbox_path=dropbox_path,
                sharepoint_path=sp_path,
                success=True,
            )

        try:
            # Download from Dropbox
            local_path = self._dbx.download_file(
                dropbox_path, tmp_dir / filename
            )
            file_size = local_path.stat().st_size

            # Upload to SharePoint
            self._docs.upload(
                self._sp_library,
                local_path,
                remote_folder=sp_folder,
                overwrite=self._overwrite,
            )

            result = ImportResult(
                dropbox_path=dropbox_path,
                sharepoint_path=sp_path,
                success=True,
                bytes_transferred=file_size,
            )

            # Write audit record
            if self._lists and self._audit_list:
                self._write_audit_record(result)

            return result

        except Exception as exc:
            logger.error("Failed to import %s: %s", dropbox_path, exc, exc_info=True)
            return ImportResult(
                dropbox_path=dropbox_path,
                sharepoint_path=sp_path,
                success=False,
                error=str(exc),
            )

    def _write_audit_record(self, result: ImportResult) -> None:
        """Write a row to the SharePoint audit list."""
        try:
            assert self._lists and self._audit_list
            self._lists.create_item(
                self._audit_list,
                {
                    "Title": Path(result.dropbox_path).name,
                    "DropboxPath": result.dropbox_path,
                    "SharePointPath": result.sharepoint_path,
                    "Status": "Success" if result.success else "Failed",
                    "ErrorMessage": result.error or "",
                    "BytesTransferred": result.bytes_transferred,
                    "ImportedAt": result.timestamp.isoformat(),
                },
            )
        except Exception as exc:
            logger.warning("Could not write audit record: %s", exc)

