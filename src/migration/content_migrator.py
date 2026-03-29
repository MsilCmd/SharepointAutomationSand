"""
src/migration/content_migrator.py

Migrate SharePoint content between:
  - Lists (within or across sites)
  - Document libraries (within or across sites)
  - Full site content

All operations are resumable: a JSON checkpoint file tracks which items
have been migrated so the process can be restarted after failure.
"""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from src.sharepoint.document_manager import DocumentManager
from src.sharepoint.list_manager import SharePointListManager

logger = logging.getLogger(__name__)


class ContentMigrator:
    """
    Migrate SharePoint list items and documents between sites or libraries.

    Example — migrate a list:
        migrator = ContentMigrator(
            source_url="https://contoso.sharepoint.com/sites/old",
            dest_url="https://contoso.sharepoint.com/sites/new",
        )
        report = migrator.migrate_list("Tasks", "ProjectTasks")

    Example — migrate a document library:
        report = migrator.migrate_library("Documents", "ArchivedDocs")
    """

    def __init__(
        self,
        source_url: str,
        dest_url: str,
        checkpoint_path: Path = Path("./migration_checkpoint.json"),
    ) -> None:
        self._src_lists = SharePointListManager(source_url)
        self._dst_lists = SharePointListManager(dest_url)
        self._src_docs = DocumentManager(source_url)
        self._dst_docs = DocumentManager(dest_url)
        self._checkpoint_path = checkpoint_path
        self._checkpoint: dict[str, set[str]] = self._load_checkpoint()

    # ── Checkpoint helpers ────────────────────────────────────────────────────

    def _load_checkpoint(self) -> dict[str, set[str]]:
        if self._checkpoint_path.exists():
            data = json.loads(self._checkpoint_path.read_text())
            return {k: set(v) for k, v in data.items()}
        return {}

    def _save_checkpoint(self) -> None:
        data = {k: list(v) for k, v in self._checkpoint.items()}
        self._checkpoint_path.write_text(json.dumps(data, indent=2))

    def _is_migrated(self, namespace: str, item_id: str) -> bool:
        return item_id in self._checkpoint.get(namespace, set())

    def _mark_migrated(self, namespace: str, item_id: str) -> None:
        self._checkpoint.setdefault(namespace, set()).add(item_id)
        self._save_checkpoint()

    # ── List migration ────────────────────────────────────────────────────────

    def migrate_list(
        self,
        source_list: str,
        dest_list: str,
        field_mapping: dict[str, str] | None = None,
        filter_query: str | None = None,
    ) -> dict[str, Any]:
        """
        Copy all items from source_list → dest_list.

        Args:
            source_list:   Name of the source SharePoint list.
            dest_list:     Name of the destination SharePoint list.
            field_mapping: Optional rename map, e.g. {"OldName": "NewName"}.
            filter_query:  Optional OData filter to migrate a subset of items.

        Returns:
            Migration report dict.
        """
        namespace = f"list:{source_list}→{dest_list}"
        report: dict[str, Any] = {
            "source": source_list,
            "destination": dest_list,
            "started_at": datetime.now(timezone.utc).isoformat(),
            "migrated": 0,
            "skipped": 0,
            "failed": 0,
            "errors": [],
        }

        items = self._src_lists.get_all_items(
            source_list, filter_query=filter_query
        )
        logger.info(
            "Migrating %d items: %s → %s", len(items), source_list, dest_list
        )

        for item in items:
            item_id = str(item["id"])
            if self._is_migrated(namespace, item_id):
                report["skipped"] += 1
                continue

            try:
                fields: dict[str, Any] = item.get("fields", {})
                # Apply field renaming
                if field_mapping:
                    fields = {
                        field_mapping.get(k, k): v for k, v in fields.items()
                    }
                # Strip read-only Graph metadata fields
                fields = {
                    k: v
                    for k, v in fields.items()
                    if not k.startswith("@") and k not in ("id", "Modified", "Created")
                }
                self._dst_lists.create_item(dest_list, fields)
                self._mark_migrated(namespace, item_id)
                report["migrated"] += 1
            except Exception as exc:
                logger.error("Failed to migrate item %s: %s", item_id, exc)
                report["failed"] += 1
                report["errors"].append({"item_id": item_id, "error": str(exc)})

        report["completed_at"] = datetime.now(timezone.utc).isoformat()
        logger.info(
            "List migration complete: %d migrated, %d skipped, %d failed",
            report["migrated"],
            report["skipped"],
            report["failed"],
        )
        return report

    # ── Document library migration ────────────────────────────────────────────

    def migrate_library(
        self,
        source_library: str,
        dest_library: str,
        source_folder: str = "/",
        dest_folder: str = "/",
        tmp_dir: Path = Path("/tmp/sp_migration"),
        extensions: list[str] | None = None,
    ) -> dict[str, Any]:
        """
        Copy all files from source_library → dest_library.

        Downloads each file to a temp directory then re-uploads.

        Args:
            source_library: Source document library name.
            dest_library:   Destination document library name.
            source_folder:  Folder path within the source library.
            dest_folder:    Folder path within the destination library.
            tmp_dir:        Local temp directory for staging files.
            extensions:     Optional file extension filter.

        Returns:
            Migration report dict.
        """
        namespace = f"lib:{source_library}→{dest_library}"
        tmp_dir.mkdir(parents=True, exist_ok=True)

        report: dict[str, Any] = {
            "source": source_library,
            "destination": dest_library,
            "started_at": datetime.now(timezone.utc).isoformat(),
            "migrated": 0,
            "skipped": 0,
            "failed": 0,
            "errors": [],
        }

        files = self._src_docs.list_files(source_library, folder_path=source_folder)
        logger.info(
            "Migrating %d files: %s/%s → %s/%s",
            len(files),
            source_library,
            source_folder,
            dest_library,
            dest_folder,
        )

        for file_item in files:
            name = file_item.get("name", "")
            if extensions and Path(name).suffix.lower() not in extensions:
                continue

            file_key = file_item.get("id", name)
            if self._is_migrated(namespace, file_key):
                report["skipped"] += 1
                continue

            remote_path = (
                f"{source_folder.strip('/')}/{name}"
                if source_folder.strip("/")
                else name
            )

            try:
                local_path = self._src_docs.download(
                    source_library, remote_path, tmp_dir / name
                )
                self._dst_docs.upload(
                    dest_library, local_path, remote_folder=dest_folder, overwrite=True
                )
                self._mark_migrated(namespace, file_key)
                report["migrated"] += 1
                # Clean up temp file immediately to save disk space
                local_path.unlink(missing_ok=True)
            except Exception as exc:
                logger.error("Failed to migrate file %s: %s", name, exc)
                report["failed"] += 1
                report["errors"].append({"file": name, "error": str(exc)})

        report["completed_at"] = datetime.now(timezone.utc).isoformat()
        logger.info(
            "Library migration complete: %d migrated, %d skipped, %d failed",
            report["migrated"],
            report["skipped"],
            report["failed"],
        )
        return report

