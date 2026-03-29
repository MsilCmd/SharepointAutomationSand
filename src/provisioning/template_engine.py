"""
src/provisioning/template_engine.py

Provision SharePoint resources from a declarative YAML or JSON template.

This makes provisioning repeatable, version-controllable, and reviewable
in pull requests — treating SharePoint config as infrastructure-as-code.

Template format (YAML):
───────────────────────
site_url: https://contoso.sharepoint.com/sites/proj

lists:
  - name: "Project Tasks"
    columns:
      - name: Status
        type: choice
        choices: [Open, In Progress, Done, Blocked]
        required: true
      - name: DueDate
        type: datetime
      - name: Assignee
        type: person
      - name: Priority
        type: number
    permissions:
      - role: write
        users: [alice@contoso.com, bob@contoso.com]
      - role: read
        users: [viewer@contoso.com]

libraries:
  - name: "Project Documents"
    folders:
      - "2024/Q1"
      - "2024/Q2"
      - "Archive"

See config/example_provision.yaml for a complete annotated example.
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

import yaml

from src.provisioning.provisioner import SharePointProvisioner

logger = logging.getLogger(__name__)


class TemplateEngine:
    """
    Apply a YAML/JSON provisioning template to a SharePoint site.

    Example:
        engine = TemplateEngine()
        report = engine.apply("config/my_site_template.yaml")
        print(report)
    """

    def apply(
        self,
        template_path: str | Path,
        dry_run: bool = False,
    ) -> dict[str, Any]:
        """
        Parse and apply a provisioning template.

        Args:
            template_path: Path to a .yaml or .json template file.
            dry_run:       Log what would be provisioned without making changes.

        Returns:
            Report dict with created resources and any errors.
        """
        template = self._load(Path(template_path))
        site_url = template.get("site_url") or ""
        if not site_url:
            raise ValueError("Template must include 'site_url'")

        provisioner = SharePointProvisioner(site_url)
        report: dict[str, Any] = {
            "site_url": site_url,
            "dry_run": dry_run,
            "created_lists": [],
            "created_folders": [],
            "errors": [],
        }

        # ── Provision lists ───────────────────────────────────────────────────
        for list_spec in template.get("lists", []):
            list_name = list_spec["name"]
            columns = list_spec.get("columns", [])
            permissions = list_spec.get("permissions", [])

            if dry_run:
                logger.info("[DRY RUN] Would create list: %s (%d columns)", list_name, len(columns))
                report["created_lists"].append({"name": list_name, "dry_run": True})
                continue

            try:
                created = provisioner.get_or_create_list(list_name, columns=columns)
                list_id = created.get("id", "")
                report["created_lists"].append({"name": list_name, "id": list_id})
                logger.info("Provisioned list '%s' (id: %s)", list_name, list_id)

                # Apply permissions
                for perm in permissions:
                    role = perm.get("role", "read")
                    users = perm.get("users", [])
                    if users and list_id:
                        provisioner.set_list_permissions(list_id, role, users)

            except Exception as exc:
                logger.error("Failed to provision list '%s': %s", list_name, exc)
                report["errors"].append({"list": list_name, "error": str(exc)})

        # ── Provision document libraries & folders ────────────────────────────
        for lib_spec in template.get("libraries", []):
            lib_name = lib_spec["name"]
            folders = lib_spec.get("folders", [])

            if dry_run:
                logger.info(
                    "[DRY RUN] Would create folder structure in '%s': %s",
                    lib_name, folders,
                )
                report["created_folders"].append({"library": lib_name, "dry_run": True})
                continue

            try:
                if folders:
                    created = provisioner.create_folder_structure(lib_name, folders)
                    report["created_folders"].append(
                        {"library": lib_name, "folders": [f["name"] for f in created]}
                    )
                    logger.info(
                        "Provisioned %d folders in library '%s'", len(created), lib_name
                    )
            except Exception as exc:
                logger.error(
                    "Failed to provision library '%s': %s", lib_name, exc
                )
                report["errors"].append({"library": lib_name, "error": str(exc)})

        return report

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _load(path: Path) -> dict[str, Any]:
        if not path.exists():
            raise FileNotFoundError(f"Template not found: {path}")
        text = path.read_text(encoding="utf-8")
        if path.suffix in (".yaml", ".yml"):
            return yaml.safe_load(text)
        if path.suffix == ".json":
            return json.loads(text)
        raise ValueError(f"Unsupported template format: {path.suffix}")

    @staticmethod
    def generate_example(output_path: str | Path = "config/example_provision.yaml") -> Path:
        """Write a fully-annotated example template to disk."""
        example = """\
# SharePoint Provisioning Template
# Apply with: python scripts/provision.py --template config/example_provision.yaml

# Full URL to the target SharePoint site
site_url: https://contoso.sharepoint.com/sites/myproject

# ── Lists ─────────────────────────────────────────────────────────────────────
lists:
  - name: "Project Tasks"
    columns:
      - name: Status
        type: choice
        choices: [Open, "In Progress", Done, Blocked]
        required: true

      - name: DueDate
        type: datetime

      - name: Assignee
        type: person

      - name: Priority
        type: number

      - name: Notes
        type: text

    # Grant permissions after list creation
    permissions:
      - role: write   # "read" | "write" | "owner"
        users:
          - alice@contoso.com
          - bob@contoso.com
      - role: read
        users:
          - viewer@contoso.com

  - name: "Import Log"
    columns:
      - name: DropboxPath
        type: text
      - name: SharePointPath
        type: text
      - name: Status
        type: choice
        choices: [Success, Failed]
      - name: ErrorMessage
        type: text
      - name: ImportedAt
        type: datetime

  - name: "Issues"
    columns:
      - name: Severity
        type: choice
        choices: [Low, Medium, High, Critical]
      - name: ReportedBy
        type: person
      - name: Resolved
        type: boolean

# ── Document Libraries ────────────────────────────────────────────────────────
libraries:
  - name: "Documents"
    folders:
      - "2024/Q1"
      - "2024/Q2"
      - "2024/Q3"
      - "2024/Q4"
      - "Archive/2023"
      - "Templates"

  - name: "Imported Files"
    folders:
      - "Dropbox Imports/2024"
      - "Dropbox Imports/Archive"
"""
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(example, encoding="utf-8")
        logger.info("Example template written to %s", out)
        return out
