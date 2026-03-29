#!/usr/bin/env python3
"""
scripts/migrate.py

CLI for content migration between SharePoint sites.

Usage:
    # Migrate a list
    python scripts/migrate.py list \
        --source "https://contoso.sharepoint.com/sites/old" \
        --dest "https://contoso.sharepoint.com/sites/new" \
        --list "Tasks" --dest-list "ProjectTasks"

    # Migrate a document library
    python scripts/migrate.py library \
        --source "https://contoso.sharepoint.com/sites/old" \
        --dest "https://contoso.sharepoint.com/sites/new" \
        --lib "Documents" --dest-lib "Archive" \
        --ext .pdf --ext .docx

    # Resume a previously interrupted migration
    python scripts/migrate.py list ... --checkpoint ./my_checkpoint.json
"""

import json
import logging
import sys
from pathlib import Path

import click
from rich.console import Console
from rich.table import Table

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.migration.content_migrator import ContentMigrator

console = Console()
logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")


@click.group()
def cli():
    """SharePoint content migration tools."""


# ── List migration command ────────────────────────────────────────────────────

@cli.command("list")
@click.option("--source", required=True, help="Source SharePoint site URL")
@click.option("--dest", required=True, help="Destination SharePoint site URL")
@click.option("--list", "list_name", required=True, help="Source list name")
@click.option("--dest-list", default=None, help="Destination list name (defaults to source name)")
@click.option("--filter", "filter_query", default=None, help="OData filter for items")
@click.option("--checkpoint", default="./migration_checkpoint.json", help="Checkpoint file path")
def migrate_list(source, dest, list_name, dest_list, filter_query, checkpoint):
    """Migrate items from one SharePoint list to another."""
    dest_list = dest_list or list_name
    migrator = ContentMigrator(source, dest, Path(checkpoint))

    console.print(f"Migrating list [cyan]{list_name}[/cyan] → [cyan]{dest_list}[/cyan]")
    report = migrator.migrate_list(list_name, dest_list, filter_query=filter_query)
    _print_report(report, "List Migration")


# ── Library migration command ─────────────────────────────────────────────────

@cli.command("library")
@click.option("--source", required=True, help="Source SharePoint site URL")
@click.option("--dest", required=True, help="Destination SharePoint site URL")
@click.option("--lib", required=True, help="Source library name")
@click.option("--dest-lib", default=None, help="Destination library name")
@click.option("--source-folder", default="/", help="Source folder path")
@click.option("--dest-folder", default="/", help="Destination folder path")
@click.option("--ext", multiple=True, help="File extensions to migrate")
@click.option("--checkpoint", default="./migration_checkpoint.json", help="Checkpoint file path")
def migrate_library(source, dest, lib, dest_lib, source_folder, dest_folder, ext, checkpoint):
    """Migrate documents from one SharePoint library to another."""
    dest_lib = dest_lib or lib
    migrator = ContentMigrator(source, dest, Path(checkpoint))
    extensions = list(ext) or None

    console.print(f"Migrating library [cyan]{lib}[/cyan] → [cyan]{dest_lib}[/cyan]")
    report = migrator.migrate_library(
        lib, dest_lib,
        source_folder=source_folder,
        dest_folder=dest_folder,
        extensions=extensions,
    )
    _print_report(report, "Library Migration")


# ── Shared output helper ──────────────────────────────────────────────────────

def _print_report(report: dict, title: str) -> None:
    table = Table(title=title, show_header=True)
    table.add_column("Metric")
    table.add_column("Value", justify="right")
    table.add_row("Migrated", f"[green]{report['migrated']}[/green]")
    table.add_row("Skipped (already done)", str(report["skipped"]))
    table.add_row("Failed", f"[red]{report['failed']}[/red]")
    console.print(table)

    if report.get("errors"):
        console.print("\n[red]Errors:[/red]")
        for err in report["errors"][:10]:
            console.print(f"  • {err}")
        if len(report["errors"]) > 10:
            console.print(f"  … and {len(report['errors']) - 10} more")
        sys.exit(1)


if __name__ == "__main__":
    cli()
