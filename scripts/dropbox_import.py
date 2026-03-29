#!/usr/bin/env python3
"""
scripts/dropbox_import.py

CLI entry point for the Dropbox → SharePoint import pipeline.

Usage:
    python scripts/dropbox_import.py \
        --site "https://contoso.sharepoint.com/sites/proj" \
        --sp-lib "Documents" \
        --dropbox-path "/Reports" \
        --sp-folder "Dropbox Imports" \
        --ext .pdf --ext .xlsx \
        --audit-list "Import Log" \
        --dry-run
"""

import logging
import sys
from pathlib import Path

import click
from rich.console import Console
from rich.table import Table

# Allow running from project root
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.dropbox.import_pipeline import DropboxToSharePointPipeline, ImportResult

console = Console()
logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")


@click.command()
@click.option("--site", required=True, help="SharePoint site URL")
@click.option("--sp-lib", required=True, help="Destination SharePoint document library")
@click.option("--dropbox-path", default="", help="Source Dropbox folder path")
@click.option("--sp-folder", default="/", help="Destination folder within SP library")
@click.option("--ext", multiple=True, help="File extensions to include (e.g. .pdf). Repeatable.")
@click.option("--audit-list", default=None, help="SharePoint list for import audit log")
@click.option("--recursive/--no-recursive", default=False)
@click.option("--dry-run/--no-dry-run", default=False)
def main(site, sp_lib, dropbox_path, sp_folder, ext, audit_list, recursive, dry_run):
    """Import files from Dropbox into a SharePoint document library."""

    extensions = list(ext) or None

    if dry_run:
        console.print("[yellow]DRY RUN — no files will be transferred[/yellow]")

    pipeline = DropboxToSharePointPipeline(
        site_url=site,
        sp_library=sp_lib,
        audit_list=audit_list,
        on_progress=lambda r: _print_result(r),
    )

    summary = pipeline.run(
        dropbox_folder=dropbox_path,
        sp_folder=sp_folder,
        extensions=extensions,
        recursive=recursive,
        dry_run=dry_run,
    )

    # Print summary table
    table = Table(title="Import Summary", show_header=True)
    table.add_column("Metric")
    table.add_column("Value", justify="right")
    table.add_row("Total files", str(summary.total))
    table.add_row("Succeeded", f"[green]{summary.succeeded}[/green]")
    table.add_row("Failed", f"[red]{summary.failed}[/red]")
    table.add_row("Success rate", f"{summary.success_rate:.1f}%")
    console.print(table)

    sys.exit(0 if summary.failed == 0 else 1)


def _print_result(result: ImportResult) -> None:
    if result.success:
        console.print(f"  [green]✓[/green] {result.dropbox_path}")
    else:
        console.print(f"  [red]✗[/red] {result.dropbox_path} — {result.error}")


if __name__ == "__main__":
    main()
