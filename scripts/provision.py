#!/usr/bin/env python3
"""
scripts/provision.py

CLI entry point for template-driven SharePoint provisioning.

Usage:
    # Apply a template
    python scripts/provision.py --template config/my_site.yaml

    # Dry-run: log what would be provisioned without touching SharePoint
    python scripts/provision.py --template config/my_site.yaml --dry-run

    # Generate an annotated example template
    python scripts/provision.py --generate-example config/example.yaml
"""

import json
import logging
import sys
from pathlib import Path

import click
from rich.console import Console
from rich.syntax import Syntax
from rich.table import Table

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.provisioning.template_engine import TemplateEngine

console = Console()
logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")


@click.command()
@click.option("--template", default=None, help="Path to YAML/JSON provisioning template")
@click.option("--dry-run/--no-dry-run", default=False, help="Preview without making changes")
@click.option("--generate-example", default=None, metavar="PATH",
              help="Write an annotated example template to PATH and exit")
def main(template, dry_run, generate_example):
    """Apply a declarative SharePoint provisioning template."""

    engine = TemplateEngine()

    if generate_example:
        out = engine.generate_example(generate_example)
        console.print(f"[green]Example template written to:[/green] {out}")
        content = Path(out).read_text()
        console.print(Syntax(content, "yaml", theme="monokai", line_numbers=True))
        sys.exit(0)

    if not template:
        console.print("[red]Error:[/red] --template is required (or use --generate-example)")
        sys.exit(1)

    if dry_run:
        console.print("[yellow]DRY RUN — no changes will be made to SharePoint[/yellow]\n")

    report = engine.apply(template, dry_run=dry_run)

    # ── Summary table ─────────────────────────────────────────────────────────
    table = Table(title="Provisioning Report", show_header=True)
    table.add_column("Resource")
    table.add_column("Name")
    table.add_column("Status")

    for lst in report.get("created_lists", []):
        status = "[yellow]dry-run[/yellow]" if lst.get("dry_run") else "[green]created[/green]"
        table.add_row("List", lst["name"], status)

    for lib in report.get("created_folders", []):
        status = "[yellow]dry-run[/yellow]" if lib.get("dry_run") else "[green]created[/green]"
        table.add_row("Library folders", lib["library"], status)

    for err in report.get("errors", []):
        resource = err.get("list") or err.get("library", "?")
        table.add_row("ERROR", resource, f"[red]{err['error'][:60]}[/red]")

    console.print(table)

    if report.get("errors"):
        sys.exit(1)


if __name__ == "__main__":
    main()

