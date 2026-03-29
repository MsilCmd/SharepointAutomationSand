#!/usr/bin/env python3
"""
scripts/generate_report.py

Generate an HTML or Excel dashboard from SharePoint list data.

Usage:
    python scripts/generate_report.py \
        --site "https://contoso.sharepoint.com/sites/proj" \
        --lists "Tasks" --lists "Issues" --lists "Import Log" \
        --format html \
        --output ./reports/dashboard.html
"""

import logging
import sys
from pathlib import Path

import click

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.reporting.dashboard import DashboardGenerator

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")


@click.command()
@click.option("--site", required=True, help="SharePoint site URL")
@click.option("--lists", multiple=True, required=True, help="List names to include. Repeatable.")
@click.option("--format", "fmt", type=click.Choice(["html", "excel"]), default="html")
@click.option("--output", default=None, help="Output file path")
@click.option("--title", default="SharePoint Dashboard", help="Dashboard title")
@click.option("--status-col", default="Status", help="Column to use for status breakdown")
def main(site, lists, fmt, output, title, status_col):
    """Generate a reporting dashboard from SharePoint list data."""

    gen = DashboardGenerator(site_url=site)

    if fmt == "html":
        out = gen.generate_list_dashboard(
            list_names=list(lists),
            output_path=Path(output) if output else None,
            title=title,
            status_column=status_col,
        )
    else:
        out = gen.generate_excel_report(
            list_names=list(lists),
            output_path=Path(output) if output else None,
        )

    click.echo(f"Report generated: {out}")


if __name__ == "__main__":
    main()
