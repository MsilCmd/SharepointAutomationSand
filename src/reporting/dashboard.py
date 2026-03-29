"""
src/reporting/dashboard.py

Generate HTML and Excel reporting dashboards from SharePoint data.

Reports include:
  - List item counts and status breakdowns
  - Document library storage usage
  - Import pipeline audit logs
  - Activity over time (item creation/modification)
"""

from __future__ import annotations

import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from jinja2 import Environment, DictLoader

from src.sharepoint.list_manager import SharePointListManager

logger = logging.getLogger(__name__)

# ── Jinja2 HTML template ──────────────────────────────────────────────────────

_HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{{ title }}</title>
  <style>
    body { font-family: Segoe UI, sans-serif; margin: 0; background: #f4f6f9; color: #333; }
    header { background: #0078d4; color: white; padding: 1.5rem 2rem; }
    header h1 { margin: 0; font-size: 1.6rem; }
    header p { margin: 0.25rem 0 0; opacity: 0.8; font-size: 0.9rem; }
    main { padding: 2rem; max-width: 1200px; margin: auto; }
    .kpi-row { display: flex; gap: 1rem; margin-bottom: 2rem; flex-wrap: wrap; }
    .kpi { background: white; border-radius: 8px; padding: 1.25rem 1.5rem;
           flex: 1; min-width: 160px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
    .kpi .value { font-size: 2.2rem; font-weight: 700; color: #0078d4; }
    .kpi .label { font-size: 0.8rem; color: #666; margin-top: 0.25rem; }
    .card { background: white; border-radius: 8px; padding: 1.5rem;
            margin-bottom: 1.5rem; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
    .card h2 { margin: 0 0 1rem; font-size: 1.1rem; color: #0078d4; }
    table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
    th { background: #f0f6ff; text-align: left; padding: 0.6rem 0.8rem; }
    td { padding: 0.5rem 0.8rem; border-bottom: 1px solid #eee; }
    tr:last-child td { border-bottom: none; }
    footer { text-align: center; padding: 1rem; font-size: 0.75rem; color: #999; }
  </style>
</head>
<body>
  <header>
    <h1>{{ title }}</h1>
    <p>Generated {{ generated_at }}</p>
  </header>
  <main>
    <div class="kpi-row">
      {% for kpi in kpis %}
      <div class="kpi">
        <div class="value">{{ kpi.value }}</div>
        <div class="label">{{ kpi.label }}</div>
      </div>
      {% endfor %}
    </div>

    {% for chart_html in charts %}
    <div class="card">{{ chart_html }}</div>
    {% endfor %}

    {% for table in tables %}
    <div class="card">
      <h2>{{ table.title }}</h2>
      <table>
        <thead><tr>{% for col in table.columns %}<th>{{ col }}</th>{% endfor %}</tr></thead>
        <tbody>
          {% for row in table.rows %}
          <tr>{% for cell in row %}<td>{{ cell }}</td>{% endfor %}</tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
    {% endfor %}
  </main>
  <footer>SharePoint Automation Suite &bull; {{ generated_at }}</footer>
</body>
</html>
"""


class DashboardGenerator:
    """
    Pull data from SharePoint and render HTML / Excel dashboards.

    Example:
        gen = DashboardGenerator(site_url="https://contoso.sharepoint.com/sites/proj")
        gen.generate_list_dashboard(
            list_names=["Tasks", "Issues", "Import Log"],
            output_path=Path("./reports/dashboard.html"),
        )
    """

    def __init__(self, site_url: str, output_dir: Path = Path("./reports")) -> None:
        self._lists = SharePointListManager(site_url)
        self._output_dir = Path(output_dir)
        self._output_dir.mkdir(parents=True, exist_ok=True)
        self._jinja = Environment(loader=DictLoader({"main": _HTML_TEMPLATE}))

    # ── Data fetching ─────────────────────────────────────────────────────────

    def _fetch_list_df(self, list_name: str) -> pd.DataFrame:
        items = self._lists.get_all_items(list_name)
        rows = []
        for item in items:
            row = {"_id": item.get("id")}
            row.update(item.get("fields", {}))
            rows.append(row)
        df = pd.DataFrame(rows)
        # Coerce datetime columns
        for col in df.columns:
            if "date" in col.lower() or "modified" in col.lower() or "created" in col.lower():
                df[col] = pd.to_datetime(df[col], errors="coerce", utc=True)
        return df

    # ── Chart helpers ─────────────────────────────────────────────────────────

    def _status_pie(self, df: pd.DataFrame, status_col: str, title: str) -> str:
        if status_col not in df.columns:
            return ""
        counts = df[status_col].value_counts().reset_index()
        counts.columns = [status_col, "count"]
        fig = px.pie(
            counts,
            names=status_col,
            values="count",
            title=title,
            color_discrete_sequence=px.colors.qualitative.Set2,
        )
        fig.update_layout(margin=dict(t=40, b=10, l=10, r=10), height=350)
        return fig.to_html(full_html=False, include_plotlyjs="cdn")

    def _timeline_bar(self, df: pd.DataFrame, date_col: str, title: str) -> str:
        if date_col not in df.columns:
            return ""
        df2 = df.copy()
        df2["_month"] = df2[date_col].dt.to_period("M").astype(str)
        counts = df2.groupby("_month").size().reset_index(name="count")
        fig = px.bar(
            counts,
            x="_month",
            y="count",
            title=title,
            labels={"_month": "Month", "count": "Items"},
            color_discrete_sequence=["#0078d4"],
        )
        fig.update_layout(margin=dict(t=40, b=10, l=10, r=10), height=300)
        return fig.to_html(full_html=False, include_plotlyjs=False)

    # ── Public API ────────────────────────────────────────────────────────────

    def generate_list_dashboard(
        self,
        list_names: list[str],
        output_path: Path | None = None,
        title: str = "SharePoint Dashboard",
        status_column: str = "Status",
    ) -> Path:
        """
        Build an HTML dashboard summarising multiple SharePoint lists.

        Args:
            list_names:    Lists to include in the report.
            output_path:   Output HTML path. Defaults to ./reports/dashboard.html.
            title:         Dashboard title shown in the header.
            status_column: Column to use for status breakdowns.

        Returns:
            Path to the generated HTML file.
        """
        if output_path is None:
            output_path = self._output_dir / "dashboard.html"

        all_data: dict[str, pd.DataFrame] = {}
        kpis: list[dict[str, Any]] = []
        charts: list[str] = []
        tables: list[dict[str, Any]] = []

        for list_name in list_names:
            try:
                df = self._fetch_list_df(list_name)
                all_data[list_name] = df
                kpis.append({"value": len(df), "label": f"{list_name} items"})

                # Status breakdown chart
                chart = self._status_pie(df, status_column, f"{list_name} by {status_column}")
                if chart:
                    charts.append(chart)

                # Activity timeline
                timeline_col = next(
                    (c for c in df.columns if "modified" in c.lower()), None
                )
                if timeline_col:
                    timeline = self._timeline_bar(
                        df, timeline_col, f"{list_name} — Activity Over Time"
                    )
                    if timeline:
                        charts.append(timeline)

                # Summary table (first 20 rows)
                preview_cols = [c for c in df.columns if not c.startswith("@") and c != "_id"][:6]
                preview = df[preview_cols].head(20)
                tables.append(
                    {
                        "title": f"{list_name} — Recent Items",
                        "columns": preview_cols,
                        "rows": preview.fillna("").values.tolist(),
                    }
                )
            except Exception as exc:
                logger.error("Could not process list '%s': %s", list_name, exc)

        # Render HTML
        template = self._jinja.get_template("main")
        html = template.render(
            title=title,
            generated_at=datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
            kpis=kpis,
            charts=charts,
            tables=tables,
        )
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(html, encoding="utf-8")
        logger.info("Dashboard written to %s", output_path)
        return output_path

    def generate_excel_report(
        self,
        list_names: list[str],
        output_path: Path | None = None,
    ) -> Path:
        """
        Export SharePoint list data to a multi-sheet Excel workbook.

        Returns:
            Path to the .xlsx file.
        """
        if output_path is None:
            output_path = self._output_dir / "sharepoint_report.xlsx"

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for list_name in list_names:
                try:
                    df = self._fetch_list_df(list_name)
                    # Excel sheet names max 31 chars
                    sheet_name = list_name[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    logger.info("Wrote sheet '%s' (%d rows)", sheet_name, len(df))
                except Exception as exc:
                    logger.error("Skipping list '%s': %s", list_name, exc)

        logger.info("Excel report written to %s", output_path)
        return output_path

