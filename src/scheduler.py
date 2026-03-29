"""
src/scheduler.py

Lightweight job scheduler for recurring SharePoint automation tasks.
Wraps the `schedule` library (pip install schedule) with:
  - Logging on every run
  - Error isolation (one failing job doesn't kill the loop)
  - Graceful shutdown on SIGINT/SIGTERM
  - Optional Power Automate alerting on job failure

Intended for simple cron-style deployments (Docker container, VM, etc.).
For enterprise scale, replace this with Azure Functions or a proper task queue.

Example usage:
    python src/scheduler.py

Or import and configure programmatically:
    from src.scheduler import JobScheduler
    sched = JobScheduler()
    sched.add_job("Dropbox Import", my_import_fn, interval_minutes=60)
    sched.run_forever()
"""

from __future__ import annotations

import logging
import signal
import sys
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any

import schedule

logger = logging.getLogger(__name__)


@dataclass
class JobResult:
    name: str
    started_at: datetime
    finished_at: datetime
    success: bool
    error: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)


class JobScheduler:
    """
    Schedule and run recurring automation jobs.

    Example:
        from src.dropbox.import_pipeline import DropboxToSharePointPipeline

        pipeline = DropboxToSharePointPipeline(
            site_url="https://contoso.sharepoint.com/sites/proj",
            sp_library="Documents",
        )

        scheduler = JobScheduler(alert_url="https://prod-12...logic.azure.com/...")

        # Import from Dropbox every hour
        scheduler.add_job(
            "Dropbox Import",
            lambda: pipeline.run(dropbox_folder="/Reports", sp_folder="Imports"),
            interval_minutes=60,
        )

        # Also run at 08:00 daily
        scheduler.add_daily_job("Morning Report", generate_report_fn, at_time="08:00")

        scheduler.run_forever()
    """

    def __init__(self, alert_url: str | None = None) -> None:
        """
        Args:
            alert_url: Optional Power Automate HTTP trigger URL for failure alerts.
        """
        self._alert_url = alert_url
        self._history: list[JobResult] = []
        self._running = True
        self._register_signals()

    def _register_signals(self) -> None:
        signal.signal(signal.SIGINT, self._shutdown)
        signal.signal(signal.SIGTERM, self._shutdown)

    def _shutdown(self, *_) -> None:
        logger.info("Scheduler shutting down…")
        self._running = False

    # ── Job registration ──────────────────────────────────────────────────────

    def add_job(
        self,
        name: str,
        fn: Callable[[], Any],
        interval_minutes: int,
        run_immediately: bool = False,
    ) -> None:
        """
        Schedule a job to run every N minutes.

        Args:
            name:             Human-readable job name (used in logs and alerts).
            fn:               Zero-argument callable to execute.
            interval_minutes: How often to run.
            run_immediately:  Whether to execute once right now before scheduling.
        """
        wrapped = self._wrap(name, fn)
        schedule.every(interval_minutes).minutes.do(wrapped)
        logger.info("Scheduled '%s' every %d min", name, interval_minutes)
        if run_immediately:
            wrapped()

    def add_daily_job(
        self,
        name: str,
        fn: Callable[[], Any],
        at_time: str = "00:00",
    ) -> None:
        """
        Schedule a job to run once per day at a specific wall-clock time.

        Args:
            at_time: 24-hour time string, e.g. "08:30".
        """
        wrapped = self._wrap(name, fn)
        schedule.every().day.at(at_time).do(wrapped)
        logger.info("Scheduled '%s' daily at %s", name, at_time)

    def add_hourly_job(self, name: str, fn: Callable[[], Any]) -> None:
        """Convenience: schedule a job every hour on the hour."""
        self.add_job(name, fn, interval_minutes=60)

    # ── Execution helpers ─────────────────────────────────────────────────────

    def _wrap(self, name: str, fn: Callable[[], Any]) -> Callable[[], None]:
        """Return a wrapped callable that logs, times, and records the result."""
        def _run() -> None:
            started = datetime.now(timezone.utc)
            logger.info("▶ Starting job: %s", name)
            try:
                fn()
                finished = datetime.now(timezone.utc)
                elapsed = (finished - started).total_seconds()
                result = JobResult(
                    name=name,
                    started_at=started,
                    finished_at=finished,
                    success=True,
                )
                self._history.append(result)
                logger.info("✓ Job '%s' completed in %.1fs", name, elapsed)
            except Exception as exc:
                finished = datetime.now(timezone.utc)
                elapsed = (finished - started).total_seconds()
                result = JobResult(
                    name=name,
                    started_at=started,
                    finished_at=finished,
                    success=False,
                    error=str(exc),
                )
                self._history.append(result)
                logger.error(
                    "✗ Job '%s' FAILED after %.1fs: %s", name, elapsed, exc, exc_info=True
                )
                self._send_alert(name, str(exc))

        return _run

    def _send_alert(self, job_name: str, error: str) -> None:
        """POST a failure alert to Power Automate (fire-and-forget)."""
        if not self._alert_url:
            return
        try:
            import threading
            import requests

            payload = {
                "event": "job_failure",
                "job_name": job_name,
                "error": error,
                "timestamp": datetime.now(timezone.utc).isoformat(),
            }
            t = threading.Thread(
                target=lambda: requests.post(self._alert_url, json=payload, timeout=10),
                daemon=True,
            )
            t.start()
        except Exception as exc:
            logger.warning("Could not send alert: %s", exc)

    # ── Run loop ──────────────────────────────────────────────────────────────

    def run_forever(self, poll_interval_seconds: int = 30) -> None:
        """
        Block and run the scheduler loop until SIGINT/SIGTERM.

        Args:
            poll_interval_seconds: How often to check for pending jobs.
        """
        logger.info("Scheduler started. Polling every %ds.", poll_interval_seconds)
        while self._running:
            schedule.run_pending()
            time.sleep(poll_interval_seconds)
        logger.info("Scheduler exited cleanly.")

    def run_all_now(self) -> None:
        """Immediately execute all registered jobs (useful for testing)."""
        schedule.run_all()

    @property
    def history(self) -> list[JobResult]:
        return list(self._history)


# ── Default scheduler (run as __main__) ──────────────────────────────────────

def _build_default_scheduler() -> JobScheduler:
    """
    Build and return the default production scheduler.
    Configure your jobs here for direct execution via `python src/scheduler.py`.
    """
    import os
    from config.settings import settings
    from src.dropbox.import_pipeline import DropboxToSharePointPipeline
    from src.reporting.dashboard import DashboardGenerator

    alert_url = os.environ.get("POWER_AUTOMATE_ALERT_URL")
    sched = JobScheduler(alert_url=alert_url)

    site_url = settings.sharepoint_site_url

    # ── Dropbox import every 2 hours ──────────────────────────────────────────
    pipeline = DropboxToSharePointPipeline(
        site_url=site_url,
        sp_library="Documents",
        audit_list="Import Log",
    )
    sched.add_job(
        "Dropbox Import",
        lambda: pipeline.run(
            dropbox_folder=os.environ.get("DROPBOX_SOURCE_FOLDER", ""),
            sp_folder="Dropbox Imports",
        ),
        interval_minutes=120,
        run_immediately=True,
    )

    # ── Daily dashboard generation at 07:00 ───────────────────────────────────
    gen = DashboardGenerator(site_url=site_url)
    sched.add_daily_job(
        "Dashboard Report",
        lambda: gen.generate_list_dashboard(
            list_names=["Import Log", "Tasks", "Issues"],
            title="Daily SharePoint Dashboard",
        ),
        at_time="07:00",
    )

    return sched


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(name)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    _build_default_scheduler().run_forever()
