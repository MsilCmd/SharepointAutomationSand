"""
src/sharepoint/power_automate_bridge.py

Trigger Power Automate (Logic Apps) flows from Python and receive
inbound webhooks from Power Automate into a local HTTP listener.

Two patterns are supported:

  1. OUTBOUND — Python calls a Power Automate HTTP trigger URL to kick off
     a cloud flow (e.g. send an approval email, kick off a Teams notification).

  2. INBOUND — A lightweight Flask/threading-based webhook server receives
     callbacks FROM Power Automate (e.g. after an approval is completed) and
     dispatches them to registered handler functions.

Power Automate side setup for pattern 1:
  - Add a "When an HTTP request is received" trigger to your flow.
  - Copy the auto-generated POST URL into POWER_AUTOMATE_TRIGGER_URL in .env.

Power Automate side setup for pattern 2:
  - Add an "HTTP" action at the end of your flow that POSTs a JSON body
    to your server's public URL (or use ngrok for local dev).
"""

from __future__ import annotations

import hashlib
import hmac
import json
import logging
import threading
from collections.abc import Callable
from http.server import BaseHTTPRequestHandler, HTTPServer
from typing import Any

import requests
from tenacity import retry, stop_after_attempt, wait_exponential

logger = logging.getLogger(__name__)


# ── Outbound: trigger Power Automate flows ────────────────────────────────────

class PowerAutomateTrigger:
    """
    Fire Power Automate HTTP-triggered flows from Python.

    Example:
        trigger = PowerAutomateTrigger(
            url="https://prod-12.westus.logic.azure.com/workflows/..."
        )
        trigger.fire({"action": "import_complete", "file_count": 42})
    """

    def __init__(self, url: str, secret: str | None = None) -> None:
        """
        Args:
            url:    The HTTP trigger URL from your Power Automate flow.
            secret: Optional shared secret for HMAC request signing.
        """
        self._url = url
        self._secret = secret

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=30))
    def fire(self, payload: dict[str, Any], timeout: int = 30) -> requests.Response:
        """
        Send a POST request to the Power Automate trigger URL.

        Args:
            payload: JSON-serialisable dict sent as the request body.
            timeout: HTTP timeout in seconds.

        Returns:
            The HTTP response from Power Automate.

        Raises:
            requests.HTTPError: If Power Automate returns a non-2xx status.
        """
        headers: dict[str, str] = {"Content-Type": "application/json"}
        body = json.dumps(payload)

        if self._secret:
            sig = hmac.new(
                self._secret.encode(), body.encode(), hashlib.sha256
            ).hexdigest()
            headers["X-Signature-SHA256"] = sig

        resp = requests.post(self._url, data=body, headers=headers, timeout=timeout)
        resp.raise_for_status()
        logger.info(
            "Power Automate flow triggered (status %d)", resp.status_code
        )
        return resp

    def fire_and_forget(self, payload: dict[str, Any]) -> None:
        """Trigger the flow in a background thread (non-blocking)."""
        t = threading.Thread(target=self.fire, args=(payload,), daemon=True)
        t.start()
        logger.debug("Power Automate trigger dispatched in background thread")


# ── Common trigger payloads ───────────────────────────────────────────────────

class FlowPayloads:
    """Pre-built payload factories for common automation scenarios."""

    @staticmethod
    def import_complete(
        file_count: int, succeeded: int, failed: int, report_url: str | None = None
    ) -> dict:
        return {
            "event": "dropbox_import_complete",
            "file_count": file_count,
            "succeeded": succeeded,
            "failed": failed,
            "report_url": report_url or "",
        }

    @staticmethod
    def list_item_created(list_name: str, item_id: str, fields: dict) -> dict:
        return {
            "event": "list_item_created",
            "list_name": list_name,
            "item_id": item_id,
            "fields": fields,
        }

    @staticmethod
    def migration_complete(
        source: str, destination: str, migrated: int, failed: int
    ) -> dict:
        return {
            "event": "migration_complete",
            "source": source,
            "destination": destination,
            "migrated": migrated,
            "failed": failed,
        }

    @staticmethod
    def alert(severity: str, message: str, context: dict | None = None) -> dict:
        return {
            "event": "alert",
            "severity": severity,  # "info" | "warning" | "error"
            "message": message,
            "context": context or {},
        }


# ── Inbound: receive callbacks from Power Automate ────────────────────────────

HandlerFn = Callable[[dict[str, Any]], None]


class WebhookServer:
    """
    Lightweight HTTP server that receives POST callbacks from Power Automate.

    Handlers are registered per event type and called synchronously.
    The server runs in a daemon thread so it doesn't block the main program.

    Example:
        server = WebhookServer(port=8080, secret="my-shared-secret")

        @server.on("import_complete")
        def handle_import(payload):
            print("Import done:", payload)

        server.start()
        # ... do other work ...
        server.stop()
    """

    def __init__(self, port: int = 8080, secret: str | None = None) -> None:
        self._port = port
        self._secret = secret
        self._handlers: dict[str, list[HandlerFn]] = {}
        self._server: HTTPServer | None = None
        self._thread: threading.Thread | None = None

    def on(self, event: str) -> Callable[[HandlerFn], HandlerFn]:
        """Decorator to register a handler for a specific event type."""
        def decorator(fn: HandlerFn) -> HandlerFn:
            self._handlers.setdefault(event, []).append(fn)
            logger.debug("Registered handler for event '%s': %s", event, fn.__name__)
            return fn
        return decorator

    def register(self, event: str, handler: HandlerFn) -> None:
        """Imperatively register a handler."""
        self._handlers.setdefault(event, []).append(handler)

    def start(self) -> None:
        """Start the webhook server in a background daemon thread."""
        handlers = self._handlers
        secret = self._secret

        class _Handler(BaseHTTPRequestHandler):
            def log_message(self, fmt, *args):  # silence default access log
                logger.debug(fmt, *args)

            def do_POST(self):
                length = int(self.headers.get("Content-Length", 0))
                raw = self.rfile.read(length)

                # Optional HMAC verification
                if secret:
                    sig = self.headers.get("X-Signature-SHA256", "")
                    expected = hmac.new(
                        secret.encode(), raw, hashlib.sha256
                    ).hexdigest()
                    if not hmac.compare_digest(sig, expected):
                        self.send_response(403)
                        self.end_headers()
                        logger.warning("Webhook signature mismatch — request rejected")
                        return

                try:
                    payload: dict = json.loads(raw)
                except json.JSONDecodeError:
                    self.send_response(400)
                    self.end_headers()
                    return

                event = payload.get("event", "__default__")
                matched = handlers.get(event, []) + handlers.get("*", [])

                for handler in matched:
                    try:
                        handler(payload)
                    except Exception as exc:
                        logger.error(
                            "Handler %s raised: %s", handler.__name__, exc, exc_info=True
                        )

                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(b'{"status":"ok"}')

        self._server = HTTPServer(("0.0.0.0", self._port), _Handler)
        self._thread = threading.Thread(
            target=self._server.serve_forever, daemon=True, name="WebhookServer"
        )
        self._thread.start()
        logger.info("Webhook server listening on port %d", self._port)

    def stop(self) -> None:
        """Gracefully shut down the server."""
        if self._server:
            self._server.shutdown()
            logger.info("Webhook server stopped")
