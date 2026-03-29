"""
tests/conftest.py — Shared pytest fixtures.
"""

import os
import pytest


@pytest.fixture(autouse=True)
def mock_env(monkeypatch):
    """
    Inject fake environment variables so Settings() initialises without a real .env.
    Applied to EVERY test automatically.
    """
    env = {
        "AZURE_TENANT_ID": "fake-tenant-id",
        "AZURE_CLIENT_ID": "fake-client-id",
        "AZURE_CLIENT_SECRET": "fake-client-secret",
        "SHAREPOINT_SITE_URL": "https://contoso.sharepoint.com/sites/test",
        "DROPBOX_APP_KEY": "fake-app-key",
        "DROPBOX_APP_SECRET": "fake-app-secret",
        "DROPBOX_REFRESH_TOKEN": "fake-refresh-token",
    }
    for key, value in env.items():
        monkeypatch.setenv(key, value)

