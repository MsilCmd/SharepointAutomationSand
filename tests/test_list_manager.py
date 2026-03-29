"""
tests/test_list_manager.py

Unit tests for SharePointListManager.
Uses `responses` to mock Graph API HTTP calls.
"""

from __future__ import annotations

import pytest
import responses as resp_mock

from src.sharepoint.list_manager import SharePointListManager

MOCK_SITE_ID = "contoso.sharepoint.com,abc123,def456"
MOCK_SITE_URL = "https://contoso.sharepoint.com/sites/test"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# ── Fixtures ──────────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def mock_auth(mocker):
    """Stub out Azure AD token acquisition."""
    mock = mocker.patch(
        "src.auth.auth_manager.AzureAuthManager.get_token",
        return_value="fake-token",
    )
    return mock


@pytest.fixture()
def mgr(mocker):
    """Return a ListManager with a stubbed site ID."""
    mocker.patch(
        "src.sharepoint.site_resolver.SiteResolver.get_site_id",
        return_value=MOCK_SITE_ID,
    )
    return SharePointListManager(MOCK_SITE_URL)


# ── Tests ─────────────────────────────────────────────────────────────────────

@resp_mock.activate
def test_get_all_items_single_page(mgr):
    """Should return all items from a single-page response."""
    resp_mock.add(
        resp_mock.GET,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items",
        json={
            "value": [
                {"id": "1", "fields": {"Title": "Task A", "Status": "Open"}},
                {"id": "2", "fields": {"Title": "Task B", "Status": "Closed"}},
            ]
        },
        status=200,
    )
    items = mgr.get_all_items("Tasks")
    assert len(items) == 2
    assert items[0]["fields"]["Title"] == "Task A"


@resp_mock.activate
def test_get_all_items_pagination(mgr):
    """Should follow @odata.nextLink to fetch all pages."""
    next_url = f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items?$skiptoken=abc"
    resp_mock.add(
        resp_mock.GET,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items",
        json={"value": [{"id": "1", "fields": {}}], "@odata.nextLink": next_url},
        status=200,
    )
    resp_mock.add(
        resp_mock.GET,
        next_url,
        json={"value": [{"id": "2", "fields": {}}]},
        status=200,
    )
    items = mgr.get_all_items("Tasks")
    assert len(items) == 2


@resp_mock.activate
def test_create_item(mgr):
    """Should POST the correct payload and return the created item."""
    resp_mock.add(
        resp_mock.POST,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items",
        json={"id": "42", "fields": {"Title": "New Task"}},
        status=201,
    )
    result = mgr.create_item("Tasks", {"Title": "New Task"})
    assert result["id"] == "42"
    # Verify payload
    sent = resp_mock.calls[0].request
    import json
    body = json.loads(sent.body)
    assert body["fields"]["Title"] == "New Task"


@resp_mock.activate
def test_update_item(mgr):
    """Should PATCH the fields endpoint."""
    resp_mock.add(
        resp_mock.PATCH,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items/5/fields",
        json={"Status": "Closed"},
        status=200,
    )
    result = mgr.update_item("Tasks", 5, {"Status": "Closed"})
    assert result["Status"] == "Closed"


@resp_mock.activate
def test_delete_item(mgr):
    """Should DELETE the item endpoint with no body."""
    resp_mock.add(
        resp_mock.DELETE,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items/7",
        status=204,
    )
    # Should not raise
    mgr.delete_item("Tasks", 7)
    assert len(resp_mock.calls) == 1


@resp_mock.activate
def test_upsert_creates_when_not_found(mgr):
    """Upsert should create a new item when no match is found."""
    # get_all_items returns empty list
    resp_mock.add(
        resp_mock.GET,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items",
        json={"value": []},
        status=200,
    )
    resp_mock.add(
        resp_mock.POST,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items",
        json={"id": "99", "fields": {"Title": "Unique"}},
        status=201,
    )
    result = mgr.upsert_item("Tasks", "Title", "Unique", {"Title": "Unique"})
    assert result["id"] == "99"


@resp_mock.activate
def test_upsert_updates_when_found(mgr):
    """Upsert should update the existing item when a match is found."""
    resp_mock.add(
        resp_mock.GET,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items",
        json={"value": [{"id": "10", "fields": {"Title": "Existing"}}]},
        status=200,
    )
    resp_mock.add(
        resp_mock.PATCH,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists/Tasks/items/10/fields",
        json={"Title": "Existing", "Status": "Done"},
        status=200,
    )
    result = mgr.upsert_item(
        "Tasks", "Title", "Existing", {"Title": "Existing", "Status": "Done"}
    )
    assert result["Status"] == "Done"
