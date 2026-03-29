"""
tests/test_import_pipeline.py

Integration-style unit tests for DropboxToSharePointPipeline.
All external I/O is mocked.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from src.dropbox.import_pipeline import DropboxToSharePointPipeline, ImportSummary


@pytest.fixture()
def pipeline(tmp_path, mocker):
    """Return a pipeline with all external deps stubbed."""
    mocker.patch("src.dropbox.import_pipeline.DropboxManager")
    mocker.patch("src.dropbox.import_pipeline.DocumentManager")
    mocker.patch("src.dropbox.import_pipeline.SharePointListManager")
    return DropboxToSharePointPipeline(
        site_url="https://contoso.sharepoint.com/sites/test",
        sp_library="Documents",
        audit_list="Import Log",
    )


def _make_file_meta(path: str, name: str):
    m = MagicMock()
    m.path_display = path
    m.name = name
    return m


def test_run_dry_run_returns_success(pipeline):
    """Dry run should report all items as success without calling download/upload."""
    pipeline._dbx.list_files.return_value = [
        _make_file_meta("/Reports/file1.pdf", "file1.pdf"),
        _make_file_meta("/Reports/file2.xlsx", "file2.xlsx"),
    ]

    summary: ImportSummary = pipeline.run(dry_run=True)

    assert summary.total == 2
    assert summary.succeeded == 2
    assert summary.failed == 0
    # No actual downloads in dry-run mode
    pipeline._dbx.download_file.assert_not_called()
    pipeline._docs.upload.assert_not_called()


def test_run_transfers_files(pipeline, tmp_path):
    """Normal run should download then upload each file."""
    fake_local = tmp_path / "file1.pdf"
    fake_local.write_bytes(b"fake pdf content")

    pipeline._dbx.list_files.return_value = [
        _make_file_meta("/Reports/file1.pdf", "file1.pdf"),
    ]
    pipeline._dbx.download_file.return_value = fake_local
    pipeline._docs.upload.return_value = {"id": "abc"}

    summary = pipeline.run()

    assert summary.total == 1
    assert summary.succeeded == 1
    assert summary.failed == 0
    pipeline._dbx.download_file.assert_called_once()
    pipeline._docs.upload.assert_called_once()


def test_run_records_failure_on_exception(pipeline, tmp_path):
    """A download error should be recorded as failed, not crash the pipeline."""
    pipeline._dbx.list_files.return_value = [
        _make_file_meta("/Reports/bad.pdf", "bad.pdf"),
    ]
    pipeline._dbx.download_file.side_effect = ConnectionError("Dropbox unreachable")

    summary = pipeline.run()

    assert summary.total == 1
    assert summary.failed == 1
    assert summary.succeeded == 0
    assert "Dropbox unreachable" in summary.results[0].error


def test_progress_callback_called(pipeline, tmp_path):
    """on_progress callback should be called once per file."""
    fake_local = tmp_path / "f.pdf"
    fake_local.write_bytes(b"x")

    pipeline._dbx.list_files.return_value = [
        _make_file_meta("/f.pdf", "f.pdf"),
    ]
    pipeline._dbx.download_file.return_value = fake_local

    calls = []
    pipeline._on_progress = lambda r: calls.append(r)
    pipeline.run()

    assert len(calls) == 1


# ── Provisioner tests ─────────────────────────────────────────────────────────

"""
tests/test_provisioner.py embedded here for brevity.
"""

import responses as resp_mock
from src.provisioning.provisioner import SharePointProvisioner

MOCK_SITE_ID = "contoso.sharepoint.com,aaa,bbb"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


@pytest.fixture()
def provisioner(mocker):
    mocker.patch(
        "src.auth.auth_manager.AzureAuthManager.get_token",
        return_value="fake-token",
    )
    mocker.patch(
        "src.sharepoint.site_resolver.SiteResolver.get_site_id",
        return_value=MOCK_SITE_ID,
    )
    return SharePointProvisioner("https://contoso.sharepoint.com/sites/test")


@resp_mock.activate
def test_create_list(provisioner):
    """Should POST to Graph lists endpoint with correct payload."""
    resp_mock.add(
        resp_mock.POST,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists",
        json={"id": "list-guid-1", "displayName": "My List"},
        status=201,
    )

    result = provisioner.create_list(
        "My List",
        columns=[{"name": "Status", "type": "choice", "choices": ["Open", "Done"]}],
    )

    assert result["id"] == "list-guid-1"
    import json
    body = json.loads(resp_mock.calls[0].request.body)
    assert body["displayName"] == "My List"
    assert "columns" in body
    assert body["columns"][0]["name"] == "Status"


@resp_mock.activate
def test_get_or_create_list_skips_existing(provisioner):
    """Should not create a list that already exists."""
    resp_mock.add(
        resp_mock.GET,
        f"{GRAPH_BASE}/sites/{MOCK_SITE_ID}/lists",
        json={"value": [{"id": "existing-id", "displayName": "Tasks"}]},
        status=200,
    )

    result = provisioner.get_or_create_list("Tasks")

    assert result["id"] == "existing-id"
    # No POST should have been made
    assert all(c.request.method == "GET" for c in resp_mock.calls)
