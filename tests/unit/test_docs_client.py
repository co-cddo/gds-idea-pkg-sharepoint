"""Unit tests for the DocsClient.

Tests document library operations using a mocked session, mirroring the
approach in test_list_client.py.
"""

from unittest.mock import MagicMock, patch

import pytest

from box2.sharepoint.docs_client import DocsClient
from box2.sharepoint.exceptions import SharePointAPIError

# ============================================================================
# Fixtures
# ============================================================================

SITE_ID = "contoso.sharepoint.com,abc-123,def-456"
DRIVE_ID = "drive-id-789"
LIBRARY_NAME = "Documents"


def _drives_response(drive_id: str = DRIVE_ID, library_name: str = LIBRARY_NAME) -> dict:
    """Build a canned Graph API drives list response."""
    return {
        "value": [
            {"id": drive_id, "name": library_name},
            {"id": "other-drive-id", "name": "Other Library"},
        ]
    }


def _file_item(
    item_id: str = "file-1",
    name: str = "report.pdf",
    last_modified: str = "2026-02-20T12:00:00Z",
    download_url: str | None = "https://cdn.example.com/report.pdf",
    drive_id: str = DRIVE_ID,
) -> dict:
    """Build a canned Graph API drive item (file) response."""
    item: dict = {
        "id": item_id,
        "name": name,
        "size": 1024,
        "lastModifiedDateTime": last_modified,
        "createdDateTime": "2026-02-20T10:00:00Z",
        "webUrl": f"https://contoso.sharepoint.com/{name}",
        "parentReference": {"driveId": drive_id, "path": "/drive/root:"},
        "file": {"mimeType": "application/pdf"},
    }
    if download_url:
        item["@microsoft.graph.downloadUrl"] = download_url
    return item


def _folder_item(item_id: str = "folder-1", name: str = "Reports") -> dict:
    """Build a canned Graph API drive item (folder) response."""
    return {
        "id": item_id,
        "name": name,
        "lastModifiedDateTime": "2026-02-20T11:00:00Z",
        "folder": {"childCount": 3},
        "parentReference": {"driveId": DRIVE_ID, "path": "/drive/root:"},
    }


def _deleted_item(item_id: str = "deleted-1", name: str = "old.txt") -> dict:
    """Build a canned Graph API deleted drive item response."""
    return {
        "id": item_id,
        "name": name,
        "file": {"mimeType": "text/plain"},
        "deleted": {"state": "deleted"},
    }


@pytest.fixture
def mock_session():
    """Create a MagicMock session that returns canned Graph API responses."""
    session = MagicMock()
    session.resolve_site_id.return_value = SITE_ID
    session.request.return_value = _drives_response()
    return session


@pytest.fixture
def client(mock_session):
    """Create a DocsClient connected to the mock session."""
    return DocsClient(mock_session, library_name=LIBRARY_NAME)


# ============================================================================
# Constructor Tests
# ============================================================================


def test_constructor_resolves_site_id(mock_session):
    """Constructor should call session.resolve_site_id()."""
    DocsClient(mock_session, library_name=LIBRARY_NAME)
    mock_session.resolve_site_id.assert_called_once()


def test_constructor_resolves_drive_id(mock_session):
    """Constructor should look up the drive by name via GET /sites/{site_id}/drives."""
    client = DocsClient(mock_session, library_name=LIBRARY_NAME)
    assert client._drive_id == DRIVE_ID


def test_constructor_defaults_to_documents_library(mock_session):
    """Constructor should default to 'Documents' library when no name is given."""
    client = DocsClient(mock_session)
    assert client.library_name == "Documents"


def test_constructor_raises_when_library_not_found(mock_session):
    """Constructor should raise SharePointAPIError when the library name doesn't match."""
    with pytest.raises(SharePointAPIError, match="not found") as exc_info:
        DocsClient(mock_session, library_name="Nonexistent Library")
    assert exc_info.value.status_code == 404
    assert exc_info.value.error_code == "driveNotFound"


def test_constructor_error_lists_available_libraries(mock_session):
    """The not-found error should include available library names."""
    with pytest.raises(SharePointAPIError, match="Documents"):
        DocsClient(mock_session, library_name="Nonexistent Library")


# ============================================================================
# resource_path and supported_change_types Tests
# ============================================================================


def test_resource_path_returns_drive_root(client):
    """resource_path should return /drives/{drive_id}/root."""
    assert client.resource_path == f"/drives/{DRIVE_ID}/root"


def test_supported_change_types_returns_updated_only(client):
    """SharePoint drives only support 'updated' change notifications."""
    assert client.supported_change_types == {"updated"}


# ============================================================================
# get_changed_files Tests
# ============================================================================


def test_get_changed_files_calls_delta_endpoint(client, mock_session):
    """get_changed_files should GET the delta endpoint with $select params."""
    mock_session.request.return_value = {"value": []}

    client.get_changed_files()

    call_args = mock_session.request.call_args
    assert call_args.args[0] == "GET"
    assert call_args.args[1] == f"/drives/{DRIVE_ID}/root/delta"
    assert "$select" in call_args.kwargs["params"]


def test_get_changed_files_uses_custom_select(client, mock_session):
    """get_changed_files should use custom select fields when provided."""
    mock_session.request.return_value = {"value": []}

    client.get_changed_files(select=["id", "name"])

    call_params = mock_session.request.call_args.kwargs["params"]
    assert call_params["$select"] == "id,name"


def test_get_changed_files_filters_out_folders(client, mock_session):
    """get_changed_files should exclude folders from results."""
    mock_session.request.return_value = {
        "value": [
            _file_item(item_id="f1", name="doc.pdf"),
            _folder_item(),
        ]
    }

    result = client.get_changed_files()

    assert len(result) == 1
    assert result[0]["name"] == "doc.pdf"


def test_get_changed_files_filters_out_deleted_items(client, mock_session):
    """get_changed_files should exclude deleted items from results."""
    mock_session.request.return_value = {
        "value": [
            _file_item(item_id="f1", name="doc.pdf"),
            _deleted_item(),
        ]
    }

    result = client.get_changed_files()

    assert len(result) == 1
    assert result[0]["name"] == "doc.pdf"


def test_get_changed_files_sorts_by_last_modified_descending(client, mock_session):
    """get_changed_files should return files sorted most-recent-first."""
    mock_session.request.return_value = {
        "value": [
            _file_item(item_id="f1", name="old.pdf", last_modified="2026-02-18T10:00:00Z"),
            _file_item(item_id="f2", name="new.pdf", last_modified="2026-02-20T15:00:00Z"),
            _file_item(item_id="f3", name="mid.pdf", last_modified="2026-02-19T12:00:00Z"),
        ]
    }

    result = client.get_changed_files()

    assert [f["name"] for f in result] == ["new.pdf", "mid.pdf", "old.pdf"]


def test_get_changed_files_returns_empty_list_when_no_files(client, mock_session):
    """get_changed_files should return an empty list when delta has no files."""
    mock_session.request.return_value = {"value": [_folder_item()]}

    result = client.get_changed_files()

    assert result == []


# ============================================================================
# get_latest_changed_file Tests
# ============================================================================


def test_get_latest_changed_file_returns_most_recent(client, mock_session):
    """get_latest_changed_file should return the most recently modified file."""
    mock_session.request.return_value = {
        "value": [
            _file_item(item_id="f1", name="old.pdf", last_modified="2026-02-18T10:00:00Z"),
            _file_item(item_id="f2", name="new.pdf", last_modified="2026-02-20T15:00:00Z"),
        ]
    }

    result = client.get_latest_changed_file()

    assert result["name"] == "new.pdf"


def test_get_latest_changed_file_raises_when_no_files(client, mock_session):
    """get_latest_changed_file should raise SharePointAPIError when no files are found."""
    mock_session.request.return_value = {"value": []}

    with pytest.raises(SharePointAPIError, match="No files found") as exc_info:
        client.get_latest_changed_file()
    assert exc_info.value.status_code == 404
    assert exc_info.value.error_code == "noFilesFound"


def test_get_latest_changed_file_passes_select(client, mock_session):
    """get_latest_changed_file should pass select through to get_changed_files."""
    mock_session.request.return_value = {"value": [_file_item()]}

    client.get_latest_changed_file(select=["id", "name"])

    call_params = mock_session.request.call_args.kwargs["params"]
    assert call_params["$select"] == "id,name"


# ============================================================================
# download_file Tests
# ============================================================================


def test_download_file_uses_pre_authenticated_url(client, tmp_path):
    """download_file should stream-download using @microsoft.graph.downloadUrl."""
    file_meta = _file_item(name="test.pdf", download_url="https://cdn.example.com/test.pdf")
    download_dir = str(tmp_path / "downloads")

    with patch("box2.sharepoint.docs_client.httpx.stream") as mock_stream:
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.iter_bytes.return_value = [b"file content"]
        mock_stream.return_value.__enter__ = MagicMock(return_value=mock_response)
        mock_stream.return_value.__exit__ = MagicMock(return_value=False)

        result = client.download_file(file_meta, download_dir=download_dir)

    mock_stream.assert_called_once_with("GET", "https://cdn.example.com/test.pdf")
    assert result.endswith("test.pdf")


def test_download_file_falls_back_to_fetching_url(client, mock_session, tmp_path):
    """download_file should fetch a fresh item when @microsoft.graph.downloadUrl is missing."""
    file_meta = _file_item(name="test.pdf", download_url=None)
    download_dir = str(tmp_path / "downloads")

    # The fallback call to get a fresh item
    mock_session.request.return_value = {
        "@microsoft.graph.downloadUrl": "https://cdn.example.com/fresh-test.pdf",
    }

    with patch("box2.sharepoint.docs_client.httpx.stream") as mock_stream:
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.iter_bytes.return_value = [b"file content"]
        mock_stream.return_value.__enter__ = MagicMock(return_value=mock_response)
        mock_stream.return_value.__exit__ = MagicMock(return_value=False)

        client.download_file(file_meta, download_dir=download_dir)

    # Should have called session.request to fetch the download URL
    mock_session.request.assert_called_with("GET", f"/drives/{DRIVE_ID}/items/file-1")
    mock_stream.assert_called_once_with("GET", "https://cdn.example.com/fresh-test.pdf")


def test_download_file_raises_when_no_url_available(client, mock_session, tmp_path):
    """download_file should raise SharePointAPIError when no download URL can be obtained."""
    file_meta = _file_item(name="test.pdf", download_url=None)
    download_dir = str(tmp_path / "downloads")

    # Fallback also returns no URL
    mock_session.request.return_value = {}

    with pytest.raises(SharePointAPIError, match="No download URL") as exc_info:
        client.download_file(file_meta, download_dir=download_dir)
    assert exc_info.value.error_code == "downloadUrlNotFound"


def test_download_file_raises_on_http_error(client, tmp_path):
    """download_file should raise SharePointAPIError when the download HTTP request fails."""
    file_meta = _file_item(name="test.pdf")
    download_dir = str(tmp_path / "downloads")

    with patch("box2.sharepoint.docs_client.httpx.stream") as mock_stream:
        mock_response = MagicMock()
        mock_response.status_code = 403
        mock_stream.return_value.__enter__ = MagicMock(return_value=mock_response)
        mock_stream.return_value.__exit__ = MagicMock(return_value=False)

        with pytest.raises(SharePointAPIError, match="download failed") as exc_info:
            client.download_file(file_meta, download_dir=download_dir)
        assert exc_info.value.status_code == 403
        assert exc_info.value.error_code == "downloadFailed"


def test_download_file_creates_directory(client, tmp_path):
    """download_file should create the download directory if it doesn't exist."""
    file_meta = _file_item(name="test.pdf")
    download_dir = str(tmp_path / "new" / "nested" / "dir")

    with patch("box2.sharepoint.docs_client.httpx.stream") as mock_stream:
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.iter_bytes.return_value = [b"data"]
        mock_stream.return_value.__enter__ = MagicMock(return_value=mock_response)
        mock_stream.return_value.__exit__ = MagicMock(return_value=False)

        result = client.download_file(file_meta, download_dir=download_dir)

    assert result == f"{download_dir}/test.pdf"


# ============================================================================
# list_files Tests
# ============================================================================


def test_list_files_calls_root_children_by_default(client, mock_session):
    """list_files should GET /drives/{id}/root/children when no folder_path is given."""
    mock_session.request.return_value = {"value": []}

    client.list_files()

    mock_session.request.assert_called_with("GET", f"/drives/{DRIVE_ID}/root/children")


def test_list_files_calls_path_children_for_subfolder(client, mock_session):
    """list_files should use the path-based URL when folder_path is provided."""
    mock_session.request.return_value = {"value": []}

    client.list_files(folder_path="Reports/2026")

    mock_session.request.assert_called_with("GET", f"/drives/{DRIVE_ID}/root:/Reports/2026:/children")


def test_list_files_returns_value_list(client, mock_session):
    """list_files should return the 'value' array from the response."""
    items = [_file_item(), _folder_item()]
    mock_session.request.return_value = {"value": items}

    result = client.list_files()

    assert result == items


# ============================================================================
# get_file Tests
# ============================================================================


def test_get_file_calls_correct_path(client, mock_session):
    """get_file should GET /drives/{id}/items/{item_id}."""
    expected = _file_item(item_id="item-42")
    mock_session.request.return_value = expected

    result = client.get_file("item-42")

    mock_session.request.assert_called_with("GET", f"/drives/{DRIVE_ID}/items/item-42")
    assert result["id"] == "item-42"


# ============================================================================
# upload_file Tests
# ============================================================================


def test_upload_file_calls_correct_path(client, mock_session):
    """upload_file should PUT to /drives/{id}/root:/{name}:/content with auth and octet-stream."""
    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.content = b'{"id": "new-file-1"}'
    mock_response.json.return_value = {"id": "new-file-1", "name": "test.txt"}
    mock_session._http.request.return_value = mock_response
    mock_session.get_token.return_value = "fake-token"

    client.upload_file("test.txt", b"hello world")

    call_args = mock_session._http.request.call_args
    assert call_args.args[0] == "PUT"
    assert call_args.args[1] == f"/drives/{DRIVE_ID}/root:/test.txt:/content"
    assert call_args.kwargs["headers"]["Content-Type"] == "application/octet-stream"
    assert call_args.kwargs["headers"]["Authorization"] == "Bearer fake-token"
    assert call_args.kwargs["content"] == b"hello world"


def test_upload_file_with_folder_path(client, mock_session):
    """upload_file should include the folder in the path when folder_path is provided."""
    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.content = b'{"id": "new-file-1"}'
    mock_response.json.return_value = {"id": "new-file-1", "name": "test.txt"}
    mock_session._http.request.return_value = mock_response
    mock_session.get_token.return_value = "fake-token"

    client.upload_file("test.txt", b"hello", folder_path="Reports/2026")

    path = mock_session._http.request.call_args.args[1]
    assert path == f"/drives/{DRIVE_ID}/root:/Reports/2026/test.txt:/content"


def test_upload_file_returns_response(client, mock_session):
    """upload_file should return the created driveItem metadata dict."""
    expected = {"id": "new-file-1", "name": "test.txt", "size": 11}
    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.content = b"..."
    mock_response.json.return_value = expected
    mock_session._http.request.return_value = mock_response
    mock_session.get_token.return_value = "fake-token"

    result = client.upload_file("test.txt", b"hello world")

    assert result == expected


def test_upload_file_raises_on_error(client, mock_session):
    """upload_file should raise SharePointAPIError on a 4xx/5xx response."""
    mock_response = MagicMock()
    mock_response.status_code = 403
    mock_response.content = b'{"error": {"code": "accessDenied", "message": "Forbidden"}}'
    mock_response.json.return_value = {"error": {"code": "accessDenied", "message": "Forbidden"}}
    mock_response.text = "Forbidden"
    mock_session._http.request.return_value = mock_response
    mock_session.get_token.return_value = "fake-token"

    with pytest.raises(SharePointAPIError, match="upload failed") as exc_info:
        client.upload_file("test.txt", b"data")
    assert exc_info.value.status_code == 403
    assert exc_info.value.error_code == "accessDenied"


# ============================================================================
# delete_file Tests
# ============================================================================


def test_delete_file_sends_delete(client, mock_session):
    """delete_file should DELETE /drives/{id}/items/{item_id}."""
    mock_session.request.return_value = {}

    client.delete_file("item-42")

    mock_session.request.assert_called_with("DELETE", f"/drives/{DRIVE_ID}/items/item-42")


# ============================================================================
# DocsClient Protocol Compliance
# ============================================================================


def test_docs_client_satisfies_subscribable_resource(mock_session):
    """DocsClient should expose resource_path and supported_change_types."""
    client = DocsClient(mock_session, library_name=LIBRARY_NAME)

    assert client.resource_path == f"/drives/{DRIVE_ID}/root"
    assert client.supported_change_types == {"updated"}
