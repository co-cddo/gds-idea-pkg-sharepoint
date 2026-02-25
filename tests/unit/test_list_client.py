from unittest.mock import MagicMock, call, patch

import pytest

from box2.sharepoint.exceptions import SharePointAPIError
from box2.sharepoint.list_client import ListClient

# ============================================================================
# Fixtures
# ============================================================================

SITE_ID = "contoso.sharepoint.com,abc-123,def-456"
LIST_ID = "list-id-789"
LIST_NAME = "My Test List"


@pytest.fixture
def mock_session():
    """Create a MagicMock session that returns canned Graph API responses."""
    session = MagicMock()
    session.resolve_site_id.return_value = SITE_ID
    session.request.return_value = {
        "value": [
            {"id": LIST_ID, "displayName": LIST_NAME},
            {"id": "other-list-id", "displayName": "Other List"},
        ]
    }
    return session


@pytest.fixture
def client(mock_session):
    """Create a ListClient connected to the mock session."""
    return ListClient(mock_session, list_name=LIST_NAME)


# ============================================================================
# Constructor Tests
# ============================================================================


def test_constructor_resolves_site_id(mock_session):
    """Constructor should call session.resolve_site_id()."""
    ListClient(mock_session, list_name=LIST_NAME)
    mock_session.resolve_site_id.assert_called_once()


def test_constructor_resolves_list_id(mock_session):
    """Constructor should look up the list by name via GET /sites/{site_id}/lists."""
    client = ListClient(mock_session, list_name=LIST_NAME)
    assert client._list_id == LIST_ID


def test_constructor_raises_when_list_not_found(mock_session):
    """Constructor should raise SharePointAPIError when the list name doesn't match."""
    with pytest.raises(SharePointAPIError, match="not found") as exc_info:
        ListClient(mock_session, list_name="Nonexistent List")
    assert exc_info.value.status_code == 404
    assert exc_info.value.error_code == "listNotFound"


def test_constructor_error_lists_available_lists(mock_session):
    """The not-found error should include available list names."""
    with pytest.raises(SharePointAPIError, match="My Test List"):
        ListClient(mock_session, list_name="Nonexistent List")


# ============================================================================
# ListClient.new Tests
# ============================================================================


def test_new_creates_list_then_connects(mock_session):
    """ListClient.new should POST to create the list, then return a connected client."""
    client = ListClient.new(mock_session, list_name=LIST_NAME)

    # First call is the POST to create, subsequent calls are from the constructor
    create_call = mock_session.request.call_args_list[0]
    assert create_call == call(
        "POST",
        f"/sites/{SITE_ID}/lists",
        json={"displayName": LIST_NAME, "list": {"template": "genericList"}},
    )
    assert client.list_name == LIST_NAME
    assert client._list_id == LIST_ID


# ============================================================================
# get_items Tests
# ============================================================================


def test_get_items_calls_correct_path(client, mock_session):
    """get_items should GET the items endpoint with $expand=fields."""
    mock_session.request.return_value = {"value": []}

    client.get_items()

    mock_session.request.assert_called_with(
        "GET",
        f"/sites/{SITE_ID}/lists/{LIST_ID}/items",
        params={"$expand": "fields"},
    )


def test_get_items_returns_value_list(client, mock_session):
    """get_items should return the 'value' array from the response."""
    items = [{"id": "1", "fields": {"Title": "A"}}, {"id": "2", "fields": {"Title": "B"}}]
    mock_session.request.return_value = {"value": items}

    result = client.get_items()

    assert result == items


def test_get_items_passes_filter(client, mock_session):
    """get_items should include $filter in params when filter_expr is provided."""
    mock_session.request.return_value = {"value": []}

    client.get_items(filter_expr="Title eq 'Test'")

    call_params = mock_session.request.call_args.kwargs["params"]
    assert call_params["$filter"] == "Title eq 'Test'"


def test_get_items_passes_select(client, mock_session):
    """get_items should include $select in params when select is provided."""
    mock_session.request.return_value = {"value": []}

    client.get_items(select=["Title", "Status"])

    call_params = mock_session.request.call_args.kwargs["params"]
    assert call_params["$select"] == "Title,Status"


# ============================================================================
# get_item Tests
# ============================================================================


def test_get_item_calls_correct_path(client, mock_session):
    """get_item should GET the specific item with $expand=fields."""
    mock_session.request.return_value = {"id": "42", "fields": {"Title": "Test"}}

    result = client.get_item("42")

    mock_session.request.assert_called_with(
        "GET",
        f"/sites/{SITE_ID}/lists/{LIST_ID}/items/42",
        params={"$expand": "fields"},
    )
    assert result["id"] == "42"


# ============================================================================
# create_item Tests
# ============================================================================


def test_create_item_posts_with_fields_wrapper(client, mock_session):
    """create_item should POST with the fields wrapped in a 'fields' key."""
    mock_session.request.return_value = {"id": "new-1", "fields": {"Title": "New"}}

    fields = {"Title": "New", "Status": "Open"}
    client.create_item(fields)

    mock_session.request.assert_called_with(
        "POST",
        f"/sites/{SITE_ID}/lists/{LIST_ID}/items",
        json={"fields": fields},
    )


def test_create_item_returns_response(client, mock_session):
    """create_item should return the created item dict from the API."""
    expected = {"id": "new-1", "fields": {"Title": "New"}}
    mock_session.request.return_value = expected

    result = client.create_item({"Title": "New"})

    assert result == expected


# ============================================================================
# update_item Tests
# ============================================================================


def test_update_item_patches_fields_subresource(client, mock_session):
    """update_item should PATCH the /fields sub-resource with the raw fields dict."""
    mock_session.request.return_value = {"Title": "Updated"}

    fields = {"Title": "Updated"}
    client.update_item("42", fields)

    mock_session.request.assert_called_with(
        "PATCH",
        f"/sites/{SITE_ID}/lists/{LIST_ID}/items/42/fields",
        json=fields,
    )


# ============================================================================
# delete_item Tests
# ============================================================================


def test_delete_item_sends_delete(client, mock_session):
    """delete_item should send a DELETE request to the item path."""
    mock_session.request.return_value = {}

    client.delete_item("42")

    mock_session.request.assert_called_with(
        "DELETE",
        f"/sites/{SITE_ID}/lists/{LIST_ID}/items/42",
    )


# ============================================================================
# delete_list Tests
# ============================================================================


def test_delete_list_sends_delete(client, mock_session):
    """delete_list should send a DELETE request to the list path."""
    mock_session.request.return_value = {}

    client.delete_list()

    mock_session.request.assert_called_with(
        "DELETE",
        f"/sites/{SITE_ID}/lists/{LIST_ID}",
    )


# ============================================================================
# resource_path and supported_change_types Tests
# ============================================================================


def test_resource_path_returns_list_path_without_items(client):
    """resource_path should return /sites/{site_id}/lists/{list_id} (no /items suffix)."""
    assert client.resource_path == f"/sites/{SITE_ID}/lists/{LIST_ID}"


def test_supported_change_types_returns_updated_only(client):
    """SharePoint lists only support 'updated' change notifications."""
    assert client.supported_change_types == {"updated"}


# ============================================================================
# get_recent Tests
# ============================================================================


def test_get_recent_calls_get_items_with_filter(client, mock_session):
    """get_recent should call get_items with a lastModifiedDateTime filter."""
    mock_session.request.return_value = {"value": []}

    client.get_recent(minutes=5)

    call_args = mock_session.request.call_args
    params = call_args.kwargs["params"]
    assert "$filter" in params
    assert "fields/Modified gt" in params["$filter"]


def test_get_recent_uses_correct_cutoff(client, mock_session):
    """get_recent should compute the cutoff as now minus the given minutes."""
    from datetime import UTC, datetime, timedelta

    mock_session.request.return_value = {"value": []}

    with patch("box2.sharepoint.list_client.datetime") as mock_dt:
        now = datetime(2026, 2, 23, 12, 0, 0, tzinfo=UTC)
        mock_dt.now.return_value = now
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)

        client.get_recent(minutes=2)

    params = mock_session.request.call_args.kwargs["params"]
    expected_cutoff = (now - timedelta(minutes=2)).strftime("%Y-%m-%dT%H:%M:%SZ")
    assert expected_cutoff in params["$filter"]


def test_get_recent_returns_items(client, mock_session):
    """get_recent should return the items from the filtered query."""
    items = [{"id": "1", "fields": {"Title": "Recent"}}]
    mock_session.request.return_value = {"value": items}

    result = client.get_recent(minutes=2)

    assert result == items


def test_get_recent_defaults_to_two_minutes(client, mock_session):
    """get_recent should default to a 2-minute lookback window."""
    mock_session.request.return_value = {"value": []}

    client.get_recent()

    params = mock_session.request.call_args.kwargs["params"]
    assert "fields/Modified gt" in params["$filter"]


def test_get_recent_rejects_zero_minutes(client):
    """get_recent should raise ValueError for zero minutes."""
    with pytest.raises(ValueError, match="must be positive"):
        client.get_recent(minutes=0)


def test_get_recent_rejects_negative_minutes(client):
    """get_recent should raise ValueError for negative minutes."""
    with pytest.raises(ValueError, match="must be positive"):
        client.get_recent(minutes=-1)
