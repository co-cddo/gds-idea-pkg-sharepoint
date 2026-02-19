"""Integration tests for SharePoint list CRUD operations.

Tests the full lifecycle of list operations against a real SharePoint site.
Each test run creates a uniquely named list and cleans up after itself.
"""

from uuid import uuid4

import pytest

from box2.sharepoint import ListClient, SharePointAPIError, SharePointSession

pytestmark = [pytest.mark.integration]

# ============================================================================
# Fixtures
# ============================================================================


@pytest.fixture
def session():
    """Create a real SharePointSession from environment variables."""
    return SharePointSession.from_env()


@pytest.fixture
def list_client(session):
    """Create a temporary SharePoint list and yield a connected client.

    The list is deleted in teardown regardless of test outcome.
    """
    list_name = f"int-test-{uuid4().hex[:8]}"
    client = ListClient.new(session, list_name=list_name)
    try:
        yield client
    finally:
        try:
            client.delete_list()
        except Exception:
            pass  # Best-effort cleanup


# ============================================================================
# List Lifecycle Tests
# ============================================================================


def test_create_and_read_item(list_client):
    """Creating an item should be readable via get_item."""
    created = list_client.create_item({"Title": "Test item"})
    item_id = created["id"]

    fetched = list_client.get_item(item_id)
    fields = fetched.get("fields", {})
    assert fields.get("Title") == "Test item"


def test_get_items_returns_created_items(list_client):
    """get_items should return all items that were created."""
    list_client.create_item({"Title": "Item A"})
    list_client.create_item({"Title": "Item B"})

    items = list_client.get_items()
    titles = [item.get("fields", {}).get("Title") for item in items]
    assert "Item A" in titles
    assert "Item B" in titles
    assert len(items) == 2


def test_update_item(list_client):
    """Updating an item should change its field values."""
    created = list_client.create_item({"Title": "Before"})
    item_id = created["id"]

    list_client.update_item(item_id, {"Title": "After"})

    fetched = list_client.get_item(item_id)
    fields = fetched.get("fields", {})
    assert fields.get("Title") == "After"


def test_delete_item(list_client):
    """Deleting an item should make it inaccessible via get_item."""
    created = list_client.create_item({"Title": "To be deleted"})
    item_id = created["id"]

    list_client.delete_item(item_id)

    with pytest.raises(SharePointAPIError) as exc_info:
        list_client.get_item(item_id)
    assert exc_info.value.status_code == 404
