"""Integration tests for ListClient.get_recent().

Creates a temporary list, adds items, and verifies that get_recent()
returns the expected items. This test is designed to surface issues with
the OData $filter expression (e.g. timestamp formatting, quoting).

Run with:
    AWS_PROFILE=aws-prototype uv run pytest tests/integration/sharepoint/test_get_recent.py -v -s
"""

import logging
import os
import time
from uuid import uuid4

import pytest
from dotenv import load_dotenv

from box2.receiver.handlers import _is_self_write
from box2.sharepoint import ListClient, SharePointSession

pytestmark = [pytest.mark.integration]

logger = logging.getLogger(__name__)


# ============================================================================
# Fixtures
# ============================================================================


@pytest.fixture
def session():
    """Create a real SharePointSession from environment variables."""
    load_dotenv()
    return SharePointSession.from_env()


@pytest.fixture
def list_client(session):
    """Create a temporary SharePoint list and yield a connected client.

    The list is deleted in teardown regardless of test outcome.
    """
    list_name = f"get-recent-test-{uuid4().hex[:8]}"
    logger.info("Creating temporary list: %s", list_name)
    client = ListClient.new(session, list_name=list_name)
    try:
        yield client
    finally:
        try:
            client.delete_list()
            logger.info("Deleted temporary list: %s", list_name)
        except Exception:
            logger.warning("Failed to delete list: %s", list_name)


# ============================================================================
# get_recent Tests
# ============================================================================


def test_get_recent_returns_newly_created_items(list_client):
    """get_recent should return items created within the lookback window."""
    # Create 3 items
    created_ids = []
    for i in range(3):
        item = list_client.create_item({"Title": f"Recent item {i}"})
        created_ids.append(item["id"])
        logger.info("Created item id=%s, title='Recent item %d'", item["id"], i)

    # Small delay to ensure Graph API propagation
    time.sleep(2)

    # get_recent with a generous window — all items should be returned
    items = list_client.get_recent(minutes=5)

    logger.info("get_recent(minutes=5) returned %d item(s)", len(items))
    for item in items:
        logger.info(
            "  id=%s, lastModifiedDateTime=%s, title=%s",
            item.get("id"),
            item.get("lastModifiedDateTime"),
            item.get("fields", {}).get("Title"),
        )

    returned_ids = [item["id"] for item in items]
    for cid in created_ids:
        assert cid in returned_ids, f"Item {cid} not found in get_recent results"


def test_get_recent_returns_updated_items(list_client):
    """get_recent should return items that were recently updated."""
    # Create an item
    item = list_client.create_item({"Title": "Before update"})
    item_id = item["id"]
    logger.info("Created item id=%s", item_id)

    time.sleep(2)

    # Update the item
    list_client.update_item(item_id, {"Title": "After update"})
    logger.info("Updated item id=%s", item_id)

    time.sleep(2)

    # get_recent should find the updated item
    items = list_client.get_recent(minutes=5)

    logger.info("get_recent(minutes=5) returned %d item(s)", len(items))
    for item in items:
        logger.info(
            "  id=%s, lastModifiedDateTime=%s, title=%s",
            item.get("id"),
            item.get("lastModifiedDateTime"),
            item.get("fields", {}).get("Title"),
        )

    returned_ids = [item["id"] for item in items]
    assert item_id in returned_ids


def test_get_recent_filter_expression_is_logged(list_client, caplog):
    """Verify the filter expression used by get_recent for debugging.

    This test creates an item and calls get_recent, capturing debug logs
    so we can inspect the exact OData $filter sent to the Graph API.
    """
    list_client.create_item({"Title": "Filter test item"})
    time.sleep(2)

    with caplog.at_level(logging.DEBUG, logger="box2.sharepoint"):
        items = list_client.get_recent(minutes=5)

    logger.info("get_recent(minutes=5) returned %d item(s)", len(items))

    # Log the raw items so we can see timestamps
    for item in items:
        logger.info(
            "  id=%s, lastModifiedDateTime=%s",
            item.get("id"),
            item.get("lastModifiedDateTime"),
        )

    # The test itself just verifies we got results — the real value is
    # inspecting the log output for the filter expression and response
    assert len(items) >= 1, "Expected at least 1 item from get_recent"


def test_get_items_unfiltered_vs_get_recent(list_client):
    """Compare get_items() (no filter) with get_recent() to spot filter issues.

    If get_items returns items but get_recent does not, the $filter
    expression is likely the problem.
    """
    # Create items
    for i in range(2):
        list_client.create_item({"Title": f"Compare item {i}"})

    time.sleep(2)

    # Unfiltered — should always return items
    all_items = list_client.get_items()
    logger.info("get_items() returned %d item(s)", len(all_items))
    for item in all_items:
        logger.info(
            "  id=%s, lastModifiedDateTime=%s, title=%s",
            item.get("id"),
            item.get("lastModifiedDateTime"),
            item.get("fields", {}).get("Title"),
        )

    # Filtered — should also return items if filter is correct
    recent_items = list_client.get_recent(minutes=5)
    logger.info("get_recent(minutes=5) returned %d item(s)", len(recent_items))
    for item in recent_items:
        logger.info(
            "  id=%s, lastModifiedDateTime=%s, title=%s",
            item.get("id"),
            item.get("lastModifiedDateTime"),
            item.get("fields", {}).get("Title"),
        )

    assert len(all_items) == len(recent_items), (
        f"Mismatch: get_items returned {len(all_items)} but get_recent returned {len(recent_items)}. "
        "The $filter expression may be incorrect."
    )


# ============================================================================
# Self-write Filter Tests
# ============================================================================


def test_app_created_item_detected_as_self_write(list_client):
    """Items created by the service principal should be detected as self-writes.

    Verifies that the Graph API populates lastModifiedBy.application.id
    for items created by the app, and that _is_self_write correctly
    identifies them.
    """
    app_identity = os.environ["SHAREPOINT_CLIENT_ID"]
    logger.info("App identity (SHAREPOINT_CLIENT_ID): %s", app_identity)

    # Create an item as the service principal
    created = list_client.create_item({"Title": "Self-write test"})
    item_id = created["id"]
    logger.info("Created item id=%s", item_id)

    # Fetch the item back with full metadata
    item = list_client.get_item(item_id)

    # Log the lastModifiedBy structure
    last_modified_by = item.get("lastModifiedBy", {})
    logger.info("lastModifiedBy: %s", last_modified_by)
    logger.info("lastModifiedBy.application.id: %s", last_modified_by.get("application", {}).get("id"))

    # _is_self_write should detect this as a self-write
    assert _is_self_write(item, app_identity) is True, (
        f"Expected _is_self_write to return True for app_identity={app_identity}, "
        f"but lastModifiedBy.application.id={last_modified_by.get('application', {}).get('id')}"
    )


def test_app_created_item_not_detected_for_different_identity(list_client):
    """Items created by the service principal should NOT match a different app identity."""
    # Create an item as the service principal
    created = list_client.create_item({"Title": "Different identity test"})
    item = list_client.get_item(created["id"])

    # _is_self_write with a different identity should return False
    assert _is_self_write(item, "some-completely-different-app-id") is False


def test_self_write_filter_with_get_recent(list_client):
    """Verify the full pipeline: get_recent + self-write filter skips app-created items.

    Simulates what dispatch_route does: call get_recent, then for each
    item check _is_self_write. All items in this test are created by the
    app, so all should be filtered out.
    """
    app_identity = os.environ["SHAREPOINT_CLIENT_ID"]

    # Create items as the service principal
    list_client.create_item({"Title": "Pipeline test 1"})
    list_client.create_item({"Title": "Pipeline test 2"})

    time.sleep(2)

    # Fetch recent items
    items = list_client.get_recent(minutes=5)
    logger.info("get_recent returned %d item(s)", len(items))
    assert len(items) >= 2, "Expected at least 2 items from get_recent"

    # All items should be detected as self-writes
    non_self_writes = [item for item in items if not _is_self_write(item, app_identity)]
    logger.info("Non-self-write items: %d", len(non_self_writes))

    assert len(non_self_writes) == 0, (
        f"Expected all items to be self-writes, but {len(non_self_writes)} were not. "
        f"Items: {[item.get('id') for item in non_self_writes]}"
    )
