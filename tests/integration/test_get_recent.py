"""Integration tests for ListClient.get_recent().

Creates a temporary list, adds items, and verifies that get_recent()
returns the expected items. This test is designed to surface issues with
the OData $filter expression (e.g. timestamp formatting, quoting).

Run with:
    AWS_PROFILE=aws-prototype uv run pytest tests/integration/sharepoint/test_get_recent.py -v -s
"""

import logging
import time
from uuid import uuid4

import pytest

from box2.sharepoint import ListClient, SharePointSession

pytestmark = [pytest.mark.integration]

logger = logging.getLogger(__name__)


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
