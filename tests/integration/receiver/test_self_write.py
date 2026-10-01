"""Integration tests for the receiver's self-write filter against live SharePoint.

Verifies that Graph populates ``lastModifiedBy.application.id`` for items
created by the service principal, so that the receiver's ``is_self_write``
check can skip the app's own writes.

Run with:
    AWS_PROFILE=aws-prototype uv run pytest tests/integration/receiver/test_self_write.py -v -s
"""

import logging
import os
import time

import pytest

from gds_idea_sharepoint.receiver.handlers import is_self_write

pytestmark = [pytest.mark.integration]

logger = logging.getLogger(__name__)


# ============================================================================
# Self-write Filter Tests
# ============================================================================


def test_app_created_item_detected_as_self_write(list_client):
    """Items created by the service principal should be detected as self-writes.

    Verifies that the Graph API populates lastModifiedBy.application.id
    for items created by the app, and that is_self_write correctly
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

    # is_self_write should detect this as a self-write
    assert is_self_write(item, app_identity) is True, (
        f"Expected is_self_write to return True for app_identity={app_identity}, "
        f"but lastModifiedBy.application.id={last_modified_by.get('application', {}).get('id')}"
    )


def test_app_created_item_not_detected_for_different_identity(list_client):
    """Items created by the service principal should NOT match a different app identity."""
    # Create an item as the service principal
    created = list_client.create_item({"Title": "Different identity test"})
    item = list_client.get_item(created["id"])

    # is_self_write with a different identity should return False
    assert is_self_write(item, "some-completely-different-app-id") is False


def test_self_write_filter_with_get_recent(list_client):
    """Verify the full pipeline: get_recent + self-write filter skips app-created items.

    Simulates what dispatch_route does: call get_recent, then for each
    item check is_self_write. All items in this test are created by the
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
    non_self_writes = [item for item in items if not is_self_write(item, app_identity)]
    logger.info("Non-self-write items: %d", len(non_self_writes))

    assert len(non_self_writes) == 0, (
        f"Expected all items to be self-writes, but {len(non_self_writes)} were not. "
        f"Items: {[item.get('id') for item in non_self_writes]}"
    )
