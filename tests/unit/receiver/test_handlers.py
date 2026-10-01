"""Unit tests for notification handler logic."""

import pytest

from gds_idea_sharepoint.receiver.config import ReceiverConfig
from gds_idea_sharepoint.receiver.dedup import InMemoryDedup
from gds_idea_sharepoint.receiver.handlers import (
    dispatch_route,
    is_self_write,
)
from gds_idea_sharepoint.receiver.models import NotificationPayload
from gds_idea_sharepoint.receiver.routes import WebhookRoute

# ============================================================================
# Fixtures
# ============================================================================

APP_IDENTITY = "test-app-id"
CONFIG = ReceiverConfig(client_state="test-secret", app_identity=APP_IDENTITY)

VALID_NOTIFICATION = {
    "subscriptionId": "sub-abc-123",
    "changeType": "updated",
    "clientState": "test-secret",
    "resource": "sites/site-id/lists/list-id",
    "tenantId": "tenant-id-999",
    "resourceData": {
        "@odata.type": "#Microsoft.Graph.ListItem",
    },
}

WRONG_STATE_NOTIFICATION = {
    **VALID_NOTIFICATION,
    "clientState": "wrong-secret",
}


def _make_item(
    item_id: str = "1",
    last_modified: str = "2026-02-23T12:00:00Z",
    modified_by_app: str | None = None,
) -> dict:
    """Build a canned list item response."""
    item = {
        "id": item_id,
        "lastModifiedDateTime": last_modified,
        "fields": {"Title": f"Item {item_id}"},
    }
    if modified_by_app:
        item["lastModifiedBy"] = {"application": {"id": modified_by_app}}
    else:
        item["lastModifiedBy"] = {"user": {"id": "human-user-123"}}
    return item


# ============================================================================
# is_self_write Tests
# ============================================================================


def test_is_self_write_returns_true_for_app_modified():
    """Items modified by the app's service principal should be detected."""
    item = _make_item(modified_by_app=APP_IDENTITY)

    assert is_self_write(item, APP_IDENTITY) is True


def test_is_self_write_returns_false_for_human_modified():
    """Items modified by a human user should not be detected as self-writes."""
    item = _make_item()

    assert is_self_write(item, APP_IDENTITY) is False


def test_is_self_write_returns_false_for_different_app():
    """Items modified by a different app should not be detected as self-writes."""
    item = _make_item(modified_by_app="other-app-id")

    assert is_self_write(item, APP_IDENTITY) is False


def test_is_self_write_handles_missing_lastmodifiedby():
    """Items without lastModifiedBy should return False (not crash)."""
    item = {"id": "1", "lastModifiedDateTime": "2026-02-23T12:00:00Z", "fields": {}}

    assert is_self_write(item, APP_IDENTITY) is False


# ============================================================================
# dispatch_route Tests
# ============================================================================


def _make_payload(*notifications: dict) -> NotificationPayload:
    """Build a NotificationPayload from notification dicts."""
    return NotificationPayload.model_validate({"value": list(notifications)})


@pytest.fixture
def dedup_store():
    """Create a fresh InMemoryDedup store."""
    return InMemoryDedup(window_seconds=300)


@pytest.mark.anyio
async def test_dispatch_route_calls_handler_for_items(dedup_store):
    """dispatch_route should call the handler for each item returned by get_items."""
    called_with = []

    async def handler(item):
        called_with.append(item)

    items = [_make_item(item_id="1"), _make_item(item_id="2")]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=False)
    payload = _make_payload(VALID_NOTIFICATION)

    dispatched = await dispatch_route(route, payload, CONFIG, dedup_store)

    assert dispatched == 2
    assert len(called_with) == 2
    assert called_with[0]["id"] == "1"
    assert called_with[1]["id"] == "2"


@pytest.mark.anyio
async def test_dispatch_route_skips_wrong_client_state(dedup_store):
    """dispatch_route should not call get_items or handler when clientState doesn't match."""
    get_items_called = False

    def get_items():
        nonlocal get_items_called
        get_items_called = True
        return []

    async def handler(item):
        pass

    route = WebhookRoute(path="/test", get_items=get_items, handler=handler, filter_self=False)
    payload = _make_payload(WRONG_STATE_NOTIFICATION)

    dispatched = await dispatch_route(route, payload, CONFIG, dedup_store)

    assert dispatched == 0
    assert get_items_called is False


@pytest.mark.anyio
async def test_dispatch_route_filters_self_writes(dedup_store):
    """dispatch_route should skip items modified by the app when filter_self=True."""
    called_with = []

    async def handler(item):
        called_with.append(item)

    items = [_make_item(item_id="1", modified_by_app=APP_IDENTITY)]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=True)
    payload = _make_payload(VALID_NOTIFICATION)

    dispatched = await dispatch_route(route, payload, CONFIG, dedup_store)

    assert dispatched == 0
    assert len(called_with) == 0


@pytest.mark.anyio
async def test_dispatch_route_no_self_filter_when_disabled(dedup_store):
    """dispatch_route should process app-modified items when filter_self=False."""
    called_with = []

    async def handler(item):
        called_with.append(item)

    items = [_make_item(item_id="1", modified_by_app=APP_IDENTITY)]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=False)
    payload = _make_payload(VALID_NOTIFICATION)

    dispatched = await dispatch_route(route, payload, CONFIG, dedup_store)

    assert dispatched == 1


@pytest.mark.anyio
async def test_dispatch_route_mixed_self_and_human_items(dedup_store):
    """dispatch_route should process human items and skip self-writes in the same batch."""
    called_with = []

    async def handler(item):
        called_with.append(item)

    items = [
        _make_item(item_id="1", modified_by_app=APP_IDENTITY),
        _make_item(item_id="2"),  # human
        _make_item(item_id="3", modified_by_app=APP_IDENTITY),
        _make_item(item_id="4"),  # human
    ]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=True)
    payload = _make_payload(VALID_NOTIFICATION)

    dispatched = await dispatch_route(route, payload, CONFIG, dedup_store)

    assert dispatched == 2
    assert [c["id"] for c in called_with] == ["2", "4"]


@pytest.mark.anyio
async def test_dispatch_route_item_level_dedup(dedup_store):
    """dispatch_route should not process the same item+timestamp twice."""
    call_count = 0

    async def handler(item):
        nonlocal call_count
        call_count += 1

    # Same items returned by both notifications
    items = [_make_item(item_id="1", last_modified="2026-02-23T12:00:00Z")]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=False)

    # Two notifications in the same payload — get_items returns same items each time
    payload = _make_payload(VALID_NOTIFICATION, {**VALID_NOTIFICATION, "subscriptionId": "sub-2"})

    await dispatch_route(route, payload, CONFIG, dedup_store)

    # Item dedup: same item + same timestamp -> handler called once
    assert call_count == 1


@pytest.mark.anyio
async def test_dispatch_route_dedup_across_calls(dedup_store):
    """dispatch_route should dedup items across separate calls (same dedup store)."""
    call_count = 0

    async def handler(item):
        nonlocal call_count
        call_count += 1

    items = [_make_item(item_id="1", last_modified="2026-02-23T12:00:00Z")]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=False)

    payload1 = _make_payload(VALID_NOTIFICATION)
    payload2 = _make_payload({**VALID_NOTIFICATION, "subscriptionId": "sub-2"})

    await dispatch_route(route, payload1, CONFIG, dedup_store)
    await dispatch_route(route, payload2, CONFIG, dedup_store)

    assert call_count == 1


@pytest.mark.anyio
async def test_dispatch_route_processes_updated_item(dedup_store):
    """dispatch_route should process the same item again if its timestamp changes."""
    called_with = []

    async def handler(item):
        called_with.append(item)

    # First call: item at time T1
    items_v1 = [_make_item(item_id="1", last_modified="2026-02-23T12:00:00Z")]
    route = WebhookRoute(path="/test", get_items=lambda: items_v1, handler=handler, filter_self=False)
    payload1 = _make_payload(VALID_NOTIFICATION)
    await dispatch_route(route, payload1, CONFIG, dedup_store)

    # Second call: same item at time T2 (human edited again)
    items_v2 = [_make_item(item_id="1", last_modified="2026-02-23T12:05:00Z")]
    route.get_items = lambda: items_v2
    payload2 = _make_payload({**VALID_NOTIFICATION, "subscriptionId": "sub-2"})
    await dispatch_route(route, payload2, CONFIG, dedup_store)

    # Both should be processed — different timestamps mean different edits
    assert len(called_with) == 2


@pytest.mark.anyio
async def test_dispatch_route_records_before_handler(dedup_store):
    """dispatch_route should record the item in dedup BEFORE calling the handler."""
    recorded_before_handler = []

    async def handler(item):
        key = f"item:/test:{item['id']}:{item['lastModifiedDateTime']}"
        # record_if_new returns False if the key was already recorded
        recorded_before_handler.append(dedup_store.record_if_new(key))

    items = [_make_item()]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=False)
    payload = _make_payload(VALID_NOTIFICATION)

    await dispatch_route(route, payload, CONFIG, dedup_store)

    # The handler should see the key as already recorded (False = duplicate)
    assert recorded_before_handler == [False]


@pytest.mark.anyio
async def test_dispatch_route_handles_get_items_failure(dedup_store):
    """dispatch_route should skip the notification if get_items raises an exception."""

    async def handler(item):
        pass

    def failing_get_items():
        raise RuntimeError("connection failed")

    route = WebhookRoute(path="/test", get_items=failing_get_items, handler=handler, filter_self=False)
    payload = _make_payload(VALID_NOTIFICATION)

    dispatched = await dispatch_route(route, payload, CONFIG, dedup_store)

    assert dispatched == 0


@pytest.mark.anyio
async def test_dispatch_route_handles_handler_failure(dedup_store):
    """dispatch_route should continue processing other items if the handler fails for one."""
    called_with = []

    async def handler(item):
        if item["id"] == "1":
            raise RuntimeError("handler failed")
        called_with.append(item)

    items = [_make_item(item_id="1"), _make_item(item_id="2")]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=False)
    payload = _make_payload(VALID_NOTIFICATION)

    dispatched = await dispatch_route(route, payload, CONFIG, dedup_store)

    # Item 1 failed, item 2 succeeded
    assert dispatched == 1
    assert called_with[0]["id"] == "2"


@pytest.mark.anyio
async def test_dispatch_route_empty_payload(dedup_store):
    """dispatch_route should handle an empty notification payload."""

    async def handler(item):
        pass

    route = WebhookRoute(path="/test", get_items=lambda: [], handler=handler, filter_self=False)
    payload = _make_payload()

    dispatched = await dispatch_route(route, payload, CONFIG, dedup_store)

    assert dispatched == 0


# ============================================================================
# WebhookRoute Validation Tests
# ============================================================================


def test_route_path_must_start_with_slash():
    """WebhookRoute should reject paths that don't start with /."""

    async def handler(item):
        pass

    with pytest.raises(ValueError, match="must start with '/'"):
        WebhookRoute(path="no_slash", get_items=lambda: [], handler=handler)


def test_route_path_accepts_valid_path():
    """WebhookRoute should accept paths that start with /."""

    async def handler(item):
        pass

    route = WebhookRoute(path="/valid", get_items=lambda: [], handler=handler)
    assert route.path == "/valid"
