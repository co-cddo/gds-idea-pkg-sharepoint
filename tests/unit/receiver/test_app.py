"""Unit tests for the FastAPI webhook receiver routes.

Uses FastAPI's TestClient (backed by httpx) to test the full HTTP contract
without a running server. Tests both the fallback /webhook endpoint (no routes)
and the route-based endpoints (with WebhookRoute).
"""

import json
from pathlib import Path
from unittest.mock import AsyncMock

import pytest

from gds_idea_sharepoint.receiver.app import create_app
from gds_idea_sharepoint.receiver.config import ReceiverConfig
from gds_idea_sharepoint.receiver.dedup import InMemoryDedup
from gds_idea_sharepoint.receiver.routes import WebhookRoute

# ============================================================================
# Fixtures
# ============================================================================

FIXTURES_DIR = Path(__file__).parent / "fixtures"
CLIENT_STATE = "test-secret"
APP_IDENTITY = "test-app-id"

VALID_NOTIFICATION = {
    "subscriptionId": "sub-abc-123",
    "changeType": "updated",
    "clientState": CLIENT_STATE,
    "resource": "sites/site-id/lists/list-id",
    "tenantId": "tenant-id-999",
    "resourceData": {
        "@odata.type": "#Microsoft.Graph.ListItem",
    },
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


@pytest.fixture
def config():
    """Create a ReceiverConfig for testing."""
    return ReceiverConfig(client_state=CLIENT_STATE, app_identity=APP_IDENTITY)


# ============================================================================
# Fallback /webhook Tests (no routes — for E2E tunnel testing)
# ============================================================================


@pytest.fixture
def fallback_client(config):
    """Create a TestClient for the fallback app (no routes)."""
    from starlette.testclient import TestClient

    store = InMemoryDedup(window_seconds=300)
    app = create_app(config, dedup_store=store)
    return TestClient(app)


def test_fallback_validation_handshake_echoes_token(fallback_client):
    """POST /webhook with validationToken should echo the token back."""
    response = fallback_client.post("/webhook?validationToken=abc-token-123")

    assert response.status_code == 200
    assert response.text == "abc-token-123"


def test_fallback_validation_handshake_returns_plain_text(fallback_client):
    """Validation response should have text/plain content type."""
    response = fallback_client.post("/webhook?validationToken=abc-token-123")

    assert "text/plain" in response.headers["content-type"]


def test_fallback_validation_handshake_with_special_characters(fallback_client):
    """Validation token with URL-encoded characters should be echoed correctly."""
    response = fallback_client.post("/webhook?validationToken=token%20with%20spaces")

    assert response.status_code == 200
    assert response.text == "token with spaces"


def test_fallback_notification_returns_202(fallback_client):
    """POST /webhook with a valid notification payload should return 202."""
    response = fallback_client.post("/webhook", json={"value": [VALID_NOTIFICATION]})

    assert response.status_code == 202
    assert response.json()["status"] == "accepted"


def test_fallback_wrong_client_state_still_returns_202(fallback_client):
    """Notifications with wrong clientState should still return 202."""
    bad = {**VALID_NOTIFICATION, "clientState": "wrong-secret"}
    response = fallback_client.post("/webhook", json={"value": [bad]})

    assert response.status_code == 202


def test_fallback_fixture_file_notification(fallback_client):
    """The list_item_created fixture should be accepted."""
    fixture_path = FIXTURES_DIR / "list_item_created.json"
    with open(fixture_path) as f:
        data = json.load(f)

    response = fallback_client.post("/webhook", json=data)

    assert response.status_code == 202


# ============================================================================
# Health Endpoint Tests
# ============================================================================


def test_health_returns_200(fallback_client):
    """GET /health should return 200."""
    response = fallback_client.get("/health")

    assert response.status_code == 200


def test_health_returns_ok_status(fallback_client):
    """GET /health should return {"status": "ok"}."""
    response = fallback_client.get("/health")

    assert response.json() == {"status": "ok"}


# ============================================================================
# Config Validation Tests
# ============================================================================


def test_config_requires_client_state():
    """ReceiverConfig should reject an empty client_state."""
    from pydantic import ValidationError

    with pytest.raises(ValidationError):
        ReceiverConfig(client_state="", app_identity=APP_IDENTITY)


def test_config_requires_app_identity():
    """ReceiverConfig should reject an empty app_identity."""
    from pydantic import ValidationError

    with pytest.raises(ValidationError):
        ReceiverConfig(client_state="secret", app_identity="")


def test_config_requires_positive_dedup_window():
    """ReceiverConfig should reject a non-positive dedup_window_seconds."""
    from pydantic import ValidationError

    with pytest.raises(ValidationError):
        ReceiverConfig(client_state="secret", app_identity=APP_IDENTITY, dedup_window_seconds=0)


# ============================================================================
# Route-based Endpoint Tests (full pipeline via TestClient)
# ============================================================================


def _make_routed_client(config, routes, store=None):
    """Create a TestClient with route-based endpoints."""
    from starlette.testclient import TestClient

    if store is None:
        store = InMemoryDedup(window_seconds=300)
    app = create_app(config, routes=routes, dedup_store=store)
    return TestClient(app), store


def test_route_validation_handshake(config):
    """Each route endpoint should handle the validation handshake."""
    handler = AsyncMock()
    route = WebhookRoute(path="/file_uploaded", get_items=lambda: [], handler=handler, filter_self=False)
    client, _ = _make_routed_client(config, [route])

    response = client.post("/file_uploaded?validationToken=my-token")

    assert response.status_code == 200
    assert response.text == "my-token"


def test_route_notification_returns_202(config):
    """POST to a route endpoint with a valid notification should return 202."""
    handler = AsyncMock()
    route = WebhookRoute(path="/test_route", get_items=lambda: [], handler=handler, filter_self=False)
    client, _ = _make_routed_client(config, [route])

    response = client.post("/test_route", json={"value": [VALID_NOTIFICATION]})

    assert response.status_code == 202


def test_route_calls_handler_with_items(config):
    """Handler should be called for each item returned by get_items."""
    handler = AsyncMock()
    items = [_make_item(item_id="1"), _make_item(item_id="2")]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=False)
    client, _ = _make_routed_client(config, [route])

    client.post("/test", json={"value": [VALID_NOTIFICATION]})

    assert handler.call_count == 2


def test_route_filters_self_writes(config):
    """Handler should not be called for items modified by the app when filter_self=True."""
    handler = AsyncMock()
    items = [_make_item(item_id="1", modified_by_app=APP_IDENTITY)]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=True)
    client, _ = _make_routed_client(config, [route])

    client.post("/test", json={"value": [VALID_NOTIFICATION]})

    handler.assert_not_called()


def test_route_no_filter_self_when_disabled(config):
    """Handler should be called for app-modified items when filter_self=False."""
    handler = AsyncMock()
    items = [_make_item(item_id="1", modified_by_app=APP_IDENTITY)]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=False)
    client, _ = _make_routed_client(config, [route])

    client.post("/test", json={"value": [VALID_NOTIFICATION]})

    handler.assert_called_once()


def test_route_wrong_client_state_skips_handler(config):
    """Handler should not be called when clientState doesn't match."""
    handler = AsyncMock()
    route = WebhookRoute(path="/test", get_items=lambda: [_make_item()], handler=handler, filter_self=False)
    client, _ = _make_routed_client(config, [route])

    bad = {**VALID_NOTIFICATION, "clientState": "wrong"}
    client.post("/test", json={"value": [bad]})

    handler.assert_not_called()


def test_route_item_level_dedup(config):
    """Same item returned by get_items across two notifications should only be processed once."""
    handler = AsyncMock()
    items = [_make_item(item_id="1", last_modified="2026-02-23T12:00:00Z")]
    route = WebhookRoute(path="/test", get_items=lambda: items, handler=handler, filter_self=False)
    client, store = _make_routed_client(config, [route])

    # Two separate requests — get_items returns same item both times
    client.post("/test", json={"value": [VALID_NOTIFICATION]})
    client.post("/test", json={"value": [{**VALID_NOTIFICATION, "subscriptionId": "sub-2"}]})

    # Item-level dedup: same item+timestamp -> handler called once
    handler.assert_called_once()


def test_multiple_routes_dispatch_to_correct_handler(config):
    """Each route should dispatch to its own handler."""
    handler_a = AsyncMock()
    handler_b = AsyncMock()

    items_a = [_make_item(item_id="1")]
    items_b = [_make_item(item_id="2")]

    routes = [
        WebhookRoute(path="/route_a", get_items=lambda: items_a, handler=handler_a, filter_self=False),
        WebhookRoute(path="/route_b", get_items=lambda: items_b, handler=handler_b, filter_self=False),
    ]
    client, _ = _make_routed_client(config, routes)

    client.post("/route_a", json={"value": [VALID_NOTIFICATION]})

    handler_a.assert_called_once()
    handler_b.assert_not_called()


def test_multiple_routes_each_get_handshake(config):
    """Each route should independently handle the validation handshake."""
    handler_a = AsyncMock()
    handler_b = AsyncMock()

    routes = [
        WebhookRoute(path="/route_a", get_items=lambda: [], handler=handler_a, filter_self=False),
        WebhookRoute(path="/route_b", get_items=lambda: [], handler=handler_b, filter_self=False),
    ]
    client, _ = _make_routed_client(config, routes)

    resp_a = client.post("/route_a?validationToken=token-a")
    resp_b = client.post("/route_b?validationToken=token-b")

    assert resp_a.status_code == 200
    assert resp_a.text == "token-a"
    assert resp_b.status_code == 200
    assert resp_b.text == "token-b"


def test_route_health_still_works(config):
    """The /health endpoint should work alongside route-based endpoints."""
    handler = AsyncMock()
    route = WebhookRoute(path="/test", get_items=lambda: [], handler=handler, filter_self=False)
    client, _ = _make_routed_client(config, [route])

    response = client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok"}
