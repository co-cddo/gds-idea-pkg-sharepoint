"""Unit tests for the WebhookClient.

Tests webhook subscription lifecycle operations using a mocked session.
"""

from datetime import UTC, datetime, timedelta
from unittest.mock import MagicMock, patch

import pytest

from box2.sharepoint.models import Subscription
from box2.sharepoint.webhook_client import MAX_EXPIRATION_MINUTES, WebhookClient

# ============================================================================
# Fixtures
# ============================================================================

SITE_ID = "contoso.sharepoint.com,abc-123,def-456"
LIST_ID = "list-id-789"
RESOURCE_PATH = f"/sites/{SITE_ID}/lists/{LIST_ID}"
NOTIFICATION_URL = "https://my-server.example.com/webhook"
CLIENT_STATE = "my-shared-secret"
SUBSCRIPTION_ID = "sub-abc-123"


def _graph_subscription(
    sub_id: str = SUBSCRIPTION_ID,
    resource: str = RESOURCE_PATH,
    change_type: str = "created,updated",
    expiration_dt: str | None = None,
) -> dict:
    """Build a canned Graph API subscription response."""
    if expiration_dt is None:
        expiration_dt = (datetime.now(UTC) + timedelta(days=7)).isoformat()
    return {
        "id": sub_id,
        "resource": resource,
        "changeType": change_type,
        "notificationUrl": NOTIFICATION_URL,
        "expirationDateTime": expiration_dt,
        "clientState": CLIENT_STATE,
    }


@pytest.fixture
def mock_session():
    """Create a MagicMock session."""
    return MagicMock()


@pytest.fixture
def mock_resource():
    """Create a mock subscribable resource that supports all change types.

    WebhookClient is generic — it works with any resource. The mock supports
    all three change types so we can test the full range of subscribe() behaviour.
    A resource that restricts change types (like ListClient) is tested separately.
    """
    resource = MagicMock()
    resource.resource_path = RESOURCE_PATH
    resource.supported_change_types = {"created", "updated", "deleted"}
    return resource


@pytest.fixture
def list_like_resource():
    """Create a mock resource that only supports 'updated' (like ListClient)."""
    resource = MagicMock()
    resource.resource_path = RESOURCE_PATH
    resource.supported_change_types = {"updated"}
    return resource


@pytest.fixture
def client(mock_session):
    """Create a WebhookClient with a mocked session."""
    return WebhookClient(mock_session)


# ============================================================================
# subscribe Tests
# ============================================================================


def test_subscribe_posts_to_subscriptions_endpoint(client, mock_session, mock_resource):
    """subscribe should POST to /subscriptions with the correct body."""
    mock_session.request.return_value = _graph_subscription()

    client.subscribe(
        resource=mock_resource,
        notification_url=NOTIFICATION_URL,
        client_state=CLIENT_STATE,
        change_types=["created", "updated"],
    )

    call_args = mock_session.request.call_args
    assert call_args.args[0] == "POST"
    assert call_args.args[1] == "/subscriptions"


def test_subscribe_sends_correct_body_fields(client, mock_session, mock_resource):
    """subscribe should include resource, changeType, notificationUrl, clientState, and expirationDateTime."""
    mock_session.request.return_value = _graph_subscription()

    client.subscribe(
        resource=mock_resource,
        notification_url=NOTIFICATION_URL,
        client_state=CLIENT_STATE,
        change_types=["created", "updated"],
    )

    body = mock_session.request.call_args.kwargs["json"]
    assert body["resource"] == RESOURCE_PATH
    assert body["changeType"] == "created,updated"
    assert body["notificationUrl"] == NOTIFICATION_URL
    assert body["clientState"] == CLIENT_STATE
    assert "expirationDateTime" in body


def test_subscribe_returns_subscription_model(client, mock_session, mock_resource):
    """subscribe should return a Subscription model, not a raw dict."""
    mock_session.request.return_value = _graph_subscription()

    result = client.subscribe(
        resource=mock_resource,
        notification_url=NOTIFICATION_URL,
        client_state=CLIENT_STATE,
        change_types=["created"],
    )

    assert isinstance(result, Subscription)
    assert result.id == SUBSCRIPTION_ID


def test_subscribe_joins_change_types(client, mock_session, mock_resource):
    """subscribe should join multiple change types with commas."""
    mock_session.request.return_value = _graph_subscription(change_type="created,updated,deleted")

    client.subscribe(
        resource=mock_resource,
        notification_url=NOTIFICATION_URL,
        client_state=CLIENT_STATE,
        change_types=["created", "updated", "deleted"],
    )

    body = mock_session.request.call_args.kwargs["json"]
    assert body["changeType"] == "created,updated,deleted"


def test_subscribe_uses_custom_expiration(client, mock_session, mock_resource):
    """subscribe should compute expiration based on the given minutes."""
    mock_session.request.return_value = _graph_subscription()

    with patch("box2.sharepoint.webhook_client.datetime") as mock_dt:
        now = datetime(2026, 2, 20, 12, 0, 0, tzinfo=UTC)
        mock_dt.now.return_value = now
        mock_dt.side_effect = lambda *a, **kw: datetime(*a, **kw)

        client.subscribe(
            resource=mock_resource,
            notification_url=NOTIFICATION_URL,
            client_state=CLIENT_STATE,
            change_types=["created"],
            expiration_minutes=1440,  # 1 day
        )

    body = mock_session.request.call_args.kwargs["json"]
    expected = (now + timedelta(minutes=1440)).isoformat()
    assert body["expirationDateTime"] == expected


# ============================================================================
# subscribe Validation Tests
# ============================================================================


def test_subscribe_rejects_empty_change_types(client, mock_resource):
    """subscribe should raise ValueError for an empty change_types list."""
    with pytest.raises(ValueError, match="must not be empty"):
        client.subscribe(
            resource=mock_resource,
            notification_url=NOTIFICATION_URL,
            client_state=CLIENT_STATE,
            change_types=[],
        )


def test_subscribe_rejects_invalid_change_type(client, mock_resource):
    """subscribe should raise ValueError for unrecognised change types."""
    with pytest.raises(ValueError, match="Invalid change types"):
        client.subscribe(
            resource=mock_resource,
            notification_url=NOTIFICATION_URL,
            client_state=CLIENT_STATE,
            change_types=["created", "modified"],
        )


def test_subscribe_rejects_unsupported_change_type_for_resource(client, list_like_resource):
    """subscribe should raise ValueError when change types are valid but unsupported by the resource."""
    with pytest.raises(ValueError, match="Unsupported change types for this resource"):
        client.subscribe(
            resource=list_like_resource,
            notification_url=NOTIFICATION_URL,
            client_state=CLIENT_STATE,
            change_types=["created", "updated"],
        )


def test_subscribe_unsupported_error_names_the_offending_types(client, list_like_resource):
    """The unsupported change types error message should list which types are unsupported."""
    with pytest.raises(ValueError, match="created") as exc_info:
        client.subscribe(
            resource=list_like_resource,
            notification_url=NOTIFICATION_URL,
            client_state=CLIENT_STATE,
            change_types=["created", "deleted"],
        )
    assert "deleted" in str(exc_info.value)


def test_subscribe_unsupported_error_shows_supported_types(client, list_like_resource):
    """The unsupported change types error message should show what the resource does support."""
    with pytest.raises(ValueError, match="updated") as exc_info:
        client.subscribe(
            resource=list_like_resource,
            notification_url=NOTIFICATION_URL,
            client_state=CLIENT_STATE,
            change_types=["created"],
        )
    assert "Supported:" in str(exc_info.value)


def test_subscribe_accepts_supported_change_type_for_resource(client, mock_session, list_like_resource):
    """subscribe should succeed when change types match what the resource supports."""
    mock_session.request.return_value = _graph_subscription(change_type="updated")

    result = client.subscribe(
        resource=list_like_resource,
        notification_url=NOTIFICATION_URL,
        client_state=CLIENT_STATE,
        change_types=["updated"],
    )

    assert isinstance(result, Subscription)


def test_subscribe_rejects_zero_expiration(client, mock_resource):
    """subscribe should raise ValueError for zero expiration_minutes."""
    with pytest.raises(ValueError, match="must be positive"):
        client.subscribe(
            resource=mock_resource,
            notification_url=NOTIFICATION_URL,
            client_state=CLIENT_STATE,
            change_types=["created"],
            expiration_minutes=0,
        )


def test_subscribe_rejects_excessive_expiration(client, mock_resource):
    """subscribe should raise ValueError when expiration exceeds the Graph API maximum."""
    with pytest.raises(ValueError, match="exceeds Graph API maximum"):
        client.subscribe(
            resource=mock_resource,
            notification_url=NOTIFICATION_URL,
            client_state=CLIENT_STATE,
            change_types=["created"],
            expiration_minutes=MAX_EXPIRATION_MINUTES + 1,
        )


# ============================================================================
# get Tests
# ============================================================================


def test_get_calls_correct_path(client, mock_session):
    """get should GET /subscriptions/{id}."""
    mock_session.request.return_value = _graph_subscription()

    client.get(SUBSCRIPTION_ID)

    mock_session.request.assert_called_with("GET", f"/subscriptions/{SUBSCRIPTION_ID}")


def test_get_returns_subscription_model(client, mock_session):
    """get should return a Subscription model."""
    mock_session.request.return_value = _graph_subscription()

    result = client.get(SUBSCRIPTION_ID)

    assert isinstance(result, Subscription)
    assert result.id == SUBSCRIPTION_ID


# ============================================================================
# list_subscriptions Tests
# ============================================================================


def test_list_subscriptions_calls_correct_path(client, mock_session):
    """list_subscriptions should GET /subscriptions."""
    mock_session.request.return_value = {"value": []}

    client.list_subscriptions()

    mock_session.request.assert_called_with("GET", "/subscriptions")


def test_list_subscriptions_returns_list_of_models(client, mock_session):
    """list_subscriptions should return a list of Subscription models."""
    mock_session.request.return_value = {
        "value": [
            _graph_subscription(sub_id="sub-1"),
            _graph_subscription(sub_id="sub-2"),
        ]
    }

    result = client.list_subscriptions()

    assert len(result) == 2
    assert all(isinstance(s, Subscription) for s in result)
    assert result[0].id == "sub-1"
    assert result[1].id == "sub-2"


def test_list_subscriptions_returns_empty_list(client, mock_session):
    """list_subscriptions should return an empty list when there are no subscriptions."""
    mock_session.request.return_value = {"value": []}

    result = client.list_subscriptions()

    assert result == []


# ============================================================================
# renew Tests
# ============================================================================


def test_renew_patches_expiration(client, mock_session):
    """renew should PATCH /subscriptions/{id} with a new expirationDateTime."""
    mock_session.request.return_value = _graph_subscription()

    client.renew(SUBSCRIPTION_ID)

    call_args = mock_session.request.call_args
    assert call_args.args[0] == "PATCH"
    assert call_args.args[1] == f"/subscriptions/{SUBSCRIPTION_ID}"
    assert "expirationDateTime" in call_args.kwargs["json"]


def test_renew_returns_subscription_model(client, mock_session):
    """renew should return the renewed Subscription."""
    mock_session.request.return_value = _graph_subscription()

    result = client.renew(SUBSCRIPTION_ID)

    assert isinstance(result, Subscription)


def test_renew_rejects_zero_expiration(client):
    """renew should raise ValueError for zero expiration_minutes."""
    with pytest.raises(ValueError, match="must be positive"):
        client.renew(SUBSCRIPTION_ID, expiration_minutes=0)


def test_renew_rejects_excessive_expiration(client):
    """renew should raise ValueError when expiration exceeds the maximum."""
    with pytest.raises(ValueError, match="exceeds Graph API maximum"):
        client.renew(SUBSCRIPTION_ID, expiration_minutes=MAX_EXPIRATION_MINUTES + 1)


# ============================================================================
# delete Tests
# ============================================================================


def test_delete_calls_correct_path(client, mock_session):
    """delete should DELETE /subscriptions/{id}."""
    mock_session.request.return_value = {}

    client.delete(SUBSCRIPTION_ID)

    mock_session.request.assert_called_with("DELETE", f"/subscriptions/{SUBSCRIPTION_ID}")


# ============================================================================
# renew_if_expiring Tests
# ============================================================================


def test_renew_if_expiring_renews_when_within_threshold(client, mock_session):
    """renew_if_expiring should renew when expiration is within the threshold."""
    # Subscription expires in 30 minutes (within default 60-min threshold)
    expiring_soon = (datetime.now(UTC) + timedelta(minutes=30)).isoformat()
    get_response = _graph_subscription(expiration_dt=expiring_soon)
    renewed_response = _graph_subscription()

    mock_session.request.side_effect = [get_response, renewed_response]

    result = client.renew_if_expiring(SUBSCRIPTION_ID)

    assert result is not None
    assert isinstance(result, Subscription)
    # Should have made 2 calls: GET (to check) then PATCH (to renew)
    assert mock_session.request.call_count == 2


def test_renew_if_expiring_skips_when_not_expiring(client, mock_session):
    """renew_if_expiring should return None when expiration is not within the threshold."""
    # Subscription expires in 7 days (well outside default 60-min threshold)
    far_future = (datetime.now(UTC) + timedelta(days=7)).isoformat()
    mock_session.request.return_value = _graph_subscription(expiration_dt=far_future)

    result = client.renew_if_expiring(SUBSCRIPTION_ID)

    assert result is None
    # Should have made only 1 call: GET (to check), no PATCH
    mock_session.request.assert_called_once()


def test_renew_if_expiring_uses_custom_threshold(client, mock_session):
    """renew_if_expiring should respect a custom threshold_minutes value."""
    # Subscription expires in 2 hours
    expiring = (datetime.now(UTC) + timedelta(hours=2)).isoformat()
    get_response = _graph_subscription(expiration_dt=expiring)
    renewed_response = _graph_subscription()

    mock_session.request.side_effect = [get_response, renewed_response]

    # With a 3-hour threshold, 2 hours remaining should trigger renewal
    result = client.renew_if_expiring(SUBSCRIPTION_ID, threshold_minutes=180)

    assert result is not None
    assert mock_session.request.call_count == 2


def test_renew_if_expiring_rejects_zero_threshold(client):
    """renew_if_expiring should raise ValueError for zero threshold_minutes."""
    with pytest.raises(ValueError, match="threshold_minutes must be positive"):
        client.renew_if_expiring(SUBSCRIPTION_ID, threshold_minutes=0)


def test_renew_if_expiring_rejects_negative_threshold(client):
    """renew_if_expiring should raise ValueError for negative threshold_minutes."""
    with pytest.raises(ValueError, match="threshold_minutes must be positive"):
        client.renew_if_expiring(SUBSCRIPTION_ID, threshold_minutes=-10)


# ============================================================================
# ListClient Protocol Compliance
# ============================================================================


def test_list_client_satisfies_subscribable_resource(mock_session):
    """ListClient should expose resource_path and supported_change_types."""
    from box2.sharepoint.list_client import ListClient

    mock_session.resolve_site_id.return_value = SITE_ID
    mock_session.request.return_value = {"value": [{"id": LIST_ID, "displayName": "Test List"}]}

    list_client = ListClient(mock_session, list_name="Test List")

    assert list_client.resource_path == f"/sites/{SITE_ID}/lists/{LIST_ID}"
    assert list_client.supported_change_types == {"updated"}
