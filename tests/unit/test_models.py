"""Unit tests for the Subscription Pydantic model."""

from datetime import UTC, datetime

import pytest
from pydantic import ValidationError

from box2.sharepoint.models import Subscription

# ============================================================================
# Fixtures
# ============================================================================

GRAPH_RESPONSE = {
    "id": "sub-abc-123",
    "resource": "/sites/site-id/lists/list-id/items",
    "changeType": "created,updated",
    "notificationUrl": "https://example.com/webhook",
    "expirationDateTime": "2026-02-27T12:00:00Z",
    "clientState": "my-shared-secret",
}


# ============================================================================
# Parsing Tests
# ============================================================================


def test_parse_full_graph_response():
    """Subscription should parse all core fields from a Graph API response."""
    sub = Subscription.model_validate(GRAPH_RESPONSE)

    assert sub.id == "sub-abc-123"
    assert sub.resource == "/sites/site-id/lists/list-id/items"
    assert sub.change_type == "created,updated"
    assert sub.notification_url == "https://example.com/webhook"
    assert sub.client_state == "my-shared-secret"
    assert sub.expiration == datetime(2026, 2, 27, 12, 0, 0, tzinfo=UTC)


def test_parse_without_client_state():
    """Subscription should allow client_state to be absent."""
    data = {k: v for k, v in GRAPH_RESPONSE.items() if k != "clientState"}
    sub = Subscription.model_validate(data)

    assert sub.client_state is None


def test_parse_preserves_extra_fields():
    """Extra fields from Graph API should be preserved in model_extra."""
    data = {
        **GRAPH_RESPONSE,
        "applicationId": "app-id-xyz",
        "creatorId": "creator-id-999",
        "latestSupportedTlsVersion": "v1_2",
    }
    sub = Subscription.model_validate(data)

    assert sub.model_extra["applicationId"] == "app-id-xyz"
    assert sub.model_extra["creatorId"] == "creator-id-999"
    assert sub.model_extra["latestSupportedTlsVersion"] == "v1_2"


def test_parse_expiration_as_datetime():
    """The expiration field should be parsed as a datetime object."""
    sub = Subscription.model_validate(GRAPH_RESPONSE)
    assert isinstance(sub.expiration, datetime)


def test_parse_expiration_preserves_timezone():
    """The expiration datetime should preserve UTC timezone info."""
    sub = Subscription.model_validate(GRAPH_RESPONSE)
    assert sub.expiration.tzinfo is not None


# ============================================================================
# Validation Tests
# ============================================================================


def test_missing_id_raises_validation_error():
    """Subscription should reject data with a missing id field."""
    data = {k: v for k, v in GRAPH_RESPONSE.items() if k != "id"}
    with pytest.raises(ValidationError):
        Subscription.model_validate(data)


def test_missing_resource_raises_validation_error():
    """Subscription should reject data with a missing resource field."""
    data = {k: v for k, v in GRAPH_RESPONSE.items() if k != "resource"}
    with pytest.raises(ValidationError):
        Subscription.model_validate(data)


def test_missing_change_type_raises_validation_error():
    """Subscription should reject data with a missing changeType field."""
    data = {k: v for k, v in GRAPH_RESPONSE.items() if k != "changeType"}
    with pytest.raises(ValidationError):
        Subscription.model_validate(data)


def test_missing_notification_url_raises_validation_error():
    """Subscription should reject data with a missing notificationUrl field."""
    data = {k: v for k, v in GRAPH_RESPONSE.items() if k != "notificationUrl"}
    with pytest.raises(ValidationError):
        Subscription.model_validate(data)


def test_missing_expiration_raises_validation_error():
    """Subscription should reject data with a missing expirationDateTime field."""
    data = {k: v for k, v in GRAPH_RESPONSE.items() if k != "expirationDateTime"}
    with pytest.raises(ValidationError):
        Subscription.model_validate(data)


# ============================================================================
# Serialisation Tests
# ============================================================================


def test_model_dump_uses_python_names():
    """model_dump() should use the Python field names by default."""
    sub = Subscription.model_validate(GRAPH_RESPONSE)
    dumped = sub.model_dump()

    assert "change_type" in dumped
    assert "notification_url" in dumped
    assert "client_state" in dumped
    assert "changeType" not in dumped


def test_model_dump_by_alias():
    """model_dump(by_alias=True) should use the Graph API field names."""
    sub = Subscription.model_validate(GRAPH_RESPONSE)
    dumped = sub.model_dump(by_alias=True)

    assert "changeType" in dumped
    assert "notificationUrl" in dumped
    assert "clientState" in dumped
