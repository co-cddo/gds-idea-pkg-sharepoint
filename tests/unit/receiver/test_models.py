"""Unit tests for the receiver notification models."""

import json
from pathlib import Path

import pytest
from pydantic import ValidationError

from gds_idea_sharepoint.receiver.models import Notification, NotificationPayload, ResourceData

# ============================================================================
# Fixtures
# ============================================================================

FIXTURES_DIR = Path(__file__).parent / "fixtures"

SINGLE_NOTIFICATION = {
    "subscriptionId": "sub-abc-123",
    "changeType": "created",
    "clientState": "test-secret",
    "resource": "sites/site-id-001/lists/list-id-002/items/42",
    "tenantId": "tenant-id-999",
    "subscriptionExpirationDateTime": "2026-02-27T12:00:00.000Z",
}


# ============================================================================
# ResourceData Model Tests
# ============================================================================


def test_parse_resource_data_with_all_fields():
    """ResourceData should parse odata_type, odata_id, and id."""
    data = {
        "@odata.type": "#Microsoft.Graph.listItem",
        "@odata.id": "sites/site-id/lists/list-id/items/42",
        "id": "42",
    }
    rd = ResourceData.model_validate(data)

    assert rd.odata_type == "#Microsoft.Graph.listItem"
    assert rd.odata_id == "sites/site-id/lists/list-id/items/42"
    assert rd.id == "42"


def test_parse_resource_data_with_id_only():
    """ResourceData should parse when only id is present."""
    rd = ResourceData.model_validate({"id": "99"})

    assert rd.id == "99"
    assert rd.odata_type is None
    assert rd.odata_id is None


def test_parse_resource_data_preserves_extra_fields():
    """ResourceData should preserve extra OData fields in model_extra."""
    data = {
        "@odata.type": "#Microsoft.Graph.listItem",
        "id": "42",
        "@odata.etag": '"abc123"',
    }
    rd = ResourceData.model_validate(data)

    assert rd.model_extra["@odata.etag"] == '"abc123"'


def test_parse_resource_data_empty():
    """ResourceData should parse an empty dict with all fields None."""
    rd = ResourceData.model_validate({})

    assert rd.odata_type is None
    assert rd.odata_id is None
    assert rd.id is None


# ============================================================================
# Notification Model Tests
# ============================================================================


def test_parse_notification_from_graph_payload():
    """Notification should parse all core fields from a Graph API notification."""
    notification = Notification.model_validate(SINGLE_NOTIFICATION)

    assert notification.subscription_id == "sub-abc-123"
    assert notification.change_type == "created"
    assert notification.client_state == "test-secret"
    assert notification.resource == "sites/site-id-001/lists/list-id-002/items/42"
    assert notification.tenant_id == "tenant-id-999"
    assert notification.subscription_expiration == "2026-02-27T12:00:00.000Z"


def test_parse_notification_without_client_state():
    """Notification should allow clientState to be absent."""
    data = {k: v for k, v in SINGLE_NOTIFICATION.items() if k != "clientState"}
    notification = Notification.model_validate(data)

    assert notification.client_state is None


def test_parse_notification_without_expiration():
    """Notification should allow subscriptionExpirationDateTime to be absent."""
    data = {k: v for k, v in SINGLE_NOTIFICATION.items() if k != "subscriptionExpirationDateTime"}
    notification = Notification.model_validate(data)

    assert notification.subscription_expiration is None


def test_parse_notification_without_resource_data():
    """Notification should allow resourceData to be absent."""
    notification = Notification.model_validate(SINGLE_NOTIFICATION)

    assert notification.resource_data is None


def test_parse_notification_with_resource_data():
    """Notification should parse resourceData into a ResourceData model."""
    data = {
        **SINGLE_NOTIFICATION,
        "resourceData": {
            "@odata.type": "#Microsoft.Graph.listItem",
            "@odata.id": "sites/site-id-001/lists/list-id-002/items/42",
            "id": "42",
        },
    }
    notification = Notification.model_validate(data)

    assert notification.resource_data is not None
    assert notification.resource_data.id == "42"
    assert notification.resource_data.odata_id == "sites/site-id-001/lists/list-id-002/items/42"
    assert notification.resource_data.odata_type == "#Microsoft.Graph.listItem"


def test_parse_notification_preserves_extra_fields():
    """Extra fields (not resourceData) should be preserved in model_extra."""
    data = {
        **SINGLE_NOTIFICATION,
        "resourceData": {"@odata.type": "#Microsoft.Graph.listItem", "id": "42"},
        "lifecycleEvent": "reauthorizationRequired",
    }
    notification = Notification.model_validate(data)

    # resourceData is now a first-class field, not in model_extra
    assert notification.resource_data is not None
    assert notification.resource_data.id == "42"
    # lifecycleEvent is still in model_extra
    assert notification.model_extra["lifecycleEvent"] == "reauthorizationRequired"


def test_parse_notification_missing_required_field():
    """Notification should reject data missing a required field."""
    data = {k: v for k, v in SINGLE_NOTIFICATION.items() if k != "subscriptionId"}
    with pytest.raises(ValidationError):
        Notification.model_validate(data)


def test_parse_notification_missing_resource():
    """Notification should reject data missing the resource field."""
    data = {k: v for k, v in SINGLE_NOTIFICATION.items() if k != "resource"}
    with pytest.raises(ValidationError):
        Notification.model_validate(data)


# ============================================================================
# NotificationPayload Tests
# ============================================================================


def test_parse_payload_with_single_notification():
    """NotificationPayload should parse a payload with one notification."""
    payload = NotificationPayload.model_validate({"value": [SINGLE_NOTIFICATION]})

    assert len(payload.value) == 1
    assert payload.value[0].subscription_id == "sub-abc-123"


def test_parse_payload_with_multiple_notifications():
    """NotificationPayload should parse a payload with multiple notifications."""
    second = {**SINGLE_NOTIFICATION, "subscriptionId": "sub-def-456", "changeType": "updated"}
    payload = NotificationPayload.model_validate({"value": [SINGLE_NOTIFICATION, second]})

    assert len(payload.value) == 2
    assert payload.value[1].subscription_id == "sub-def-456"
    assert payload.value[1].change_type == "updated"


def test_parse_payload_with_empty_value():
    """NotificationPayload should accept an empty value list."""
    payload = NotificationPayload.model_validate({"value": []})

    assert payload.value == []


def test_parse_fixture_file():
    """NotificationPayload should parse the list_item_created fixture."""
    fixture_path = FIXTURES_DIR / "list_item_created.json"
    with open(fixture_path) as f:
        data = json.load(f)

    payload = NotificationPayload.model_validate(data)

    assert len(payload.value) == 1
    assert payload.value[0].change_type == "created"
    assert payload.value[0].client_state == "test-secret"
    assert payload.value[0].resource_data is not None
    assert payload.value[0].resource_data.id == "42"
