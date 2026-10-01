"""Pydantic models for Microsoft Graph change notification payloads.

These models are resource-agnostic — they handle notifications from any
subscribable resource (list items, drives, sites, etc.). The ``resource``
field is a raw string that identifies what changed; routing to specific
handlers based on resource type is the responsibility of the handler layer.
"""

from pydantic import BaseModel, ConfigDict, Field


class ResourceData(BaseModel):
    """Lightweight resource identity sent with every Graph change notification.

    Microsoft Graph includes this block even for basic (non-rich)
    subscriptions. It identifies the specific item that changed, allowing
    the receiver to fetch that single item by ID rather than querying for
    all recently changed items.

    Attributes:
        odata_type: OData entity type (e.g. ``"#Microsoft.Graph.listItem"``).
        odata_id: Full Graph API path to the changed item
            (e.g. ``"sites/site-id/lists/list-id/items/42"``).
        id: Short identifier for the changed item (e.g. ``"42"``).
    """

    model_config = ConfigDict(extra="allow", populate_by_name=True)

    odata_type: str | None = Field(default=None, alias="@odata.type")
    odata_id: str | None = Field(default=None, alias="@odata.id")
    id: str | None = None


class Notification(BaseModel):
    """A single change notification from Microsoft Graph.

    Represents one entry in the ``value`` array of a Graph notification POST.
    Extra fields from the Graph payload (e.g. ``encryptedContent``,
    ``lifecycleEvent``) are preserved in ``model_extra``.

    Attributes:
        subscription_id: ID of the subscription that triggered this notification.
        client_state: Shared secret for validation (should match receiver config).
        change_type: Type of change — "created", "updated", or "deleted".
        resource: Graph API resource path that changed.
        resource_data: Lightweight identity of the specific item that changed.
        tenant_id: Azure AD tenant ID that owns the subscription.
        subscription_expiration: When the subscription expires (ISO 8601 string).
    """

    model_config = ConfigDict(extra="allow", populate_by_name=True)

    subscription_id: str = Field(alias="subscriptionId")
    client_state: str | None = Field(default=None, alias="clientState")
    change_type: str = Field(alias="changeType")
    resource: str
    resource_data: ResourceData | None = Field(default=None, alias="resourceData")
    tenant_id: str = Field(alias="tenantId")
    subscription_expiration: str | None = Field(default=None, alias="subscriptionExpirationDateTime")


class NotificationPayload(BaseModel):
    """The top-level payload Microsoft Graph POSTs to the notification URL.

    Attributes:
        value: List of individual notifications.
    """

    value: list[Notification]
