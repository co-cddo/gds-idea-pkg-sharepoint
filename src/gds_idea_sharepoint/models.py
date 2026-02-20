"""Pydantic models for SharePoint Graph API responses.

Structured models for Microsoft Graph API objects returned by SharePoint
operations. Models use ``extra="allow"`` so that any additional fields from
the API response are preserved in ``model_extra`` without breaking parsing.
"""

from datetime import datetime

from pydantic import BaseModel, ConfigDict, Field


class Subscription(BaseModel):
    """A Microsoft Graph webhook subscription.

    Core fields are typed explicitly. Any additional fields returned by the
    Graph API (e.g. ``applicationId``, ``creatorId``,
    ``latestSupportedTlsVersion``) are accessible via ``model_extra``.

    Attributes:
        id: Unique identifier for the subscription.
        resource: The Graph API resource being monitored.
        change_type: Comma-separated change types (e.g. "created,updated").
        notification_url: URL that receives POST notifications.
        expiration: When the subscription expires (UTC).
        client_state: Shared secret sent with each notification for validation.
    """

    model_config = ConfigDict(extra="allow", populate_by_name=True)

    id: str
    resource: str
    change_type: str = Field(alias="changeType")
    notification_url: str = Field(alias="notificationUrl")
    expiration: datetime = Field(alias="expirationDateTime")
    client_state: str | None = Field(default=None, alias="clientState")
