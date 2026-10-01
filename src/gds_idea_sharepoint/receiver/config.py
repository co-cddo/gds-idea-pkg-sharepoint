"""Receiver configuration.

Centralised configuration for the webhook receiver. Uses Pydantic for
validation so that misconfigured deployments fail fast at startup.
"""

from pydantic import BaseModel, Field


class ReceiverConfig(BaseModel):
    """Configuration for the webhook receiver.

    Attributes:
        client_state: Shared secret that Microsoft sends with each notification.
            The receiver rejects any notification whose ``clientState`` does not
            match this value.
        app_identity: The Azure AD application (service principal) ID used by
            this backend. Used to filter out self-triggered notifications —
            items where ``lastModifiedBy.application.id`` matches this value
            are skipped when ``filter_self=True`` on a route.
        dedup_window_seconds: Time window (in seconds) during which duplicate
            notifications for the same resource change are suppressed.
            Defaults to 300 (5 minutes).
    """

    client_state: str = Field(min_length=1, description="Shared secret for notification validation")
    app_identity: str = Field(min_length=1, description="Azure AD service principal app ID for self-write filtering")
    dedup_window_seconds: int = Field(default=300, gt=0, description="Deduplication window in seconds")
