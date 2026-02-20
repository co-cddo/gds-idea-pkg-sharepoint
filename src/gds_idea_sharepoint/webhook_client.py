"""SharePoint webhook subscription management via Microsoft Graph API.

Provides lifecycle management for Microsoft Graph change notification
subscriptions. The client is session-level — it can manage subscriptions
across any subscribable resource (lists, drives, etc.).

Usage::

    from box2.sharepoint import SharePointSession, ListClient, WebhookClient

    session = SharePointSession.from_env()
    lists = ListClient(session, list_name="Correspondence")
    webhooks = WebhookClient(session)

    # Subscribe to all changes on the list
    sub = webhooks.subscribe(
        resource=lists,
        notification_url="https://my-server.example.com/webhook",
        client_state="my-shared-secret",
        change_types=["created", "updated", "deleted"],
    )

    # Renew before expiry
    sub = webhooks.renew(sub.id)

    # Or let the client decide whether renewal is needed
    sub = webhooks.renew_if_expiring(sub.id, threshold_minutes=120)

    # Clean up
    webhooks.delete(sub.id)
"""

import logging
from datetime import UTC, datetime, timedelta

from box2.sharepoint.models import Subscription
from box2.sharepoint.protocols import SubscribableResource
from box2.sharepoint.session import SharePointSession

logger = logging.getLogger(__name__)

VALID_CHANGE_TYPES = {"created", "updated", "deleted"}

# Microsoft Graph maximum subscription lifetime for list items (30 days).
MAX_EXPIRATION_MINUTES = 43200


class WebhookClient:
    """Client for managing Microsoft Graph webhook subscriptions.

    Session-level client that can create and manage subscriptions on any
    resource that satisfies the ``SubscribableResource`` protocol (e.g.
    ``ListClient``, future ``DriveClient``).
    """

    def __init__(self, session: SharePointSession):
        self._session = session
        logger.info("WebhookClient ready")

    def subscribe(
        self,
        resource: SubscribableResource,
        notification_url: str,
        client_state: str,
        change_types: list[str],
        expiration_minutes: int = 10080,
    ) -> Subscription:
        """Create a new webhook subscription on a SharePoint resource.

        Microsoft will send a validation request to ``notification_url``
        before the subscription is confirmed. The receiver must respond
        with the ``validationToken`` query parameter.

        Args:
            resource: A subscribable SharePoint resource (e.g. a ``ListClient``).
            notification_url: Public URL that receives POST notifications.
            client_state: Shared secret sent with each notification for validation.
            change_types: List of change types to subscribe to. Valid values are
                ``"created"``, ``"updated"``, and ``"deleted"``.
            expiration_minutes: Minutes until the subscription expires.
                Defaults to 10080 (7 days). Maximum is 43200 (30 days).

        Returns:
            The created Subscription.

        Raises:
            ValueError: If change_types contains invalid values or expiration is
                out of range.
            SharePointAPIError: If the Graph API call fails (e.g. notification URL
                validation fails).
        """
        self._validate_change_types(change_types)
        self._validate_expiration(expiration_minutes)

        expiration_dt = datetime.now(UTC) + timedelta(minutes=expiration_minutes)

        body = {
            "changeType": ",".join(change_types),
            "notificationUrl": notification_url,
            "resource": resource.resource_path,
            "expirationDateTime": expiration_dt.isoformat(),
            "clientState": client_state,
        }

        logger.info(
            "Creating subscription on %s (change_types=%s, expires=%s)",
            resource.resource_path,
            change_types,
            expiration_dt.isoformat(),
        )
        data = self._session.request("POST", "/subscriptions", json=body)
        subscription = Subscription.model_validate(data)
        logger.info("Subscription created (id=%s)", subscription.id)
        return subscription

    def get(self, subscription_id: str) -> Subscription:
        """Get details of an existing subscription.

        Args:
            subscription_id: The subscription ID.

        Returns:
            The Subscription.

        Raises:
            SharePointAPIError: If the subscription is not found or the API call fails.
        """
        data = self._session.request("GET", f"/subscriptions/{subscription_id}")
        return Subscription.model_validate(data)

    def list_subscriptions(self) -> list[Subscription]:
        """List all active subscriptions for this application.

        Returns:
            List of Subscription objects.

        Raises:
            SharePointAPIError: If the API call fails.
        """
        data = self._session.request("GET", "/subscriptions")
        return [Subscription.model_validate(item) for item in data.get("value", [])]

    def renew(
        self,
        subscription_id: str,
        expiration_minutes: int = 10080,
    ) -> Subscription:
        """Renew an existing subscription by extending its expiration.

        Args:
            subscription_id: The subscription ID to renew.
            expiration_minutes: Minutes from now until new expiration.
                Defaults to 10080 (7 days). Maximum is 43200 (30 days).

        Returns:
            The renewed Subscription with updated expiration.

        Raises:
            ValueError: If expiration is out of range.
            SharePointAPIError: If the subscription is not found or the API call fails.
        """
        self._validate_expiration(expiration_minutes)

        expiration_dt = datetime.now(UTC) + timedelta(minutes=expiration_minutes)

        logger.info("Renewing subscription %s (new expiry=%s)", subscription_id, expiration_dt.isoformat())
        data = self._session.request(
            "PATCH",
            f"/subscriptions/{subscription_id}",
            json={"expirationDateTime": expiration_dt.isoformat()},
        )
        subscription = Subscription.model_validate(data)
        logger.info("Subscription %s renewed (expires=%s)", subscription_id, subscription.expiration)
        return subscription

    def delete(self, subscription_id: str) -> None:
        """Delete an existing subscription.

        Args:
            subscription_id: The subscription ID to delete.

        Raises:
            SharePointAPIError: If the subscription is not found or the API call fails.
        """
        logger.info("Deleting subscription %s", subscription_id)
        self._session.request("DELETE", f"/subscriptions/{subscription_id}")
        logger.info("Subscription %s deleted", subscription_id)

    def renew_if_expiring(
        self,
        subscription_id: str,
        threshold_minutes: int = 60,
        expiration_minutes: int = 10080,
    ) -> Subscription | None:
        """Renew a subscription only if it is close to expiring.

        Fetches the current subscription, checks whether its expiration is
        within ``threshold_minutes`` of now, and renews it if so.

        Args:
            subscription_id: The subscription ID to check.
            threshold_minutes: Renew if expiration is within this many minutes.
                Defaults to 60.
            expiration_minutes: Minutes from now for the new expiration if renewed.
                Defaults to 10080 (7 days).

        Returns:
            The renewed Subscription if renewal was performed, or ``None`` if
            the subscription is not yet close to expiring.

        Raises:
            ValueError: If expiration or threshold is out of range.
            SharePointAPIError: If the subscription is not found or the API call fails.
        """
        if threshold_minutes <= 0:
            raise ValueError(f"threshold_minutes must be positive, got {threshold_minutes}")

        subscription = self.get(subscription_id)
        now = datetime.now(UTC)
        remaining = (subscription.expiration - now).total_seconds() / 60

        if remaining <= threshold_minutes:
            logger.info(
                "Subscription %s expires in %.0f min (threshold=%d min) — renewing",
                subscription_id,
                remaining,
                threshold_minutes,
            )
            return self.renew(subscription_id, expiration_minutes)

        logger.info(
            "Subscription %s has %.0f min remaining (threshold=%d min) — no renewal needed",
            subscription_id,
            remaining,
            threshold_minutes,
        )
        return None

    @staticmethod
    def _validate_change_types(change_types: list[str]) -> None:
        """Validate that all change types are recognised by the Graph API.

        Args:
            change_types: List of change type strings.

        Raises:
            ValueError: If the list is empty or contains invalid values.
        """
        if not change_types:
            raise ValueError("change_types must not be empty")
        invalid = set(change_types) - VALID_CHANGE_TYPES
        if invalid:
            raise ValueError(f"Invalid change types: {invalid}. Valid values: {VALID_CHANGE_TYPES}")

    @staticmethod
    def _validate_expiration(expiration_minutes: int) -> None:
        """Validate that the expiration duration is within Graph API limits.

        Args:
            expiration_minutes: Requested duration in minutes.

        Raises:
            ValueError: If out of range.
        """
        if expiration_minutes <= 0:
            raise ValueError(f"expiration_minutes must be positive, got {expiration_minutes}")
        if expiration_minutes > MAX_EXPIRATION_MINUTES:
            raise ValueError(
                f"expiration_minutes ({expiration_minutes}) exceeds Graph API maximum ({MAX_EXPIRATION_MINUTES})"
            )
