"""Pluggable item-level deduplication for webhook processing.

After a notification arrives the receiver fetches the specific item from
the Graph API and checks whether it has already been processed. The dedup
key combines the item ID and its ``lastModifiedDateTime`` so that the
same edit is not handled twice, while a new edit on the same item (with a
newer timestamp) is treated as a fresh event.

The ``DeduplicationStore`` protocol uses a single atomic
``record_if_new`` method that checks and records in one step. This
design maps cleanly onto DynamoDB conditional writes for Lambda
deployments while keeping the in-memory implementation simple.

Two implementations are provided:

- ``InMemoryDedup`` — dict-based, suitable for local dev and testing.
- ``DynamoDedup`` — DynamoDB-backed, provides atomic cross-invocation
  deduplication for concurrent Lambda executions.
"""

import logging
import time
from typing import Any, Protocol

import boto3

logger = logging.getLogger(__name__)


# ============================================================================
# Key Builder
# ============================================================================


def build_item_dedup_key(route_path: str, item: dict[str, Any]) -> str:
    """Build a deduplication key for a list/drive item scoped to a route.

    The key combines the route path, item ID, and ``lastModifiedDateTime``
    so that the same edit is not processed twice, but a new edit on the
    same item (with a newer timestamp) is treated as a fresh event. The
    route path ensures that items with the same ID in different lists or
    drives do not collide.

    Args:
        route_path: The webhook route path (e.g. ``"/invitation_updated"``).
        item: An item dict as returned by the Graph API.

    Returns:
        A string key suitable for deduplication lookups.
    """
    item_id = item.get("id", "unknown")
    modified = item.get("lastModifiedDateTime", "unknown")
    key = f"item:{route_path}:{item_id}:{modified}"
    logger.debug("Built item dedup key: %s (route=%s, item_id=%s, modified=%s)", key, route_path, item_id, modified)
    return key


# ============================================================================
# Protocol
# ============================================================================


class DeduplicationStore(Protocol):
    """Interface for notification deduplication backends.

    Implementations must provide a single atomic check-and-record
    operation. This avoids race conditions between separate "check"
    and "record" calls when multiple Lambda invocations process
    notifications concurrently.
    """

    def record_if_new(self, key: str) -> bool:
        """Atomically check whether a key exists and record it if not.

        If the key has not been seen within the dedup window, it is
        recorded and the method returns ``True``. If the key already
        exists (i.e. it is a duplicate), the method returns ``False``
        without modifying the store.

        Args:
            key: A unique identifier for the notification or item event.

        Returns:
            ``True`` if the key was newly recorded (not a duplicate).
            ``False`` if the key already existed (duplicate).
        """
        ...


class InMemoryDedup:
    """In-memory deduplication store using a dict with TTL-based expiry.

    Suitable for single-instance deployments and local development.
    For Lambda or multi-instance deployments, use ``DynamoDedup``
    which provides atomic cross-invocation deduplication.

    Args:
        window_seconds: Time window during which a key is considered a duplicate.
    """

    def __init__(self, window_seconds: int = 300):
        self._window_seconds = window_seconds
        self._seen: dict[str, float] = {}

    def record_if_new(self, key: str) -> bool:
        """Atomically check and record a dedup key.

        Performs lazy cleanup of expired entries before checking.
        If the key exists and is within the dedup window, returns
        ``False``. Otherwise records the key and returns ``True``.

        Args:
            key: A unique identifier for the notification or item event.

        Returns:
            ``True`` if the key was newly recorded (not a duplicate).
            ``False`` if the key already existed (duplicate).
        """
        self._cleanup()
        now = time.monotonic()

        if key in self._seen:
            elapsed = now - self._seen[key]
            if elapsed < self._window_seconds:
                logger.debug("Duplicate key detected: %s (%.0fs ago)", key, elapsed)
                return False

        self._seen[key] = now
        logger.debug("Recorded new key: %s", key)
        return True

    def _cleanup(self) -> None:
        """Remove entries older than the dedup window."""
        now = time.monotonic()
        expired = [k for k, ts in self._seen.items() if now - ts >= self._window_seconds]
        for k in expired:
            del self._seen[k]
        if expired:
            logger.debug("Cleaned up %d expired dedup entries", len(expired))


class DynamoDedup:
    """DynamoDB-backed deduplication store with atomic conditional writes.

    Uses ``put_item`` with a condition expression to achieve at-most-once
    semantics across concurrent Lambda invocations. DynamoDB's TTL feature
    handles automatic cleanup of expired entries.

    Table schema:
        - ``pk`` (String, partition key): the dedup key.
        - ``ttl`` (Number): Unix epoch timestamp for DynamoDB TTL expiry.

    Enable TTL on the ``ttl`` attribute when creating the table::

        aws dynamodb update-time-to-live \\
            --table-name my-dedup-table \\
            --time-to-live-specification "Enabled=true, AttributeName=ttl"

    Args:
        table_name: Name of the DynamoDB table.
        window_seconds: Time window during which a key is considered a
            duplicate. Also used to set the TTL on new records.
        table: Optional pre-configured boto3 DynamoDB Table resource
            (useful for testing or custom session configuration).
    """

    def __init__(
        self,
        table_name: str,
        window_seconds: int = 300,
        table: Any = None,
    ):
        self._window_seconds = window_seconds
        if table is not None:
            self._table = table
        else:
            dynamodb = boto3.resource("dynamodb")
            self._table = dynamodb.Table(table_name)

    def record_if_new(self, key: str) -> bool:
        """Atomically check and record a dedup key in DynamoDB.

        Attempts a conditional ``put_item`` that succeeds only if the key
        does not already exist or has expired (TTL elapsed). This provides
        atomic at-most-once semantics even when multiple Lambda invocations
        race on the same key.

        Note:
            DynamoDB TTL deletion is asynchronous and may lag by up to
            48 hours. The condition expression checks the ``ttl`` value
            directly so that expired-but-not-yet-deleted items are treated
            as available for re-recording.

        Args:
            key: A unique identifier for the notification or item event.

        Returns:
            ``True`` if the key was newly recorded (not a duplicate).
            ``False`` if the key already existed (duplicate).
        """
        now = int(time.time())
        ttl = now + self._window_seconds

        try:
            self._table.put_item(
                Item={"pk": key, "ttl": ttl},
                ConditionExpression="attribute_not_exists(pk) OR #t < :now",
                ExpressionAttributeNames={"#t": "ttl"},
                ExpressionAttributeValues={":now": now},
            )
            logger.debug("DynamoDB: recorded new key: %s (ttl=%d)", key, ttl)
            return True
        except self._table.meta.client.exceptions.ConditionalCheckFailedException:
            logger.debug("DynamoDB: duplicate key detected: %s", key)
            return False
