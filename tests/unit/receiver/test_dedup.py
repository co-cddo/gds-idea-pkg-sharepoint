"""Unit tests for deduplication stores and key builders."""

import time
from unittest.mock import MagicMock, patch

import pytest

from gds_idea_sharepoint.receiver.dedup import DynamoDedup, InMemoryDedup, build_item_dedup_key

# ============================================================================
# build_item_dedup_key Tests
# ============================================================================


def test_item_dedup_key_format():
    """Key should combine route path, item ID, and lastModifiedDateTime."""
    item = {"id": "item-1", "lastModifiedDateTime": "2026-01-01T00:00:00Z"}

    key = build_item_dedup_key("/test_route", item)

    assert key == "item:/test_route:item-1:2026-01-01T00:00:00Z"


def test_item_dedup_key_different_versions_differ():
    """Different timestamps on the same item should produce different keys."""
    item_v1 = {"id": "item-1", "lastModifiedDateTime": "2026-01-01T00:00:00Z"}
    item_v2 = {"id": "item-1", "lastModifiedDateTime": "2026-01-01T01:00:00Z"}

    assert build_item_dedup_key("/test", item_v1) != build_item_dedup_key("/test", item_v2)


def test_item_dedup_key_same_data_same_key():
    """Identical item data on the same route should produce the same key."""
    item_a = {"id": "item-1", "lastModifiedDateTime": "2026-01-01T00:00:00Z"}
    item_b = {"id": "item-1", "lastModifiedDateTime": "2026-01-01T00:00:00Z"}

    assert build_item_dedup_key("/test", item_a) == build_item_dedup_key("/test", item_b)


def test_item_dedup_key_different_routes_differ():
    """Same item on different routes should produce different keys."""
    item = {"id": "item-1", "lastModifiedDateTime": "2026-01-01T00:00:00Z"}

    key_a = build_item_dedup_key("/route_a", item)
    key_b = build_item_dedup_key("/route_b", item)

    assert key_a != key_b


# ============================================================================
# InMemoryDedup Tests
# ============================================================================


def test_in_memory_first_key_returns_true():
    """A key that has never been seen should return True (new)."""
    store = InMemoryDedup(window_seconds=300)

    assert store.record_if_new("key-1") is True


def test_in_memory_same_key_returns_false():
    """A key seen again within the window should return False (duplicate)."""
    store = InMemoryDedup(window_seconds=300)

    store.record_if_new("key-1")

    assert store.record_if_new("key-1") is False


def test_in_memory_expired_key_returns_true():
    """A key seen again after the window should return True (expired)."""
    store = InMemoryDedup(window_seconds=1)

    store.record_if_new("key-1")

    with patch("gds_idea_sharepoint.receiver.dedup.time.monotonic", return_value=time.monotonic() + 2):
        assert store.record_if_new("key-1") is True


def test_in_memory_different_keys_are_independent():
    """Different keys should be tracked independently."""
    store = InMemoryDedup(window_seconds=300)

    assert store.record_if_new("key-1") is True
    assert store.record_if_new("key-2") is True
    assert store.record_if_new("key-1") is False
    assert store.record_if_new("key-2") is False


def test_in_memory_cleanup_removes_expired_entries():
    """Expired entries should be removed during cleanup."""
    store = InMemoryDedup(window_seconds=1)

    store.record_if_new("key-1")
    store.record_if_new("key-2")

    future_time = time.monotonic() + 2
    with patch("gds_idea_sharepoint.receiver.dedup.time.monotonic", return_value=future_time):
        # record_if_new triggers cleanup
        store.record_if_new("key-new")

    # key-1 and key-2 should have been cleaned up; only key-new remains
    assert len(store._seen) == 1
    assert "key-new" in store._seen


# ============================================================================
# DynamoDedup Tests
# ============================================================================


@pytest.fixture
def mock_table():
    """Create a mock DynamoDB table with a ConditionalCheckFailedException."""
    table = MagicMock()

    # Set up the exception class on the mock client
    exc_class = type("ConditionalCheckFailedException", (Exception,), {})
    table.meta.client.exceptions.ConditionalCheckFailedException = exc_class

    return table


def test_dynamo_new_key_returns_true(mock_table):
    """put_item succeeding means the key is new — should return True."""
    mock_table.put_item.return_value = {}
    store = DynamoDedup(table_name="test-table", window_seconds=300, table=mock_table)

    assert store.record_if_new("key-1") is True


def test_dynamo_new_key_calls_put_item(mock_table):
    """record_if_new should call put_item with correct parameters."""
    mock_table.put_item.return_value = {}
    store = DynamoDedup(table_name="test-table", window_seconds=300, table=mock_table)

    with patch("gds_idea_sharepoint.receiver.dedup.time.time", return_value=1000000):
        store.record_if_new("key-1")

    mock_table.put_item.assert_called_once()
    call_kwargs = mock_table.put_item.call_args[1]

    assert call_kwargs["Item"] == {"pk": "key-1", "ttl": 1000300}
    assert call_kwargs["ConditionExpression"] == "attribute_not_exists(pk) OR #t < :now"
    assert call_kwargs["ExpressionAttributeNames"] == {"#t": "ttl"}
    assert call_kwargs["ExpressionAttributeValues"] == {":now": 1000000}


def test_dynamo_duplicate_key_returns_false(mock_table):
    """ConditionalCheckFailedException means the key exists — should return False."""
    exc_class = mock_table.meta.client.exceptions.ConditionalCheckFailedException
    mock_table.put_item.side_effect = exc_class("Item already exists")
    store = DynamoDedup(table_name="test-table", window_seconds=300, table=mock_table)

    assert store.record_if_new("key-1") is False


def test_dynamo_ttl_uses_window_seconds(mock_table):
    """TTL should be current time + window_seconds."""
    mock_table.put_item.return_value = {}
    store = DynamoDedup(table_name="test-table", window_seconds=600, table=mock_table)

    with patch("gds_idea_sharepoint.receiver.dedup.time.time", return_value=2000000):
        store.record_if_new("key-1")

    call_kwargs = mock_table.put_item.call_args[1]
    assert call_kwargs["Item"]["ttl"] == 2000600


def test_dynamo_unexpected_error_propagates(mock_table):
    """Non-conditional-check exceptions should propagate, not be swallowed."""
    mock_table.put_item.side_effect = RuntimeError("connection lost")
    store = DynamoDedup(table_name="test-table", window_seconds=300, table=mock_table)

    with pytest.raises(RuntimeError, match="connection lost"):
        store.record_if_new("key-1")


def test_dynamo_default_table_created_from_boto3():
    """Without an injected table, DynamoDedup should create one via boto3."""
    mock_dynamodb = MagicMock()
    mock_table = MagicMock()
    mock_dynamodb.Table.return_value = mock_table

    with patch("gds_idea_sharepoint.receiver.dedup.boto3.resource", return_value=mock_dynamodb):
        store = DynamoDedup(table_name="my-table", window_seconds=300)

    mock_dynamodb.Table.assert_called_once_with("my-table")
    assert store._table is mock_table
