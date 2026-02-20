"""Integration tests for SharePoint webhook subscriptions.

Tests the full lifecycle of webhook subscription operations against a real
SharePoint site. These tests require:

1. A publicly accessible notification URL for Microsoft's validation handshake
2. AWS credentials with access to the SharePoint federation role
3. The standard SharePoint environment variables

Set the WEBHOOK_NOTIFICATION_URL environment variable to a publicly reachable
endpoint that responds to Microsoft's validation handshake (echoes back the
validationToken query parameter as text/plain).

These tests are skipped unless both the integration environment and the
webhook notification URL are available.
"""

import os
from uuid import uuid4

import pytest

from box2.sharepoint import ListClient, SharePointSession, WebhookClient
from box2.sharepoint.models import Subscription

pytestmark = [pytest.mark.integration, pytest.mark.webhook]

NOTIFICATION_URL = os.environ.get("WEBHOOK_NOTIFICATION_URL", "")
CLIENT_STATE = f"int-test-{uuid4().hex[:8]}"


@pytest.fixture
def session():
    """Create a real SharePointSession from environment variables."""
    return SharePointSession.from_env()


@pytest.fixture
def webhook_client(session):
    """Create a WebhookClient with a real session."""
    return WebhookClient(session)


@pytest.fixture
def list_client(session):
    """Create a temporary SharePoint list for subscription tests.

    The list is deleted in teardown regardless of test outcome.
    """
    list_name = f"int-test-{uuid4().hex[:8]}"
    client = ListClient.new(session, list_name=list_name)
    try:
        yield client
    finally:
        try:
            client.delete_list()
        except Exception:
            pass  # Best-effort cleanup


# ============================================================================
# Subscription Lifecycle Tests
# ============================================================================


@pytest.mark.skipif(not NOTIFICATION_URL, reason="WEBHOOK_NOTIFICATION_URL not set")
def test_create_subscription(webhook_client, list_client):
    """Creating a subscription should return a Subscription with a valid ID."""
    sub = webhook_client.subscribe(
        resource=list_client,
        notification_url=NOTIFICATION_URL,
        client_state=CLIENT_STATE,
        change_types=["created", "updated"],
    )
    try:
        assert isinstance(sub, Subscription)
        assert sub.id
        assert sub.resource == list_client.resource_path
    finally:
        try:
            webhook_client.delete(sub.id)
        except Exception:
            pass


@pytest.mark.skipif(not NOTIFICATION_URL, reason="WEBHOOK_NOTIFICATION_URL not set")
def test_list_includes_created_subscription(webhook_client, list_client):
    """A newly created subscription should appear in list_subscriptions."""
    sub = webhook_client.subscribe(
        resource=list_client,
        notification_url=NOTIFICATION_URL,
        client_state=CLIENT_STATE,
        change_types=["created"],
    )
    try:
        subs = webhook_client.list_subscriptions()
        sub_ids = [s.id for s in subs]
        assert sub.id in sub_ids
    finally:
        try:
            webhook_client.delete(sub.id)
        except Exception:
            pass


@pytest.mark.skipif(not NOTIFICATION_URL, reason="WEBHOOK_NOTIFICATION_URL not set")
def test_renew_subscription(webhook_client, list_client):
    """Renewing a subscription should update its expiration."""
    sub = webhook_client.subscribe(
        resource=list_client,
        notification_url=NOTIFICATION_URL,
        client_state=CLIENT_STATE,
        change_types=["created"],
        expiration_minutes=60,
    )
    try:
        renewed = webhook_client.renew(sub.id, expiration_minutes=10080)
        assert renewed.expiration > sub.expiration
    finally:
        try:
            webhook_client.delete(sub.id)
        except Exception:
            pass


@pytest.mark.skipif(not NOTIFICATION_URL, reason="WEBHOOK_NOTIFICATION_URL not set")
def test_delete_subscription(webhook_client, list_client):
    """Deleting a subscription should remove it from the list."""
    sub = webhook_client.subscribe(
        resource=list_client,
        notification_url=NOTIFICATION_URL,
        client_state=CLIENT_STATE,
        change_types=["created"],
    )

    webhook_client.delete(sub.id)

    subs = webhook_client.list_subscriptions()
    sub_ids = [s.id for s in subs]
    assert sub.id not in sub_ids
