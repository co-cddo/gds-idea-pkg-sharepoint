"""Fixtures for receiver integration tests that need a live SharePoint list."""

import logging
from uuid import uuid4

import pytest

from gds_idea_sharepoint import ListClient, SharePointSession

logger = logging.getLogger(__name__)


@pytest.fixture
def session():
    """Create a real SharePointSession from environment variables."""
    return SharePointSession.from_env()


@pytest.fixture
def list_client(session):
    """Create a temporary SharePoint list and yield a connected client.

    The list is deleted in teardown regardless of test outcome.
    """
    list_name = f"self-write-test-{uuid4().hex[:8]}"
    logger.info("Creating temporary list: %s", list_name)
    client = ListClient.new(session, list_name=list_name)
    try:
        yield client
    finally:
        try:
            client.delete_list()
            logger.info("Deleted temporary list: %s", list_name)
        except Exception:
            logger.warning("Failed to delete list: %s", list_name)
