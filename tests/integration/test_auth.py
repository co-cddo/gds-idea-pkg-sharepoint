"""Integration tests for SharePoint authentication.

Tests the full auth chain against real AWS and Azure AD services.
Requires AWS credentials and SharePoint environment variables.
"""

import pytest

from box2.sharepoint import SharePointSession

pytestmark = [pytest.mark.integration]

# ============================================================================
# Fixtures
# ============================================================================


@pytest.fixture
def session():
    """Create a real SharePointSession from environment variables."""
    return SharePointSession.from_env()


# ============================================================================
# Authentication Tests (from_env)
# ============================================================================


def test_session_creation(session):
    """Session should be created successfully from environment variables."""
    assert session.tenant_id
    assert session.client_id
    assert session.site_host
    assert session.site_path


def test_token_acquisition(session):
    """get_token should return a non-empty access token string."""
    token = session.get_token()
    assert isinstance(token, str)
    assert len(token) > 0


def test_site_resolution(session):
    """resolve_site_id should return a site ID string containing the hostname."""
    site_id = session.resolve_site_id()
    assert isinstance(site_id, str)
    assert len(site_id) > 0
    # Graph API site IDs include the hostname as the first component
    assert session.site_host in site_id


# ============================================================================
# Authentication Tests (from_secret)
# ============================================================================


def test_from_secret_session_creation():
    """from_secret should build a valid session from the dev Secrets Manager secret."""
    session = SharePointSession.from_secret("sharepoint-app")
    assert session.tenant_id
    assert session.client_id
    assert session.site_host
    assert session.site_path


def test_from_secret_token_acquisition():
    """from_secret session should acquire a valid Graph API token."""
    session = SharePointSession.from_secret("sharepoint-app")
    token = session.get_token()
    assert isinstance(token, str)
    assert len(token) > 0


def test_from_secret_site_resolution():
    """from_secret session should resolve the SharePoint site ID."""
    session = SharePointSession.from_secret("sharepoint-app")
    site_id = session.resolve_site_id()
    assert isinstance(site_id, str)
    assert session.site_host in site_id
