from unittest.mock import MagicMock, patch

import pytest

from box2.sharepoint.exceptions import SharePointAPIError, SharePointAuthError, SharePointConfigError
from box2.sharepoint.session import SharePointSession

# ============================================================================
# Fixtures
# ============================================================================

REQUIRED_ENV = {
    "SHAREPOINT_TENANT_ID": "test-tenant-id",
    "SHAREPOINT_CLIENT_ID": "test-client-id",
    "SHAREPOINT_SITE_HOST": "contoso.sharepoint.com",
    "SHAREPOINT_SITE_PATH": "/sites/test-site",
    "SHAREPOINT_ROLE_ARN": "arn:aws:iam::123456789012:role/test-role",
}


@pytest.fixture
def session():
    """Create a SharePointSession with mocked external dependencies."""
    with (
        patch("box2.sharepoint.session.ClientAssertionCredential"),
        patch("box2.sharepoint.session.httpx.Client"),
    ):
        return SharePointSession(
            tenant_id="test-tenant-id",
            client_id="test-client-id",
            site_host="contoso.sharepoint.com",
            site_path="/sites/test-site",
            role_arn="arn:aws:iam::123456789012:role/test-role",
        )


# ============================================================================
# from_env Tests
# ============================================================================


@patch.dict("os.environ", REQUIRED_ENV, clear=True)
def test_from_env_succeeds_with_all_vars():
    """from_env should create a session when all required vars are set."""
    with (
        patch("box2.sharepoint.session.ClientAssertionCredential"),
        patch("box2.sharepoint.session.httpx.Client"),
    ):
        session = SharePointSession.from_env()
        assert session.tenant_id == "test-tenant-id"
        assert session.client_id == "test-client-id"
        assert session.site_host == "contoso.sharepoint.com"
        assert session.site_path == "/sites/test-site"


@patch.dict("os.environ", {}, clear=True)
def test_from_env_raises_when_all_vars_missing():
    """from_env should raise SharePointConfigError when all vars are missing."""
    with pytest.raises(SharePointConfigError, match="Missing required environment variables"):
        SharePointSession.from_env()


@patch.dict("os.environ", {"SHAREPOINT_TENANT_ID": "t", "SHAREPOINT_CLIENT_ID": "c"}, clear=True)
def test_from_env_lists_missing_vars():
    """from_env should name each missing variable in the error message."""
    with pytest.raises(SharePointConfigError) as exc_info:
        SharePointSession.from_env()
    message = str(exc_info.value)
    assert "SHAREPOINT_SITE_HOST" in message
    assert "SHAREPOINT_SITE_PATH" in message
    assert "SHAREPOINT_ROLE_ARN" in message
    # These should NOT appear since they were provided
    assert "SHAREPOINT_TENANT_ID" not in message
    assert "SHAREPOINT_CLIENT_ID" not in message


@patch.dict("os.environ", {**REQUIRED_ENV, "AWS_REGION": "us-east-1"}, clear=True)
def test_from_env_reads_optional_aws_region():
    """from_env should use AWS_REGION if set, otherwise default to eu-west-2."""
    with (
        patch("box2.sharepoint.session.ClientAssertionCredential"),
        patch("box2.sharepoint.session.httpx.Client"),
    ):
        session = SharePointSession.from_env()
        assert session._aws_region == "us-east-1"


@patch.dict("os.environ", REQUIRED_ENV, clear=True)
def test_from_env_defaults_aws_region():
    """from_env should default AWS_REGION to eu-west-2."""
    with (
        patch("box2.sharepoint.session.ClientAssertionCredential"),
        patch("box2.sharepoint.session.httpx.Client"),
    ):
        session = SharePointSession.from_env()
        assert session._aws_region == "eu-west-2"


# ============================================================================
# resolve_site_id Tests
# ============================================================================


def test_resolve_site_id_calls_request(session):
    """resolve_site_id should call request with the correct site reference path."""
    session.request = MagicMock(return_value={"id": "site-id-abc"})

    result = session.resolve_site_id()

    assert result == "site-id-abc"
    session.request.assert_called_once_with("GET", "/sites/contoso.sharepoint.com:/sites/test-site")


def test_resolve_site_id_caches_result(session):
    """resolve_site_id should only call request once even when called multiple times."""
    session.request = MagicMock(return_value={"id": "site-id-abc"})

    first = session.resolve_site_id()
    second = session.resolve_site_id()

    assert first == second == "site-id-abc"
    session.request.assert_called_once()


# ============================================================================
# request Tests
# ============================================================================


def test_request_injects_auth_header(session):
    """request should add an Authorization header with the Bearer token."""
    session.get_token = MagicMock(return_value="fake-token-123")
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"result": "ok"}
    session._http.request = MagicMock(return_value=mock_response)

    session.request("GET", "/sites/123")

    session._http.request.assert_called_once_with(
        "GET",
        "/sites/123",
        headers={"Authorization": "Bearer fake-token-123"},
        json=None,
        params=None,
    )


def test_request_returns_json_on_success(session):
    """request should return parsed JSON on a 2xx response."""
    session.get_token = MagicMock(return_value="token")
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"value": [1, 2, 3]}
    session._http.request = MagicMock(return_value=mock_response)

    result = session.request("GET", "/test")

    assert result == {"value": [1, 2, 3]}


def test_request_returns_empty_dict_on_204(session):
    """request should return an empty dict on a 204 No Content response."""
    session.get_token = MagicMock(return_value="token")
    mock_response = MagicMock()
    mock_response.status_code = 204
    session._http.request = MagicMock(return_value=mock_response)

    result = session.request("DELETE", "/items/1")

    assert result == {}


def test_request_raises_api_error_on_4xx(session):
    """request should raise SharePointAPIError with status and error code on 4xx."""
    session.get_token = MagicMock(return_value="token")
    mock_response = MagicMock()
    mock_response.status_code = 404
    mock_response.content = b'{"error": {"code": "itemNotFound", "message": "Item not found"}}'
    mock_response.json.return_value = {"error": {"code": "itemNotFound", "message": "Item not found"}}
    mock_response.text = "Item not found"
    session._http.request = MagicMock(return_value=mock_response)

    with pytest.raises(SharePointAPIError) as exc_info:
        session.request("GET", "/items/999")
    assert exc_info.value.status_code == 404
    assert exc_info.value.error_code == "itemNotFound"


def test_request_raises_api_error_on_5xx(session):
    """request should raise SharePointAPIError on 5xx server errors."""
    session.get_token = MagicMock(return_value="token")
    mock_response = MagicMock()
    mock_response.status_code = 500
    mock_response.content = b'{"error": {"code": "generalException", "message": "Internal error"}}'
    mock_response.json.return_value = {"error": {"code": "generalException", "message": "Internal error"}}
    mock_response.text = "Internal error"
    session._http.request = MagicMock(return_value=mock_response)

    with pytest.raises(SharePointAPIError) as exc_info:
        session.request("POST", "/lists")
    assert exc_info.value.status_code == 500
    assert exc_info.value.error_code == "generalException"


def test_request_handles_empty_error_body(session):
    """request should handle error responses with no JSON body."""
    session.get_token = MagicMock(return_value="token")
    mock_response = MagicMock()
    mock_response.status_code = 502
    mock_response.content = b""
    mock_response.json.return_value = {}
    mock_response.text = "Bad Gateway"
    session._http.request = MagicMock(return_value=mock_response)

    with pytest.raises(SharePointAPIError) as exc_info:
        session.request("GET", "/test")
    assert exc_info.value.status_code == 502
    assert exc_info.value.error_code is None


def test_request_passes_json_body(session):
    """request should forward the json parameter to the HTTP client."""
    session.get_token = MagicMock(return_value="token")
    mock_response = MagicMock()
    mock_response.status_code = 201
    mock_response.json.return_value = {"id": "new-item"}
    session._http.request = MagicMock(return_value=mock_response)

    body = {"fields": {"Title": "Test"}}
    session.request("POST", "/items", json=body)

    call_kwargs = session._http.request.call_args
    assert call_kwargs.kwargs["json"] == body


def test_request_passes_params(session):
    """request should forward query parameters to the HTTP client."""
    session.get_token = MagicMock(return_value="token")
    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {"value": []}
    session._http.request = MagicMock(return_value=mock_response)

    params = {"$expand": "fields", "$filter": "Title eq 'Test'"}
    session.request("GET", "/items", params=params)

    call_kwargs = session._http.request.call_args
    assert call_kwargs.kwargs["params"] == params


# ============================================================================
# get_token Tests
# ============================================================================


def test_get_token_wraps_credential_failure(session):
    """get_token should wrap credential exceptions in SharePointAuthError."""
    session._credential.get_token = MagicMock(side_effect=RuntimeError("credential expired"))

    with pytest.raises(SharePointAuthError, match="Failed to acquire Graph API token"):
        session.get_token()
