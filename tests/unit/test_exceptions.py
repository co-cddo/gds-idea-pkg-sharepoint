import pytest

from box2.sharepoint.exceptions import (
    SharePointAPIError,
    SharePointAuthError,
    SharePointConfigError,
    SharePointError,
)

# ============================================================================
# Exception Hierarchy Tests
# ============================================================================


def test_all_exceptions_inherit_from_base():
    """All SharePoint exceptions should be subclasses of SharePointError."""
    assert issubclass(SharePointConfigError, SharePointError)
    assert issubclass(SharePointAuthError, SharePointError)
    assert issubclass(SharePointAPIError, SharePointError)


def test_base_inherits_from_exception():
    """SharePointError should be a subclass of the built-in Exception."""
    assert issubclass(SharePointError, Exception)


# ============================================================================
# SharePointAPIError Tests
# ============================================================================


def test_api_error_stores_status_code():
    """SharePointAPIError should store the HTTP status code."""
    error = SharePointAPIError("Not found", status_code=404, error_code="itemNotFound")
    assert error.status_code == 404


def test_api_error_stores_error_code():
    """SharePointAPIError should store the Graph error code."""
    error = SharePointAPIError("Not found", status_code=404, error_code="itemNotFound")
    assert error.error_code == "itemNotFound"


def test_api_error_preserves_message():
    """SharePointAPIError should preserve the error message via str()."""
    error = SharePointAPIError("Something went wrong", status_code=500)
    assert str(error) == "Something went wrong"


def test_api_error_with_no_error_code():
    """SharePointAPIError should accept error_code=None (the default)."""
    error = SharePointAPIError("Forbidden", status_code=403)
    assert error.status_code == 403
    assert error.error_code is None


# ============================================================================
# Simple Exception Tests
# ============================================================================


def test_config_error_preserves_message():
    """SharePointConfigError should preserve the error message."""
    error = SharePointConfigError("Missing SHAREPOINT_TENANT_ID")
    assert "Missing SHAREPOINT_TENANT_ID" in str(error)


def test_auth_error_preserves_message():
    """SharePointAuthError should preserve the error message."""
    error = SharePointAuthError("Token acquisition failed")
    assert "Token acquisition failed" in str(error)


def test_api_error_is_catchable_as_base():
    """Catching SharePointError should also catch SharePointAPIError."""
    with pytest.raises(SharePointError):
        raise SharePointAPIError("fail", status_code=500)
