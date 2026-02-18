"""Custom exceptions for SharePoint operations."""


class SharePointError(Exception):
    """Base exception for all SharePoint operations."""


class SharePointConfigError(SharePointError):
    """Missing or invalid configuration (e.g. environment variables)."""


class SharePointAuthError(SharePointError):
    """Authentication failure — JWT vending or token exchange."""


class SharePointAPIError(SharePointError):
    """Microsoft Graph API returned an error.

    Attributes:
        status_code: HTTP status code from Graph API.
        error_code: Graph error code (e.g. "itemNotFound"), if available.
    """

    def __init__(self, message: str, status_code: int, error_code: str | None = None):
        self.status_code = status_code
        self.error_code = error_code
        super().__init__(message)
