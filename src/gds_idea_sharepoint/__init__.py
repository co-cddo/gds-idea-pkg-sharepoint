"""box2.sharepoint — SharePoint operations via Microsoft Graph API."""

from box2.sharepoint.docs_client import DocsClient
from box2.sharepoint.exceptions import (
    SharePointAPIError,
    SharePointAuthError,
    SharePointConfigError,
    SharePointError,
)
from box2.sharepoint.graph_api_schema import generate_graph_schema
from box2.sharepoint.list_client import ListClient
from box2.sharepoint.models import Subscription
from box2.sharepoint.protocols import SubscribableResource
from box2.sharepoint.session import SharePointSession
from box2.sharepoint.webhook_client import WebhookClient

__all__ = [
    "SharePointSession",
    "ListClient",
    "DocsClient",
    "WebhookClient",
    "Subscription",
    "SubscribableResource",
    "SharePointError",
    "SharePointConfigError",
    "SharePointAuthError",
    "SharePointAPIError",
    "generate_graph_schema",
]
