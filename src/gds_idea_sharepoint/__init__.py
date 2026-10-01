"""gds_idea_sharepoint — SharePoint operations via Microsoft Graph API."""

from gds_idea_sharepoint.docs_client import DocsClient
from gds_idea_sharepoint.exceptions import (
    SharePointAPIError,
    SharePointAuthError,
    SharePointConfigError,
    SharePointError,
)
from gds_idea_sharepoint.graph_api_schema import contains_url_type, generate_graph_schema, unwrap_optional
from gds_idea_sharepoint.list_client import ListClient, list_existing
from gds_idea_sharepoint.models import Subscription
from gds_idea_sharepoint.protocols import SubscribableResource
from gds_idea_sharepoint.session import SharePointSession
from gds_idea_sharepoint.webhook_client import WebhookClient

__all__ = [
    "SharePointSession",
    "ListClient",
    "list_existing",
    "DocsClient",
    "WebhookClient",
    "Subscription",
    "SubscribableResource",
    "SharePointError",
    "SharePointConfigError",
    "SharePointAuthError",
    "SharePointAPIError",
    "generate_graph_schema",
    "unwrap_optional",
    "contains_url_type",
]
