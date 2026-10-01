"""gds_idea_sharepoint.receiver — FastAPI webhook receiver for Microsoft Graph notifications.

This is an optional submodule and is never imported by ``import gds_idea_sharepoint``.
Install the dependencies it needs with ``pip install gds-idea-sharepoint[receiver]``
(FastAPI). To run it on AWS Lambda behind API Gateway, use ``gds-idea-sharepoint[lambda]``,
which adds Mangum. Serving it locally additionally needs an ASGI server such as uvicorn.

Subscriptions are created with :class:`gds_idea_sharepoint.WebhookClient`; this module
receives the notifications Graph then sends to them.

Usage::

    from gds_idea_sharepoint.receiver import create_app, ReceiverConfig, WebhookRoute

    config = ReceiverConfig(
        client_state="my-shared-secret",
        app_identity="<service-principal-app-id>",
    )

    app = create_app(
        config=config,
        routes=[
            WebhookRoute(
                path="/file_uploaded",
                get_items=lambda: docs_client.get_recent(minutes=2),
                handler=process_new_file,
                filter_self=False,
            ),
        ],
    )
"""

from gds_idea_sharepoint.receiver import _require  # noqa: F401  (fails fast if the extra is missing)
from gds_idea_sharepoint.receiver.app import create_app
from gds_idea_sharepoint.receiver.config import ReceiverConfig
from gds_idea_sharepoint.receiver.dedup import (
    DeduplicationStore,
    DynamoDedup,
    InMemoryDedup,
    build_item_dedup_key,
)
from gds_idea_sharepoint.receiver.handlers import is_self_write
from gds_idea_sharepoint.receiver.models import Notification, NotificationPayload, ResourceData
from gds_idea_sharepoint.receiver.routes import WebhookRoute

__all__ = [
    "create_app",
    "ReceiverConfig",
    "WebhookRoute",
    "DeduplicationStore",
    "DynamoDedup",
    "InMemoryDedup",
    "build_item_dedup_key",
    "is_self_write",
    "Notification",
    "NotificationPayload",
    "ResourceData",
]
