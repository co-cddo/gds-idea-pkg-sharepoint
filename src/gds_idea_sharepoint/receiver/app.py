"""FastAPI webhook receiver application.

Provides a factory function that creates a configured FastAPI app for
receiving Microsoft Graph change notifications. The app dynamically
registers one endpoint per ``WebhookRoute``, each handling the full
notification-to-handler pipeline.

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
            WebhookRoute(
                path="/item_reviewed",
                get_items=lambda: processing_list.get_recent(minutes=2),
                handler=process_human_edit,
                filter_self=True,
            ),
        ],
    )
"""

import json
import logging
from typing import Annotated

from fastapi import BackgroundTasks, Depends, FastAPI, Request, Response
from fastapi.responses import JSONResponse, PlainTextResponse

from gds_idea_sharepoint.receiver.config import ReceiverConfig
from gds_idea_sharepoint.receiver.dedup import DeduplicationStore, InMemoryDedup
from gds_idea_sharepoint.receiver.handlers import dispatch_route
from gds_idea_sharepoint.receiver.models import NotificationPayload
from gds_idea_sharepoint.receiver.routes import WebhookRoute

logger = logging.getLogger(__name__)


def create_app(
    config: ReceiverConfig,
    routes: list[WebhookRoute] | None = None,
    dedup_store: DeduplicationStore | None = None,
) -> FastAPI:
    """Create a configured FastAPI webhook receiver.

    Registers one POST endpoint per route. Each endpoint handles the
    Microsoft Graph validation handshake and delegates notification
    processing to the route's pipeline (client_state check, item query
    via ``get_items()``, self-write filtering, item-level dedup,
    handler call).

    If no routes are provided, a single ``/webhook`` endpoint is created
    that accepts notifications but only logs them (useful for E2E testing
    of the tunnel and subscription handshake).

    Args:
        config: Receiver configuration (client_state, app_identity, etc.).
        routes: List of ``WebhookRoute`` definitions. Each becomes a POST
            endpoint in the app.
        dedup_store: Optional deduplication store. If not provided, an
            ``InMemoryDedup`` is created using the config's dedup window.

    Returns:
        A FastAPI application instance with webhook and health routes.
    """
    if dedup_store is None:
        dedup_store = InMemoryDedup(window_seconds=config.dedup_window_seconds)

    app = FastAPI(title="Webhook Receiver", docs_url=None, redoc_url=None)

    def get_config() -> ReceiverConfig:
        return config

    def get_dedup_store() -> DeduplicationStore:
        return dedup_store

    if routes:
        for route in routes:
            _register_route(app, route, get_config, get_dedup_store)
            logger.info("Registered webhook route: POST %s", route.path)
    else:
        _register_fallback_route(app, get_config, get_dedup_store)
        logger.info("Registered fallback webhook route: POST /webhook")

    @app.get("/health")
    async def health() -> dict:
        """Health check endpoint."""
        return {"status": "ok"}

    return app


def _register_route(
    app: FastAPI,
    route: WebhookRoute,
    get_config: callable,
    get_dedup_store: callable,
) -> None:
    """Register a POST endpoint for a single WebhookRoute.

    Args:
        app: The FastAPI application.
        route: The route to register.
        get_config: Dependency provider for ReceiverConfig.
        get_dedup_store: Dependency provider for DeduplicationStore.
    """

    # Capture route in a closure variable to avoid late-binding issues
    # in the for-loop. A default parameter (e.g. _route=route) would be
    # treated by FastAPI as a request body field, causing 422 errors.
    bound_route = route

    @app.post(bound_route.path)
    async def webhook_endpoint(
        request: Request,
        background_tasks: BackgroundTasks,
        receiver_config: Annotated[ReceiverConfig, Depends(get_config)],
        store: Annotated[DeduplicationStore, Depends(get_dedup_store)],
    ) -> Response:
        """Handle Microsoft Graph webhook notifications for a route.

        Returns 202 immediately and processes the notification in a
        background task. This prevents Graph from timing out when
        handlers take a long time (e.g. LLM calls). With Mangum,
        background tasks complete before the Lambda invocation ends.
        """
        # Validation handshake
        validation_token = request.query_params.get("validationToken")
        if validation_token:
            logger.info("Validation handshake on %s — echoing token", bound_route.path)
            return PlainTextResponse(content=validation_token, status_code=200)

        # Parse notification payload
        body = await request.json()
        if not body:
            return JSONResponse(content={"status": "no data"}, status_code=400)

        logger.debug("Raw request body on %s: %s", bound_route.path, json.dumps(body, default=str))

        payload = NotificationPayload.model_validate(body)

        # Dispatch in background — response goes back to Graph immediately
        handler_name = getattr(bound_route.handler, "__name__", repr(bound_route.handler))
        logger.info("Route %s: dispatching background task → %s", bound_route.path, handler_name)
        background_tasks.add_task(dispatch_route, bound_route, payload, receiver_config, store)

        return JSONResponse(content={"status": "accepted"}, status_code=202)


def _register_fallback_route(
    app: FastAPI,
    get_config: callable,
    get_dedup_store: callable,
) -> None:
    """Register a fallback /webhook endpoint for E2E testing.

    This endpoint accepts notifications and logs them but does not query
    any list or call any handler. Useful for testing the tunnel and
    subscription handshake without wiring up real routes.

    Args:
        app: The FastAPI application.
        get_config: Dependency provider for ReceiverConfig.
        get_dedup_store: Dependency provider for DeduplicationStore.
    """
    from gds_idea_sharepoint.receiver.handlers import _log_notification

    @app.post("/webhook")
    async def webhook_fallback(
        request: Request,
        receiver_config: Annotated[ReceiverConfig, Depends(get_config)],
        store: Annotated[DeduplicationStore, Depends(get_dedup_store)],
    ) -> Response:
        """Fallback webhook endpoint — logs notifications without processing."""
        validation_token = request.query_params.get("validationToken")
        if validation_token:
            logger.info("Validation handshake received — echoing token")
            return PlainTextResponse(content=validation_token, status_code=200)

        body = await request.json()
        if not body:
            return JSONResponse(content={"status": "no data"}, status_code=400)

        logger.debug("Raw request body on /webhook: %s", json.dumps(body, default=str))

        payload = NotificationPayload.model_validate(body)
        logged = 0

        for notification in payload.value:
            if notification.client_state != receiver_config.client_state:
                continue

            _log_notification(notification)
            logged += 1

        logger.info("Fallback: logged %d notification(s)", logged)
        return JSONResponse(content={"status": "accepted"}, status_code=202)
