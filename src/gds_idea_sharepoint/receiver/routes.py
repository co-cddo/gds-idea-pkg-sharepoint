"""Webhook route configuration.

Defines the ``WebhookRoute`` dataclass used to declare endpoints in the
receiver app factory. Each route maps a URL path to a callable that fetches
candidate items and a handler that processes each one.
"""

from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field
from typing import Any


@dataclass
class WebhookRoute:
    """A webhook endpoint definition for the receiver app factory.

    Each route becomes a POST endpoint in the FastAPI app. When a notification
    arrives, the factory calls ``get_items()`` to fetch candidate items, applies
    self-write filtering and item-level dedup, then calls ``handler(item)`` for
    each remaining item.

    Attributes:
        path: URL path for this endpoint, e.g. ``"/file_uploaded"``.
            Must start with ``/``.
        get_items: Callable that returns a list of candidate items to consider.
            Typically a lambda wrapping a client's ``get_recent()`` method,
            e.g. ``lambda: list_client.get_recent(minutes=2)``.
        handler: Async callable invoked once per matching item. Receives a
            single item dict (the full Graph API response with fields expanded).
        filter_self: If ``True``, items where ``lastModifiedBy.application.id``
            matches the config's ``app_identity`` are skipped. Set to ``False``
            for resources the app never writes to (e.g. a files list where
            only humans upload). Defaults to ``True``.

    Example::

        WebhookRoute(
            path="/file_uploaded",
            get_items=lambda: docs_client.get_recent(minutes=2),
            handler=process_new_file,
            filter_self=False,
        )
    """

    path: str
    get_items: Callable[[], list[dict[str, Any]]]
    handler: Callable[[dict[str, Any]], Awaitable[None]]
    filter_self: bool = field(default=True)

    def __post_init__(self) -> None:
        """Validate route configuration."""
        if not self.path.startswith("/"):
            raise ValueError(f"Route path must start with '/', got '{self.path}'")
