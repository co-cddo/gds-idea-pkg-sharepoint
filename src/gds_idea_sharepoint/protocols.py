"""Structural interfaces for SharePoint resource types.

Protocols define the structural contracts that SharePoint clients must satisfy
to participate in cross-cutting operations like webhook subscriptions. Any
client that exposes the required properties is compatible — no inheritance needed.
"""

from typing import Protocol


class SubscribableResource(Protocol):
    """A SharePoint resource that can be subscribed to for change notifications.

    Any client that exposes a ``resource_path`` property satisfies this protocol.
    For example, ``ListClient`` returns ``/sites/{id}/lists/{id}/items`` and a
    future ``DriveClient`` would return ``/sites/{id}/drives/{id}/root``.
    """

    @property
    def resource_path(self) -> str:
        """Graph API resource path used in subscription creation.

        Returns:
            A path relative to the Graph API root, e.g.
            ``/sites/{site_id}/lists/{list_id}/items``.
        """
        ...
