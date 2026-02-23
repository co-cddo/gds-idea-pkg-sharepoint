"""Structural interfaces for SharePoint resource types.

Protocols define the structural contracts that SharePoint clients must satisfy
to participate in cross-cutting operations like webhook subscriptions. Any
client that exposes the required properties is compatible — no inheritance needed.
"""

from typing import Protocol


class SubscribableResource(Protocol):
    """A SharePoint resource that can be subscribed to for change notifications.

    Any client that exposes ``resource_path`` and ``supported_change_types``
    properties satisfies this protocol. For example, ``ListClient`` returns
    ``/sites/{id}/lists/{id}`` and ``{"updated"}``, while ``DocsClient``
    returns ``/drives/{id}/root`` and ``{"updated"}``.
    """

    @property
    def resource_path(self) -> str:
        """Graph API resource path used in subscription creation.

        Returns:
            A path relative to the Graph API root, e.g.
            ``/sites/{site_id}/lists/{list_id}`` or ``/drives/{drive_id}/root``.
        """
        ...

    @property
    def supported_change_types(self) -> set[str]:
        """Change types supported by this resource for webhook subscriptions.

        Microsoft Graph supports different change types depending on the
        resource. Both SharePoint lists and drive root items only support
        ``"updated"`` change notifications.

        Returns:
            A set of valid change type strings for this resource.
        """
        ...
