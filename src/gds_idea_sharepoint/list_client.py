"""SharePoint list operations via Microsoft Graph API.

Provides CRUD operations on SharePoint lists. Requires an authenticated
SharePointSession which handles the auth chain and HTTP requests.

Usage::

    from box2.sharepoint import SharePointSession, ListClient

    session = SharePointSession.from_env()

    # Connect to an existing list
    client = ListClient(session, list_name="My List")

    # Or create a new list
    client = ListClient.new(session, list_name="My New List")

    items = client.get_items()
    item = client.create_item({"Title": "New item", "Status": "Open"})
    client.update_item(item["id"], {"Status": "Closed"})
    client.delete_item(item["id"])
    client.delete_list()
"""

import logging
from typing import Any

from box2.sharepoint.exceptions import SharePointAPIError
from box2.sharepoint.session import SharePointSession

logger = logging.getLogger(__name__)


class ListClient:
    """Client for CRUD operations on a single SharePoint list.

    Resolves the list ID by name at construction time and caches it.
    All operations delegate HTTP calls to the session's ``request()`` method.
    """

    def __init__(self, session: SharePointSession, list_name: str):
        self._session = session
        self.list_name = list_name

        self._site_id = session.resolve_site_id()
        self._list_id = self._resolve_list_id()
        logger.info("ListClient ready (list=%s, list_id=%s)", self.list_name, self._list_id)

    @classmethod
    def new(cls, session: SharePointSession, list_name: str) -> "ListClient":
        """Create a new SharePoint list and return a connected client.

        Creates a generic (custom) list with the given name. If the Graph API
        rejects the request (e.g. the list already exists), a
        SharePointAPIError is raised.

        Args:
            session: An authenticated SharePointSession.
            list_name: Display name for the new list.

        Returns:
            A connected ListClient instance for the newly created list.

        Raises:
            SharePointAPIError: If list creation fails.
        """
        site_id = session.resolve_site_id()

        logger.info("Creating list '%s'", list_name)
        session.request(
            "POST",
            f"/sites/{site_id}/lists",
            json={
                "displayName": list_name,
                "list": {"template": "genericList"},
            },
        )
        logger.info("List '%s' created", list_name)

        return cls(session, list_name)

    def _resolve_list_id(self) -> str:
        """Resolve the list ID by name from the SharePoint site.

        Returns:
            The Graph API list ID string.

        Raises:
            SharePointAPIError: If the list is not found or the API call fails.
        """
        logger.debug("Resolving list ID for '%s'", self.list_name)
        data = self._session.request("GET", f"/sites/{self._site_id}/lists")

        lists = data.get("value", [])
        match = next((lst for lst in lists if lst.get("displayName") == self.list_name), None)

        if match is None:
            available = [lst.get("displayName") for lst in lists]
            raise SharePointAPIError(
                f"List '{self.list_name}' not found. Available: {available}",
                status_code=404,
                error_code="listNotFound",
            )

        list_id = match["id"]
        logger.debug("Resolved list '%s' -> %s", self.list_name, list_id)
        return list_id

    def get_items(
        self,
        select: list[str] | None = None,
        filter_expr: str | None = None,
    ) -> list[dict[str, Any]]:
        """Get all items from the list.

        Args:
            select: Optional list of field names to return.
            filter_expr: Optional OData $filter expression.

        Returns:
            List of item dicts, each containing ``id`` and ``fields``.

        Raises:
            SharePointAPIError: If the Graph API call fails.
        """
        params: dict[str, str] = {"$expand": "fields"}
        if select:
            params["$select"] = ",".join(select)
        if filter_expr:
            params["$filter"] = filter_expr

        data = self._session.request(
            "GET",
            f"/sites/{self._site_id}/lists/{self._list_id}/items",
            params=params,
        )
        return data.get("value", [])

    def get_item(self, item_id: str) -> dict[str, Any]:
        """Get a single item by ID.

        Args:
            item_id: The list item ID.

        Returns:
            Item dict containing ``id`` and ``fields``.

        Raises:
            SharePointAPIError: If the item is not found or the API call fails.
        """
        return self._session.request(
            "GET",
            f"/sites/{self._site_id}/lists/{self._list_id}/items/{item_id}",
            params={"$expand": "fields"},
        )

    def create_item(self, fields: dict[str, Any]) -> dict[str, Any]:
        """Create a new item in the list.

        Args:
            fields: Dict of column name -> value pairs for the new item.

        Returns:
            The created item dict as returned by the Graph API.

        Raises:
            SharePointAPIError: If the API call fails.
        """
        logger.info("Creating item in list '%s'", self.list_name)
        return self._session.request(
            "POST",
            f"/sites/{self._site_id}/lists/{self._list_id}/items",
            json={"fields": fields},
        )

    def update_item(self, item_id: str, fields: dict[str, Any]) -> dict[str, Any]:
        """Update an existing item's fields.

        Args:
            item_id: The list item ID to update.
            fields: Dict of column name -> value pairs to update.

        Returns:
            The updated fields dict as returned by the Graph API.

        Raises:
            SharePointAPIError: If the item is not found or the API call fails.
        """
        logger.info("Updating item %s in list '%s'", item_id, self.list_name)
        return self._session.request(
            "PATCH",
            f"/sites/{self._site_id}/lists/{self._list_id}/items/{item_id}/fields",
            json=fields,
        )

    def delete_item(self, item_id: str) -> None:
        """Delete an item from the list.

        Args:
            item_id: The list item ID to delete.

        Raises:
            SharePointAPIError: If the item is not found or the API call fails.
        """
        logger.info("Deleting item %s from list '%s'", item_id, self.list_name)
        self._session.request(
            "DELETE",
            f"/sites/{self._site_id}/lists/{self._list_id}/items/{item_id}",
        )

    def delete_list(self) -> None:
        """Delete the entire list from the SharePoint site.

        After calling this method the client is no longer usable.

        Raises:
            SharePointAPIError: If the API call fails.
        """
        logger.info("Deleting list '%s' (id=%s)", self.list_name, self._list_id)
        self._session.request(
            "DELETE",
            f"/sites/{self._site_id}/lists/{self._list_id}",
        )
