"""SharePoint list operations via Microsoft Graph API.

Provides CRUD operations on SharePoint lists. Requires an authenticated
SharePointSession which handles the auth chain and HTTP requests.

Usage::

    from box2.sharepoint import SharePointSession, ListClient, list_existing

    session = SharePointSession.from_env()

    # See what lists exist
    names = list_existing(session)

    # Connect to an existing list
    client = ListClient(session, list_name="My List")

    # Or create a new list
    client = ListClient.new(session, list_name="My New List")

    # Or create with a Pydantic-derived schema
    client = ListClient.new_with_schema(session, "Invitations", SharepointInvitation)

    # Or ensure a list exists (create if missing, connect if present)
    client = ListClient.ensure(session, "Invitations", SharepointInvitation)

    items = client.get_items()
    item = client.create_item({"Title": "New item", "Status": "Open"})
    client.upsert_item({"Title": "New item", "Status": "Closed"})  # updates existing
    client.upsert_item({"Title": "Another item"})                  # creates new
    client.update_item(item["id"], {"Status": "Closed"})
    client.delete_item(item["id"])
    client.delete_list()
"""

import logging
from datetime import UTC, datetime, timedelta
from typing import Any

from pydantic import BaseModel

from box2.sharepoint.exceptions import SharePointAPIError
from box2.sharepoint.graph_api_schema import generate_graph_schema
from box2.sharepoint.session import SharePointSession

logger = logging.getLogger(__name__)


def list_existing(session: SharePointSession) -> list[str]:
    """Return display names of all lists on the SharePoint site.

    Args:
        session: An authenticated SharePointSession.

    Returns:
        Sorted list of display names for all lists on the site.

    Raises:
        SharePointAPIError: If the Graph API call fails.
    """
    site_id = session.resolve_site_id()
    resp = session.request("GET", f"/sites/{site_id}/lists")
    return sorted(lst["displayName"] for lst in resp.get("value", []))


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

    @classmethod
    def new_with_schema(
        cls,
        session: SharePointSession,
        list_name: str,
        model: type[BaseModel],
    ) -> "ListClient":
        """Create a new SharePoint list with columns from a Pydantic model.

        Generates the Graph API column schema from the model's field
        definitions using ``generate_graph_schema()``, creates the list,
        and returns a connected client.

        Args:
            session: An authenticated SharePointSession.
            list_name: Display name for the new list.
            model: A Pydantic BaseModel subclass defining the list columns.

        Returns:
            A connected ListClient instance for the newly created list.

        Raises:
            SharePointAPIError: If list creation fails (e.g. list already exists).
        """
        site_id = session.resolve_site_id()
        payload = generate_graph_schema(model, list_name)

        logger.info("Creating list '%s' with schema from %s", list_name, model.__name__)
        session.request("POST", f"/sites/{site_id}/lists", json=payload)
        logger.info("List '%s' created with schema", list_name)

        return cls(session, list_name)

    @classmethod
    def ensure(
        cls,
        session: SharePointSession,
        list_name: str,
        model: type[BaseModel],
    ) -> "ListClient":
        """Ensure a SharePoint list exists, creating it if missing.

        Idempotent: if the list already exists, connects to it and returns
        a client. If it doesn't exist, creates it with columns derived
        from the Pydantic model via ``new_with_schema()``, then returns a
        connected client.

        Args:
            session: An authenticated SharePointSession.
            list_name: Display name for the list.
            model: A Pydantic BaseModel subclass defining the list columns.
                Only used if the list needs to be created.

        Returns:
            A connected ListClient instance.

        Raises:
            SharePointAPIError: If the Graph API calls fail.
        """
        existing = list_existing(session)
        if list_name in existing:
            logger.info("List '%s' already exists, connecting", list_name)
            return cls(session, list_name)

        logger.info("List '%s' not found, creating with schema from %s", list_name, model.__name__)
        return cls.new_with_schema(session, list_name, model)

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

    @property
    def resource_path(self) -> str:
        """Graph API resource path for webhook subscriptions.

        Returns:
            The resource path for this list, suitable for use with
            ``WebhookClient.subscribe()``.
        """
        return f"/sites/{self._site_id}/lists/{self._list_id}"

    @property
    def supported_change_types(self) -> set[str]:
        """Change types supported by SharePoint list subscriptions.

        SharePoint lists only support ``"updated"`` notifications via the
        Microsoft Graph API. Item creation and deletion are surfaced as
        ``"updated"`` change notifications on the list resource.

        Returns:
            A set containing ``"updated"``.
        """
        return {"updated"}

    def get_items(
        self,
        select: list[str] | None = None,
        filter_expr: str | None = None,
        prefer_unindexed: bool = False,
    ) -> list[dict[str, Any]]:
        """Get all items from the list.

        Args:
            select: Optional list of field names to return.
            filter_expr: Optional OData $filter expression.
            prefer_unindexed: If True, sends the
                ``Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`` header,
                allowing filters on non-indexed columns. Safe for small lists;
                may fail unpredictably on large ones.

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

        extra_headers = {"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"} if prefer_unindexed else None

        data = self._session.request(
            "GET",
            f"/sites/{self._site_id}/lists/{self._list_id}/items",
            params=params,
            extra_headers=extra_headers,
        )
        return data.get("value", [])

    def get_recent(self, minutes: int = 2) -> list[dict[str, Any]]:
        """Get items modified in the last N minutes.

        Convenience wrapper around ``get_items()`` that builds the OData
        ``$filter`` expression for recent modifications. Useful for webhook
        handlers that need to find what changed since the last notification.

        Args:
            minutes: Lookback window in minutes. Defaults to 2.

        Returns:
            List of item dicts with fields expanded.

        Raises:
            ValueError: If minutes is not positive.
            SharePointAPIError: If the Graph API call fails.
        """
        if minutes <= 0:
            raise ValueError(f"minutes must be positive, got {minutes}")

        cutoff = (datetime.now(UTC) - timedelta(minutes=minutes)).strftime("%Y-%m-%dT%H:%M:%SZ")
        return self.get_items(filter_expr=f"fields/Modified gt '{cutoff}'")

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

    def upsert_item(self, fields: dict[str, Any], *, key_field: str = "Title") -> dict[str, Any]:
        """Create or update a list item, matching on a key field.

        Queries for an existing item where ``key_field`` matches the value in
        ``fields``. Creates the item if none is found; updates it if exactly one
        is found. Raises if the key is ambiguous (multiple matches) or missing.

        Args:
            fields: Dict of column name -> value pairs. Must include ``key_field``.
            key_field: Column name used to look up the existing item. Defaults
                to ``"Title"``.

        Returns:
            The item dict with ``id`` and ``fields`` keys, reflecting the
            state after the create or update.

        Raises:
            ValueError: If ``key_field`` is not present in ``fields``.
            SharePointAPIError: If multiple items match the key (ambiguous),
                or if any underlying Graph API call fails.
        """
        if key_field not in fields:
            raise ValueError(f"key_field '{key_field}' must be present in fields")

        key_value = str(fields[key_field]).replace("'", "''")  # OData single-quote escape
        existing = self.get_items(filter_expr=f"fields/{key_field} eq '{key_value}'", prefer_unindexed=True)

        if len(existing) > 1:
            raise SharePointAPIError(
                f"upsert_item: {len(existing)} items match {key_field}='{fields[key_field]}' in list '{self.list_name}'",
                status_code=409,
                error_code="ambiguousKey",
            )

        if not existing:
            logger.info("upsert_item: no match on %s — creating", key_field)
            return self.create_item(fields)

        match = existing[0]
        item_id = match["id"]
        logger.info("upsert_item: match %s — updating", item_id)
        updated_fields = self.update_item(item_id, fields)
        return {"id": item_id, "fields": updated_fields}

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
