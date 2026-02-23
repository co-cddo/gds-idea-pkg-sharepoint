"""SharePoint document library operations via Microsoft Graph API.

Provides file operations on a SharePoint document library (drive). Requires
an authenticated SharePointSession which handles the auth chain and HTTP
requests.

Usage::

    from box2.sharepoint import SharePointSession, DocsClient, WebhookClient

    session = SharePointSession.from_env()

    # Connect to the default "Documents" library
    docs = DocsClient(session)

    # Or connect to a named library
    docs = DocsClient(session, library_name="Shared Documents")

    # Subscribe to file-upload notifications via WebhookClient
    webhooks = WebhookClient(session)
    sub = webhooks.subscribe(
        resource=docs,
        notification_url="https://my-server.example.com/webhook",
        client_state="my-shared-secret",
        change_types=["updated"],
    )

    # Retrieve changed files (delta query)
    files = docs.get_changed_files()
    latest = docs.get_latest_changed_file()

    # Download a file locally
    local_path = docs.download_file(latest)

    # List files in a folder
    children = docs.list_files()
    children = docs.list_files(folder_path="Reports/2026")

    # Get a single file's metadata
    meta = docs.get_file("item-id-123")
"""

import logging
import os
from typing import Any

import httpx

from box2.sharepoint.exceptions import SharePointAPIError
from box2.sharepoint.session import SharePointSession

logger = logging.getLogger(__name__)

# Default fields to request from the delta endpoint.
_DELTA_SELECT_FIELDS = [
    "id",
    "name",
    "size",
    "webUrl",
    "createdDateTime",
    "createdBy",
    "lastModifiedDateTime",
    "parentReference",
    "file",
    "folder",
]


class DocsClient:
    """Client for file operations on a SharePoint document library (drive).

    Resolves the drive ID by library name at construction time and caches it.
    All operations delegate HTTP calls to the session's ``request()`` method.

    Implements the ``SubscribableResource`` protocol so it can be passed
    directly to ``WebhookClient.subscribe()``.
    """

    def __init__(self, session: SharePointSession, library_name: str = "Documents"):
        self._session = session
        self.library_name = library_name

        self._site_id = session.resolve_site_id()
        self._drive_id = self._resolve_drive_id()
        logger.info("DocsClient ready (library=%s, drive_id=%s)", self.library_name, self._drive_id)

    # ── SubscribableResource protocol ─────────────────────────────────────────

    @property
    def resource_path(self) -> str:
        """Graph API resource path for webhook subscriptions.

        Returns:
            The resource path for this drive root, suitable for use with
            ``WebhookClient.subscribe()``.
        """
        return f"/drives/{self._drive_id}/root"

    @property
    def supported_change_types(self) -> set[str]:
        """Change types supported by SharePoint drive subscriptions.

        Drive root change notifications only support ``"updated"`` via the
        Microsoft Graph API. File creation and deletion are surfaced as
        ``"updated"`` change notifications on the drive root resource.

        Returns:
            A set containing ``"updated"``.
        """
        return {"updated"}

    # ── Drive resolution ──────────────────────────────────────────────────────

    def _resolve_drive_id(self) -> str:
        """Resolve the drive ID by library name from the SharePoint site.

        Returns:
            The Graph API drive ID string.

        Raises:
            SharePointAPIError: If the library is not found or the API call fails.
        """
        logger.debug("Resolving drive ID for '%s'", self.library_name)
        data = self._session.request("GET", f"/sites/{self._site_id}/drives")

        drives = data.get("value", [])
        match = next((d for d in drives if d.get("name") == self.library_name), None)

        if match is None:
            available = [d.get("name") for d in drives]
            raise SharePointAPIError(
                f"Library '{self.library_name}' not found. Available: {available}",
                status_code=404,
                error_code="driveNotFound",
            )

        drive_id = match["id"]
        logger.debug("Resolved library '%s' -> %s", self.library_name, drive_id)
        return drive_id

    # ── File operations ───────────────────────────────────────────────────────

    def get_changed_files(self, select: list[str] | None = None) -> list[dict[str, Any]]:
        """Get changed files across the drive using the delta query.

        Calls the delta endpoint to retrieve all changes, then filters out
        folders and deleted items, returning only actual files sorted by
        ``lastModifiedDateTime`` descending (most recent first).

        Args:
            select: Optional list of field names to request. Defaults to a
                standard set including id, name, size, timestamps, and
                parent reference.

        Returns:
            List of file metadata dicts, sorted most-recent-first.

        Raises:
            SharePointAPIError: If the Graph API call fails.
        """
        fields = select or _DELTA_SELECT_FIELDS
        select_param = ",".join(fields)

        data = self._session.request(
            "GET",
            f"/drives/{self._drive_id}/root/delta",
            params={"$select": select_param},
        )

        items = data.get("value", [])

        # Keep only actual files — filter out folders and deleted items
        files = [item for item in items if "file" in item and "deleted" not in item]

        # Sort by lastModifiedDateTime descending (most recent first)
        files.sort(key=lambda x: x.get("lastModifiedDateTime", ""), reverse=True)

        return files

    def get_latest_changed_file(self, select: list[str] | None = None) -> dict[str, Any]:
        """Get the most recently modified file across the drive.

        Convenience wrapper around ``get_changed_files()`` that returns only
        the single most recently modified file.

        Args:
            select: Optional list of field names to request.

        Returns:
            Metadata dict for the most recently modified file.

        Raises:
            SharePointAPIError: If no files are found in the drive delta or
                the Graph API call fails.
        """
        files = self.get_changed_files(select=select)

        if not files:
            raise SharePointAPIError(
                "No files found in drive delta",
                status_code=404,
                error_code="noFilesFound",
            )

        return files[0]

    def download_file(self, file_metadata: dict[str, Any], download_dir: str = "downloads") -> str:
        """Download a file from SharePoint to local disk.

        Uses the ``@microsoft.graph.downloadUrl`` pre-authenticated URL from
        the file metadata when available. Falls back to fetching a fresh item
        to obtain the download URL if it is missing.

        Args:
            file_metadata: File metadata dict as returned by
                ``get_changed_files()`` or ``get_latest_changed_file()``.
                Must contain ``name`` and either
                ``@microsoft.graph.downloadUrl`` or ``id`` and
                ``parentReference.driveId``.
            download_dir: Local directory to save the file. Created if it
                does not exist. Defaults to ``"downloads"``.

        Returns:
            The local file path of the downloaded file.

        Raises:
            SharePointAPIError: If the download URL cannot be obtained or
                the download fails.
        """
        os.makedirs(download_dir, exist_ok=True)
        file_name = file_metadata["name"]

        download_url = file_metadata.get("@microsoft.graph.downloadUrl")

        if not download_url:
            download_url = self._fetch_download_url(file_metadata)

        if not download_url:
            raise SharePointAPIError(
                f"No download URL available for '{file_name}'",
                status_code=404,
                error_code="downloadUrlNotFound",
            )

        logger.info("Downloading '%s' to %s/", file_name, download_dir)
        local_path = os.path.join(download_dir, file_name)

        # Stream download via httpx — the pre-authenticated URL needs no
        # auth headers, so we use a standalone client.
        with httpx.stream("GET", download_url) as response:
            if response.status_code >= 400:
                raise SharePointAPIError(
                    f"File download failed ({response.status_code})",
                    status_code=response.status_code,
                    error_code="downloadFailed",
                )
            with open(local_path, "wb") as f:
                for chunk in response.iter_bytes(chunk_size=8192):
                    f.write(chunk)

        logger.info("Downloaded '%s' -> %s", file_name, local_path)
        return local_path

    def _fetch_download_url(self, file_metadata: dict[str, Any]) -> str | None:
        """Fetch a fresh download URL for a file when the metadata lacks one.

        Args:
            file_metadata: File metadata dict containing ``id`` and
                ``parentReference.driveId``.

        Returns:
            The download URL string, or ``None`` if still unavailable.

        Raises:
            SharePointAPIError: If the Graph API call fails.
        """
        item_id = file_metadata.get("id")
        drive_id = file_metadata.get("parentReference", {}).get("driveId", self._drive_id)

        if not item_id:
            return None

        data = self._session.request("GET", f"/drives/{drive_id}/items/{item_id}")
        return data.get("@microsoft.graph.downloadUrl")

    def list_files(self, folder_path: str = "") -> list[dict[str, Any]]:
        """List files and folders in a drive folder.

        Args:
            folder_path: Path relative to the drive root. Use ``""`` or omit
                for the root folder. Example: ``"Reports/2026"``.

        Returns:
            List of child item dicts (files and folders).

        Raises:
            SharePointAPIError: If the folder is not found or the API call fails.
        """
        if folder_path:
            path = f"/drives/{self._drive_id}/root:/{folder_path}:/children"
        else:
            path = f"/drives/{self._drive_id}/root/children"

        data = self._session.request("GET", path)
        return data.get("value", [])

    def get_file(self, item_id: str) -> dict[str, Any]:
        """Get metadata for a single file by item ID.

        Args:
            item_id: The drive item ID.

        Returns:
            File metadata dict as returned by the Graph API.

        Raises:
            SharePointAPIError: If the item is not found or the API call fails.
        """
        return self._session.request("GET", f"/drives/{self._drive_id}/items/{item_id}")

    def upload_file(self, file_name: str, content: bytes, folder_path: str = "") -> dict[str, Any]:
        """Upload a file to the document library.

        Uses the Graph API simple upload endpoint (``PUT .../content``) for
        files up to 4 MB. For larger files, use the resumable upload session
        API instead.

        This method accesses the session's internal HTTP client directly
        because ``session.request()`` always sends a JSON body, whereas
        file upload requires ``application/octet-stream`` with raw bytes.

        Args:
            file_name: Name for the file in SharePoint (e.g. ``"report.pdf"``).
            content: File content as bytes.
            folder_path: Optional folder path relative to the drive root.
                Use ``""`` for the root folder. Example: ``"Reports/2026"``.

        Returns:
            The created driveItem metadata dict.

        Raises:
            SharePointAPIError: If the upload fails.
        """
        if folder_path:
            path = f"/drives/{self._drive_id}/root:/{folder_path}/{file_name}:/content"
        else:
            path = f"/drives/{self._drive_id}/root:/{file_name}:/content"

        logger.info("Uploading '%s' (%d bytes)", file_name, len(content))

        token = self._session.get_token()
        response = self._session._http.request(
            "PUT",
            path,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/octet-stream",
            },
            content=content,
        )

        if response.status_code >= 400:
            body = response.json() if response.content else {}
            error = body.get("error", {})
            error_code = error.get("code")
            error_message = error.get("message", response.text)
            raise SharePointAPIError(
                f"File upload failed ({response.status_code}): {error_message}",
                status_code=response.status_code,
                error_code=error_code,
            )

        if response.status_code == 204:
            return {}

        logger.info("Uploaded '%s' successfully", file_name)
        return response.json()

    def delete_file(self, item_id: str) -> None:
        """Delete a file from the document library by item ID.

        Args:
            item_id: The drive item ID to delete.

        Raises:
            SharePointAPIError: If the item is not found or the API call fails.
        """
        logger.info("Deleting file item_id=%s", item_id)
        self._session.request("DELETE", f"/drives/{self._drive_id}/items/{item_id}")
        logger.info("Deleted file item_id=%s", item_id)
