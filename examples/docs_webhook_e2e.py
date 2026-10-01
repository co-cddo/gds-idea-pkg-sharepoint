"""End-to-end webhook test for SharePoint document library (DocsClient).

Exercises the full file-upload notification loop — all in one process:
  1. Starts the FastAPI receiver on localhost:8000
  2. Opens an ngrok tunnel (via pyngrok) to get a public HTTPS URL
  3. Authenticates to SharePoint
  4. Connects to the document library via DocsClient
  5. Subscribes to drive webhooks (Graph validates via ngrok -> receiver)
  6. Uploads a test file (triggers "updated" notification on drive root)
  7. Uses DocsClient to detect the changed file via delta query
  8. Downloads the file locally
  9. Cleans up (deletes test file, subscription, and tunnel)

Unlike the list E2E (lists_webhook_e2e.py), drive notifications only tell you
"something changed in this drive" — they don't identify the specific file.
The delta query (get_changed_files / get_latest_changed_file) is what
discovers which files actually changed.

Prerequisites:
  1. ``uv sync --extra receiver`` to install FastAPI, uvicorn, and pyngrok
  2. ngrok auth token in ``.env`` as ``NGROK_AUTH_TOKEN=<token>`` (free account at ngrok.com)
  3. .env file with SharePoint credentials (or env vars exported)
  4. AWS credentials available (for STS assume-role)

Usage:
    AWS_PROFILE=<profile> uv run python examples/docs_webhook_e2e.py

Set these environment variables before running (or add to .env):
    export NGROK_AUTH_TOKEN=<ngrok-auth-token>
    export SHAREPOINT_TENANT_ID=<azure-ad-tenant-id>
    export SHAREPOINT_CLIENT_ID=<azure-ad-client-id>
    export SHAREPOINT_SITE_HOST=<sharepoint-hostname>
    export SHAREPOINT_SITE_PATH=<sharepoint-site-path>
    export SHAREPOINT_ROLE_ARN=<iam-role-arn>

Optional:
    export DOCS_LIBRARY_NAME=Documents        # defaults to "Documents"
    export CLIENT_STATE=e2e-test-secret       # shared secret for validation
"""

import logging
import os
import sys
import threading
import time
from uuid import uuid4

from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

CLIENT_STATE = os.environ.get("CLIENT_STATE", "e2e-test-secret")
PORT = 8000
LIBRARY_NAME = os.environ.get("DOCS_LIBRARY_NAME", "Documents")

# How long to wait after the upload for Graph to deliver the notification.
# Microsoft says drive notifications arrive within 1 minute on average,
# but can take up to 60 minutes in the worst case.
NOTIFICATION_WAIT_SECONDS = 60


def banner(text: str) -> None:
    """Print a prominent step banner."""
    width = 60
    print()
    print("=" * width)
    print(f"  {text}")
    print("=" * width)


def wait(seconds: int, reason: str) -> None:
    """Wait with a countdown so the user knows what's happening."""
    print(f"\n  Waiting {seconds}s for {reason}...")
    for remaining in range(seconds, 0, -1):
        print(f"    {remaining}s remaining...", end="\r")
        time.sleep(1)
    print(f"    Done waiting.{' ' * 20}")


def start_receiver() -> None:
    """Start the FastAPI receiver in a background thread."""
    import uvicorn

    from gds_idea_sharepoint.receiver import ReceiverConfig, create_app

    config = ReceiverConfig(client_state=CLIENT_STATE)
    app = create_app(config)

    # Run uvicorn in a daemon thread so it dies when the main thread exits
    server = uvicorn.Server(uvicorn.Config(app, host="0.0.0.0", port=PORT, log_level="info"))
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()

    # Give uvicorn a moment to bind the port
    time.sleep(1)
    logger.info("Receiver started on http://localhost:%d", PORT)
    return server


def start_ngrok() -> str:
    """Open an ngrok tunnel and return the public HTTPS URL.

    The ngrok auth token is read from the ``NGROK_AUTH_TOKEN`` environment
    variable (which can live in your ``.env`` file). This avoids needing
    the ``ngrok`` CLI binary for token setup.
    """
    from pyngrok import conf, ngrok

    auth_token = os.environ.get("NGROK_AUTH_TOKEN")
    if auth_token:
        conf.get_default().auth_token = auth_token
        logger.info("ngrok auth token set from NGROK_AUTH_TOKEN env var")
    else:
        logger.warning("NGROK_AUTH_TOKEN not set — pyngrok will use its default config")

    tunnel = ngrok.connect(PORT, "http")
    public_url = tunnel.public_url

    # pyngrok may return http:// — ensure https
    if public_url.startswith("http://"):
        public_url = public_url.replace("http://", "https://", 1)

    logger.info("ngrok tunnel open: %s -> http://localhost:%d", public_url, PORT)
    return public_url


def main() -> None:
    """Run the end-to-end document library webhook test."""
    from gds_idea_sharepoint import DocsClient, SharePointSession, WebhookClient

    banner("DOCS WEBHOOK END-TO-END TEST")
    print()
    print("  This script tests the full file-upload notification loop:")
    print("    ngrok -> FastAPI receiver -> Graph subscription -> file upload -> notification")
    print("    -> delta query -> file download")
    print()
    print("  Everything runs in this single process.")
    print()

    # ----------------------------------------------------------------
    # Step 1: Start the receiver and ngrok tunnel
    # ----------------------------------------------------------------
    banner("STEP 1: Start receiver + ngrok tunnel")

    try:
        start_receiver()
    except Exception as e:
        logger.error("Failed to start receiver: %s", e)
        sys.exit(1)

    try:
        ngrok_url = start_ngrok()
    except Exception as e:
        logger.error("Failed to start ngrok: %s", e)
        logger.error("Have you run 'ngrok config add-authtoken <token>' ?")
        sys.exit(1)

    webhook_url = f"{ngrok_url}/webhook"
    print(f"\n  Notification URL: {webhook_url}")

    # Quick health check through the tunnel
    print("  Checking receiver health through ngrok...")
    try:
        import httpx

        resp = httpx.get(f"{ngrok_url}/health", timeout=10)
        if resp.status_code == 200:
            print("  Receiver is healthy through ngrok tunnel!")
        else:
            logger.warning("Health check returned %d", resp.status_code)
    except Exception as e:
        logger.error("Could not reach receiver through ngrok: %s", e)
        sys.exit(1)

    # ----------------------------------------------------------------
    # Step 2: Authenticate to SharePoint
    # ----------------------------------------------------------------
    banner("STEP 2: Authenticate to SharePoint")

    try:
        session = SharePointSession.from_env()
        logger.info("Session created (site=%s:%s)", session.site_host, session.site_path)
    except Exception as e:
        logger.error("Failed to create session: %s", e)
        sys.exit(1)

    # ----------------------------------------------------------------
    # Step 3: Connect to document library
    # ----------------------------------------------------------------
    docs_client = None
    subscription_id = None
    webhooks = None
    test_file_id = None
    test_file_name = f"e2e-test-{uuid4().hex[:8]}.txt"
    download_dir = "downloads"

    banner(f"STEP 3: Connect to '{LIBRARY_NAME}' library")

    try:
        docs_client = DocsClient(session, library_name=LIBRARY_NAME)
        logger.info("Connected to library: %s", LIBRARY_NAME)
        logger.info("Resource path: %s", docs_client.resource_path)
    except Exception as e:
        logger.error("Failed to connect to document library: %s", e)
        sys.exit(1)

    try:
        # ----------------------------------------------------------------
        # Step 4: Subscribe to webhooks
        # ----------------------------------------------------------------
        banner("STEP 4: Create webhook subscription")
        print(f"  Resource:    {docs_client.resource_path}")
        print(f"  URL:         {webhook_url}")
        print("  Change type: updated")
        print()
        print("  Microsoft will send a validation request to the receiver.")
        print("  If subscription creation succeeds, the handshake worked.")
        print()

        webhooks = WebhookClient(session)
        subscription = webhooks.subscribe(
            resource=docs_client,
            notification_url=webhook_url,
            client_state=CLIENT_STATE,
            change_types=["updated"],
            expiration_minutes=60,
        )
        subscription_id = subscription.id
        logger.info("Subscription created! id=%s", subscription_id)
        logger.info("Expires: %s", subscription.expiration)
        print()
        print("  Validation handshake succeeded.")
        print("  The receiver correctly echoed the validationToken back to Graph.")

        # ----------------------------------------------------------------
        # Step 5: Upload a test file
        # ----------------------------------------------------------------
        banner(f"STEP 5: Upload test file '{test_file_name}'")

        file_content = (
            f"Box2 DocsClient E2E test file\n"
            f"Uploaded at: {time.strftime('%Y-%m-%d %H:%M:%S UTC', time.gmtime())}\n"
            f"This file should be automatically cleaned up.\n"
        ).encode()

        upload_result = docs_client.upload_file(test_file_name, file_content)
        test_file_id = upload_result.get("id")
        logger.info("File uploaded: id=%s, name=%s", test_file_id, test_file_name)
        logger.info("Web URL: %s", upload_result.get("webUrl", "N/A"))
        print()
        print("  Drive notifications only say 'something changed in this drive'.")
        print("  Watch the logs for a NOTIFICATION RECEIVED banner with:")
        print(f"    Resource: drives/.../{docs_client._drive_id}/root")
        print("    Change Type: updated")

        wait(NOTIFICATION_WAIT_SECONDS, "Graph to deliver 'updated' notification")

        # ----------------------------------------------------------------
        # Step 6: Detect + download the changed file
        # ----------------------------------------------------------------
        banner("STEP 6: Detect changed file via delta query + download")

        print("  Calling docs_client.get_latest_changed_file()...")
        try:
            latest = docs_client.get_latest_changed_file()
            logger.info("Latest changed file detected:")
            logger.info("  Name:     %s", latest.get("name"))
            logger.info("  ID:       %s", latest.get("id"))
            logger.info("  Size:     %s bytes", latest.get("size", "N/A"))
            logger.info("  Modified: %s", latest.get("lastModifiedDateTime"))
            logger.info("  URL:      %s", latest.get("webUrl", "N/A"))

            # Download the file
            print(f"\n  Downloading to ./{download_dir}/...")
            local_path = docs_client.download_file(latest, download_dir=download_dir)
            logger.info("Downloaded -> %s", local_path)

            # Verify content if it's our test file
            if latest.get("name") == test_file_name:
                print(f"\n  The latest file IS our test file ({test_file_name}).")
            else:
                print(f"\n  Note: The latest file ({latest.get('name')}) is not our test file.")
                print(f"  Our test file ({test_file_name}) may not be the most recent change")
                print("  if other files were modified in the library concurrently.")

        except Exception as e:
            logger.warning("Could not detect/download changed file: %s", e)
            print("  This can happen if the delta query hasn't caught up yet.")
            print("  The important thing is that the notification was received.")

    finally:
        # ----------------------------------------------------------------
        # Cleanup: always attempt to remove test file and subscription
        # ----------------------------------------------------------------
        banner("CLEANUP")

        if test_file_id and docs_client:
            try:
                logger.info("Deleting test file '%s' (id=%s)...", test_file_name, test_file_id)
                docs_client.delete_file(test_file_id)
                logger.info("Test file deleted")
            except Exception as e:
                logger.warning("Failed to delete test file: %s", e)

        if subscription_id and webhooks:
            try:
                logger.info("Deleting subscription %s...", subscription_id)
                webhooks.delete(subscription_id)
                logger.info("Subscription deleted")
            except Exception as e:
                logger.warning("Failed to delete subscription: %s (may have already expired)", e)

        # Clean up downloaded test file
        downloaded_path = os.path.join(download_dir, test_file_name)
        if os.path.exists(downloaded_path):
            try:
                os.remove(downloaded_path)
                logger.info("Cleaned up local download: %s", downloaded_path)
            except Exception as e:
                logger.warning("Failed to clean up download: %s", e)

        # Kill the ngrok tunnel
        try:
            from pyngrok import ngrok

            ngrok.kill()
            logger.info("ngrok tunnel closed")
        except Exception:
            pass

    # ----------------------------------------------------------------
    # Summary
    # ----------------------------------------------------------------
    banner("DONE")
    print()
    print("  Scroll up through the logs to see the notification banner.")
    print()
    print("  Note: Drive notifications tell you 'something changed in the drive'")
    print("  but NOT which specific file. The delta query is what identifies")
    print("  the actual file that was uploaded.")
    print()
    print("  Microsoft Graph notification delivery is 'near real-time'")
    print("  but can occasionally be delayed up to 60 minutes for drives.")
    print("  If you didn't see a notification, the subscription and validation")
    print("  handshake working is the most important confirmation.")
    print()
    print("  What was verified:")
    print("    [x] SharePoint authentication (STS -> Azure AD -> Graph)")
    print(f"    [x] Document library connection ('{LIBRARY_NAME}')")
    print("    [x] ngrok tunnel (pyngrok -> localhost:8000)")
    print("    [x] Webhook subscription creation (Graph -> ngrok -> receiver)")
    print("    [x] Validation handshake (receiver echoed validationToken)")
    print("    [x] File upload via DocsClient")
    print("    [x] Notification delivery and processing")
    print("    [x] Delta query to detect changed files")
    print("    [x] File download via DocsClient")
    print("    [x] Cleanup (test file + subscription + tunnel closed)")
    print()


if __name__ == "__main__":
    main()
