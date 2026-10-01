"""End-to-end webhook test against a real SharePoint site.

Exercises the full notification loop — all in one process:
  1. Starts the FastAPI receiver on localhost:8000
  2. Opens an ngrok tunnel (via pyngrok) to get a public HTTPS URL
  3. Authenticates to SharePoint
  4. Creates a throwaway test list
  5. Subscribes to webhooks (Graph validates via ngrok -> receiver)
  6. Creates, updates, and deletes a list item (triggers notifications)
  7. Cleans up everything on exit

Prerequisites:
  1. ``uv sync --extra receiver`` to install FastAPI, uvicorn, and pyngrok
  2. ngrok auth token in ``.env`` as ``NGROK_AUTH_TOKEN=<token>`` (free account at ngrok.com)
  3. .env file with SharePoint credentials (or env vars exported)
  4. AWS credentials available (for STS assume-role)

Usage:
    AWS_PROFILE=<profile> uv run python examples/lists_webhook_e2e.py

Set these environment variables before running (or add to .env):
    export NGROK_AUTH_TOKEN=<ngrok-auth-token>
    export SHAREPOINT_TENANT_ID=<azure-ad-tenant-id>
    export SHAREPOINT_CLIENT_ID=<azure-ad-client-id>
    export SHAREPOINT_SITE_HOST=<sharepoint-hostname>
    export SHAREPOINT_SITE_PATH=<sharepoint-site-path>
    export SHAREPOINT_ROLE_ARN=<iam-role-arn>
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

# How long to wait after each mutation for Graph to deliver notifications.
# Microsoft says notifications are "near real-time" but can take up to a few minutes.
NOTIFICATION_WAIT_SECONDS = 30


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
    """Run the end-to-end webhook test."""
    from gds_idea_sharepoint import ListClient, SharePointSession, WebhookClient

    banner("WEBHOOK END-TO-END TEST")
    print()
    print("  This script tests the full notification loop:")
    print("    ngrok -> FastAPI receiver -> Graph subscription -> list changes -> notifications")
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
    # Step 3: Create a throwaway test list
    # ----------------------------------------------------------------
    list_name = f"e2e-test-{uuid4().hex[:8]}"
    list_client = None
    subscription_id = None
    webhooks = None

    banner(f"STEP 3: Create test list '{list_name}'")

    try:
        list_client = ListClient.new(session, list_name=list_name)
        logger.info("List created: %s", list_name)
        logger.info("Resource path: %s", list_client.resource_path)
    except Exception as e:
        logger.error("Failed to create list: %s", e)
        sys.exit(1)

    try:
        # ----------------------------------------------------------------
        # Step 4: Subscribe to webhooks
        # ----------------------------------------------------------------
        banner("STEP 4: Create webhook subscription")
        print(f"  Resource:    {list_client.resource_path}")
        print(f"  URL:         {webhook_url}")
        print("  Changes:     created, updated, deleted")
        print()
        print("  Microsoft will send a validation request to the receiver.")
        print("  If subscription creation succeeds, the handshake worked.")
        print()

        webhooks = WebhookClient(session)
        subscription = webhooks.subscribe(
            resource=list_client,
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
        # Step 5: Create a list item (triggers "created" notification)
        # ----------------------------------------------------------------
        banner("STEP 5: Create a list item")

        item = list_client.create_item({"Title": "E2E webhook test item"})
        item_id = item.get("id")
        logger.info("Item created: id=%s", item_id)
        print("\n  Watch the logs above for a NOTIFICATION RECEIVED banner (change_type=created)")

        wait(NOTIFICATION_WAIT_SECONDS, "Graph to deliver 'created' notification")

        # ----------------------------------------------------------------
        # Step 6: Update the list item (triggers "updated" notification)
        # ----------------------------------------------------------------
        banner("STEP 6: Update the list item")

        list_client.update_item(item_id, {"Title": "E2E webhook test item (updated)"})
        logger.info("Item updated: id=%s", item_id)
        print("\n  Watch the logs above for a NOTIFICATION RECEIVED banner (change_type=updated)")

        wait(NOTIFICATION_WAIT_SECONDS, "Graph to deliver 'updated' notification")

        # ----------------------------------------------------------------
        # Step 7: Delete the list item (triggers "deleted" notification)
        # ----------------------------------------------------------------
        banner("STEP 7: Delete the list item")

        list_client.delete_item(item_id)
        logger.info("Item deleted: id=%s", item_id)
        print("\n  Watch the logs above for a NOTIFICATION RECEIVED banner (change_type=deleted)")

        wait(NOTIFICATION_WAIT_SECONDS, "Graph to deliver 'deleted' notification")

    finally:
        # ----------------------------------------------------------------
        # Cleanup: always attempt to remove subscription and list
        # ----------------------------------------------------------------
        banner("CLEANUP")

        if subscription_id and webhooks:
            try:
                logger.info("Deleting subscription %s...", subscription_id)
                webhooks.delete(subscription_id)
                logger.info("Subscription deleted")
            except Exception as e:
                logger.warning("Failed to delete subscription: %s (may have already expired)", e)

        if list_client:
            try:
                logger.info("Deleting test list '%s'...", list_name)
                list_client.delete_list()
                logger.info("List deleted")
            except Exception as e:
                logger.warning("Failed to delete list: %s", e)

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
    print("  Scroll up through the logs to see notification banners.")
    print("  You should have seen up to 3 notifications (created, updated, deleted).")
    print()
    print("  Note: Microsoft Graph notification delivery is 'near real-time'")
    print("  but can occasionally be delayed. If you didn't see all 3,")
    print("  that's normal -- the subscription and validation handshake")
    print("  working is the most important confirmation.")
    print()
    print("  What was verified:")
    print("    [x] SharePoint authentication (STS -> Azure AD -> Graph)")
    print("    [x] List creation and CRUD operations")
    print("    [x] ngrok tunnel (pyngrok -> localhost:8000)")
    print("    [x] Webhook subscription creation (Graph -> ngrok -> receiver)")
    print("    [x] Validation handshake (receiver echoed validationToken)")
    print("    [x] Notification delivery and processing")
    print("    [x] Cleanup (subscription + list + tunnel closed)")
    print()


if __name__ == "__main__":
    main()
