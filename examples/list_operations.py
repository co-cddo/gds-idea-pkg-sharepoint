"""Test SharePoint list CRUD operations — full lifecycle.

Exercises every ListClient operation against a real SharePoint site:
  1. Create a new list
  2. Create items
  3. Read items back
  4. Update an item
  5. Confirm the update
  6. Delete the item
  7. Delete the list

The script is self-contained: it creates everything it needs and cleans
up after itself.

Set these environment variables before running (or add to .env):
    export AWS_PROFILE=<your-profile>
    export SHAREPOINT_TENANT_ID=<azure-ad-tenant-id>
    export SHAREPOINT_CLIENT_ID=<azure-ad-client-id>
    export SHAREPOINT_SITE_HOST=<sharepoint-hostname>
    export SHAREPOINT_SITE_PATH=<sharepoint-site-path>
    export SHAREPOINT_ROLE_ARN=<iam-role-arn>

Usage:
    uv run python examples/test_sharepoint_list.py
"""

import logging
import sys

from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

LIST_NAME = "integration-test-list"


def main():
    from box2.sharepoint import ListClient, SharePointSession

    # Step 1: Create session
    logger.info("Creating SharePoint session...")
    try:
        session = SharePointSession.from_env()
        logger.info("Session created (site=%s:%s)", session.site_host, session.site_path)
    except Exception as e:
        logger.error("Failed to create session: %s", e)
        sys.exit(1)

    # Step 2: Create a new list
    logger.info("Creating list '%s'...", LIST_NAME)
    try:
        client = ListClient.new(session, list_name=LIST_NAME)
    except Exception as e:
        logger.error("Failed to create list: %s", e)
        sys.exit(1)

    # Step 3: Create a test item
    logger.info("Creating a test item...")
    try:
        new_item = client.create_item({"Title": "box2 integration test"})
        new_id = new_item.get("id")
        logger.info("Created item id=%s", new_id)
    except Exception as e:
        logger.error("Failed to create item: %s", e)
        sys.exit(1)

    # Step 4: Read all items
    logger.info("Reading items...")
    try:
        items = client.get_items()
        logger.info("Found %d items", len(items))
        for item in items:
            fields = item.get("fields", {})
            logger.info("  - id=%s, Title=%s", item.get("id"), fields.get("Title", "N/A"))
    except Exception as e:
        logger.error("Failed to read items: %s", e)
        sys.exit(1)

    # Step 5: Update the item
    logger.info("Updating item %s...", new_id)
    try:
        updated_fields = client.update_item(new_id, {"Title": "box2 integration test (updated)"})
        logger.info("  Updated Title: %s", updated_fields.get("Title"))
    except Exception as e:
        logger.error("Failed to update item: %s", e)
        sys.exit(1)

    # Step 6: Read it back to confirm
    logger.info("Confirming update...")
    try:
        confirmed = client.get_item(new_id)
        fields = confirmed.get("fields", {})
        logger.info("  Confirmed Title: %s", fields.get("Title"))
    except Exception as e:
        logger.error("Failed to confirm update: %s", e)
        sys.exit(1)

    # Step 7: Delete the item
    logger.info("Deleting item %s...", new_id)
    try:
        client.delete_item(new_id)
        logger.info("Item deleted")
    except Exception as e:
        logger.error("Failed to delete item: %s", e)
        sys.exit(1)

    # Step 8: Delete the list
    logger.info("Deleting list '%s'...", LIST_NAME)
    try:
        client.delete_list()
        logger.info("List deleted")
    except Exception as e:
        logger.error("Failed to delete list: %s", e)
        sys.exit(1)

    logger.info("All list operations successful!")


if __name__ == "__main__":
    main()
