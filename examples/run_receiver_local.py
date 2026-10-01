"""Run the webhook receiver locally with mock routes.

Starts a uvicorn server with route-based endpoints that use fake data
and logging-only handlers. No SharePoint credentials or ngrok required.

This exercises the full dispatch pipeline:
  - Validation handshake (validationToken echo)
  - Client state check
  - Notification-level dedup
  - Item fetching (canned data)
  - Self-write filtering (on /item_reviewed only)
  - Item-level dedup
  - Handler invocation (logs a banner)

Usage:
    uv run python examples/run_receiver_local.py

Then in another terminal, run the smoke test:
    bash examples/smoke_test_receiver.sh
"""

import logging
import sys
from datetime import UTC, datetime, timedelta

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

PORT = 8000
CLIENT_STATE = "local-test-secret"
APP_IDENTITY = "mock-app-identity-001"


# ============================================================================
# Mock data — simulates what get_recent() would return from SharePoint
# ============================================================================

# Timestamps are fixed at startup so that item-level dedup works correctly:
# the same item+timestamp pair is recognised as a duplicate on the second call.
_TIMESTAMP_1_MIN_AGO = (datetime.now(UTC) - timedelta(minutes=1)).strftime("%Y-%m-%dT%H:%M:%SZ")
_TIMESTAMP_NOW = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def mock_files_list_items() -> list[dict]:
    """Simulate items from a Files list (new uploads by humans)."""
    return [
        {
            "id": "file-001",
            "lastModifiedDateTime": _TIMESTAMP_1_MIN_AGO,
            "createdDateTime": _TIMESTAMP_1_MIN_AGO,
            "lastModifiedBy": {"user": {"id": "human-user-alice", "displayName": "Alice"}},
            "fields": {
                "Title": "Ministerial Briefing - AI Policy.docx",
                "FileLeafRef": "AI_Policy_Briefing.docx",
            },
        },
        {
            "id": "file-002",
            "lastModifiedDateTime": _TIMESTAMP_NOW,
            "createdDateTime": _TIMESTAMP_NOW,
            "lastModifiedBy": {"user": {"id": "human-user-bob", "displayName": "Bob"}},
            "fields": {
                "Title": "Constituent Letter - Housing.pdf",
                "FileLeafRef": "Housing_Letter.pdf",
            },
        },
    ]


def mock_processing_list_items() -> list[dict]:
    """Simulate items from a Processing list (mix of app-created and human-edited).

    The first item was last modified by the app (should be filtered by self-write
    filtering). The second was last modified by a human reviewer.
    """
    return [
        {
            "id": "proc-001",
            "lastModifiedDateTime": _TIMESTAMP_1_MIN_AGO,
            "lastModifiedBy": {"application": {"id": APP_IDENTITY}},
            "fields": {
                "Title": "AI Policy Briefing — triage result",
                "Status": "Pending Review",
                "TriageDecision": "Priority",
            },
        },
        {
            "id": "proc-002",
            "lastModifiedDateTime": _TIMESTAMP_NOW,
            "lastModifiedBy": {"user": {"id": "human-reviewer-carol", "displayName": "Carol"}},
            "fields": {
                "Title": "Housing Letter — triage result",
                "Status": "Reviewed",
                "ReviewerComment": "Approved, please draft response.",
            },
        },
    ]


# ============================================================================
# Placeholder handlers — log what would be processed
# ============================================================================


async def handle_new_file(item: dict) -> None:
    """Placeholder handler for newly uploaded files."""
    file_name = item.get("fields", {}).get("FileLeafRef", "unknown")
    banner = (
        "\n"
        "------------------------------------------------------------\n"
        "  HANDLER: handle_new_file\n"
        f"    Item ID:   {item['id']}\n"
        f"    File:      {file_name}\n"
        f"    Action:    Would run triage + extraction pipeline\n"
        "------------------------------------------------------------"
    )
    logger.info(banner)


async def handle_human_review(item: dict) -> None:
    """Placeholder handler for human-reviewed processing items."""
    status = item.get("fields", {}).get("Status", "unknown")
    comment = item.get("fields", {}).get("ReviewerComment", "(none)")
    banner = (
        "\n"
        "------------------------------------------------------------\n"
        "  HANDLER: handle_human_review\n"
        f"    Item ID:   {item['id']}\n"
        f"    Status:    {status}\n"
        f"    Comment:   {comment}\n"
        f"    Action:    Would run redrafting pipeline\n"
        "------------------------------------------------------------"
    )
    logger.info(banner)


# ============================================================================
# Main
# ============================================================================


def main() -> None:
    """Start the local receiver with mock routes."""
    try:
        import uvicorn
    except ImportError:
        logger.error("uvicorn not found. Run: uv sync --extra receiver")
        sys.exit(1)

    from gds_idea_sharepoint.receiver import ReceiverConfig, WebhookRoute, create_app

    config = ReceiverConfig(client_state=CLIENT_STATE, app_identity=APP_IDENTITY)

    routes = [
        WebhookRoute(
            path="/file_uploaded",
            get_items=mock_files_list_items,
            handler=handle_new_file,
            filter_self=False,
        ),
        WebhookRoute(
            path="/item_reviewed",
            get_items=mock_processing_list_items,
            handler=handle_human_review,
            filter_self=True,
        ),
    ]

    app = create_app(config, routes=routes)

    print()
    print("=" * 64)
    print("  Webhook Receiver — LOCAL MODE (mock data)")
    print("=" * 64)
    print()
    print(f"  Server:        http://localhost:{PORT}")
    print(f"  Client state:  {CLIENT_STATE}")
    print(f"  App identity:  {APP_IDENTITY}")
    print()
    print("  Routes:")
    print("    POST /file_uploaded   — new file handler (filter_self=False)")
    print("    POST /item_reviewed   — human review handler (filter_self=True)")
    print("    GET  /health          — health check")
    print()
    print("  Smoke test (run in another terminal):")
    print()
    print("    bash examples/smoke_test_receiver.sh")
    print()
    print("=" * 64)
    print()

    uvicorn.run(app, host="0.0.0.0", port=PORT, log_level="info")


if __name__ == "__main__":
    main()
