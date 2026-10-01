"""Run the webhook receiver locally.

Starts a uvicorn server with the FastAPI receiver app configured for
local E2E testing. The client_state is read from the CLIENT_STATE
environment variable (or defaults to "e2e-test-secret").

Usage:
    # Terminal 2 (after starting ngrok in Terminal 1):
    uv run python examples/run_receiver.py

    # Or with a custom client_state:
    CLIENT_STATE=my-secret uv run python examples/run_receiver.py
"""

import logging
import os
import sys

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

PORT = 8000
DEFAULT_CLIENT_STATE = "e2e-test-secret"


def main() -> None:
    """Start the receiver on localhost."""
    try:
        import uvicorn
    except ImportError:
        logger.error("uvicorn not found. Run: uv sync --extra receiver")
        sys.exit(1)

    from gds_idea_sharepoint.receiver import ReceiverConfig, create_app

    client_state = os.environ.get("CLIENT_STATE", DEFAULT_CLIENT_STATE)

    config = ReceiverConfig(client_state=client_state)
    app = create_app(config)

    logger.info("Starting receiver on http://localhost:%d", PORT)
    logger.info("  client_state = %s", client_state)
    logger.info("  POST /webhook  — notification endpoint")
    logger.info("  GET  /health   — health check")
    logger.info("")
    logger.info("Waiting for notifications...")

    uvicorn.run(app, host="0.0.0.0", port=PORT, log_level="info")


if __name__ == "__main__":
    main()
