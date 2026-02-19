"""Test SharePoint authentication.

Verifies the full auth chain:
  AWS STS (assume role) → STS JWT → Azure AD token exchange → Microsoft Graph API call

Set these environment variables before running (or add to .env):
    export AWS_PROFILE=<your-profile>
    export SHAREPOINT_TENANT_ID=<azure-ad-tenant-id>
    export SHAREPOINT_CLIENT_ID=<azure-ad-client-id>
    export SHAREPOINT_SITE_HOST=<sharepoint-hostname>    # e.g. contoso.sharepoint.com
    export SHAREPOINT_SITE_PATH=<sharepoint-site-path>   # e.g. /sites/my-site
    export SHAREPOINT_ROLE_ARN=<iam-role-arn>

Usage:
    uv run python examples/test_sharepoint_auth.py
"""

import logging
import sys

from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


def main():
    from box2.sharepoint import SharePointSession

    # Step 1: Create session from environment variables
    logger.info("Creating SharePoint session from environment variables...")
    try:
        session = SharePointSession.from_env()
        logger.info(
            "Session created (tenant=%s, site=%s:%s)",
            session.tenant_id,
            session.site_host,
            session.site_path,
        )
        logger.info("Will assume role: %s", session._role_arn)
    except Exception as e:
        logger.error("Failed to create session: %s", e)
        sys.exit(1)

    # Step 2: Get a Graph API token
    logger.info("Acquiring Microsoft Graph token...")
    try:
        token = session.get_token()
        # Show first/last few chars only — don't log the full token
        logger.info("Token acquired: %s...%s", token[:10], token[-10:])
    except Exception as e:
        logger.error("Failed to acquire token: %s", e)
        sys.exit(1)

    # Step 3: Make a test API call — resolve site by host + path
    site_ref = f"{session.site_host}:{session.site_path}"
    logger.info("Testing Graph API call: GET /sites/%s ...", site_ref)
    try:
        response = session._http.get(
            f"/sites/{site_ref}",
            headers={"Authorization": f"Bearer {token}"},
        )
        response.raise_for_status()
        site = response.json()
        logger.info("Site ID: %s", site.get("id", "unknown"))
        logger.info("Site name: %s", site.get("displayName", "unknown"))
        logger.info("Site URL: %s", site.get("webUrl", "unknown"))
        logger.info("Authentication successful!")
    except Exception as e:
        logger.error("Graph API call failed: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
