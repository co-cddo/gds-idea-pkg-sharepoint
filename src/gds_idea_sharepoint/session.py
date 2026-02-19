"""SharePoint session — authenticated access to Microsoft Graph API.

The auth flow:
1. AWS STS vends a JWT via outbound identity federation
2. azure-identity exchanges that JWT for a Microsoft Graph token
   via Azure AD's client_assertion grant
3. The token is cached and auto-refreshed by azure-identity

Required environment variables:
    SHAREPOINT_TENANT_ID  — Azure AD tenant ID
    SHAREPOINT_CLIENT_ID  — Azure AD app registration client ID
    SHAREPOINT_SITE_HOST  — SharePoint site hostname (e.g. contoso.sharepoint.com)
    SHAREPOINT_SITE_PATH  — SharePoint site path (e.g. /sites/my-site)
    SHAREPOINT_ROLE_ARN   — IAM role ARN to assume before STS JWT vending

Optional:
    AWS_REGION            — AWS region for STS calls (default: eu-west-2)
"""

import logging
import os

import boto3
import httpx
from azure.identity import ClientAssertionCredential

from box2.sharepoint.exceptions import SharePointAuthError, SharePointConfigError

logger = logging.getLogger(__name__)

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class SharePointSession:
    """Authenticated session for Microsoft Graph / SharePoint operations.

    Handles the AWS STS → Azure AD → Graph API auth chain.
    Token caching and refresh are handled by azure-identity.
    """

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        site_host: str,
        site_path: str,
        role_arn: str,
        aws_region: str = "eu-west-2",
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.site_host = site_host
        self.site_path = site_path
        self._role_arn = role_arn
        self._aws_region = aws_region

        self._credential = ClientAssertionCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            func=self._vend_aws_jwt,
        )

        self._http = httpx.Client(
            base_url=GRAPH_BASE_URL,
            timeout=30.0,
        )

    @classmethod
    def from_env(cls) -> "SharePointSession":
        """Create a session from environment variables.

        Raises:
            SharePointConfigError: If required environment variables are missing.
        """
        required = {
            "SHAREPOINT_TENANT_ID": "Azure AD tenant ID",
            "SHAREPOINT_CLIENT_ID": "Azure AD app registration client ID",
            "SHAREPOINT_SITE_HOST": "SharePoint site hostname (e.g. contoso.sharepoint.com)",
            "SHAREPOINT_SITE_PATH": "SharePoint site path (e.g. /sites/my-site)",
            "SHAREPOINT_ROLE_ARN": "IAM role ARN to assume before STS JWT vending",
        }
        missing = [f"{k} ({desc})" for k, desc in required.items() if not os.environ.get(k)]
        if missing:
            raise SharePointConfigError("Missing required environment variables:\n  " + "\n  ".join(missing))

        return cls(
            tenant_id=os.environ["SHAREPOINT_TENANT_ID"],
            client_id=os.environ["SHAREPOINT_CLIENT_ID"],
            site_host=os.environ["SHAREPOINT_SITE_HOST"],
            site_path=os.environ["SHAREPOINT_SITE_PATH"],
            role_arn=os.environ["SHAREPOINT_ROLE_ARN"],
            aws_region=os.environ.get("AWS_REGION", "eu-west-2"),
        )

    def get_token(self) -> str:
        """Get a valid Microsoft Graph access token.

        Token caching and refresh are handled automatically by azure-identity.

        Raises:
            SharePointAuthError: If token acquisition fails.
        """
        try:
            token = self._credential.get_token(GRAPH_SCOPE)
            logger.debug("Graph token acquired (expires: %s)", token.expires_on)
            return token.token
        except Exception as e:
            raise SharePointAuthError(f"Failed to acquire Graph API token: {e}") from e

    def _vend_aws_jwt(self) -> str:
        """Get a JWT from AWS STS outbound identity federation.

        Assumes the configured IAM role, then calls STS to vend a JWT.
        This JWT is used as a client_assertion to Azure AD.
        """
        try:
            sts = boto3.client("sts", region_name=self._aws_region)

            logger.debug("Assuming role %s before JWT vending", self._role_arn)
            assumed = sts.assume_role(
                RoleArn=self._role_arn,
                RoleSessionName="box2-sharepoint",
            )
            creds = assumed["Credentials"]
            sts = boto3.client(
                "sts",
                region_name=self._aws_region,
                aws_access_key_id=creds["AccessKeyId"],
                aws_secret_access_key=creds["SecretAccessKey"],
                aws_session_token=creds["SessionToken"],
            )

            resp = sts.get_web_identity_token(
                Audience=["api://AzureADTokenExchange"],
                DurationSeconds=300,
                SigningAlgorithm="RS256",
            )
            logger.debug("AWS STS JWT vended successfully")
            return resp["WebIdentityToken"]
        except SharePointAuthError:
            raise
        except Exception as e:
            raise SharePointAuthError(f"AWS STS JWT vending failed: {e}") from e
