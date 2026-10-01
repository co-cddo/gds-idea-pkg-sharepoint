"""SharePoint session — authenticated access to Microsoft Graph API.

The auth flow:
1. AWS STS vends a JWT via outbound identity federation
2. azure-identity exchanges that JWT for a Microsoft Graph token
   via Azure AD's client_assertion grant
3. The token is cached and auto-refreshed by azure-identity

Two factory methods are provided:

``SharePointSession.from_env()``
    Reads config from environment variables (see below).

``SharePointSession.from_secret(secret_name)``
    Reads config from an AWS Secrets Manager secret. The secret must be a
    JSON object with keys: ``tenant_id``, ``client_id``, ``site_host``,
    ``site_path``, and ``role_arn`` (optionally ``role_session_name``).
    Eliminates the boilerplate of fetching
    and injecting secrets before calling ``from_env()``.

Required environment variables (for ``from_env``):
    SHAREPOINT_TENANT_ID  — Azure AD tenant ID
    SHAREPOINT_CLIENT_ID  — Azure AD app registration client ID
    SHAREPOINT_SITE_HOST  — SharePoint site hostname (e.g. contoso.sharepoint.com)
    SHAREPOINT_SITE_PATH  — SharePoint site path (e.g. /sites/my-site)
    SHAREPOINT_ROLE_ARN   — IAM role ARN to assume before STS JWT vending

Optional:
    AWS_REGION                    — AWS region for STS calls (default: eu-west-2)
    SHAREPOINT_ROLE_SESSION_NAME  — ``RoleSessionName`` sent to STS AssumeRole
                                    (default: ``box2-sharepoint``)
"""

import json
import logging
import os
import re

import boto3
import httpx
from azure.identity import ClientAssertionCredential

from gds_idea_sharepoint.exceptions import SharePointAPIError, SharePointAuthError, SharePointConfigError

logger = logging.getLogger(__name__)

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Default RoleSessionName sent to STS AssumeRole. It appears in CloudTrail and can be
# matched by IAM trust policy conditions (sts:RoleSessionName), so changing it can
# break existing roles. Override with ``role_session_name`` / SHAREPOINT_ROLE_SESSION_NAME.
DEFAULT_ROLE_SESSION_NAME = "box2-sharepoint"

# STS constraint: 2-64 characters from [\w+=,.@-]
_ROLE_SESSION_NAME_PATTERN = re.compile(r"[\w+=,.@-]{2,64}", re.ASCII)


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
        role_session_name: str = DEFAULT_ROLE_SESSION_NAME,
    ):
        """Create a session.

        Args:
            tenant_id: Azure AD tenant ID.
            client_id: Azure AD app registration client ID.
            site_host: SharePoint site hostname, e.g. ``contoso.sharepoint.com``.
            site_path: SharePoint site path, e.g. ``/sites/my-site``.
            role_arn: IAM role ARN to assume before STS JWT vending.
            aws_region: AWS region for STS calls.
            role_session_name: ``RoleSessionName`` sent to STS AssumeRole. Shows up in
                CloudTrail and can be matched by IAM trust policy conditions.

        Raises:
            SharePointConfigError: If ``role_session_name`` is not a valid STS session
                name (2-64 characters from ``[\\w+=,.@-]``).
        """
        if not _ROLE_SESSION_NAME_PATTERN.fullmatch(role_session_name):
            raise SharePointConfigError(
                f"Invalid role_session_name {role_session_name!r}: must be 2-64 characters from [\\w+=,.@-]"
            )

        self.tenant_id = tenant_id
        self.client_id = client_id
        self.site_host = site_host
        self.site_path = site_path
        self._role_arn = role_arn
        self._aws_region = aws_region
        self._role_session_name = role_session_name
        self._site_id: str | None = None

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
            role_session_name=os.environ.get("SHAREPOINT_ROLE_SESSION_NAME") or DEFAULT_ROLE_SESSION_NAME,
        )

    @classmethod
    def from_secret(
        cls,
        secret_name: str,
        role_arn: str | None = None,
        region: str | None = None,
        role_session_name: str | None = None,
    ) -> "SharePointSession":
        """Create a session by reading config from an AWS Secrets Manager secret.

        The secret must be a JSON object with the following keys:

        .. code-block:: json

            {
                "tenant_id": "...",
                "client_id": "...",
                "site_host": "...",
                "site_path": "...",
                "role_arn":  "..."
            }

        ``role_arn`` may be omitted from the secret if the ``role_arn``
        parameter is supplied instead; the secret JSON takes precedence.

        The secret may also contain an optional ``role_session_name`` key. The
        same precedence applies: secret JSON, then the ``role_session_name``
        parameter, then ``DEFAULT_ROLE_SESSION_NAME``.

        Args:
            secret_name: Name or ARN of the Secrets Manager secret.
            role_arn: Fallback IAM role ARN for JWT vending, used only when
                the secret JSON does not contain a ``role_arn`` key.
            region: AWS region for the Secrets Manager client. Falls back to
                the ``AWS_REGION`` environment variable, then ``"eu-west-2"``.
            role_session_name: Fallback ``RoleSessionName`` for STS AssumeRole, used
                only when the secret JSON does not contain ``role_session_name``.

        Raises:
            SharePointConfigError: If the secret cannot be fetched, is not
                valid JSON, is missing required keys, or no ``role_arn`` can
                be resolved.
        """
        region = region or os.environ.get("AWS_REGION", "eu-west-2")
        logger.debug("Loading SharePoint config from secret '%s' (region=%s)", secret_name, region)

        try:
            sm = boto3.client("secretsmanager", region_name=region)
            response = sm.get_secret_value(SecretId=secret_name)
            secret = json.loads(response["SecretString"])
        except json.JSONDecodeError as e:
            raise SharePointConfigError(f"Secret '{secret_name}' is not valid JSON: {e}") from e
        except Exception as e:
            raise SharePointConfigError(f"Failed to fetch secret '{secret_name}': {e}") from e

        required_keys = ["tenant_id", "client_id", "site_host", "site_path"]
        missing = [k for k in required_keys if not secret.get(k)]
        if missing:
            raise SharePointConfigError(f"Secret '{secret_name}' is missing required keys: {', '.join(missing)}")

        resolved_role_arn = secret.get("role_arn") or role_arn
        if not resolved_role_arn:
            raise SharePointConfigError(
                f"Secret '{secret_name}' does not contain 'role_arn' and no role_arn argument was provided"
            )

        return cls(
            tenant_id=secret["tenant_id"],
            client_id=secret["client_id"],
            site_host=secret["site_host"],
            site_path=secret["site_path"],
            role_arn=resolved_role_arn,
            aws_region=region,
            role_session_name=secret.get("role_session_name") or role_session_name or DEFAULT_ROLE_SESSION_NAME,
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

    def resolve_site_id(self) -> str:
        """Resolve and cache the Graph API site ID from site_host and site_path.

        Calls ``GET /sites/{site_host}:{site_path}`` on first invocation and
        caches the result for subsequent calls.

        Returns:
            The Graph API site ID string.

        Raises:
            SharePointAPIError: If the Graph API call fails.
        """
        if self._site_id is not None:
            return self._site_id

        site_ref = f"{self.site_host}:{self.site_path}"
        logger.debug("Resolving site ID for %s", site_ref)
        data = self.request("GET", f"/sites/{site_ref}")
        self._site_id = data["id"]
        logger.debug("Resolved site ID: %s", self._site_id)
        return self._site_id

    def request(
        self,
        method: str,
        path: str,
        json: dict | None = None,
        params: dict | None = None,
        extra_headers: dict | None = None,
    ) -> dict:
        """Make an authenticated request to the Microsoft Graph API.

        Handles token injection and raises structured errors on failure.

        Args:
            method: HTTP method (GET, POST, PATCH, DELETE).
            path: URL path relative to ``https://graph.microsoft.com/v1.0``.
            json: Optional JSON body for POST/PATCH requests.
            params: Optional query parameters.
            extra_headers: Optional additional headers merged into the request.
                The ``Authorization`` header is always set from the token and
                cannot be overridden here.

        Returns:
            Parsed JSON response as a dict.

        Raises:
            SharePointAPIError: If the Graph API returns a non-2xx status.
            SharePointAuthError: If token acquisition fails.
        """
        token = self.get_token()
        headers = {"Authorization": f"Bearer {token}", **(extra_headers or {})}

        response = self._http.request(method, path, headers=headers, json=json, params=params)

        if response.status_code >= 400:
            body = response.json() if response.content else {}
            error = body.get("error", {})
            error_code = error.get("code")
            error_message = error.get("message", response.text)
            raise SharePointAPIError(
                f"Graph API {method} {path} failed ({response.status_code}): {error_message}",
                status_code=response.status_code,
                error_code=error_code,
            )

        if response.status_code == 204:
            return {}

        return response.json()

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
                RoleSessionName=self._role_session_name,
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
