# gds-idea-sharepoint

SharePoint access via the Microsoft Graph API: lists, document libraries, webhook (change notification) subscriptions, and an optional FastAPI receiver for those notifications.

Authentication uses AWS outbound identity federation. AWS STS vends a JWT, and `azure-identity` exchanges it for a Microsoft Graph token through Azure AD's client-assertion grant. No Azure client secret is stored.

Extracted from [`gds-idea-box2.0`](https://github.com/co-cddo/gds-idea-box2.0), where it was the `box2.sharepoint` module.

## Install

Published to the GDS IDEA package index:

```bash
pip install gds-idea-sharepoint --extra-index-url https://co-cddo.github.io/gds-idea-pypi/simple/
```

With `uv`, add the index to your `pyproject.toml` and depend on the package as normal:

```toml
[[tool.uv.index]]
name = "gds-idea"
url = "https://co-cddo.github.io/gds-idea-pypi/simple/"
```

The webhook receiver is optional and brings extra dependencies:

```bash
pip install "gds-idea-sharepoint[receiver]"   # adds FastAPI
pip install "gds-idea-sharepoint[lambda]"     # receiver + Mangum, for AWS Lambda behind API Gateway
```

Importing `gds_idea_sharepoint` never imports the receiver, so the base install does not need FastAPI.

## Usage

```python
from gds_idea_sharepoint import DocsClient, ListClient, SharePointSession

session = SharePointSession.from_env()          # or SharePointSession.from_secret("my-secret")

items = ListClient(session, list_name="Invitations")
items.create_item({"Title": "Hello"})
recent = items.get_recent(minutes=2)

docs = DocsClient(session, library_name="Documents")
for file in docs.get_recent(minutes=2):
    docs.download_file(file, download_dir="/tmp/downloads")
```

Main classes:

| Class | Purpose |
|---|---|
| `SharePointSession` | Authenticated Graph session (`from_env`, `from_secret`) |
| `ListClient` | List CRUD, `upsert_item`, `get_recent`, list creation from a Pydantic model |
| `DocsClient` | Document library listing, upload, download and `get_recent` |
| `WebhookClient` | Create, renew and delete Graph change-notification subscriptions |
| `generate_graph_schema` | Build a Graph list schema from a Pydantic model |

### Webhook receiver

`WebhookClient` registers a subscription; `gds_idea_sharepoint.receiver` is the endpoint Graph then calls. The receiver handles the validation handshake and the `clientState` check, fetches recently changed items, drops the app's own writes, de-duplicates, and calls your handler once per item.

```python
from gds_idea_sharepoint import DocsClient, ListClient, SharePointSession, WebhookClient
from gds_idea_sharepoint.receiver import ReceiverConfig, WebhookRoute, create_app

session = SharePointSession.from_env()
docs = DocsClient(session, library_name="Documents")
reviews = ListClient(session, list_name="Reviews")


async def process_new_file(item: dict) -> None: ...
async def process_review(item: dict) -> None: ...


app = create_app(
    config=ReceiverConfig(client_state="shared-secret", app_identity="<service-principal-app-id>"),
    routes=[
        WebhookRoute(path="/file_uploaded", get_items=lambda: docs.get_recent(minutes=2),
                     handler=process_new_file, filter_self=False),
        WebhookRoute(path="/review_updated", get_items=lambda: reviews.get_recent(minutes=2),
                     handler=process_review, filter_self=True),
    ],
)

# Each route is its own subscription URL:
WebhookClient(session).subscribe(
    resource=reviews,
    notification_url="https://example.com/review_updated",
    client_state="shared-secret",
    change_types=["updated"],
)
```

- **Routes:** one endpoint per subscription. `filter_self=True` skips items whose `lastModifiedBy.application.id` equals `app_identity`.
- **Deduplication:** the dedup record is written before the handler runs, giving at-most-once handling (handlers may call LLMs and are not idempotent). `InMemoryDedup` is for local use only; on AWS Lambda use `DynamoDedup`, because concurrent invocations share no memory.
- **Lambda:** wrap the app with `Mangum(app, lifespan="off")`.

### Configuration

`SharePointSession.from_env()` reads:

| Variable | Description |
|---|---|
| `SHAREPOINT_TENANT_ID` | Azure AD tenant ID |
| `SHAREPOINT_CLIENT_ID` | Azure AD app registration client ID |
| `SHAREPOINT_SITE_HOST` | Site hostname, e.g. `contoso.sharepoint.com` |
| `SHAREPOINT_SITE_PATH` | Site path, e.g. `/sites/my-site` |
| `SHAREPOINT_ROLE_ARN` | IAM role to assume before vending the STS JWT |
| `AWS_REGION` | Optional, defaults to `eu-west-2` |
| `SHAREPOINT_ROLE_SESSION_NAME` | Optional `RoleSessionName` for STS `AssumeRole`, defaults to `box2-sharepoint` |

`SharePointSession.from_secret(name)` reads the same values from an AWS Secrets Manager JSON secret with the keys `tenant_id`, `client_id`, `site_host`, `site_path` and `role_arn`, plus an optional `role_session_name`.

The STS session name appears in CloudTrail and can be matched by `sts:RoleSessionName` conditions in an IAM trust policy, so the default (`box2-sharepoint`) is kept for compatibility. It can also be passed directly as `role_session_name=` to `SharePointSession(...)` or `SharePointSession.from_secret(...)`. It must be 2-64 characters from `[A-Za-z0-9_+=,.@-]`.

## Development

```bash
uv sync
uv run pre-commit install

uv run pytest                                   # unit tests (uv sync installs all extras via the dev group)
AWS_PROFILE=<profile> uv run pytest tests/integration/ -v   # needs AWS credentials and a live SharePoint site

uv run ruff check src/ tests/
uv run ruff format --check src/ tests/
```

Integration tests are skipped automatically when no AWS credentials are available. The webhook integration tests additionally need a public notification URL.

Example scripts are in [`examples/`](examples/).

## Versioning

Versions come from git tags via [hatch-vcs](https://github.com/ofek/hatch-vcs) and are never set by hand. On merge to `main`, the release workflow tags a new version and publishes it to the package index. The bump level is set by the PR label:

- `bump:major`: major version bump
- `bump:minor`: minor version bump
- no label: patch version bump

## Licence

[MIT License](LICENCE)
