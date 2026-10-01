# gds-idea-sharepoint

SharePoint access via the Microsoft Graph API: lists, document libraries and webhook (change notification) subscriptions.

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

`SharePointSession.from_secret(name)` reads the same values from an AWS Secrets Manager JSON secret with the keys `tenant_id`, `client_id`, `site_host`, `site_path` and `role_arn`.

## Development

```bash
uv sync
uv run pre-commit install

uv run pytest                                   # unit tests
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
