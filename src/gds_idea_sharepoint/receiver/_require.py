"""Fail fast, with an actionable message, when the ``receiver`` extra is not installed."""

try:
    import fastapi  # noqa: F401
except ImportError as e:
    raise ImportError(
        "gds_idea_sharepoint.receiver requires FastAPI. Install it with: pip install 'gds-idea-sharepoint[receiver]'"
    ) from e
