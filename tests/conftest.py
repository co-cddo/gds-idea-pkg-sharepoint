"""
Pytest configuration and shared fixtures.

Integration tests (marked @pytest.mark.integration) require AWS credentials
to call live LLM endpoints. They are:
- Excluded in CI via: pytest -m "not integration"
- Auto-skipped locally if AWS credentials are not configured
"""

import os

import pytest


def _has_aws_credentials() -> bool:
    """Check if AWS credentials are available for integration tests."""
    # Check for explicit profile (preferred)
    if os.environ.get("AWS_PROFILE"):
        return True
    # Check for explicit key-based credentials
    if os.environ.get("AWS_ACCESS_KEY_ID") and os.environ.get(
        "AWS_SECRET_ACCESS_KEY"
    ):
        return True
    return False


def pytest_collection_modifyitems(config, items):
    """Auto-skip integration tests when AWS credentials are not available."""
    if _has_aws_credentials():
        return

    skip_integration = pytest.mark.skip(
        reason="AWS credentials not available (set AWS_PROFILE or AWS_ACCESS_KEY_ID)"
    )
    for item in items:
        if "integration" in item.keywords:
            item.add_marker(skip_integration)
