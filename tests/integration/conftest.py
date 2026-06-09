"""
Pytest configuration for integration tests.

Integration tests require AWS credentials to call live LLM endpoints.
Without credentials, tests are automatically skipped with a helpful message.
"""

import os

import pytest
from dotenv import load_dotenv

load_dotenv()


def _has_aws_credentials() -> bool:
    """Check if AWS credentials are available for integration tests."""
    if os.environ.get("AWS_PROFILE"):
        return True
    if os.environ.get("AWS_ACCESS_KEY_ID") and os.environ.get("AWS_SECRET_ACCESS_KEY"):
        return True
    return False


def pytest_collection_modifyitems(config, items):
    """Auto-skip integration tests when AWS credentials are not available."""
    if _has_aws_credentials():
        return

    skip_integration = pytest.mark.skip(reason="AWS credentials not available (set AWS_PROFILE or AWS_ACCESS_KEY_ID)")
    for item in items:
        if "integration" in item.keywords:
            item.add_marker(skip_integration)
