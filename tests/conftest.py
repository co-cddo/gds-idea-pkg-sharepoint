"""
Root pytest configuration and shared fixtures.

Test layout:
  tests/unit/          — fast tests with no external dependencies
  tests/integration/   — tests requiring AWS credentials and live LLM

Integration test auto-skip logic lives in tests/integration/conftest.py.
"""
