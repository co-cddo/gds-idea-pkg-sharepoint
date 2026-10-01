"""Tests that the receiver is an optional extra and never leaks into the base import."""

import subprocess
import sys
import textwrap


def _run_without_optional_deps(body: str) -> subprocess.CompletedProcess[str]:
    """Run *body* in a fresh interpreter in which fastapi and mangum cannot be imported."""
    prelude = textwrap.dedent(
        """
        import sys
        sys.modules["fastapi"] = None  # makes `import fastapi` raise ImportError
        sys.modules["mangum"] = None
        """
    )
    return subprocess.run(
        [sys.executable, "-c", prelude + textwrap.dedent(body)],
        capture_output=True,
        text=True,
        check=False,
    )


def test_base_import_works_without_receiver_extra():
    """`import gds_idea_sharepoint` must not need FastAPI, and must not import the receiver."""
    result = _run_without_optional_deps(
        """
        import gds_idea_sharepoint
        assert "gds_idea_sharepoint.receiver" not in sys.modules
        assert gds_idea_sharepoint.SharePointSession
        """
    )
    assert result.returncode == 0, result.stderr


def test_receiver_import_without_extra_gives_actionable_error():
    """Importing the receiver without FastAPI explains which extra to install."""
    result = _run_without_optional_deps(
        """
        try:
            import gds_idea_sharepoint.receiver
        except ImportError as e:
            print(e)
        else:
            raise SystemExit("expected ImportError")
        """
    )
    assert result.returncode == 0, result.stderr
    assert "gds-idea-sharepoint[receiver]" in result.stdout


def test_receiver_imports_with_extra_installed():
    """With the extra installed the receiver's public API is importable."""
    from gds_idea_sharepoint.receiver import ReceiverConfig, WebhookRoute, create_app, is_self_write

    assert all([ReceiverConfig, WebhookRoute, create_app, is_self_write])
