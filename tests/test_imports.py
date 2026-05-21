"""Import sanity checks — modules load on any platform.

We do not import `outlook` (which requires pywin32) or `server` (which
imports the MCP runtime) here so the test suite is platform-portable. The
safety scan is the load-bearing CI gate; these tests confirm the static
structure parses cleanly.
"""

from __future__ import annotations


def test_safety_module_imports() -> None:
    from outlook_readonly_mcp import safety

    assert safety.FORBIDDEN_METHODS
    assert safety.FORBIDDEN_PROPERTIES
    assert safety.ALL_FORBIDDEN
    assert safety.OlDefaultFolders.olFolderInbox == 6


def test_models_module_imports() -> None:
    from outlook_readonly_mcp import models

    e = models.EmailSummary(entry_id="x", store_id="y", subject="hi")
    assert e.subject == "hi"
    assert e.unread is None


def test_package_version() -> None:
    import outlook_readonly_mcp

    assert outlook_readonly_mcp.__version__
