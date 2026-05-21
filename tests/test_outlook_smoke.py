"""Smoke tests against the live Outlook desktop client.

These are skipped automatically when Outlook is not reachable (e.g. on
non-Windows CI or when the desktop client isn't installed). Run locally
with Outlook running to validate the read paths.
"""

from __future__ import annotations

import sys

import pytest


def _outlook_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except Exception:
        return False
    try:
        from outlook_readonly_mcp import outlook

        with outlook.com_apartment():
            outlook.get_application()
        return True
    except Exception:
        return False


outlook_required = pytest.mark.skipif(
    not _outlook_available(), reason="Outlook desktop / pywin32 not available"
)


@outlook_required
def test_list_accounts_returns_at_least_one() -> None:
    from outlook_readonly_mcp.tools import mail

    accounts = mail.list_accounts()
    assert accounts, "Expected at least one configured account"


@outlook_required
def test_list_folders_includes_inbox() -> None:
    from outlook_readonly_mcp.tools import mail

    folders = mail.list_folders(max_depth=3)
    names = {f.name.lower() for f in folders}
    assert "inbox" in names, f"Expected an Inbox in folder list, got: {sorted(names)[:20]}..."


@outlook_required
def test_list_emails_returns_a_list_and_unread_state_unchanged() -> None:
    from outlook_readonly_mcp import outlook as ol
    from outlook_readonly_mcp.safety import OlDefaultFolders
    from outlook_readonly_mcp.tools import mail

    with ol.com_apartment():
        inbox = ol.get_default_folder(OlDefaultFolders.olFolderInbox)
        before_unread = inbox.UnReadItemCount

    emails = mail.list_emails(limit=5)
    assert isinstance(emails, list)

    with ol.com_apartment():
        inbox = ol.get_default_folder(OlDefaultFolders.olFolderInbox)
        after_unread = inbox.UnReadItemCount

    assert before_unread == after_unread, (
        "Listing emails must not change unread count "
        f"(was {before_unread}, now {after_unread})"
    )


@outlook_required
def test_get_email_roundtrip_preserves_unread_flag() -> None:
    from outlook_readonly_mcp.tools import mail

    summaries = mail.list_emails(limit=1)
    if not summaries:
        pytest.skip("Inbox is empty")
    s = summaries[0]
    before_unread = s.unread
    detail = mail.get_email(entry_id=s.entry_id, store_id=s.store_id)
    assert detail.entry_id == s.entry_id
    # Re-read summary to verify state didn't flip.
    again = mail.list_emails(limit=1)
    assert again and again[0].entry_id == s.entry_id
    assert again[0].unread == before_unread, "UnRead flag must not change on read"
