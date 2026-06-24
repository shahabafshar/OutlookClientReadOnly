"""Smoke tests against the live Outlook desktop client.

These are skipped automatically when Outlook is not reachable (e.g. on
non-Windows CI or when the desktop client isn't installed). Run locally
with Outlook running to validate the read paths.
"""

from __future__ import annotations

import os
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


# ---------------------------------------------------------------------------
# Live draft creation — DOUBLE-gated and skipped by default.
#
# This actually creates a draft in your mailbox, and the server CANNOT delete
# it (Delete is forbidden). So it only runs when you explicitly opt in with
# BOTH env vars, and it prints the new EntryID so you can remove it by hand.
# ---------------------------------------------------------------------------

_DRAFT_TEST_ENABLED = (
    os.environ.get("OUTLOOK_MCP_ENABLE_DRAFTS", "").strip().lower()
    in {"1", "true", "yes", "on"}
    and os.environ.get("OUTLOOK_MCP_TEST_CREATE_DRAFT", "").strip().lower()
    in {"1", "true", "yes", "on"}
)

draft_test_required = pytest.mark.skipif(
    not (_outlook_available() and _DRAFT_TEST_ENABLED),
    reason=(
        "Live draft test is opt-in. Set OUTLOOK_MCP_ENABLE_DRAFTS=1 and "
        "OUTLOOK_MCP_TEST_CREATE_DRAFT=1 to run it (creates a real draft you "
        "must delete manually)."
    ),
)


@draft_test_required
def test_create_draft_lands_unsent_in_drafts(capsys) -> None:
    from outlook_readonly_mcp.tools import drafts

    result = drafts.create_draft(
        subject="[outlook-readonly-mcp self-test] safe to delete",
        body="This unsent draft was created by the MCP test suite.",
        to=["nobody@example.invalid"],
    )
    # The guarantee: created, unsent, sitting in Drafts.
    assert result.created is True
    assert result.sent is False
    assert result.entry_id
    assert "draft" in (result.folder or "").lower()

    with capsys.disabled():
        print(
            f"\n[draft self-test] created UNSENT draft EntryID={result.entry_id} "
            f"in folder {result.folder!r}. Delete it manually in Outlook."
        )
