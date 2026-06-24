"""MCP server entry point.

Wires the tool functions into FastMCP and runs over stdio. The read tools are
always registered. The single write tool — `create_draft` — is registered
ONLY when the environment gate `OUTLOOK_MCP_ENABLE_DRAFTS` is set truthy; the
server is otherwise pure read-only. Under no configuration can the server
send, reply to, forward, move, or delete mail: those symbols are statically
forbidden everywhere (see safety.SEND_FAMILY and tests/test_safety_no_writes).
"""

from __future__ import annotations

import os
import sys
from datetime import datetime
from typing import Optional

from mcp.server.fastmcp import FastMCP

from . import __version__
from .tools import calendar as cal_tools
from .tools import contacts as contact_tools
from .tools import mail as mail_tools
from .tools import tasks as task_tools


def _drafts_enabled() -> bool:
    """True when draft creation has been explicitly enabled via env gate."""
    return os.environ.get("OUTLOOK_MCP_ENABLE_DRAFTS", "").strip().lower() in {
        "1", "true", "yes", "on",
    }


DRAFTS_ENABLED = _drafts_enabled()


mcp = FastMCP(
    name="outlook-readonly",
    instructions=(
        "Access to the local Microsoft Outlook desktop client. This server can "
        "list folders, accounts, emails, calendar events, contacts, tasks, and "
        "notes; search them; and read attachments. Reads never mark items as "
        "read or unread. It CANNOT send, reply to, forward, move, or delete "
        "mail under any configuration. When draft creation is enabled "
        "(OUTLOOK_MCP_ENABLE_DRAFTS=1) it can additionally create UNSENT drafts "
        "in the Drafts folder — but even then it cannot send them; only a human "
        "in Outlook can."
    ),
)


# ---- Mail ----------------------------------------------------------------

@mcp.tool()
def list_accounts() -> list[dict]:
    """List every Outlook account/store configured on this machine."""
    return [a.model_dump() for a in mail_tools.list_accounts()]


@mcp.tool()
def list_folders(max_depth: int = 4) -> list[dict]:
    """Enumerate all mail/calendar/etc. folders across every store, up to a depth."""
    return [f.model_dump() for f in mail_tools.list_folders(max_depth=max_depth)]


@mcp.tool()
def list_emails(
    folder_path: Optional[str] = None,
    limit: int = 50,
    offset: int = 0,
    unread_only: bool = False,
    has_attachment: Optional[bool] = None,
    sender_contains: Optional[str] = None,
    subject_contains: Optional[str] = None,
    received_after: Optional[datetime] = None,
    received_before: Optional[datetime] = None,
) -> list[dict]:
    """List emails in a folder (Inbox by default), newest first. All filters optional.

    folder_path: '/'-separated, first segment is the store/account name. Pass
    None for the primary Inbox. Use list_folders to discover paths.
    """
    return [
        e.model_dump()
        for e in mail_tools.list_emails(
            folder_path=folder_path,
            limit=limit,
            offset=offset,
            unread_only=unread_only,
            has_attachment=has_attachment,
            sender_contains=sender_contains,
            subject_contains=subject_contains,
            received_after=received_after,
            received_before=received_before,
        )
    ]


@mcp.tool()
def search_emails(
    query: str,
    folder_path: Optional[str] = None,
    limit: int = 50,
    scope: str = "subject_and_sender",
) -> list[dict]:
    """Phrase-match search.

    scope: 'subject_only' | 'sender_only' | 'subject_and_sender' (default) |
    'body' | 'all'. Body/all searches are slower as they touch every body.
    """
    return [
        e.model_dump()
        for e in mail_tools.search_emails(
            query=query, folder_path=folder_path, limit=limit, scope=scope
        )
    ]


@mcp.tool()
def get_email(entry_id: str, store_id: Optional[str] = None) -> dict:
    """Fetch one email by EntryID, including recipients, body (text + HTML),
    headers, and attachment metadata. Does not modify the item (read flag
    is preserved)."""
    return mail_tools.get_email(entry_id=entry_id, store_id=store_id).model_dump()


@mcp.tool()
def list_attachments(entry_id: str, store_id: Optional[str] = None) -> list[dict]:
    """List attachment metadata for an email without extracting contents."""
    return [
        a.model_dump()
        for a in mail_tools.list_attachments(entry_id=entry_id, store_id=store_id)
    ]


@mcp.tool()
def get_attachment(
    entry_id: str,
    attachment_index: int,
    save_to_dir: Optional[str] = None,
    return_base64: bool = False,
    max_bytes: int = 25 * 1024 * 1024,
    store_id: Optional[str] = None,
) -> dict:
    """Extract one attachment by 1-based index.

    Provide save_to_dir (returns saved_to_path), return_base64=True
    (returns content_base64), or both. Does NOT modify the parent email.
    """
    return mail_tools.get_attachment(
        entry_id=entry_id,
        attachment_index=attachment_index,
        save_to_dir=save_to_dir,
        return_base64=return_base64,
        max_bytes=max_bytes,
        store_id=store_id,
    ).model_dump()


# ---- Calendar ------------------------------------------------------------

@mcp.tool()
def list_calendar_folders() -> list[dict]:
    """List the default calendar folder for each configured Outlook store."""
    return [f.model_dump() for f in cal_tools.list_calendar_folders()]


@mcp.tool()
def list_events(
    start: datetime,
    end: datetime,
    folder_path: Optional[str] = None,
    include_recurrences: bool = True,
    limit: int = 200,
) -> list[dict]:
    """List calendar events overlapping [start, end]. Recurrences expanded
    when include_recurrences=True."""
    return [
        e.model_dump()
        for e in cal_tools.list_events(
            start=start,
            end=end,
            folder_path=folder_path,
            include_recurrences=include_recurrences,
            limit=limit,
        )
    ]


@mcp.tool()
def get_event(entry_id: str, store_id: Optional[str] = None) -> dict:
    """Fetch one calendar event by EntryID."""
    return cal_tools.get_event(entry_id=entry_id, store_id=store_id).model_dump()


# ---- Contacts ------------------------------------------------------------

@mcp.tool()
def list_contact_folders() -> list[dict]:
    """List the default contacts folder for each configured Outlook store."""
    return [f.model_dump() for f in contact_tools.list_contact_folders()]


@mcp.tool()
def list_contacts(
    folder_path: Optional[str] = None,
    limit: int = 200,
    offset: int = 0,
) -> list[dict]:
    """List contacts in the default (or specified) contacts folder."""
    return [
        c.model_dump()
        for c in contact_tools.list_contacts(
            folder_path=folder_path, limit=limit, offset=offset
        )
    ]


@mcp.tool()
def search_contacts(
    query: str,
    folder_path: Optional[str] = None,
    limit: int = 100,
) -> list[dict]:
    """Phrase-match contacts by name, file-as, company, or email."""
    return [
        c.model_dump()
        for c in contact_tools.search_contacts(
            query=query, folder_path=folder_path, limit=limit
        )
    ]


@mcp.tool()
def get_contact(entry_id: str, store_id: Optional[str] = None) -> dict:
    """Fetch one contact by EntryID."""
    return contact_tools.get_contact(entry_id=entry_id, store_id=store_id).model_dump()


# ---- Tasks & Notes -------------------------------------------------------

@mcp.tool()
def list_tasks(
    folder_path: Optional[str] = None,
    include_completed: bool = False,
    limit: int = 200,
    offset: int = 0,
) -> list[dict]:
    """List tasks in the default (or specified) tasks folder."""
    return [
        t.model_dump()
        for t in task_tools.list_tasks(
            folder_path=folder_path,
            include_completed=include_completed,
            limit=limit,
            offset=offset,
        )
    ]


@mcp.tool()
def get_task(entry_id: str, store_id: Optional[str] = None) -> dict:
    """Fetch one task by EntryID."""
    return task_tools.get_task(entry_id=entry_id, store_id=store_id).model_dump()


@mcp.tool()
def list_notes(
    folder_path: Optional[str] = None,
    limit: int = 200,
    offset: int = 0,
) -> list[dict]:
    """List sticky notes in the default (or specified) notes folder."""
    return [
        n.model_dump()
        for n in task_tools.list_notes(
            folder_path=folder_path, limit=limit, offset=offset
        )
    ]


@mcp.tool()
def get_note(entry_id: str, store_id: Optional[str] = None) -> dict:
    """Fetch one sticky note by EntryID."""
    return task_tools.get_note(entry_id=entry_id, store_id=store_id).model_dump()


# ---- Drafts (write) — registered ONLY when the env gate is set -----------

if DRAFTS_ENABLED:
    from .tools import drafts as draft_tools

    @mcp.tool()
    def create_draft(
        subject: str = "",
        body: Optional[str] = None,
        html_body: Optional[str] = None,
        to: Optional[list[str]] = None,
        cc: Optional[list[str]] = None,
        bcc: Optional[list[str]] = None,
    ) -> dict:
        """Create an UNSENT draft email in the Drafts folder.

        IMPORTANT: This server can NEVER send the draft — sending is statically
        impossible in its source. Only a human opening Outlook can send it.
        The draft is saved to Drafts and left there. Provide either `body`
        (plain text) or `html_body`, not both. `to`/`cc`/`bcc` are lists of
        addresses or "Name <addr>" strings. Attachments and reply/forward
        drafts are not supported.
        """
        return draft_tools.create_draft(
            subject=subject,
            body=body,
            html_body=html_body,
            to=to,
            cc=cc,
            bcc=bcc,
        ).model_dump()


# ---- Server info ---------------------------------------------------------

# Operations that are forbidden under EVERY configuration — these never become
# available, drafts gate or not. Mirrors safety.SEND_FAMILY in spirit.
_ALWAYS_FORBIDDEN = [
    "send", "submit", "reply", "reply_all", "forward", "deliver_now",
    "delete", "move", "copy", "mark_as_read", "mark_as_unread",
    "display", "print", "send_draft",
]


@mcp.tool()
def server_info() -> dict:
    """Return server name, version, and capability summary."""
    capabilities = [
        "mail.list_accounts",
        "mail.list_folders",
        "mail.list_emails",
        "mail.search_emails",
        "mail.get_email",
        "mail.list_attachments",
        "mail.get_attachment",
        "calendar.list_calendar_folders",
        "calendar.list_events",
        "calendar.get_event",
        "contacts.list_contact_folders",
        "contacts.list_contacts",
        "contacts.search_contacts",
        "contacts.get_contact",
        "tasks.list_tasks",
        "tasks.get_task",
        "notes.list_notes",
        "notes.get_note",
    ]
    if DRAFTS_ENABLED:
        capabilities.append("drafts.create_draft")

    return {
        "name": "outlook-readonly",
        "version": __version__,
        # Pure read-only unless drafts have been explicitly enabled.
        "read_only": not DRAFTS_ENABLED,
        "drafts_enabled": DRAFTS_ENABLED,
        # The server cannot send under ANY configuration — this is invariant.
        "can_send": False,
        "capabilities": capabilities,
        "forbidden_operations": _ALWAYS_FORBIDDEN,
        "notes": (
            "Draft creation is "
            + ("ENABLED" if DRAFTS_ENABLED else "disabled")
            + ". Even when enabled, the server can only save UNSENT drafts to "
            "the Drafts folder; it can never send, and only a human in Outlook "
            "can send a draft."
        ),
    }


def main() -> int:
    """Entry point installed as `outlook-readonly-mcp`."""
    try:
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
