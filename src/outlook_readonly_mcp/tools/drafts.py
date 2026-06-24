"""Draft creation — the ONLY module in this server permitted to write.

What this module may do (and nothing more): create a new `MailItem`, set the
draft fields, and `Save()` it into the Drafts folder.

What this module is *statically forbidden* from doing — enforced by the AST
scan in `tests/test_safety_no_writes.py`, which scans this file against the
full forbidden set minus only `{CreateItem, Save, Subject, Body, HTMLBody,
To, CC, BCC}`:

    Send, Submit, DeliverNow, SaveSentMessageFolder, Reply, ReplyAll,
    ReplyAllEx, Forward, Move, Display, GetInspector, ExecuteMso, SendKeys,
    Execute, Post, SetProperty, SetProperties, Delete, Copy, ...

i.e. every symbol in `safety.SEND_FAMILY` remains forbidden here. The scan
also rejects any `getattr(x, "Send")`-style dynamic lookup of those names.

Why this means the server can never send
----------------------------------------
A draft becomes a sent message only when something calls a send-family method
that moves it Drafts -> Outbox. Outlook's send pipeline drains only the
Outbox; it never auto-sends items in Drafts. Since no send-family symbol can
appear in this (or any) source file, and the only mutating calls available
here are CreateItem + Save (which land the item in Drafts), a draft created
here stays a draft.

Containment: the live COM `MailItem` never leaves this function. Callers
receive a plain `DraftResult` data model, so no caller can invoke `.Send()`
on it either.

Honest limit: once the draft exists it is an ordinary Outlook item. A human
(or other software) can open Outlook and send it. This module guarantees the
*server* cannot send; it cannot make the draft itself unsendable — Outlook
offers no such thing.
"""

from __future__ import annotations

from typing import Optional

from .. import outlook
from ..models import DraftResult
from ..safety import OlDefaultFolders, OlItemType


def _join_recipients(values: Optional[list[str]]) -> str:
    if not values:
        return ""
    cleaned = [v.strip() for v in values if v and v.strip()]
    return "; ".join(cleaned)


def create_draft(
    subject: str = "",
    body: Optional[str] = None,
    html_body: Optional[str] = None,
    to: Optional[list[str]] = None,
    cc: Optional[list[str]] = None,
    bcc: Optional[list[str]] = None,
) -> DraftResult:
    """Create an UNSENT draft email and save it to the Drafts folder.

    Returns a DraftResult with the new item's EntryID. Never sends. Raises if
    the created item does not end up as an unsent item in the Drafts folder.

    Attachments and reply/forward drafts are intentionally unsupported: adding
    attachments needs `Attachments.Add` and replies need `Reply*`, all of
    which stay forbidden. Compose body content directly instead.
    """
    if body is not None and html_body is not None:
        raise ValueError("Provide either body or html_body, not both.")

    to_str = _join_recipients(to)
    cc_str = _join_recipients(cc)
    bcc_str = _join_recipients(bcc)

    with outlook.com_apartment():
        app = outlook.get_application()
        ns = app.GetNamespace("MAPI")
        drafts_folder = ns.GetDefaultFolder(OlDefaultFolders.olFolderDrafts)
        drafts_entry_id = drafts_folder.EntryID

        # Create + compose. The only writes in this whole server live here.
        item = app.CreateItem(OlItemType.olMailItem)
        if subject:
            item.Subject = subject
        if html_body is not None:
            item.HTMLBody = html_body
        elif body is not None:
            item.Body = body
        if to_str:
            item.To = to_str
        if cc_str:
            item.CC = cc_str
        if bcc_str:
            item.BCC = bcc_str

        # Save -> Drafts. This does NOT send (Send would route to the Outbox).
        item.Save()

        # ---- Defensive read-back guards (reads only) -------------------
        # If any of these fail we refuse to report success. We cannot delete
        # the stray item (Delete is forbidden), but we surface the anomaly.
        entry_id = item.EntryID or ""
        parent = item.Parent
        parent_name = ""
        parent_entry_id = ""
        store_id = ""
        if parent is not None:
            parent_name = getattr(parent, "Name", "") or ""
            parent_entry_id = getattr(parent, "EntryID", "") or ""
            store = getattr(parent, "Store", None)
            if store is not None:
                store_id = getattr(store, "StoreID", "") or ""

        # Guard 1: the item must report itself unsent.
        if bool(getattr(item, "Sent", False)):
            raise RuntimeError(
                "Aborting: the created item reports Sent=True, which must never "
                "happen for a draft."
            )

        # Guard 2: the item must be sitting in the Drafts folder, not the Outbox.
        if parent_entry_id != drafts_entry_id:
            raise RuntimeError(
                f"Aborting: draft landed in folder {parent_name!r} rather than "
                "Drafts. The server will not treat this as a successful draft."
            )

        return DraftResult(
            entry_id=entry_id,
            store_id=store_id,
            subject=subject,
            folder=parent_name,
            to=to or [],
            cc=cc or [],
            bcc=bcc or [],
            sent=False,
            created=True,
        )
