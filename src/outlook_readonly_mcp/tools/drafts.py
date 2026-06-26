"""Draft management — the ONLY module in this server permitted to write.

Operations (all confined to UNSENT drafts in the Drafts folder):
  - create_draft            : compose and save a new unsent draft
  - update_draft            : edit fields of an existing draft
  - delete_draft            : remove a draft (-> Deleted Items, recoverable)
  - add_draft_attachment    : attach a file to a draft
  - remove_draft_attachment : detach an attachment from a draft

Two layers of safety
--------------------
1. STATIC (can never send). This file is scanned (tests/test_safety_no_writes)
   against the full forbidden set minus only
   `{CreateItem, Save, Delete, Add, Remove, Subject, Body, HTMLBody, To, CC,
   BCC}`. Every symbol in `safety.SEND_FAMILY` — Send, Submit, DeliverNow,
   Reply*, Forward, Move, Display, GetInspector, ExecuteMso, SendKeys, Execute,
   Post, SetProperty* — stays forbidden here too, including `getattr("Send")`
   tricks. A draft therefore can never be sent or moved to the Outbox by this
   server. (Outlook only sends from the Outbox; we have no path there.)

2. RUNTIME (only ever touch drafts). `Delete`/`Add`/`Remove` are generic
   mutators, so the static scan no longer proves we mutate *only* drafts. That
   is enforced by `_resolve_editable_draft`: it resolves the item by EntryID
   and refuses anything that is not an UNSENT item physically located in the
   Drafts folder. Every mutating op below calls it first, so the server can
   never edit or delete a sent message, a received mail, or any non-draft item.

Honest limits
-------------
- A draft, once created, is an ordinary Outlook item. A human can open Outlook
  and send it. This module guarantees the *server* cannot send.
- `delete_draft` moves the item to Deleted Items (recoverable), it does not
  permanently destroy it (`PermanentlyDelete` stays forbidden).
- The live COM item never leaves a function; callers only receive plain models.
"""

from __future__ import annotations

import os
from typing import Any, Optional

from .. import outlook
from ..converters import attachment_info_list
from ..models import AttachmentInfo, DraftDeleteResult, DraftResult
from ..safety import OlDefaultFolders, OlItemType

# Sentinel distinguishing "leave this field unchanged" from "set to empty".
_UNCHANGED = object()


class NotAnEditableDraftError(RuntimeError):
    """Raised when a target item is not an unsent draft in the Drafts folder."""


def _join_recipients(values: Optional[list[str]]) -> str:
    if not values:
        return ""
    return "; ".join(v.strip() for v in values if v and v.strip())


def _drafts_entry_id_for_item(item: Any) -> str:
    """EntryID of the Drafts default folder of the item's own store."""
    parent = item.Parent
    store = getattr(parent, "Store", None) if parent is not None else None
    if store is None:
        return ""
    drafts = store.GetDefaultFolder(OlDefaultFolders.olFolderDrafts)
    return getattr(drafts, "EntryID", "") or ""


def _resolve_editable_draft(entry_id: str, store_id: Optional[str]) -> Any:
    """Resolve an item by EntryID and verify it is an editable draft.

    Refuses (raises NotAnEditableDraftError) unless ALL hold:
      - the item is a mail item (IPM.Note)
      - it reports Sent == False
      - it physically lives in the Drafts default folder of its own store

    This is the runtime guarantee that edit/delete/attachment operations can
    only ever affect unsent drafts — never sent or received mail.
    """
    if not entry_id:
        raise ValueError("entry_id is required")

    item = outlook.get_item_by_entry_id(entry_id, store_id)

    message_class = getattr(item, "MessageClass", "") or ""
    if not message_class.startswith("IPM.Note"):
        raise NotAnEditableDraftError(
            f"Item is not a mail item (MessageClass={message_class!r}); refusing."
        )

    if bool(getattr(item, "Sent", True)):
        raise NotAnEditableDraftError(
            "Item reports Sent=True; only UNSENT drafts may be modified or "
            "deleted. Refusing."
        )

    parent = item.Parent
    parent_entry_id = getattr(parent, "EntryID", "") or "" if parent is not None else ""
    drafts_entry_id = _drafts_entry_id_for_item(item)
    if not drafts_entry_id or parent_entry_id != drafts_entry_id:
        parent_name = getattr(parent, "Name", "") if parent is not None else ""
        raise NotAnEditableDraftError(
            f"Item is in folder {parent_name!r}, not the Drafts folder; refusing "
            "to modify or delete it."
        )

    return item


def _draft_result(item: Any, *, created: bool = False) -> DraftResult:
    """Build a DraftResult from a (verified) draft item — reads only."""
    parent = item.Parent
    parent_name = getattr(parent, "Name", "") if parent is not None else ""
    store_id = ""
    account_name = ""
    if parent is not None:
        store = getattr(parent, "Store", None)
        if store is not None:
            store_id = getattr(store, "StoreID", "") or ""
            account_name = getattr(store, "DisplayName", "") or ""

    # Defensive: never report a draft that somehow looks sent.
    if bool(getattr(item, "Sent", False)):
        raise RuntimeError("Aborting: item reports Sent=True, which must never "
                           "happen for a draft.")

    return DraftResult(
        entry_id=getattr(item, "EntryID", "") or "",
        store_id=store_id,
        account=account_name,
        subject=getattr(item, "Subject", "") or "",
        folder=parent_name,
        to=_split(getattr(item, "To", "")),
        cc=_split(getattr(item, "CC", "")),
        bcc=_split(getattr(item, "BCC", "")),
        attachments=attachment_info_list(item),
        sent=False,
        created=created,
    )


def _split(value: Optional[str]) -> list[str]:
    if not value:
        return []
    return [p.strip() for p in str(value).replace(",", ";").split(";") if p.strip()]


# ---------------------------------------------------------------------------
# Create
# ---------------------------------------------------------------------------

def _resolve_drafts_folder(ns: Any, account: Optional[str]) -> Any:
    """Return the Drafts default folder for the requested account.

    `account` may be an SMTP address or an account/store display name (e.g.
    "safshar@iastate.edu"). None selects the default store's Drafts. Raises if
    no match is found — we never silently fall back to a different mailbox than
    the caller asked for (that is exactly the "wrong account" bug this guards
    against). All operations here are reads; no item is moved.
    """
    if not account or not account.strip():
        return ns.GetDefaultFolder(OlDefaultFolders.olFolderDrafts)
    target = account.strip().lower()

    accounts = ns.Accounts
    for i in range(accounts.Count):
        acc = accounts.Item(i + 1)
        smtp = (getattr(acc, "SmtpAddress", "") or "").lower()
        disp = (getattr(acc, "DisplayName", "") or "").lower()
        if target in (smtp, disp):
            store = getattr(acc, "DeliveryStore", None)
            if store is not None:
                return store.GetDefaultFolder(OlDefaultFolders.olFolderDrafts)

    stores = ns.Stores
    for i in range(stores.Count):
        st = stores.Item(i + 1)
        if (getattr(st, "DisplayName", "") or "").lower() == target:
            return st.GetDefaultFolder(OlDefaultFolders.olFolderDrafts)

    raise ValueError(
        f"No Outlook account or store matches {account!r}; refusing to guess a "
        "mailbox. Use list_accounts to see valid names."
    )


def create_draft(
    subject: str = "",
    body: Optional[str] = None,
    html_body: Optional[str] = None,
    to: Optional[list[str]] = None,
    cc: Optional[list[str]] = None,
    bcc: Optional[list[str]] = None,
    attachments: Optional[list[str]] = None,
    account: Optional[str] = None,
) -> DraftResult:
    """Create an UNSENT draft and save it to a Drafts folder.

    `account` selects which mailbox's Drafts folder to use (SMTP address or
    account display name, e.g. "safshar@iastate.edu"); None uses the default
    store. `attachments` is a list of local file paths. Provide either `body`
    (plain text) or `html_body`, not both. Never sends; the item is created
    directly in the target Drafts folder (via Items.Add — no Move is used),
    and verified unsent and in Drafts before returning.
    """
    if body is not None and html_body is not None:
        raise ValueError("Provide either body or html_body, not both.")

    attachment_paths = _validate_attachment_paths(attachments)
    to_str, cc_str, bcc_str = (
        _join_recipients(to), _join_recipients(cc), _join_recipients(bcc),
    )

    with outlook.com_apartment():
        app = outlook.get_application()
        ns = app.GetNamespace("MAPI")
        drafts_folder = _resolve_drafts_folder(ns, account)
        drafts_entry_id = drafts_folder.EntryID

        # Born directly in the target account's Drafts. Items.Add (permitted
        # in this module) means the draft starts life in the correct mailbox,
        # so no Move (forbidden) is ever needed to relocate it.
        item = drafts_folder.Items.Add("IPM.Note")
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
        for path in attachment_paths:
            item.Attachments.Add(path)

        item.Save()  # -> Drafts. Never sends.

        # Guards (reads only).
        if bool(getattr(item, "Sent", False)):
            raise RuntimeError("Aborting: created item reports Sent=True.")
        parent = item.Parent
        parent_entry_id = getattr(parent, "EntryID", "") if parent is not None else ""
        if parent_entry_id != drafts_entry_id:
            parent_name = getattr(parent, "Name", "") if parent is not None else ""
            raise RuntimeError(
                f"Aborting: draft landed in {parent_name!r}, not Drafts."
            )

        return _draft_result(item, created=True)


# ---------------------------------------------------------------------------
# Update / edit
# ---------------------------------------------------------------------------

def update_draft(
    entry_id: str,
    store_id: Optional[str] = None,
    subject: Any = _UNCHANGED,
    body: Any = _UNCHANGED,
    html_body: Any = _UNCHANGED,
    to: Any = _UNCHANGED,
    cc: Any = _UNCHANGED,
    bcc: Any = _UNCHANGED,
) -> DraftResult:
    """Edit fields of an existing draft. Only provided fields change.

    Pass a value to set it; omit it (or pass the sentinel) to leave it as is.
    For text fields, "" clears them; for to/cc/bcc, [] clears recipients.
    `body` and `html_body` are mutually exclusive in a single call.
    """
    if body is not _UNCHANGED and html_body is not _UNCHANGED:
        raise ValueError("Set either body or html_body in a single call, not both.")

    with outlook.com_apartment():
        item = _resolve_editable_draft(entry_id, store_id)

        if subject is not _UNCHANGED:
            item.Subject = subject or ""
        if html_body is not _UNCHANGED:
            item.HTMLBody = html_body or ""
        if body is not _UNCHANGED:
            item.Body = body or ""
        if to is not _UNCHANGED:
            item.To = _join_recipients(to)
        if cc is not _UNCHANGED:
            item.CC = _join_recipients(cc)
        if bcc is not _UNCHANGED:
            item.BCC = _join_recipients(bcc)

        item.Save()
        return _draft_result(item)


# ---------------------------------------------------------------------------
# Delete
# ---------------------------------------------------------------------------

def delete_draft(entry_id: str, store_id: Optional[str] = None) -> DraftDeleteResult:
    """Delete a draft (moves it to Deleted Items; recoverable).

    Refuses anything that is not an unsent draft in the Drafts folder.
    """
    with outlook.com_apartment():
        item = _resolve_editable_draft(entry_id, store_id)
        subject = getattr(item, "Subject", "") or ""
        item.Delete()
        return DraftDeleteResult(entry_id=entry_id, deleted=True, subject=subject)


# ---------------------------------------------------------------------------
# Attachments
# ---------------------------------------------------------------------------

def add_draft_attachment(
    entry_id: str,
    file_path: str,
    store_id: Optional[str] = None,
) -> DraftResult:
    """Attach a local file to a draft, then save."""
    paths = _validate_attachment_paths([file_path])
    with outlook.com_apartment():
        item = _resolve_editable_draft(entry_id, store_id)
        item.Attachments.Add(paths[0])
        item.Save()
        return _draft_result(item)


def remove_draft_attachment(
    entry_id: str,
    attachment_index: int,
    store_id: Optional[str] = None,
) -> DraftResult:
    """Remove an attachment from a draft by 1-based index, then save."""
    if attachment_index < 1:
        raise ValueError("attachment_index is 1-based and must be >= 1.")
    with outlook.com_apartment():
        item = _resolve_editable_draft(entry_id, store_id)
        count = item.Attachments.Count
        if attachment_index > count:
            raise IndexError(
                f"attachment_index {attachment_index} out of range "
                f"(draft has {count} attachment(s))."
            )
        item.Attachments.Remove(attachment_index)
        item.Save()
        return _draft_result(item)


def list_draft_attachments(
    entry_id: str,
    store_id: Optional[str] = None,
) -> list[AttachmentInfo]:
    """List attachments on a draft (read-only; verifies it is a draft)."""
    with outlook.com_apartment():
        item = _resolve_editable_draft(entry_id, store_id)
        return attachment_info_list(item)


def _validate_attachment_paths(paths: Optional[list[str]]) -> list[str]:
    out: list[str] = []
    for p in paths or []:
        if not p or not p.strip():
            continue
        full = os.path.abspath(p)
        if not os.path.isfile(full):
            raise FileNotFoundError(f"Attachment file not found: {p}")
        out.append(full)
    return out
