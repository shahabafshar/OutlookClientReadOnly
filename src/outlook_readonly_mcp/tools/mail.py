"""Read-only mail tools: accounts, folders, listing, search, get, attachments."""

from __future__ import annotations

import base64
import os
import re
import tempfile
from datetime import datetime
from typing import Optional

from .. import outlook
from ..converters import (
    attachment_info_list,
    email_detail,
    email_summary,
    folder_to_model,
)
from ..models import (
    AccountInfo,
    AttachmentContent,
    AttachmentInfo,
    EmailDetail,
    EmailSummary,
    FolderInfo,
)
from ..safety import OlDefaultFolders


def list_accounts() -> list[AccountInfo]:
    """Return every configured Outlook account/store."""
    with outlook.com_apartment():
        app = outlook.get_application()
        ns = app.GetNamespace("MAPI")
        out: list[AccountInfo] = []
        accounts = ns.Accounts
        for i in range(accounts.Count):
            acc = accounts.Item(i + 1)
            try:
                delivery_store_name = acc.DeliveryStore.DisplayName
            except Exception:
                delivery_store_name = None
            out.append(
                AccountInfo(
                    name=getattr(acc, "DisplayName", "") or "",
                    smtp_address=getattr(acc, "SmtpAddress", None),
                    account_type=getattr(acc, "AccountType", None),
                    delivery_store_name=delivery_store_name,
                )
            )
        return out


def list_folders(max_depth: int = 4) -> list[FolderInfo]:
    """Enumerate every folder across every store up to `max_depth` levels."""
    with outlook.com_apartment():
        results: list[FolderInfo] = []
        for root in outlook.get_store_folders():
            for f, depth in outlook.iter_folders_recursive(root, max_depth=max_depth):
                results.append(folder_to_model(f, depth=depth))
        return results


def _resolve_folder(folder_path: Optional[str], default_id: int):
    if folder_path:
        return outlook.find_folder_by_path(folder_path)
    return outlook.get_default_folder(default_id)


def _escape_dasl_string(s: str) -> str:
    return s.replace("'", "''")


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
) -> list[EmailSummary]:
    """List emails in a folder (Inbox by default) with optional filters.

    All filters are evaluated by Outlook via `Items.Restrict` (DASL) where
    possible, with the remainder applied client-side. No item is modified.
    """
    if limit <= 0 or limit > 1000:
        raise ValueError("limit must be in (0, 1000]")
    if offset < 0:
        raise ValueError("offset must be >= 0")

    with outlook.com_apartment():
        folder = _resolve_folder(folder_path, OlDefaultFolders.olFolderInbox)
        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        clauses: list[str] = []
        if unread_only:
            clauses.append("[UnRead] = True")
        if has_attachment is True:
            clauses.append('"urn:schemas:httpmail:hasattachment" = 1')
        if has_attachment is False:
            clauses.append('"urn:schemas:httpmail:hasattachment" = 0')
        if subject_contains:
            esc = _escape_dasl_string(subject_contains)
            clauses.append(f"\"urn:schemas:httpmail:subject\" ci_phrasematch '{esc}'")
        if sender_contains:
            esc = _escape_dasl_string(sender_contains)
            clauses.append(f"\"urn:schemas:httpmail:fromemail\" ci_phrasematch '{esc}'")
        if received_after:
            clauses.append(
                "\"urn:schemas:httpmail:datereceived\" >= '"
                + received_after.strftime("%m/%d/%Y %H:%M %p") + "'"
            )
        if received_before:
            clauses.append(
                "\"urn:schemas:httpmail:datereceived\" < '"
                + received_before.strftime("%m/%d/%Y %H:%M %p") + "'"
            )

        if clauses:
            try:
                restricted = items.Restrict("@SQL=" + " AND ".join(clauses))
            except Exception:
                restricted = items
        else:
            restricted = items

        try:
            restricted.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        out: list[EmailSummary] = []
        total = restricted.Count
        # Skip `offset` items, take `limit`.
        idx = 0
        taken = 0
        # Outlook Items collection is 1-based.
        for i in range(1, total + 1):
            try:
                item = restricted.Item(i)
            except Exception:
                continue
            mc = getattr(item, "MessageClass", "") or ""
            if not mc.startswith("IPM.Note"):
                continue
            if idx < offset:
                idx += 1
                continue
            out.append(email_summary(item))
            taken += 1
            if taken >= limit:
                break
            idx += 1
        return out


def search_emails(
    query: str,
    folder_path: Optional[str] = None,
    limit: int = 50,
    scope: str = "subject_and_sender",
) -> list[EmailSummary]:
    """Phrase-match search over a folder.

    scope is one of: 'subject_only', 'sender_only', 'subject_and_sender',
    'body', 'all'. Body and all-scope searches are slower as they touch
    every item's body.
    """
    if limit <= 0 or limit > 500:
        raise ValueError("limit must be in (0, 500]")
    if not query.strip():
        raise ValueError("query must not be empty")

    with outlook.com_apartment():
        folder = _resolve_folder(folder_path, OlDefaultFolders.olFolderInbox)
        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        esc = _escape_dasl_string(query)
        clauses: list[str] = []
        if scope in ("subject_only", "subject_and_sender", "all"):
            clauses.append(f"\"urn:schemas:httpmail:subject\" ci_phrasematch '{esc}'")
        if scope in ("sender_only", "subject_and_sender", "all"):
            clauses.append(f"\"urn:schemas:httpmail:fromname\" ci_phrasematch '{esc}'")
            clauses.append(f"\"urn:schemas:httpmail:fromemail\" ci_phrasematch '{esc}'")
        if scope in ("body", "all"):
            clauses.append(
                f"\"urn:schemas:httpmail:textdescription\" ci_phrasematch '{esc}'"
            )
        if not clauses:
            raise ValueError(f"Unknown scope {scope!r}")

        sql = "@SQL=" + " OR ".join(f"({c})" for c in clauses)
        try:
            restricted = items.Restrict(sql)
        except Exception:
            # Fall back to client-side filter
            restricted = items

        try:
            restricted.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        pattern = re.compile(re.escape(query), re.IGNORECASE)
        out: list[EmailSummary] = []
        for i in range(1, restricted.Count + 1):
            try:
                item = restricted.Item(i)
            except Exception:
                continue
            mc = getattr(item, "MessageClass", "") or ""
            if not mc.startswith("IPM.Note"):
                continue
            # If DASL fell back, validate client-side.
            if restricted is items:
                hay = " ".join(
                    str(x or "")
                    for x in (
                        getattr(item, "Subject", ""),
                        getattr(item, "SenderName", ""),
                        getattr(item, "SenderEmailAddress", ""),
                    )
                )
                if scope in ("body", "all"):
                    hay += " " + str(getattr(item, "Body", "") or "")
                if not pattern.search(hay):
                    continue
            out.append(email_summary(item))
            if len(out) >= limit:
                break
        return out


def get_email(entry_id: str, store_id: Optional[str] = None) -> EmailDetail:
    """Fetch full detail for a single email by EntryID."""
    if not entry_id:
        raise ValueError("entry_id is required")
    with outlook.com_apartment():
        item = outlook.get_item_by_entry_id(entry_id, store_id)
        mc = getattr(item, "MessageClass", "") or ""
        if not mc.startswith("IPM.Note"):
            raise ValueError(f"Item with EntryID is not a mail item (class={mc!r})")
        return email_detail(item)


def list_attachments(entry_id: str, store_id: Optional[str] = None) -> list[AttachmentInfo]:
    """Return metadata for every attachment on a mail item without extracting."""
    if not entry_id:
        raise ValueError("entry_id is required")
    with outlook.com_apartment():
        item = outlook.get_item_by_entry_id(entry_id, store_id)
        return attachment_info_list(item)


_SAFE_NAME = re.compile(r"[^A-Za-z0-9._-]+")


def _safe_filename(name: str) -> str:
    base = _SAFE_NAME.sub("_", name).strip("._")
    return base or "attachment.bin"


def get_attachment(
    entry_id: str,
    attachment_index: int,
    save_to_dir: Optional[str] = None,
    return_base64: bool = False,
    max_bytes: int = 25 * 1024 * 1024,
    store_id: Optional[str] = None,
) -> AttachmentContent:
    """Extract one attachment by 1-based index.

    Either saves to `save_to_dir` (returns path) and/or returns its bytes as
    base64. Outlook's `SaveAsFile` only writes the attachment bytes to disk;
    it does not modify the parent mail item. To stay strictly within the
    read-only invariant of this server, we route through a temp file when
    base64 is requested and discard it after reading.
    """
    if attachment_index < 1:
        raise ValueError("attachment_index is 1-based and must be >= 1")
    if not return_base64 and not save_to_dir:
        raise ValueError("Either save_to_dir or return_base64=True is required")

    with outlook.com_apartment():
        item = outlook.get_item_by_entry_id(entry_id, store_id)
        atts = item.Attachments
        if attachment_index > atts.Count:
            raise IndexError(
                f"attachment_index {attachment_index} out of range "
                f"(item has {atts.Count} attachments)"
            )
        att = atts.Item(attachment_index)
        size = int(getattr(att, "Size", 0) or 0)
        if size > max_bytes:
            raise ValueError(
                f"Attachment size {size} exceeds max_bytes={max_bytes}; "
                "raise max_bytes if intentional."
            )

        filename = _safe_filename(getattr(att, "FileName", "") or "attachment.bin")
        content_type = None
        try:
            content_type = att.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x370E001E"
            )
        except Exception:
            pass

        out_path: Optional[str] = None
        b64: Optional[str] = None

        if save_to_dir:
            os.makedirs(save_to_dir, exist_ok=True)
            target = os.path.join(save_to_dir, filename)
            # Outlook's SaveAsFile on Attachment writes the bytes to disk.
            # It does not mutate the mail item itself.
            att.SaveAsFile(target)
            out_path = target

        if return_base64:
            if out_path:
                with open(out_path, "rb") as f:
                    data = f.read()
            else:
                # Round-trip through a temp file to read bytes.
                with tempfile.TemporaryDirectory() as td:
                    temp_target = os.path.join(td, filename)
                    att.SaveAsFile(temp_target)
                    with open(temp_target, "rb") as f:
                        data = f.read()
            b64 = base64.b64encode(data).decode("ascii")

        return AttachmentContent(
            filename=filename,
            size_bytes=size,
            content_type=content_type,
            content_base64=b64,
            saved_to_path=out_path,
        )
