"""Pure conversion functions from Outlook COM items to read-only Pydantic models.

Every function here only reads properties. No assignments, no method calls
on COM items that mutate state. The test
`tests/test_safety_no_writes.py` verifies this by AST scan.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any, Optional

from .models import (
    AttachmentInfo,
    CalendarEvent,
    Contact,
    EmailDetail,
    EmailSummary,
    FolderInfo,
    NoteItem,
    TaskItem,
)


def _safe(getter, default=None):
    """Read a COM property and swallow any access error."""
    try:
        return getter()
    except Exception:
        return default


def _pydt(value: Any) -> Optional[datetime]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    try:
        # pywin32 returns pywintypes.datetime which is a datetime subclass
        return datetime.fromisoformat(str(value))
    except Exception:
        return None


def _store_id_of(item: Any) -> str:
    parent = _safe(lambda: item.Parent)
    if parent is not None:
        store = _safe(lambda: parent.Store)
        if store is not None:
            sid = _safe(lambda: store.StoreID)
            if sid:
                return sid
    return ""


def folder_to_model(folder: Any, depth: int = 0) -> FolderInfo:
    path_parts: list[str] = []
    cursor = folder
    # Walk up using FolderPath if available, otherwise reconstruct.
    folder_path = _safe(lambda: folder.FolderPath, "")
    if not folder_path:
        # fall back to walking Parent chain
        seen = 0
        while cursor is not None and seen < 32:
            name = _safe(lambda c=cursor: c.Name)
            if not name:
                break
            path_parts.insert(0, name)
            parent = _safe(lambda c=cursor: c.Parent)
            if parent is None or _safe(lambda p=parent: p.Name) is None:
                break
            cursor = parent
            seen += 1
        folder_path = "/".join(path_parts)

    store = _safe(lambda: folder.Store)
    return FolderInfo(
        name=_safe(lambda: folder.Name) or "",
        path=folder_path,
        entry_id=_safe(lambda: folder.EntryID) or "",
        store_id=_safe(lambda: store.StoreID) if store else "",
        item_count=_safe(lambda: folder.Items.Count) or 0,
        unread_count=_safe(lambda: folder.UnReadItemCount) or 0,
        default_item_type=_safe(lambda: folder.DefaultItemType),
        depth=depth,
    )


def _recipient_addresses(item: Any, kind: int) -> list[str]:
    """kind: 1=To, 2=CC, 3=BCC for MailItem; matches OlMailRecipientType."""
    out: list[str] = []
    recipients = _safe(lambda: item.Recipients)
    if recipients is None:
        return out
    count = _safe(lambda: recipients.Count) or 0
    for i in range(count):
        r = _safe(lambda i=i: recipients.Item(i + 1))
        if r is None:
            continue
        if _safe(lambda r=r: r.Type) != kind:
            continue
        name = _safe(lambda r=r: r.Name) or ""
        addr = _safe(lambda r=r: r.Address) or ""
        # Prefer SMTP via AddressEntry if Address is an EX/legacyDN
        if addr and not _is_smtp(addr):
            ae = _safe(lambda r=r: r.AddressEntry)
            if ae is not None:
                smtp = _safe(lambda ae=ae: ae.GetExchangeUser().PrimarySmtpAddress)
                if smtp:
                    addr = smtp
        out.append(f"{name} <{addr}>" if name and addr else (addr or name))
    return out


def _is_smtp(s: str) -> bool:
    return "@" in s and "/" not in s


def email_summary(item: Any) -> EmailSummary:
    sender_email = _safe(lambda: item.SenderEmailAddress)
    if sender_email and not _is_smtp(sender_email):
        sender = _safe(lambda: item.Sender)
        if sender is not None:
            smtp = _safe(lambda: sender.GetExchangeUser().PrimarySmtpAddress)
            if smtp:
                sender_email = smtp

    return EmailSummary(
        entry_id=_safe(lambda: item.EntryID) or "",
        store_id=_store_id_of(item),
        subject=_safe(lambda: item.Subject) or "",
        sender_name=_safe(lambda: item.SenderName),
        sender_email=sender_email,
        received_time=_pydt(_safe(lambda: item.ReceivedTime)),
        sent_time=_pydt(_safe(lambda: item.SentOn)),
        unread=_safe(lambda: bool(item.UnRead)),
        has_attachments=bool(_safe(lambda: item.Attachments.Count) or 0),
        size_bytes=_safe(lambda: item.Size),
        importance=_safe(lambda: item.Importance),
        conversation_topic=_safe(lambda: item.ConversationTopic),
    )


def attachment_info_list(item: Any) -> list[AttachmentInfo]:
    out: list[AttachmentInfo] = []
    atts = _safe(lambda: item.Attachments)
    if atts is None:
        return out
    count = _safe(lambda: atts.Count) or 0
    for i in range(count):
        a = _safe(lambda i=i: atts.Item(i + 1))
        if a is None:
            continue
        out.append(
            AttachmentInfo(
                index=i + 1,
                filename=_safe(lambda a=a: a.FileName) or "",
                display_name=_safe(lambda a=a: a.DisplayName),
                size_bytes=_safe(lambda a=a: a.Size),
                content_type=_safe(lambda a=a: a.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x370E001E"
                )),
            )
        )
    return out


def email_detail(item: Any) -> EmailDetail:
    base = email_summary(item).model_dump()
    headers = _safe(lambda: item.PropertyAccessor.GetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    ))
    return EmailDetail(
        **base,
        to_recipients=_recipient_addresses(item, 1),
        cc_recipients=_recipient_addresses(item, 2),
        bcc_recipients=_recipient_addresses(item, 3),
        reply_to=_safe(lambda: item.ReplyRecipientNames),
        body_text=_safe(lambda: item.Body),
        body_html=_safe(lambda: item.HTMLBody),
        headers=headers,
        categories=_safe(lambda: item.Categories),
        attachments=attachment_info_list(item),
    )


def event_to_model(item: Any) -> CalendarEvent:
    return CalendarEvent(
        entry_id=_safe(lambda: item.EntryID) or "",
        store_id=_store_id_of(item),
        subject=_safe(lambda: item.Subject) or "",
        organizer=_safe(lambda: item.Organizer),
        location=_safe(lambda: item.Location),
        start=_pydt(_safe(lambda: item.Start)),
        end=_pydt(_safe(lambda: item.End)),
        all_day=bool(_safe(lambda: item.AllDayEvent) or False),
        is_recurring=bool(_safe(lambda: item.IsRecurring) or False),
        busy_status=_safe(lambda: item.BusyStatus),
        required_attendees=_safe(lambda: item.RequiredAttendees),
        optional_attendees=_safe(lambda: item.OptionalAttendees),
        body_text=_safe(lambda: item.Body),
        categories=_safe(lambda: item.Categories),
    )


def contact_to_model(item: Any) -> Contact:
    return Contact(
        entry_id=_safe(lambda: item.EntryID) or "",
        store_id=_store_id_of(item),
        full_name=_safe(lambda: item.FullName) or "",
        file_as=_safe(lambda: item.FileAs),
        company_name=_safe(lambda: item.CompanyName),
        job_title=_safe(lambda: item.JobTitle),
        department=_safe(lambda: item.Department),
        email1=_safe(lambda: item.Email1Address),
        email2=_safe(lambda: item.Email2Address),
        email3=_safe(lambda: item.Email3Address),
        business_phone=_safe(lambda: item.BusinessTelephoneNumber),
        home_phone=_safe(lambda: item.HomeTelephoneNumber),
        mobile_phone=_safe(lambda: item.MobileTelephoneNumber),
        business_address=_safe(lambda: item.BusinessAddress),
        home_address=_safe(lambda: item.HomeAddress),
        categories=_safe(lambda: item.Categories),
    )


def task_to_model(item: Any) -> TaskItem:
    return TaskItem(
        entry_id=_safe(lambda: item.EntryID) or "",
        store_id=_store_id_of(item),
        subject=_safe(lambda: item.Subject) or "",
        due_date=_pydt(_safe(lambda: item.DueDate)),
        start_date=_pydt(_safe(lambda: item.StartDate)),
        date_completed=_pydt(_safe(lambda: item.DateCompleted)),
        complete=bool(_safe(lambda: item.Complete) or False),
        percent_complete=int(_safe(lambda: item.PercentComplete) or 0),
        status=_safe(lambda: item.Status),
        importance=_safe(lambda: item.Importance),
        owner=_safe(lambda: item.Owner),
        body_text=_safe(lambda: item.Body),
        categories=_safe(lambda: item.Categories),
    )


def note_to_model(item: Any) -> NoteItem:
    return NoteItem(
        entry_id=_safe(lambda: item.EntryID) or "",
        store_id=_store_id_of(item),
        subject=_safe(lambda: item.Subject) or "",
        body_text=_safe(lambda: item.Body),
        last_modification_time=_pydt(_safe(lambda: item.LastModificationTime)),
        color=_safe(lambda: item.Color),
        categories=_safe(lambda: item.Categories),
    )
