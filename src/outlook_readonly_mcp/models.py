"""Pydantic models for read-only Outlook payloads.

These models are pure data carriers; they never reference COM objects and
never contain any mutation methods.
"""

from __future__ import annotations

from datetime import datetime
from typing import Optional

from pydantic import BaseModel, Field


class FolderInfo(BaseModel):
    name: str
    path: str
    entry_id: str
    store_id: str
    item_count: int
    unread_count: int = 0
    default_item_type: Optional[int] = None
    depth: int = 0


class AttachmentInfo(BaseModel):
    index: int
    filename: str
    display_name: Optional[str] = None
    size_bytes: Optional[int] = None
    content_type: Optional[str] = None


class EmailSummary(BaseModel):
    entry_id: str
    store_id: str
    subject: str = ""
    sender_name: Optional[str] = None
    sender_email: Optional[str] = None
    received_time: Optional[datetime] = None
    sent_time: Optional[datetime] = None
    unread: Optional[bool] = None
    has_attachments: bool = False
    size_bytes: Optional[int] = None
    importance: Optional[int] = None
    conversation_topic: Optional[str] = None


class EmailDetail(EmailSummary):
    to_recipients: list[str] = Field(default_factory=list)
    cc_recipients: list[str] = Field(default_factory=list)
    bcc_recipients: list[str] = Field(default_factory=list)
    reply_to: Optional[str] = None
    body_text: Optional[str] = None
    body_html: Optional[str] = None
    headers: Optional[str] = None
    categories: Optional[str] = None
    attachments: list[AttachmentInfo] = Field(default_factory=list)


class AttachmentContent(BaseModel):
    filename: str
    size_bytes: int
    content_type: Optional[str] = None
    content_base64: Optional[str] = None
    saved_to_path: Optional[str] = None


class CalendarEvent(BaseModel):
    entry_id: str
    store_id: str
    subject: str = ""
    organizer: Optional[str] = None
    location: Optional[str] = None
    start: Optional[datetime] = None
    end: Optional[datetime] = None
    all_day: bool = False
    is_recurring: bool = False
    busy_status: Optional[int] = None
    required_attendees: Optional[str] = None
    optional_attendees: Optional[str] = None
    body_text: Optional[str] = None
    categories: Optional[str] = None


class Contact(BaseModel):
    entry_id: str
    store_id: str
    full_name: str = ""
    file_as: Optional[str] = None
    company_name: Optional[str] = None
    job_title: Optional[str] = None
    department: Optional[str] = None
    email1: Optional[str] = None
    email2: Optional[str] = None
    email3: Optional[str] = None
    business_phone: Optional[str] = None
    home_phone: Optional[str] = None
    mobile_phone: Optional[str] = None
    business_address: Optional[str] = None
    home_address: Optional[str] = None
    categories: Optional[str] = None


class TaskItem(BaseModel):
    entry_id: str
    store_id: str
    subject: str = ""
    due_date: Optional[datetime] = None
    start_date: Optional[datetime] = None
    date_completed: Optional[datetime] = None
    complete: bool = False
    percent_complete: int = 0
    status: Optional[int] = None
    importance: Optional[int] = None
    owner: Optional[str] = None
    body_text: Optional[str] = None
    categories: Optional[str] = None


class NoteItem(BaseModel):
    entry_id: str
    store_id: str
    subject: str = ""
    body_text: Optional[str] = None
    last_modification_time: Optional[datetime] = None
    color: Optional[int] = None
    categories: Optional[str] = None


class AccountInfo(BaseModel):
    name: str
    smtp_address: Optional[str] = None
    account_type: Optional[int] = None
    delivery_store_name: Optional[str] = None


class EmailPage(BaseModel):
    """Envelope for list/search results.

    `resolved_folder` echoes the canonical path of the folder that was
    actually queried (e.g. '\\\\zfard@iastate.edu\\Sent Items') so callers can
    verify scoping instead of trusting it blindly.
    """

    resolved_folder: str
    resolved_store: str = ""
    count: int = 0
    emails: list[EmailSummary] = Field(default_factory=list)


class DraftResult(BaseModel):
    """Result of creating an UNSENT draft.

    `sent` is always False — the server verifies the created item reports
    itself unsent and is sitting in the Drafts folder before returning, and
    aborts otherwise. The live COM MailItem is never exposed; only this data
    model leaves the writer module.
    """

    entry_id: str
    store_id: str
    account: str = ""
    subject: str = ""
    folder: str = ""
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    attachments: list[AttachmentInfo] = Field(default_factory=list)
    sent: bool = False
    created: bool = True


class DraftDeleteResult(BaseModel):
    """Result of deleting a draft. The item is moved to Deleted Items (it is
    recoverable there) — it is never permanently destroyed."""

    entry_id: str
    deleted: bool = True
    subject: str = ""
