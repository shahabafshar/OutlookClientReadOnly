"""Read-only calendar tools."""

from __future__ import annotations

from datetime import datetime
from typing import Optional

from .. import outlook
from ..converters import event_to_model
from ..models import CalendarEvent, FolderInfo
from ..safety import OlDefaultFolders


def list_calendar_folders() -> list[FolderInfo]:
    """Return the default calendar folder for each configured store."""
    from ..converters import folder_to_model

    with outlook.com_apartment():
        results: list[FolderInfo] = []
        ns = outlook.get_namespace()
        for i in range(ns.Folders.Count):
            store_root = ns.Folders.Item(i + 1)
            try:
                store = store_root.Store
                cal = store.GetDefaultFolder(OlDefaultFolders.olFolderCalendar)
            except Exception:
                continue
            results.append(folder_to_model(cal))
        return results


def list_events(
    start: datetime,
    end: datetime,
    folder_path: Optional[str] = None,
    include_recurrences: bool = True,
    limit: int = 200,
) -> list[CalendarEvent]:
    """List calendar events overlapping [start, end].

    If `include_recurrences=True`, recurring series are expanded into
    individual occurrences within the window (via `Items.IncludeRecurrences`,
    which is a view setting, not a mutation).
    """
    if end <= start:
        raise ValueError("end must be after start")
    if limit <= 0 or limit > 2000:
        raise ValueError("limit must be in (0, 2000]")

    with outlook.com_apartment():
        if folder_path:
            folder = outlook.find_folder_by_path(folder_path)
        else:
            folder = outlook.get_default_folder(OlDefaultFolders.olFolderCalendar)
        items = folder.Items
        if include_recurrences:
            items.IncludeRecurrences = True
        items.Sort("[Start]")

        # DASL between Start and End
        fmt = "%m/%d/%Y %H:%M %p"
        restriction = (
            f"[Start] <= '{end.strftime(fmt)}' AND [End] >= '{start.strftime(fmt)}'"
        )
        try:
            restricted = items.Restrict(restriction)
        except Exception:
            restricted = items

        out: list[CalendarEvent] = []
        for i in range(1, restricted.Count + 1):
            try:
                item = restricted.Item(i)
            except Exception:
                continue
            mc = getattr(item, "MessageClass", "") or ""
            if not mc.startswith("IPM.Appointment"):
                continue
            out.append(event_to_model(item))
            if len(out) >= limit:
                break
        return out


def get_event(entry_id: str, store_id: Optional[str] = None) -> CalendarEvent:
    """Fetch full detail for a single calendar event by EntryID."""
    if not entry_id:
        raise ValueError("entry_id is required")
    with outlook.com_apartment():
        item = outlook.get_item_by_entry_id(entry_id, store_id)
        mc = getattr(item, "MessageClass", "") or ""
        if not mc.startswith("IPM.Appointment"):
            raise ValueError(f"Item with EntryID is not an appointment (class={mc!r})")
        return event_to_model(item)
