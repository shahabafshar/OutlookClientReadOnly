"""Read-only task and note tools."""

from __future__ import annotations

from typing import Optional

from .. import outlook
from ..converters import note_to_model, task_to_model
from ..models import NoteItem, TaskItem
from ..safety import OlDefaultFolders


def list_tasks(
    folder_path: Optional[str] = None,
    include_completed: bool = False,
    limit: int = 200,
    offset: int = 0,
) -> list[TaskItem]:
    """List tasks (default Tasks folder)."""
    if limit <= 0 or limit > 2000:
        raise ValueError("limit must be in (0, 2000]")
    if offset < 0:
        raise ValueError("offset must be >= 0")

    with outlook.com_apartment():
        if folder_path:
            folder = outlook.find_folder_by_path(folder_path)
        else:
            folder = outlook.get_default_folder(OlDefaultFolders.olFolderTasks)
        items = folder.Items
        try:
            items.Sort("[DueDate]")
        except Exception:
            pass

        out: list[TaskItem] = []
        idx = 0
        for i in range(1, items.Count + 1):
            try:
                item = items.Item(i)
            except Exception:
                continue
            mc = getattr(item, "MessageClass", "") or ""
            if not mc.startswith("IPM.Task"):
                continue
            if not include_completed and bool(getattr(item, "Complete", False)):
                continue
            if idx < offset:
                idx += 1
                continue
            out.append(task_to_model(item))
            if len(out) >= limit:
                break
            idx += 1
        return out


def get_task(entry_id: str, store_id: Optional[str] = None) -> TaskItem:
    """Fetch a single task by EntryID."""
    if not entry_id:
        raise ValueError("entry_id is required")
    with outlook.com_apartment():
        item = outlook.get_item_by_entry_id(entry_id, store_id)
        mc = getattr(item, "MessageClass", "") or ""
        if not mc.startswith("IPM.Task"):
            raise ValueError(f"Item with EntryID is not a task (class={mc!r})")
        return task_to_model(item)


def list_notes(
    folder_path: Optional[str] = None,
    limit: int = 200,
    offset: int = 0,
) -> list[NoteItem]:
    """List sticky notes (default Notes folder)."""
    if limit <= 0 or limit > 2000:
        raise ValueError("limit must be in (0, 2000]")
    if offset < 0:
        raise ValueError("offset must be >= 0")

    with outlook.com_apartment():
        if folder_path:
            folder = outlook.find_folder_by_path(folder_path)
        else:
            folder = outlook.get_default_folder(OlDefaultFolders.olFolderNotes)
        items = folder.Items
        try:
            items.Sort("[LastModificationTime]", True)
        except Exception:
            pass

        out: list[NoteItem] = []
        idx = 0
        for i in range(1, items.Count + 1):
            try:
                item = items.Item(i)
            except Exception:
                continue
            mc = getattr(item, "MessageClass", "") or ""
            if not mc.startswith("IPM.StickyNote"):
                continue
            if idx < offset:
                idx += 1
                continue
            out.append(note_to_model(item))
            if len(out) >= limit:
                break
            idx += 1
        return out


def get_note(entry_id: str, store_id: Optional[str] = None) -> NoteItem:
    """Fetch a single note by EntryID."""
    if not entry_id:
        raise ValueError("entry_id is required")
    with outlook.com_apartment():
        item = outlook.get_item_by_entry_id(entry_id, store_id)
        mc = getattr(item, "MessageClass", "") or ""
        if not mc.startswith("IPM.StickyNote"):
            raise ValueError(f"Item with EntryID is not a note (class={mc!r})")
        return note_to_model(item)
