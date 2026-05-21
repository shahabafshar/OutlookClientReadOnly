"""Read-only contact tools."""

from __future__ import annotations

import re
from typing import Optional

from .. import outlook
from ..converters import contact_to_model, folder_to_model
from ..models import Contact, FolderInfo
from ..safety import OlDefaultFolders


def list_contact_folders() -> list[FolderInfo]:
    """Return the default contacts folder for each configured store."""
    with outlook.com_apartment():
        results: list[FolderInfo] = []
        ns = outlook.get_namespace()
        for i in range(ns.Folders.Count):
            store_root = ns.Folders.Item(i + 1)
            try:
                store = store_root.Store
                cf = store.GetDefaultFolder(OlDefaultFolders.olFolderContacts)
            except Exception:
                continue
            results.append(folder_to_model(cf))
        return results


def list_contacts(
    folder_path: Optional[str] = None,
    limit: int = 200,
    offset: int = 0,
) -> list[Contact]:
    """List contacts in a folder (default Contacts) with pagination."""
    if limit <= 0 or limit > 2000:
        raise ValueError("limit must be in (0, 2000]")
    if offset < 0:
        raise ValueError("offset must be >= 0")

    with outlook.com_apartment():
        if folder_path:
            folder = outlook.find_folder_by_path(folder_path)
        else:
            folder = outlook.get_default_folder(OlDefaultFolders.olFolderContacts)
        items = folder.Items
        try:
            items.Sort("[FileAs]")
        except Exception:
            pass

        out: list[Contact] = []
        idx = 0
        for i in range(1, items.Count + 1):
            try:
                item = items.Item(i)
            except Exception:
                continue
            mc = getattr(item, "MessageClass", "") or ""
            if not mc.startswith("IPM.Contact"):
                continue
            if idx < offset:
                idx += 1
                continue
            out.append(contact_to_model(item))
            if len(out) >= limit:
                break
            idx += 1
        return out


def get_contact(entry_id: str, store_id: Optional[str] = None) -> Contact:
    """Fetch a single contact by EntryID."""
    if not entry_id:
        raise ValueError("entry_id is required")
    with outlook.com_apartment():
        item = outlook.get_item_by_entry_id(entry_id, store_id)
        mc = getattr(item, "MessageClass", "") or ""
        if not mc.startswith("IPM.Contact"):
            raise ValueError(f"Item with EntryID is not a contact (class={mc!r})")
        return contact_to_model(item)


def search_contacts(
    query: str,
    folder_path: Optional[str] = None,
    limit: int = 100,
) -> list[Contact]:
    """Phrase-match contacts by full name, file-as, company, or email."""
    if not query.strip():
        raise ValueError("query must not be empty")
    if limit <= 0 or limit > 500:
        raise ValueError("limit must be in (0, 500]")

    with outlook.com_apartment():
        if folder_path:
            folder = outlook.find_folder_by_path(folder_path)
        else:
            folder = outlook.get_default_folder(OlDefaultFolders.olFolderContacts)
        items = folder.Items
        pattern = re.compile(re.escape(query), re.IGNORECASE)

        out: list[Contact] = []
        for i in range(1, items.Count + 1):
            try:
                item = items.Item(i)
            except Exception:
                continue
            mc = getattr(item, "MessageClass", "") or ""
            if not mc.startswith("IPM.Contact"):
                continue
            hay = " ".join(
                str(x or "")
                for x in (
                    getattr(item, "FullName", ""),
                    getattr(item, "FileAs", ""),
                    getattr(item, "CompanyName", ""),
                    getattr(item, "Email1Address", ""),
                    getattr(item, "Email2Address", ""),
                    getattr(item, "Email3Address", ""),
                )
            )
            if pattern.search(hay):
                out.append(contact_to_model(item))
                if len(out) >= limit:
                    break
        return out
