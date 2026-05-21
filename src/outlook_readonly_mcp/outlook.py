"""Read-only connection to the local Microsoft Outlook desktop client.

This module wraps `pywin32` COM access to Outlook. By construction it only
exposes read paths:

- Connect to a running (or auto-started) Outlook.Application via Dispatch.
- Resolve the MAPI Namespace and traverse Folders/Items by enumeration.
- Look up specific items by their stable EntryID via GetItemFromID.

No code path in this module calls any mutating method on a COM object
(see `safety.FORBIDDEN_METHODS`). The companion test
`tests/test_safety_no_writes.py` enforces this with an AST scan.
"""

from __future__ import annotations

import threading
from contextlib import contextmanager
from typing import Any, Iterator

try:
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
    _HAS_PYWIN32 = True
except ImportError:  # pragma: no cover - only on non-Windows
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    _HAS_PYWIN32 = False


_THREAD_LOCAL = threading.local()


class OutlookUnavailableError(RuntimeError):
    """Raised when Outlook is not installed or COM is unreachable."""


def _require_pywin32() -> None:
    if not _HAS_PYWIN32:
        raise OutlookUnavailableError(
            "pywin32 is not installed. This MCP server runs only on Windows "
            "with Outlook desktop and pywin32 installed."
        )


@contextmanager
def com_apartment() -> Iterator[None]:
    """Initialize a COM apartment for the current thread for the duration.

    Clears any cached COM dispatch on entry and exit, because COM proxies
    are apartment-bound: a cached Application from a previous apartment is
    a dead proxy once `CoUninitialize` has run.
    """
    _require_pywin32()
    if hasattr(_THREAD_LOCAL, "app"):
        del _THREAD_LOCAL.app
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        if hasattr(_THREAD_LOCAL, "app"):
            del _THREAD_LOCAL.app
        pythoncom.CoUninitialize()


def get_application() -> Any:
    """Return a cached Outlook.Application dispatch for the current thread.

    pywin32 COM objects are apartment-bound, so we cache per-thread. We use
    `gencache.EnsureDispatch` so that the type library is generated on first
    use — without it, late-bound `Dispatch` can't resolve `GetNamespace` and
    other early-bound methods reliably.
    """
    _require_pywin32()
    app = getattr(_THREAD_LOCAL, "app", None)
    if app is None:
        try:
            app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        except AttributeError:
            # Stale generated cache — rebuild and retry once.
            import shutil

            gen_path = getattr(win32com, "__gen_path__", None)
            if gen_path:
                shutil.rmtree(gen_path, ignore_errors=True)
            app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        _THREAD_LOCAL.app = app
    return app


def get_namespace() -> Any:
    """Return the MAPI namespace for the active Outlook session."""
    app = get_application()
    return app.GetNamespace("MAPI")


def get_default_folder(folder_id: int) -> Any:
    """Return the default folder for the primary store by olDefaultFolders id."""
    ns = get_namespace()
    return ns.GetDefaultFolder(folder_id)


def get_store_folders() -> list[Any]:
    """Return the top-level folders of every configured Outlook store/account."""
    ns = get_namespace()
    stores = ns.Folders
    return [stores.Item(i + 1) for i in range(stores.Count)]


def iter_folders_recursive(folder: Any, max_depth: int = 6) -> Iterator[tuple[Any, int]]:
    """Yield (folder, depth) for `folder` and every descendant up to max_depth.

    Pure traversal — no folder mutations are performed.
    """
    yield folder, 0
    if max_depth <= 0:
        return
    subs = folder.Folders
    for i in range(subs.Count):
        child = subs.Item(i + 1)
        yield from _iter_with_depth(child, depth=1, max_depth=max_depth)


def _iter_with_depth(folder: Any, depth: int, max_depth: int) -> Iterator[tuple[Any, int]]:
    yield folder, depth
    if depth >= max_depth:
        return
    subs = folder.Folders
    for i in range(subs.Count):
        child = subs.Item(i + 1)
        yield from _iter_with_depth(child, depth=depth + 1, max_depth=max_depth)


def find_folder_by_path(path: str) -> Any:
    """Resolve a folder by '/'-separated path, e.g. 'me@example.com/Inbox/Sub'.

    The first segment must match a store/top-level folder name.
    """
    parts = [p for p in path.replace("\\", "/").split("/") if p]
    if not parts:
        raise ValueError("Empty folder path")

    ns = get_namespace()
    current = None
    roots = ns.Folders
    for i in range(roots.Count):
        f = roots.Item(i + 1)
        if f.Name == parts[0]:
            current = f
            break
    if current is None:
        raise LookupError(f"No top-level folder named {parts[0]!r}")

    for seg in parts[1:]:
        sub_found = None
        subs = current.Folders
        for j in range(subs.Count):
            f = subs.Item(j + 1)
            if f.Name == seg:
                sub_found = f
                break
        if sub_found is None:
            raise LookupError(f"No subfolder {seg!r} under {current.Name!r}")
        current = sub_found
    return current


def get_item_by_entry_id(entry_id: str, store_id: str | None = None) -> Any:
    """Resolve any Outlook item by its EntryID.

    EntryIDs are stable identifiers for items in a given store. This is the
    only item-lookup path the server exposes; no item is ever created.
    """
    ns = get_namespace()
    if store_id:
        return ns.GetItemFromID(entry_id, store_id)
    return ns.GetItemFromID(entry_id)


def iter_items_sorted_desc(folder: Any, by: str = "ReceivedTime") -> Iterator[Any]:
    """Yield items in a folder sorted descending by a property.

    Uses `Items.Sort` which only reorders the in-memory view; it does not
    modify items on disk.
    """
    items = folder.Items
    try:
        items.Sort(f"[{by}]", True)
    except Exception:
        # Some folder item types (e.g. contacts) don't support ReceivedTime.
        pass
    count = items.Count
    for i in range(count):
        yield items.Item(i + 1)


def restrict_items(folder: Any, dasl_or_filter: str) -> Any:
    """Return an Items collection restricted by an Outlook filter expression.

    This is a pure query — `Items.Restrict` returns a filtered view without
    altering the underlying items.
    """
    return folder.Items.Restrict(dasl_or_filter)
