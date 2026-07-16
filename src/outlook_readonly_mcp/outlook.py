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

import os
import shutil
import sys
import tempfile
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


def _looks_like_gen_py_corruption(exc: BaseException) -> bool:
    """Detect the pywin32 gen_py stale-cache symptom.

    When two processes race on `gencache.EnsureDispatch`, one of them can
    end up with a half-written typelib stub module that loads but lacks
    `CLSIDToPackageMap` and other internal tables. The first real method
    invocation then raises `AttributeError: module 'win32com.gen_py.<guid>'
    has no attribute 'CLSIDToPackageMap'`. We also catch the late-bound
    fallback symptom `Outlook.Application.GetNamespace`.
    """
    if not isinstance(exc, AttributeError):
        return False
    msg = str(exc)
    return (
        "CLSIDToPackageMap" in msg
        or "gen_py" in msg
        or "GetNamespace" in msg
    )


def _purge_gen_py_cache() -> None:
    """Nuke every gen_py location pywin32 might be reading from.

    Caches can live in `win32com.__gen_path__`, the per-user temp dir
    (`%TEMP%\\gen_py`), or the package install location. After deleting
    on disk we also drop any already-imported stub modules so the next
    `EnsureDispatch` cannot reuse them.
    """
    paths: list[str] = []
    gen_path = getattr(win32com, "__gen_path__", None)
    if gen_path:
        paths.append(gen_path)
    paths.append(os.path.join(tempfile.gettempdir(), "gen_py"))
    try:
        import win32com.gen_py as _gen  # type: ignore[import-not-found]

        pkg_path = os.path.dirname(getattr(_gen, "__file__", "") or "")
        if pkg_path:
            paths.append(pkg_path)
    except Exception:
        pass

    seen: set[str] = set()
    for p in paths:
        if not p or p in seen:
            continue
        seen.add(p)
        shutil.rmtree(p, ignore_errors=True)

    for mod_name in [m for m in list(sys.modules) if m.startswith("win32com.gen_py")]:
        sys.modules.pop(mod_name, None)


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


def _ensure_dispatch_with_probe() -> Any:
    """`EnsureDispatch` an Outlook.Application and verify it actually works.

    A successful dispatch is not enough — the gen_py typelib stub can be
    half-written, in which case the first real method call (here a probe
    to `GetNamespace`) raises `AttributeError: ... CLSIDToPackageMap`.
    When that happens we purge every gen_py location, drop cached stubs
    from `sys.modules`, and re-dispatch exactly once.
    """
    for attempt in range(2):
        try:
            app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            # Probe a real method so a corrupt stub fails here, not later
            # inside an MCP tool call.
            _ = app.GetNamespace("MAPI")
            return app
        except AttributeError as exc:
            if attempt == 0 and _looks_like_gen_py_corruption(exc):
                _purge_gen_py_cache()
                continue
            raise
    raise OutlookUnavailableError(
        "Outlook COM dispatch failed even after clearing the gen_py cache."
    )


def get_application() -> Any:
    """Return a cached Outlook.Application dispatch for the current thread.

    pywin32 COM objects are apartment-bound, so we cache per-thread. We use
    `gencache.EnsureDispatch` so that the type library is generated on first
    use — without it, late-bound `Dispatch` can't resolve `GetNamespace` and
    other early-bound methods reliably. The dispatch is probed before being
    cached so a corrupt typelib stub is detected and self-healed here, not
    later inside an MCP tool call.
    """
    _require_pywin32()
    app = getattr(_THREAD_LOCAL, "app", None)
    if app is None:
        app = _ensure_dispatch_with_probe()
        _THREAD_LOCAL.app = app
    return app


def get_namespace() -> Any:
    """Return the MAPI namespace for the active Outlook session."""
    app = get_application()
    try:
        return app.GetNamespace("MAPI")
    except AttributeError as exc:
        # If a long-lived cache went stale mid-flight, self-heal once.
        if not _looks_like_gen_py_corruption(exc):
            raise
        _purge_gen_py_cache()
        if hasattr(_THREAD_LOCAL, "app"):
            del _THREAD_LOCAL.app
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


# ---------------------------------------------------------------------------
# Explicit store / folder resolution (no silent fallback — ever)
# ---------------------------------------------------------------------------

class FolderResolutionError(LookupError):
    """The requested account/folder could not be resolved.

    Raised instead of any silent fallback: returning plausible data from the
    wrong mailbox is strictly worse than an error.
    """


# Well-known mail folder names -> olDefaultFolders ids. Matching by *type*
# (GetDefaultFolder) makes these work even when the display names are
# localized; a literal name walk below covers everything else.
MAIL_DEFAULT_FOLDER_ALIASES: dict[str, int] = {
    "inbox": 6,
    "sent items": 5,
    "sent": 5,
    "drafts": 16,
    "deleted items": 3,
    "deleted": 3,
    "trash": 3,
    "outbox": 4,
    "junk email": 23,
    "junk": 23,
    "spam": 23,
}


def _norm(value: Any) -> str:
    return (str(value) if value is not None else "").strip().lower()


def resolve_store(ns: Any, account: str) -> Any:
    """Resolve an account identifier to its Store — by SMTP address, account
    display name, or store display name. Raises FolderResolutionError listing
    the available names when nothing matches. Reads only.
    """
    target = _norm(account)
    if not target:
        raise ValueError("account must be a non-empty string")

    available: list[str] = []
    accounts = ns.Accounts
    for i in range(accounts.Count):
        acc = accounts.Item(i + 1)
        smtp = _norm(getattr(acc, "SmtpAddress", ""))
        disp = _norm(getattr(acc, "DisplayName", ""))
        if smtp:
            available.append(smtp)
        if target in (smtp, disp):
            store = getattr(acc, "DeliveryStore", None)
            if store is not None:
                return store

    stores = ns.Stores
    for i in range(stores.Count):
        st = stores.Item(i + 1)
        name = getattr(st, "DisplayName", "") or ""
        available.append(name)
        if _norm(name) == target:
            return st

    raise FolderResolutionError(
        f"No Outlook account or store matches {account!r}. "
        f"Available: {sorted(set(available))}"
    )


def folder_path_of(folder: Any) -> str:
    """Best-effort canonical path of a folder (e.g. '\\\\store\\Inbox')."""
    try:
        p = folder.FolderPath
        if p:
            return str(p)
    except Exception:
        pass
    return getattr(folder, "Name", "") or ""


def _looks_like_entry_id(value: str) -> bool:
    v = value.strip()
    return len(v) >= 32 and all(c in "0123456789abcdefABCDEF" for c in v)


def _walk_children_ci(parent: Any, segment: str) -> tuple[Any | None, list[str]]:
    """Case-insensitive child lookup. Returns (child_or_None, sibling_names)."""
    target = _norm(segment)
    names: list[str] = []
    subs = parent.Folders
    for j in range(subs.Count):
        f = subs.Item(j + 1)
        name = getattr(f, "Name", "") or ""
        names.append(name)
        if _norm(name) == target:
            return f, names
    return None, names


def resolve_mail_folder(
    folder: str | None = None,
    account: str | None = None,
    default_folder_id: int = 6,
) -> Any:
    """Resolve (folder, account) to a Folder object. NEVER falls back silently.

    Accepted `folder` forms:
      - None                        -> the default folder (Inbox unless
                                       overridden) of `account`'s store, or of
                                       the default store when account is None
      - a folder EntryID            -> ns.GetFolderFromID (with the account's
                                       StoreID when given)
      - '\\\\store\\folder\\sub' or
        'store/folder/sub'          -> explicit store-qualified path walk
      - 'Sent Items' / 'Drafts' ... -> well-known folder resolved by TYPE in
                                       the target store (localization-safe),
                                       else by case-insensitive name among the
                                       store root's children

    Any resolution miss raises FolderResolutionError naming what IS available.
    """
    ns = get_namespace()

    # --- EntryID form ----------------------------------------------------
    if folder and _looks_like_entry_id(folder):
        store_id = None
        if account:
            store_id = getattr(resolve_store(ns, account), "StoreID", None)
        try:
            if store_id:
                return ns.GetFolderFromID(folder.strip(), store_id)
            return ns.GetFolderFromID(folder.strip())
        except Exception as exc:
            raise FolderResolutionError(
                f"Folder EntryID could not be resolved"
                f"{' in account ' + account if account else ''}: {exc}"
            ) from exc

    parts = [p for p in (folder or "").replace("\\", "/").split("/") if p]

    # --- store-qualified path form ----------------------------------------
    if len(parts) >= 2:
        roots = ns.Folders
        current = None
        root_names: list[str] = []
        for i in range(roots.Count):
            f = roots.Item(i + 1)
            name = getattr(f, "Name", "") or ""
            root_names.append(name)
            if _norm(name) == _norm(parts[0]):
                current = f
                break
        if current is not None:
            remaining = parts[1:]
        elif account:
            # Path did not start with a store name; treat the whole path as
            # relative to the requested account's store.
            current = resolve_store(ns, account).GetRootFolder()
            remaining = parts
        else:
            raise FolderResolutionError(
                f"No store named {parts[0]!r}. Available stores: {root_names}"
            )
        for seg in remaining:
            child, names = _walk_children_ci(current, seg)
            if child is None:
                raise FolderResolutionError(
                    f"No subfolder {seg!r} under {folder_path_of(current)!r}. "
                    f"Available: {names}"
                )
            current = child
        return current

    # --- single name (or None) within the target store ---------------------
    store = resolve_store(ns, account) if account else ns.DefaultStore

    if not parts:
        return store.GetDefaultFolder(default_folder_id)

    alias_id = MAIL_DEFAULT_FOLDER_ALIASES.get(_norm(parts[0]))
    if alias_id is not None:
        try:
            return store.GetDefaultFolder(alias_id)
        except Exception:
            pass  # fall through to the literal name walk

    root = store.GetRootFolder()
    child, names = _walk_children_ci(root, parts[0])
    if child is not None:
        return child

    raise FolderResolutionError(
        f"No folder named {parts[0]!r} in store "
        f"{getattr(store, 'DisplayName', '?')!r}. Available: {sorted(names)}"
    )


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
