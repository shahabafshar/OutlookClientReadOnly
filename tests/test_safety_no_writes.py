"""CI-enforced safety test: scan source for any mutating / sending COM call.

This is the load-bearing guarantee of the project. It parses every Python
file under `src/outlook_readonly_mcp/` (except `safety.py`, which legitimately
declares the forbidden names as data) and looks for, per file:

1. Call expressions `<expr>.NAME(...)` where NAME is forbidden for that file.
2. Assignment targets `<expr>.NAME = ...` where NAME is forbidden for that file.
3. `getattr(<expr>, "NAME")(...)` immediate dynamic calls of a forbidden name.
4. `getattr(<expr>, "NAME")` literals of any `SEND_FAMILY` name (called or not).

The per-file forbidden sets come from `safety.forbidden_calls_for()` /
`safety.forbidden_properties_for()`. Every file is scanned against the full
forbidden set EXCEPT the single draft writer `tools/drafts.py`, which is
permitted only `{CreateItem, Save, Subject, Body, HTMLBody, To, CC, BCC}`.

Critically, the send family (`safety.SEND_FAMILY`: Send, Submit, Reply*,
Forward, Move, Display, DeliverNow, GetInspector, ExecuteMso, SendKeys,
Execute, Post, SetProperty*) is forbidden in EVERY file, drafts.py included.
So the scan statically proves no source file can transition a draft to a sent
state — the "can never send" guarantee.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

from outlook_readonly_mcp.safety import (
    DRAFT_ALLOWED,
    DRAFT_ALLOWED_METHODS,
    DRAFT_MODULE_BASENAME,
    FORBIDDEN_APPLICATION_METHODS,
    FORBIDDEN_METHODS,
    FORBIDDEN_PROPERTIES,
    SEND_FAMILY,
    forbidden_calls_for,
    forbidden_properties_for,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src" / "outlook_readonly_mcp"

# safety.py legitimately holds the forbidden names as string data.
EXCLUDED_FILES = {SRC_ROOT / "safety.py"}


def _iter_source_files() -> list[Path]:
    return [p for p in SRC_ROOT.rglob("*.py") if p not in EXCLUDED_FILES]


class _Scanner(ast.NodeVisitor):
    """AST scanner parameterised by the effective forbidden sets for a file."""

    def __init__(
        self,
        forbidden_calls: frozenset[str],
        forbidden_props: frozenset[str],
        send_family: frozenset[str] = SEND_FAMILY,
    ) -> None:
        self.forbidden_calls = forbidden_calls
        self.forbidden_props = forbidden_props
        self.send_family = send_family
        self.violations: list[tuple[int, str, str]] = []

    # --- direct calls and dynamic getattr calls --------------------------
    def visit_Call(self, node: ast.Call) -> None:  # noqa: N802
        func = node.func

        # `<expr>.NAME(...)`
        if isinstance(func, ast.Attribute) and func.attr in self.forbidden_calls:
            self.violations.append((node.lineno, "call", func.attr))

        # `getattr(<expr>, "NAME")(...)` — immediate dynamic call
        if isinstance(func, ast.Call):
            name = _getattr_literal(func)
            if name is not None and name in self.forbidden_calls:
                self.violations.append((node.lineno, "getattr-call", name))

        # `getattr(<expr>, "NAME")` of a SEND_FAMILY name, called or not.
        gname = _getattr_literal(node)
        if gname is not None and gname in self.send_family:
            self.violations.append((node.lineno, "getattr-name", gname))

        self.generic_visit(node)

    # --- assignments -----------------------------------------------------
    def visit_Assign(self, node: ast.Assign) -> None:  # noqa: N802
        for target in node.targets:
            self._check_assign_target(target, node.lineno)
        self.generic_visit(node)

    def visit_AugAssign(self, node: ast.AugAssign) -> None:  # noqa: N802
        self._check_assign_target(node.target, node.lineno)
        self.generic_visit(node)

    def visit_AnnAssign(self, node: ast.AnnAssign) -> None:  # noqa: N802
        if node.value is not None:
            self._check_assign_target(node.target, node.lineno)
        self.generic_visit(node)

    def _check_assign_target(self, target: ast.expr, lineno: int) -> None:
        if isinstance(target, ast.Attribute) and target.attr in self.forbidden_props:
            self.violations.append((lineno, "assign", target.attr))


def _getattr_literal(node: ast.AST) -> str | None:
    """If `node` is `getattr(x, "NAME")` return NAME, else None."""
    if (
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Name)
        and node.func.id == "getattr"
        and len(node.args) >= 2
        and isinstance(node.args[1], ast.Constant)
        and isinstance(node.args[1].value, str)
    ):
        return node.args[1].value
    return None


def _scan(path: Path) -> list[tuple[int, str, str]]:
    scanner = _Scanner(
        forbidden_calls=forbidden_calls_for(path.name),
        forbidden_props=forbidden_properties_for(path.name),
    )
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    scanner.visit(tree)
    return scanner.violations


# ---------------------------------------------------------------------------
# Core scan
# ---------------------------------------------------------------------------

def test_no_forbidden_symbols_in_source() -> None:
    """Every file is free of the COM symbols forbidden for that file."""
    all_violations: list[str] = []
    files_scanned = 0
    for path in _iter_source_files():
        files_scanned += 1
        for lineno, kind, name in _scan(path):
            rel = path.relative_to(REPO_ROOT)
            all_violations.append(f"{rel}:{lineno}: forbidden {kind} `.{name}`")

    assert files_scanned > 0, "No source files were scanned — test setup is broken."
    assert not all_violations, (
        "Safety invariant broken — forbidden COM symbols appear in source:\n"
        + "\n".join(all_violations)
    )


# ---------------------------------------------------------------------------
# The "can never send" guarantee
# ---------------------------------------------------------------------------

def test_send_family_never_appears_anywhere() -> None:
    """No SEND_FAMILY symbol may appear in ANY file — drafts.py included.

    This is the explicit, human-readable form of the can-never-send guarantee.
    It scans every file against ONLY the send family (so it is independent of
    the per-file carve-out logic) and asserts zero hits.
    """
    hits: list[str] = []
    for path in _iter_source_files():
        scanner = _Scanner(
            forbidden_calls=SEND_FAMILY,
            forbidden_props=frozenset(),
            send_family=SEND_FAMILY,
        )
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        scanner.visit(tree)
        for lineno, kind, name in scanner.violations:
            rel = path.relative_to(REPO_ROOT)
            hits.append(f"{rel}:{lineno}: SEND_FAMILY {kind} `.{name}`")
    assert not hits, "Server could send — SEND_FAMILY symbol found:\n" + "\n".join(hits)


def test_send_family_contains_the_critical_names() -> None:
    """Tripwire: the send family must not be silently narrowed."""
    for required in (
        "Send", "Submit", "Reply", "ReplyAll", "Forward",
        "DeliverNow", "Move", "Display", "GetInspector",
        "ExecuteMso", "SendKeys", "Execute", "Post",
        "SetProperty", "SetProperties",
    ):
        assert required in SEND_FAMILY, f"{required} missing from SEND_FAMILY"
    # The draft carve-out must never intersect the send family.
    assert DRAFT_ALLOWED.isdisjoint(SEND_FAMILY)


def test_draft_carveout_permits_management_but_never_send() -> None:
    """drafts.py may create/save/delete/attach — but never any send symbol.

    And every newly-permitted mutator stays forbidden in all other files.
    """
    drafts_calls = forbidden_calls_for("drafts.py")
    for allowed in ("CreateItem", "Save", "Delete", "Add", "Remove"):
        assert allowed not in drafts_calls, f"{allowed} should be permitted in drafts.py"
    for blocked in SEND_FAMILY:
        assert blocked in drafts_calls, f"{blocked} must stay forbidden in drafts.py"

    other_calls = forbidden_calls_for("mail.py")
    for must_block in ("CreateItem", "Save", "Delete", "Add", "Remove"):
        assert must_block in other_calls, (
            f"{must_block} must remain forbidden outside the draft writer"
        )


# ---------------------------------------------------------------------------
# Draft writer containment
# ---------------------------------------------------------------------------

def test_drafts_is_the_sole_writer() -> None:
    """CreateItem / Save may appear ONLY in the draft writer module."""
    offenders: list[str] = []
    for path in _iter_source_files():
        if path.name == DRAFT_MODULE_BASENAME:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if (
                isinstance(node, ast.Call)
                and isinstance(node.func, ast.Attribute)
                and node.func.attr in DRAFT_ALLOWED_METHODS
            ):
                rel = path.relative_to(REPO_ROOT)
                offenders.append(f"{rel}:{node.lineno}: `.{node.func.attr}`")
    assert not offenders, (
        "Write call outside the draft writer module:\n" + "\n".join(offenders)
    )


def test_draft_writer_still_cannot_send() -> None:
    """Scan drafts.py specifically and assert zero send-family symbols."""
    drafts = SRC_ROOT / "tools" / DRAFT_MODULE_BASENAME
    if not drafts.exists():
        pytest.skip("drafts.py not present")
    violations = _scan(drafts)
    sendish = [v for v in violations if v[2] in SEND_FAMILY]
    assert not sendish, f"drafts.py contains a send-family symbol: {sendish}"


# ---------------------------------------------------------------------------
# Meta-tests: prove the scanner actually fires
# ---------------------------------------------------------------------------

def test_scanner_catches_planted_violations(tmp_path) -> None:
    bad = tmp_path / "bad.py"
    bad.write_text(
        "def f(x):\n"
        "    x.Delete()\n"
        "    x.UnRead = False\n",
        encoding="utf-8",
    )
    scanner = _Scanner(
        forbidden_calls=FORBIDDEN_METHODS | FORBIDDEN_APPLICATION_METHODS,
        forbidden_props=FORBIDDEN_PROPERTIES,
    )
    scanner.visit(ast.parse(bad.read_text(encoding="utf-8")))
    kinds = {(k, n) for _, k, n in scanner.violations}
    assert ("call", "Delete") in kinds
    assert ("assign", "UnRead") in kinds


def test_scanner_catches_send_even_with_draft_carveout(tmp_path) -> None:
    """A drafts-like file that tries to Send (directly or via getattr) is caught."""
    bad = tmp_path / "drafts.py"  # named to receive the draft carve-out
    bad.write_text(
        "def f(item):\n"
        "    item.Save()\n"            # allowed in drafts.py
        "    item.Subject = 'x'\n"     # allowed in drafts.py
        "    item.Send()\n"            # FORBIDDEN even in drafts.py
        "    getattr(item, 'Send')()\n"  # dynamic send — also forbidden
        "    f2 = getattr(item, 'Submit')\n",  # deferred send ref — forbidden
        encoding="utf-8",
    )
    scanner = _Scanner(
        forbidden_calls=forbidden_calls_for("drafts.py"),
        forbidden_props=forbidden_properties_for("drafts.py"),
    )
    scanner.visit(ast.parse(bad.read_text(encoding="utf-8")))
    kinds = {(k, n) for _, k, n in scanner.violations}
    # Save + Subject are allowed → must NOT be flagged.
    assert ("call", "Save") not in kinds
    assert ("assign", "Subject") not in kinds
    # Send/Submit must be flagged via every channel.
    assert ("call", "Send") in kinds
    assert ("getattr-call", "Send") in kinds
    assert ("getattr-name", "Send") in kinds
    assert ("getattr-name", "Submit") in kinds


# ---------------------------------------------------------------------------
# Tripwires on the forbidden lists themselves
# ---------------------------------------------------------------------------

def test_forbidden_lists_are_non_empty() -> None:
    assert len(FORBIDDEN_METHODS) >= 25, "FORBIDDEN_METHODS appears truncated"
    assert len(FORBIDDEN_PROPERTIES) >= 10, "FORBIDDEN_PROPERTIES appears truncated"
    assert len(SEND_FAMILY) >= 12, "SEND_FAMILY appears truncated"
    for name in ("Send", "Delete", "Display", "MarkAsTask", "SetProperty"):
        assert name in FORBIDDEN_METHODS, name
    assert "CreateItem" in FORBIDDEN_APPLICATION_METHODS
    assert "UnRead" in FORBIDDEN_PROPERTIES
    assert "Body" in FORBIDDEN_PROPERTIES
