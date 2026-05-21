"""CI-enforced safety test: scan source for any mutating COM call.

This test parses every Python file under `src/outlook_readonly_mcp/` (except
`safety.py` itself, which legitimately declares the forbidden symbols as
data) and looks for two patterns via AST:

1. Call expressions of the form `<expr>.NAME(...)` where NAME is in
   FORBIDDEN_METHODS or FORBIDDEN_APPLICATION_METHODS.
2. Assignment targets of the form `<expr>.NAME` where NAME is in
   FORBIDDEN_PROPERTIES.

If any match is found, the test fails with a precise file:line report.
This is the build-time guarantee that the server is structurally
incapable of mutating Outlook state.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

from outlook_readonly_mcp.safety import (
    FORBIDDEN_APPLICATION_METHODS,
    FORBIDDEN_METHODS,
    FORBIDDEN_PROPERTIES,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src" / "outlook_readonly_mcp"

# The safety module legitimately holds the forbidden names as string data;
# it must be excluded so the test isn't false-positive on itself.
EXCLUDED_FILES = {SRC_ROOT / "safety.py"}

FORBIDDEN_CALLS = FORBIDDEN_METHODS | FORBIDDEN_APPLICATION_METHODS


def _iter_source_files() -> list[Path]:
    return [p for p in SRC_ROOT.rglob("*.py") if p not in EXCLUDED_FILES]


class _Scanner(ast.NodeVisitor):
    def __init__(self, source_path: Path) -> None:
        self.path = source_path
        self.violations: list[tuple[int, str, str]] = []

    def visit_Call(self, node: ast.Call) -> None:  # noqa: N802
        if isinstance(node.func, ast.Attribute):
            name = node.func.attr
            if name in FORBIDDEN_CALLS:
                self.violations.append((node.lineno, "call", name))
        self.generic_visit(node)

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
        if isinstance(target, ast.Attribute) and target.attr in FORBIDDEN_PROPERTIES:
            self.violations.append((lineno, "assign", target.attr))


def test_no_forbidden_symbols_in_source() -> None:
    """Every source file must be free of forbidden COM calls and assignments."""
    all_violations: list[str] = []
    files_scanned = 0
    for path in _iter_source_files():
        files_scanned += 1
        try:
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        except SyntaxError as e:  # pragma: no cover - sanity
            pytest.fail(f"Could not parse {path}: {e}")
        scanner = _Scanner(path)
        scanner.visit(tree)
        for lineno, kind, name in scanner.violations:
            rel = path.relative_to(REPO_ROOT)
            all_violations.append(f"{rel}:{lineno}: forbidden {kind} `.{name}`")

    assert files_scanned > 0, "No source files were scanned — test setup is broken."
    assert not all_violations, (
        "Read-only invariant broken — these forbidden COM symbols appear in source:\n"
        + "\n".join(all_violations)
    )


def test_forbidden_lists_are_non_empty() -> None:
    """Tripwire: the forbidden lists must remain populated, or the scan is a no-op."""
    assert len(FORBIDDEN_METHODS) >= 20, "FORBIDDEN_METHODS appears truncated"
    assert len(FORBIDDEN_PROPERTIES) >= 10, "FORBIDDEN_PROPERTIES appears truncated"
    assert "Send" in FORBIDDEN_METHODS
    assert "Delete" in FORBIDDEN_METHODS
    assert "Display" in FORBIDDEN_METHODS
    assert "MarkAsTask" in FORBIDDEN_METHODS
    assert "CreateItem" in FORBIDDEN_APPLICATION_METHODS
    assert "UnRead" in FORBIDDEN_PROPERTIES
    assert "Body" in FORBIDDEN_PROPERTIES


def test_scanner_catches_a_planted_violation(tmp_path) -> None:
    """Meta-test: confirm the scanner actually flags a known-bad snippet."""
    bad = tmp_path / "bad.py"
    bad.write_text("def f(x):\n    x.Delete()\n    x.UnRead = False\n", encoding="utf-8")
    tree = ast.parse(bad.read_text(encoding="utf-8"), filename=str(bad))
    scanner = _Scanner(bad)
    scanner.visit(tree)
    kinds = {(kind, name) for _, kind, name in scanner.violations}
    assert ("call", "Delete") in kinds
    assert ("assign", "UnRead") in kinds


def test_no_create_item_call() -> None:
    """Explicit double-check: CreateItem never appears as a call anywhere in src."""
    found: list[str] = []
    for path in _iter_source_files():
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if (
                isinstance(node, ast.Call)
                and isinstance(node.func, ast.Attribute)
                and node.func.attr in {"CreateItem", "CreateItemFromTemplate"}
            ):
                rel = path.relative_to(REPO_ROOT)
                found.append(f"{rel}:{node.lineno}")
    assert not found, "CreateItem* call found (server must never create items): " + ", ".join(found)
