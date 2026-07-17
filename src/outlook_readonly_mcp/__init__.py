"""Read-only MCP server for Microsoft Outlook desktop client.

This package is structurally incapable of modifying Outlook state:
- No mutating COM methods are wired anywhere.
- A unit test (tests/test_safety_no_writes.py) scans the source tree and
  fails the build if any forbidden COM symbol appears.
- All item lookups go through GetItemFromID / GetDefaultFolder; no
  CreateItem path exists.
"""

__version__ = "0.3.0"
