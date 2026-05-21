# outlook-readonly-mcp

A **strictly read-only** [MCP server](https://modelcontextprotocol.io) for the locally installed Microsoft Outlook desktop client. It lets an MCP client (Claude Code, Claude Desktop, etc.) query, search, and read mail, calendar events, contacts, tasks, and sticky notes — and it is **structurally incapable** of creating, modifying, deleting, sending, replying, forwarding, or even marking items as read.

> Platform: **Windows only.** Requires Outlook desktop client (classic Outlook with COM automation) and Python 3.10+.

---

## Read-only guarantee

The "read-only" promise is enforced three ways, in layers:

1. **Design discipline.** Every code path uses only Outlook COM getters and item-resolution methods (`GetItemFromID`, `GetDefaultFolder`, `Folders.Item`, `Items.Restrict`, `Items.Sort`). Nothing creates, sends, saves, moves, copies, deletes, or marks items as read.
2. **CI-enforced AST scan.** [`tests/test_safety_no_writes.py`](tests/test_safety_no_writes.py) parses every source file under [`src/outlook_readonly_mcp/`](src/outlook_readonly_mcp/) and fails the build if any forbidden Outlook COM symbol appears — see [`src/outlook_readonly_mcp/safety.py`](src/outlook_readonly_mcp/safety.py) for the full denylist (`Send`, `Delete`, `Save`, `SaveAs`, `Move`, `Copy`, `Reply*`, `Forward`, `Display`, `MarkAsTask`, `CreateItem*`, plus property assignments to `UnRead`, `Subject`, `Body`, recipients, etc.). The scan is two-axis: it flags both `.Method(` *call* expressions and `.Property = ...` *assignment* targets. A meta-test plants a known-bad snippet to prove the scanner actually fires.
3. **Live unread-flag preservation test.** [`tests/test_outlook_smoke.py`](tests/test_outlook_smoke.py) records the Inbox `UnReadItemCount` before and after listing and reading emails and asserts it does not change.

The single Outlook write-capable method the server is allowed to call is `Attachment.SaveAsFile` (used by `get_attachment`) — this writes attachment **bytes to disk** and does not modify the parent mail item in any way. It is constrained to attachment extraction only.

### What the server **cannot do, by design**

| Operation | Why it cannot happen |
|---|---|
| Send / Submit / Reply / ReplyAll / Forward an email | Method denylist + AST scan |
| Delete, move, or copy any item | Method denylist + AST scan |
| Create a new mail, appointment, contact, task, or note | `CreateItem*` denylist + AST scan |
| Modify subject, body, recipients, importance, categories, etc. | Property assignment denylist + AST scan |
| Mark as read / unread | `UnRead` is in the property denylist; live test asserts count is invariant |
| Open an item in the UI (which can flip UnRead) | `Display` is in the method denylist |
| Print or save an item file | `PrintOut`, `Save`, `SaveAs` are in the method denylist |
| Add or remove folders / search folders / rules | `Add`, `Remove`, `AddSearchFolder` are in the method denylist |

---

## Capabilities

19 MCP tools, all read-only:

- **Accounts & Folders:** `list_accounts`, `list_folders`, `list_calendar_folders`, `list_contact_folders`
- **Mail:** `list_emails` (filter by unread, attachment, sender, subject, date range), `search_emails` (subject/sender/body), `get_email`, `list_attachments`, `get_attachment` (save to disk and/or base64)
- **Calendar:** `list_events` (date-range, recurrences expanded), `get_event`
- **Contacts:** `list_contacts`, `search_contacts`, `get_contact`
- **Tasks & Notes:** `list_tasks`, `get_task`, `list_notes`, `get_note`
- **Meta:** `server_info` (declares read-only and lists forbidden operations)

Items are identified by Outlook **EntryID** (stable per-store). `store_id` is optional and only needed when an EntryID could be ambiguous across multiple configured accounts.

---

## Install

```powershell
# From the project directory
python -m pip install -e .
```

This installs the `outlook-readonly-mcp` console entry point and the `outlook_readonly_mcp` Python package.

For development:

```powershell
python -m pip install -e ".[dev]"
python -m pytest
```

The test suite has three groups:
- **Safety scan tests** — run on every platform, are the load-bearing CI gate.
- **Import sanity tests** — run on every platform.
- **Live Outlook smoke tests** — auto-skipped when Outlook isn't reachable.

---

## Wire it up to an MCP client

The server speaks MCP over **stdio**. Configure your MCP client to launch the entry point.

### Claude Desktop / Claude Code

Add to your MCP servers config (e.g. `claude_desktop_config.json` or a `.mcp.json` at the workspace root):

```json
{
  "mcpServers": {
    "outlook-readonly": {
      "command": "outlook-readonly-mcp"
    }
  }
}
```

If the script isn't on `PATH`, use the full Python invocation:

```json
{
  "mcpServers": {
    "outlook-readonly": {
      "command": "python",
      "args": ["-m", "outlook_readonly_mcp.server"]
    }
  }
}
```

Restart the client. The 19 tools above will appear under the `outlook-readonly` server.

---

## Project layout

```
src/outlook_readonly_mcp/
  __init__.py
  safety.py          # forbidden-symbols denylist + Outlook constant mirrors
  outlook.py         # COM connection, apartment management, folder/item resolution
  converters.py      # COM item -> Pydantic model (only reads)
  models.py          # Pydantic data carriers
  server.py          # FastMCP entry point, tool registration
  tools/
    mail.py          # accounts, folders, emails, attachments
    calendar.py      # calendar folders + events
    contacts.py      # contacts + search
    tasks.py         # tasks + sticky notes
tests/
  test_safety_no_writes.py  # AST scan -- the read-only gate
  test_imports.py           # platform-portable smoke
  test_outlook_smoke.py     # live Outlook tests (auto-skip if unavailable)
```

---

## Security notes

- **Local-only.** The server runs as a child process of the MCP client. It does not open network ports. Outlook's own data exposure rules (Exchange permissions, on-disk PST/OST ACLs) still apply.
- **Outlook may show a security prompt** the first time a process talks to it via COM. This is normal — Outlook is verifying that an external program is allowed to read mail.
- **Attachment extraction** uses `Attachment.SaveAsFile`, which writes attachment bytes to disk and does not modify the parent mail item. By default the server caps extraction at 25 MiB per attachment; raise via the tool parameter if needed.
- **No telemetry.** The server logs nothing externally and stores nothing on disk except temporary attachment files (cleaned up after base64 encoding).

---

## License

MIT
