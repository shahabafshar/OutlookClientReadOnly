# outlook-readonly-mcp

A read-only-by-default [MCP server](https://modelcontextprotocol.io) for the locally installed Microsoft Outlook desktop client. It lets an MCP client (Claude Code, Claude Desktop, etc.) query, search, and read mail, calendar events, contacts, tasks, and sticky notes. It can **never send, reply to, forward, move, or delete** mail under any configuration, and never marks items as read.

Optionally — only when you set the env gate `OUTLOOK_MCP_ENABLE_DRAFTS=1` — it can additionally create **unsent drafts** in the Drafts folder. Even then it **cannot send them**: sending is statically impossible in the source. See [Optional: draft creation](#optional-draft-creation-cannot-send).

> Platform: **Windows only.** Requires Outlook desktop client (classic Outlook with COM automation) and Python 3.10+.

---

## The "can never send" guarantee

The headline invariant — *the server cannot send, reply, forward, move, or delete* — is enforced in layers:

1. **Design discipline.** Read paths use only Outlook COM getters and item-resolution methods (`GetItemFromID`, `GetDefaultFolder`, `Folders.Item`, `Items.Restrict`, `Items.Sort`).
2. **CI-enforced AST scan.** [`tests/test_safety_no_writes.py`](tests/test_safety_no_writes.py) parses every source file under [`src/outlook_readonly_mcp/`](src/outlook_readonly_mcp/) and fails the build if any forbidden Outlook COM symbol appears — see [`src/outlook_readonly_mcp/safety.py`](src/outlook_readonly_mcp/safety.py). It is **four-axis**: it flags `.Method(` *calls*, `.Property = …` *assignments*, `getattr(x, "Method")()` *dynamic calls*, and any `getattr(x, "Send")` *literal* of a send-family name. Meta-tests plant known-bad snippets (including a `Send()` inside a drafts-like file) to prove the scanner fires.
3. **`SEND_FAMILY` is forbidden in *every* file** — including the draft writer. This set (`Send`, `Submit`, `DeliverNow`, `Reply*`, `Forward`, `Move`, `Display`, `GetInspector`, `ExecuteMso`, `SendKeys`, `Execute`, `Post`, `SetProperty*`) is never carved out for anyone, so no source file can transition a draft from Drafts → Outbox (the only way Outlook sends). [`test_send_family_never_appears_anywhere`](tests/test_safety_no_writes.py) asserts zero occurrences across the whole tree.
4. **Live unread-flag preservation test.** [`tests/test_outlook_smoke.py`](tests/test_outlook_smoke.py) records the Inbox `UnReadItemCount` before and after listing/reading and asserts it does not change.

Two Outlook write-capable methods exist in the codebase, both tightly scoped:
- `Attachment.SaveAsFile` (in `get_attachment`) — writes attachment **bytes to disk**; does not modify the mail item.
- `CreateItem` + `Save` (in `tools/drafts.py` **only**, and only when the env gate is on) — create and save an **unsent** draft. [`test_drafts_is_the_sole_writer`](tests/test_safety_no_writes.py) asserts these appear in no other file.

### What the server **cannot do, under any configuration**

| Operation | Why it cannot happen |
|---|---|
| Send / Submit / Reply / ReplyAll / Forward an email | In `SEND_FAMILY` → forbidden in every file + AST scan |
| Move an item (e.g. into the Outbox) or delete it | `Move` / `Delete` denylist + AST scan |
| Mark as read / unread | `UnRead` in property denylist; live test asserts count invariant |
| Open an item in the UI (can flip UnRead / enable Ctrl+Enter send) | `Display`, `GetInspector` denylist |
| Send via ribbon, synthetic keystrokes, or MAPI submit flag | `ExecuteMso`, `SendKeys`, `SetProperty*` denylist |
| Modify an existing item's subject, body, recipients, etc. | Property-assignment denylist (drafts.py may only set fields on a **new** item) |
| Create appointments, contacts, tasks, notes | `CreateItem` carve-out is mail-only and confined to drafts.py |

**Honest limit on drafts.** A draft, once created, is an ordinary Outlook item. A human (or other software) can open Outlook and send it. This server guarantees *it* can never send; it cannot make the draft itself unsendable — Outlook offers no such thing. If your requirement is "no one can ever send this," don't enable drafts.

---

## Capabilities

**19 read tools (always on):**

- **Accounts & Folders:** `list_accounts`, `list_folders`, `list_calendar_folders`, `list_contact_folders`
- **Mail:** `list_emails` (filter by unread, attachment, sender, subject, date range), `search_emails` (subject/sender/body), `get_email`, `list_attachments`, `get_attachment` (save to disk and/or base64)
- **Calendar:** `list_events` (date-range, recurrences expanded), `get_event`
- **Contacts:** `list_contacts`, `search_contacts`, `get_contact`
- **Tasks & Notes:** `list_tasks`, `get_task`, `list_notes`, `get_note`
- **Meta:** `server_info` (reports `read_only`, `drafts_enabled`, `can_send: false`, and the forbidden-operations list)

**1 write tool (only when `OUTLOOK_MCP_ENABLE_DRAFTS=1`):** `create_draft`.

Items are identified by Outlook **EntryID** (stable per-store). `store_id` is optional and only needed when an EntryID could be ambiguous across multiple configured accounts.

---

## Optional: draft creation (cannot send)

Disabled by default. To turn it on, set the environment variable in your MCP client config:

```json
{
  "mcpServers": {
    "outlook-readonly": {
      "command": "python",
      "args": ["-m", "outlook_readonly_mcp.server"],
      "env": { "OUTLOOK_MCP_ENABLE_DRAFTS": "1" }
    }
  }
}
```

When enabled, one tool appears:

```
create_draft(subject="", body=None, html_body=None, to=[], cc=[], bcc=[])
```

- Saves an **unsent** draft to the default **Drafts** folder and returns its `entry_id`.
- Provide either `body` (plain text) **or** `html_body`, not both.
- `to` / `cc` / `bcc` are lists of addresses or `"Name <addr>"` strings.
- Before returning, the writer reads the item back and **aborts** unless it reports `Sent == False` and is physically in the Drafts folder.
- **Not supported on purpose:** attachments (needs `Attachments.Add`) and reply/forward drafts (need `Reply*`) — those symbols stay forbidden.

The created COM object never leaves the writer function; callers only ever receive a plain `DraftResult`, so nothing downstream can call `.Send()` on it either.

### Verifying draft creation locally (opt-in, leaves an artifact)

There is a live test that actually creates a draft. It is **double-gated** and skipped by default, because the server cannot delete what it creates:

```powershell
$env:OUTLOOK_MCP_ENABLE_DRAFTS = "1"
$env:OUTLOOK_MCP_TEST_CREATE_DRAFT = "1"
python -m pytest tests/test_outlook_smoke.py::test_create_draft_lands_unsent_in_drafts -s
```

It prints the new draft's EntryID and asserts it is unsent and in Drafts. Delete the draft manually in Outlook afterward.

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

Restart the client. The 19 read tools appear under the `outlook-readonly` server (plus `create_draft` if the draft gate is set).

---

## Project layout

```
src/outlook_readonly_mcp/
  __init__.py
  safety.py          # forbidden-symbols denylist, SEND_FAMILY, scoped carve-out helpers
  outlook.py         # COM connection, apartment management, folder/item resolution
  converters.py      # COM item -> Pydantic model (only reads)
  models.py          # Pydantic data carriers
  server.py          # FastMCP entry point, tool registration, draft env-gate
  tools/
    mail.py          # accounts, folders, emails, attachments
    calendar.py      # calendar folders + events
    contacts.py      # contacts + search
    tasks.py         # tasks + sticky notes
    drafts.py        # THE ONLY WRITER: create unsent drafts (env-gated). Never sends.
tests/
  test_safety_no_writes.py  # AST scan -- the can-never-send gate (4-axis, scoped)
  test_imports.py           # platform-portable smoke
  test_outlook_smoke.py     # live Outlook tests + opt-in draft test (auto-skip if unavailable)
```

---

## Security notes

- **Local-only.** The server runs as a child process of the MCP client. It does not open network ports. Outlook's own data exposure rules (Exchange permissions, on-disk PST/OST ACLs) still apply.
- **Outlook may show a security prompt** the first time a process talks to it via COM. This is normal — Outlook is verifying that an external program is allowed to read mail.
- **Attachment extraction** uses `Attachment.SaveAsFile`, which writes attachment bytes to disk and does not modify the parent mail item. By default the server caps extraction at 25 MiB per attachment; raise via the tool parameter if needed.
- **Drafts are off by default** and require the explicit `OUTLOOK_MCP_ENABLE_DRAFTS=1` env gate. Even when on, the server can only create *unsent* drafts and can never send — see [the can-never-send guarantee](#the-can-never-send-guarantee). The server cannot delete a draft it created (`Delete` is forbidden); remove drafts manually in Outlook.
- **No telemetry.** The server logs nothing externally and stores nothing on disk except temporary attachment files (cleaned up after base64 encoding).

---

## License

MIT
