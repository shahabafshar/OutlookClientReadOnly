# outlook-readonly-mcp

A read-only-by-default [MCP server](https://modelcontextprotocol.io) for the locally installed Microsoft Outlook desktop client. It lets an MCP client (Claude Code, Claude Desktop, etc.) query, search, and read mail, calendar events, contacts, tasks, and sticky notes. It can **never send, reply to, forward, move, or delete** mail under any configuration, and never marks items as read.

By default it can additionally create **unsent drafts** in the Drafts folder (disable with `OUTLOOK_MCP_ENABLE_DRAFTS=0`). Even when drafting, it **cannot send them**: sending is statically impossible in the source. See [Draft creation](#draft-creation-cannot-send).

> Platform: **Windows only.** Requires Outlook desktop client (classic Outlook with COM automation) and Python 3.10+.

---

## The "can never send" guarantee

The headline invariant — *the server cannot send, reply, forward, move, or delete* — is enforced in layers:

1. **Design discipline.** Read paths use only Outlook COM getters and item-resolution methods (`GetItemFromID`, `GetDefaultFolder`, `Folders.Item`, `Items.Restrict`, `Items.Sort`).
2. **CI-enforced AST scan.** [`tests/test_safety_no_writes.py`](tests/test_safety_no_writes.py) parses every source file under [`src/outlook_readonly_mcp/`](src/outlook_readonly_mcp/) and fails the build if any forbidden Outlook COM symbol appears — see [`src/outlook_readonly_mcp/safety.py`](src/outlook_readonly_mcp/safety.py). It is **four-axis**: it flags `.Method(` *calls*, `.Property = …` *assignments*, `getattr(x, "Method")()` *dynamic calls*, and any `getattr(x, "Send")` *literal* of a send-family name. Meta-tests plant known-bad snippets (including a `Send()` inside a drafts-like file) to prove the scanner fires.
3. **`SEND_FAMILY` is forbidden in *every* file** — including the draft writer. This set (`Send`, `Submit`, `DeliverNow`, `Reply*`, `Forward`, `Move`, `Display`, `GetInspector`, `ExecuteMso`, `SendKeys`, `Execute`, `Post`, `SetProperty*`) is never carved out for anyone, so no source file can transition a draft from Drafts → Outbox (the only way Outlook sends). [`test_send_family_never_appears_anywhere`](tests/test_safety_no_writes.py) asserts zero occurrences across the whole tree.
4. **Live unread-flag preservation test.** [`tests/test_outlook_smoke.py`](tests/test_outlook_smoke.py) records the Inbox `UnReadItemCount` before and after listing/reading and asserts it does not change.

The write-capable methods in the codebase are all tightly scoped:
- `Attachment.SaveAsFile` (in `get_attachment`) — writes attachment **bytes to disk**; does not modify the mail item.
- `CreateItem`, `Save`, `Delete`, `Attachments.Add`, `Attachments.Remove` (in `tools/drafts.py` **only**, and only when the env gate is on) — create/edit/delete drafts and manage their attachments. [`test_drafts_is_the_sole_writer`](tests/test_safety_no_writes.py) asserts these appear in no other file. None of them is in `SEND_FAMILY`, so they cannot send.

### Two distinct guarantees

The draft-management surface has **two different kinds of protection** — it's important to know which is which:

1. **Can never send — STATIC & absolute.** Proven by the AST scan: no `SEND_FAMILY` symbol (`Send`, `Submit`, `Reply*`, `Forward`, `Move`, `DeliverNow`, `Display`, `GetInspector`, `ExecuteMso`, `SendKeys`, `Execute`, `Post`, `SetProperty*`) appears in *any* file, drafts.py included. Outlook only sends from the Outbox, and there is no `Move`/`Send` path to get a draft there. This holds no matter what the code does at runtime.
2. **Only ever touches drafts — RUNTIME guard.** Because `Delete`/`Add`/`Remove` are generic mutators, the scan no longer proves drafts.py mutates *only* drafts. Instead, every edit/delete/attachment op first calls `_resolve_editable_draft`, which **refuses** any item that is not an *unsent* mail item physically located in the **Drafts** folder. So the server can never edit or delete sent or received mail — but this is defense-by-validation, not defense-by-construction.

### What the server **cannot do, under any configuration**

| Operation | Why it cannot happen |
|---|---|
| Send / Submit / Reply / ReplyAll / Forward an email | In `SEND_FAMILY` → statically forbidden in every file |
| Move an item into the Outbox (the only send path) | `Move` in `SEND_FAMILY` → statically forbidden |
| Send via UI, ribbon, synthetic keystrokes, or MAPI submit flag | `Display`, `GetInspector`, `ExecuteMso`, `SendKeys`, `SetProperty*` forbidden |
| Modify or delete a **sent** or **received** email | Runtime guard rejects any non-draft / sent item |
| Mark as read / unread | `UnRead` in property denylist; live test asserts count invariant |
| Create appointments, contacts, tasks, notes | `CreateItem` carve-out is mail-only and confined to drafts.py |

**What the server CAN do (only when drafts are enabled):** create, edit, and delete **unsent drafts**, and add/remove their attachments — all confined to the Drafts folder at runtime.

**Honest limits on drafts.**
- A draft, once created, is an ordinary Outlook item. A human (or other software) can open Outlook and send it. This server guarantees *it* can never send; it cannot make the draft itself unsendable — Outlook offers no such thing.
- `delete_draft` moves the draft to **Deleted Items** (recoverable); it does not permanently destroy it (`PermanentlyDelete` stays forbidden).

---

## Capabilities

**19 read tools (always on):**

- **Accounts & Folders:** `list_accounts`, `list_folders`, `list_calendar_folders`, `list_contact_folders`
- **Mail:** `list_emails` (filter by unread, attachment, sender, subject, date range), `search_emails` (subject/sender/body), `get_email`, `list_attachments`, `get_attachment` (save to disk and/or base64)
- **Calendar:** `list_events` (date-range, recurrences expanded), `get_event`
- **Contacts:** `list_contacts`, `search_contacts`, `get_contact`
- **Tasks & Notes:** `list_tasks`, `get_task`, `list_notes`, `get_note`
- **Meta:** `server_info` (reports `read_only`, `drafts_enabled`, `can_send: false`, and the forbidden-operations list)

**6 draft-management tools (enabled by default; disable with `OUTLOOK_MCP_ENABLE_DRAFTS=0`):** `create_draft`, `update_draft`, `delete_draft`, `add_draft_attachment`, `remove_draft_attachment`, `list_draft_attachments`. All confined to unsent drafts in the Drafts folder; none can send.

Items are identified by Outlook **EntryID** (stable per-store). `store_id` is optional and only needed when an EntryID could be ambiguous across multiple configured accounts.

---

## Draft creation (cannot send)

**Enabled by default.** To turn it OFF, set the environment variable in your MCP client config:

```json
{
  "mcpServers": {
    "outlook-readonly": {
      "command": "python",
      "args": ["-m", "outlook_readonly_mcp.server"],
      "env": { "OUTLOOK_MCP_ENABLE_DRAFTS": "0" }
    }
  }
}
```

When enabled (the default), six draft-management tools appear:

```
create_draft(subject="", body=None, html_body=None, to=[], cc=[], bcc=[], attachments=[])
update_draft(entry_id, store_id=None, subject=None, body=None, html_body=None, to=None, cc=None, bcc=None)
delete_draft(entry_id, store_id=None)
add_draft_attachment(entry_id, file_path, store_id=None)
remove_draft_attachment(entry_id, attachment_index, store_id=None)
list_draft_attachments(entry_id, store_id=None)
```

- **create_draft** — saves an **unsent** draft to the **Drafts** folder and returns its `entry_id`. Provide `body` (plain text) **or** `html_body`, not both. `to`/`cc`/`bcc` are lists of addresses or `"Name <addr>"` strings. `attachments` is a list of local file paths. Before returning it reads the item back and **aborts** unless it reports `Sent == False` and is in Drafts.
- **update_draft** — edits an existing draft. Only the fields you pass change; omit a field to leave it. `""` clears a text field; `[]` clears recipients.
- **delete_draft** — moves the draft to **Deleted Items** (recoverable).
- **add/remove/list_draft_attachments** — manage a draft's attachments (`add` takes a local file path; `remove` takes a 1-based index).
- **Every edit/delete/attachment op first verifies the target is an unsent draft in the Drafts folder, and refuses otherwise** — so none of these can touch sent or received mail.
- **Reply/forward drafts are not supported** (they need `Reply*`, which stays forbidden). Compose content directly.

The live COM item never leaves the writer functions; callers only receive a plain `DraftResult`/`DraftDeleteResult`, so nothing downstream can call `.Send()` on it either.

### Verifying the draft lifecycle locally (opt-in, self-cleaning)

A live test runs the full lifecycle — create → edit → add attachment → remove attachment → verify the guard refuses a non-draft → **delete (cleanup)** — leaving nothing behind. It is **double-gated** and skipped by default:

```powershell
$env:OUTLOOK_MCP_ENABLE_DRAFTS = "1"
$env:OUTLOOK_MCP_TEST_CREATE_DRAFT = "1"
python -m pytest tests/test_outlook_smoke.py::test_draft_full_lifecycle_self_cleaning -v
```

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

Restart the client. The 19 read tools appear under the `outlook-readonly` server, plus `create_draft` (on by default; disable with `OUTLOOK_MCP_ENABLE_DRAFTS=0`).

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
    drafts.py        # THE ONLY WRITER: create/edit/delete unsent drafts + attachments
                     #   (env-gated). Runtime-guarded to Drafts only. Never sends.
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
- **Drafts are on by default** (disable with `OUTLOOK_MCP_ENABLE_DRAFTS=0`). Either way, the server can only create *unsent* drafts and can never send — see [the can-never-send guarantee](#the-can-never-send-guarantee). The server cannot delete a draft it created (`Delete` is forbidden); remove drafts manually in Outlook.
- **No telemetry.** The server logs nothing externally and stores nothing on disk except temporary attachment files (cleaned up after base64 encoding).

---

## License

MIT
