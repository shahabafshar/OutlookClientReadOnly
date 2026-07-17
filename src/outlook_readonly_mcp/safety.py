"""Safety constants for the Outlook MCP server.

This module declares the complete set of Outlook / COM symbols this server
is forbidden from using. The accompanying test
`tests/test_safety_no_writes.py` scans every source file under `src/` and
fails CI if any forbidden symbol is referenced.

The forbidden set covers anything that could mutate Outlook state: send,
delete, save, move/copy, mark-as-read, reply, forward, print, create-item,
display (which can flip UnRead), low-level MAPI writes, and direct
collection mutations.

Draft management (optional, env-gated) is the *only* write surface, confined
to a single module — `tools/drafts.py`. That file is granted a carve-out
(`CreateItem`, `Reply`, `ReplyAll`, `Forward` as COMPOSE methods, `Save`,
`Delete`, attachment `Add`/`Remove`, and the property setters needed to
compose a draft) via `forbidden_calls_for()` / `forbidden_properties_for()`.
Crucially, the carve-out does NOT include any symbol in `SEND_FAMILY`: send,
submit, move, display, deliver-now, MAPI submit-flag writes, etc. remain
forbidden in drafts.py exactly as everywhere else. The scan therefore
statically proves that no source file — drafts.py included — can transition a
draft from Drafts to a sent state.
"""

from __future__ import annotations

import os

# Methods on MailItem / AppointmentItem / ContactItem / TaskItem / NoteItem
# that mutate state. Includes anything Outlook treats as a write or that
# can trigger an implicit write (e.g. Display marks the item as read).
FORBIDDEN_METHODS: frozenset[str] = frozenset({
    # mail send / submit
    "Send", "Submit", "ClearConversationIndex",
    # delete / discard
    "Delete", "PermanentlyDelete", "Restore",
    # save / persist
    "Save", "SaveAs", "SaveSentMessageFolder",
    # move / copy / cut
    "Move", "Copy", "CopyToShared", "Cut",
    # ui actions that mutate state (Display flips UnRead via reading-pane)
    "Display", "PrintOut", "Close",
    # reply / forward
    "Reply", "ReplyAll", "ReplyAllEx", "Forward",
    # creation
    "CreateItem", "CreateItemFromTemplate", "CreateSharingItem",
    "CreateRecipient", "Add", "AddPicture", "AddCustomAction",
    # task / flag mutations
    "MarkAsTask", "ClearTaskFlag", "MarkComplete",
    # category / properties mutations
    "AssignToCategories", "ShowCategoriesDialog",
    # collections that mutate
    "Remove", "RemoveAll", "Import",
    # rule / search-folder creation
    "AddSearchFolder",
    # account / store changes
    "DeliverNow", "SynchronizeBookings",
    # low-level MAPI writes (PropertyAccessor.SetProperty* can set submit flags)
    "SetProperty", "SetProperties",
    # UI / automation paths that could trigger a send (inspector + ribbon +
    # synthetic keystrokes), and the Actions/Execute reply-forward pipeline
    "GetInspector", "ExecuteMso", "SendKeys", "Execute", "Post",
})

# Property setters or boolean state we must never assign to. The test
# enforces this by detecting any assignment statement of the form
# `<expr>.<NAME> = ...` where NAME appears in this set.
FORBIDDEN_PROPERTIES: frozenset[str] = frozenset({
    "UnRead", "ReadReceiptRequested", "OriginatorDeliveryReportRequested",
    "Subject", "Body", "HTMLBody", "RTFBody",
    "To", "CC", "BCC", "SentOnBehalfOfName", "SendUsingAccount",
    "Importance", "Sensitivity", "FlagStatus", "FlagRequest",
    "ReminderSet", "ReminderTime", "Categories",
    "DeleteAfterSubmit", "AutoForwarded", "MessageClass",
    "Start", "End", "Location", "AllDayEvent", "Status",
    "DueDate", "StartDate", "Complete", "PercentComplete",
    "FullName", "FileAs", "Email1Address", "BusinessTelephoneNumber",
})

# Application-level methods that initialize new items must also never appear.
FORBIDDEN_APPLICATION_METHODS: frozenset[str] = frozenset({
    "CreateItem", "CreateItemFromTemplate", "CreateObject",
    "Quit", "ActiveExplorer",  # ActiveExplorer can trigger UI side effects
})

# The complete set we scan for.
ALL_FORBIDDEN: frozenset[str] = (
    FORBIDDEN_METHODS | FORBIDDEN_PROPERTIES | FORBIDDEN_APPLICATION_METHODS
)

# All symbols that can cause an email to be SENT, or are a precondition to a
# send (UI display + ribbon + synthetic keystrokes, relocation to the Outbox,
# MAPI submit-flag writes). These are forbidden in EVERY file — including the
# draft writer `tools/drafts.py`. This is the load-bearing "can never send"
# guarantee: the scan proves none of these names appears anywhere in source,
# as a method call, an assignment, or a `getattr` string literal.
#
# NOTE on Reply/ReplyAll/Forward: these are NOT send operations in the Outlook
# object model — item.Reply() merely CONSTRUCTS a new MailItem pre-wired with
# thread identity (ConversationIndex/Topic), recipients, and quoted history.
# Nothing leaves the machine unless Send/Submit/Move-to-Outbox happens later,
# and all of those stay in this set. They are therefore classified as compose
# methods in the drafts carve-out (drafts.py only), not send-family.
# `ReplyAllEx` and the Actions `Execute` pipeline remain forbidden everywhere.
SEND_FAMILY: frozenset[str] = frozenset({
    "Send", "Submit", "DeliverNow", "SaveSentMessageFolder",
    "ReplyAllEx",
    "Post", "Execute", "ExecuteMso", "SendKeys", "GetInspector",
    "Display", "Move", "SetProperty", "SetProperties",
})

# ---------------------------------------------------------------------------
# Draft-writer carve-out
# ---------------------------------------------------------------------------
# The ONE module permitted to write. Keyed on basename so any *other* file is
# always scanned against the full forbidden set.
DRAFT_MODULE_BASENAME = "drafts.py"

# The mutating calls the draft writer may use. ALL of these are disjoint from
# SEND_FAMILY, so permitting them in drafts.py does not weaken the can-never-
# send guarantee one bit (Move/Send/Submit/Display/SetProperty* stay forbidden):
#   CreateItem  - make a new mail item
#   Reply       - CONSTRUCT a reply MailItem (correct thread identity,
#   ReplyAll      recipients, quoted history). Does not send; the result is
#                 immediately Save()d into Drafts and verified unsent.
#   Forward     - CONSTRUCT a forward MailItem (carries original attachments).
#   Save        - persist to Drafts (never sends; Send would route to Outbox)
#   Delete      - remove a draft (runtime-guarded to drafts only; recoverable
#                 via Deleted Items)
#   Add         - add an attachment to a draft (Attachments.Add)
#   Remove      - remove an attachment from a draft (Attachments.Remove)
#
# IMPORTANT: Delete/Add/Remove are generic mutators. The static scan no longer
# proves drafts.py only mutates *drafts* — that is enforced at RUNTIME by
# `tools/drafts._resolve_editable_draft`, which refuses any item that is not an
# unsent item physically in the Drafts folder. Edit/delete/attachment ops MUST
# go through that guard. Reply/ReplyAll/Forward read the source item without
# modifying it and their result MUST be routed through the saved-draft guard.
DRAFT_ALLOWED_METHODS: frozenset[str] = frozenset({
    "CreateItem", "Reply", "ReplyAll", "Forward", "Save", "Delete", "Add", "Remove",
})

# The only properties the draft writer may assign: the fields of a draft.
DRAFT_ALLOWED_PROPERTIES: frozenset[str] = frozenset({
    "Subject", "Body", "HTMLBody", "To", "CC", "BCC",
})

DRAFT_ALLOWED: frozenset[str] = DRAFT_ALLOWED_METHODS | DRAFT_ALLOWED_PROPERTIES

# Sanity guards on the policy itself (also asserted by the test suite): the
# draft carve-out must never overlap the send family.
assert DRAFT_ALLOWED.isdisjoint(SEND_FAMILY), "draft carve-out leaks a send symbol"
assert SEND_FAMILY <= (FORBIDDEN_METHODS | FORBIDDEN_PROPERTIES), (
    "SEND_FAMILY must be a subset of the forbidden symbols"
)


def _is_draft_writer(filename: str) -> bool:
    return os.path.basename(filename) == DRAFT_MODULE_BASENAME


def forbidden_calls_for(filename: str) -> frozenset[str]:
    """Effective set of forbidden *call* names for a given source file.

    Every file is forbidden the full method + application-method set, except
    the draft writer, which is permitted DRAFT_ALLOWED_METHODS only. The
    send family is never removed, so it stays forbidden in drafts.py too.
    """
    base = FORBIDDEN_METHODS | FORBIDDEN_APPLICATION_METHODS
    if _is_draft_writer(filename):
        return base - DRAFT_ALLOWED_METHODS
    return base


def forbidden_properties_for(filename: str) -> frozenset[str]:
    """Effective set of forbidden *assignment* targets for a given file.

    The draft writer may assign the draft fields (Subject/Body/To/...); all
    other files may assign none of the forbidden properties.
    """
    if _is_draft_writer(filename):
        return FORBIDDEN_PROPERTIES - DRAFT_ALLOWED_PROPERTIES
    return FORBIDDEN_PROPERTIES


# Outlook constants (subset we actually use). Mirrored here so we don't
# need to import the COM enum dispatch at import time.
class OlDefaultFolders:
    olFolderDeletedItems = 3
    olFolderOutbox = 4
    olFolderSentMail = 5
    olFolderInbox = 6
    olFolderCalendar = 9
    olFolderContacts = 10
    olFolderJournal = 11
    olFolderNotes = 12
    olFolderTasks = 13
    olFolderDrafts = 16
    olFolderJunk = 23


class OlItemType:
    olMailItem = 0
    olAppointmentItem = 1
    olContactItem = 2
    olTaskItem = 3
    olJournalItem = 4
    olNoteItem = 5
    olPostItem = 6
    olDistributionListItem = 7


# Item class names returned by Outlook MessageClass strings.
MAIL_CLASSES = ("IPM.Note",)
APPT_CLASSES = ("IPM.Appointment",)
CONTACT_CLASSES = ("IPM.Contact",)
TASK_CLASSES = ("IPM.Task",)
NOTE_CLASSES = ("IPM.StickyNote",)
