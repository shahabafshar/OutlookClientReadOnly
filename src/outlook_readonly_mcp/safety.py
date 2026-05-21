"""Safety constants for the read-only Outlook MCP server.

This module declares the complete set of Outlook / COM symbols this server
is forbidden from ever using. The accompanying test
`tests/test_safety_no_writes.py` scans every source file under `src/` and
fails CI if any forbidden symbol is referenced.

The forbidden set covers anything that could mutate Outlook state: send,
delete, save, move/copy, mark-as-read, reply, forward, print, create-item,
display (which can flip UnRead), and direct collection mutations.
"""

from __future__ import annotations

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
