---
name: make-note
description: "Capture a user-requested work note in the daily Markdown file under JobNotes. Trigger when the user says “make note” or asks to make, add, or record a note."
argument-hint: "The detail to add to today's job notes"
user-invocable: true
---

# Make Note

Record the detail the user asks to note in the repository's daily job notes.

## Procedure

1. Use the current local date to locate `JobNotes\YYYY-MM-DD.md` from the repository root.
2. If the daily file does not exist, create it with a date heading and a `## Notes` section. Do not create a second notes folder.
3. Append the requested detail under `## Notes` as a bullet, prefixed with the current local time (`HH:mm`).
4. Preserve all existing content. Do not rewrite or remove other notes, work-log entries, or follow-ups.
5. If the user says only “make note” without specifying the detail and it is not clear from the immediate context, ask what they want recorded instead of guessing.
6. Confirm the note was added and link to the daily file.
