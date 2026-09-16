# Changelog

## Excel VBA edition — recurring schedules

- Added daily and weekly recurring tasks driven by Excel `Application.OnTime`.
- Added strict missed-occurrence handling: a planned time that has passed is recorded as “已错过”, never sent late, and the next occurrence is calculated automatically.
- Added persisted task status, next/last planned time, duplicate-occurrence protection, and no automatic retry after failed or uncertain sends.
- Added confirmation before a task is enabled; scheduled occurrences send without a later prompt.
- Added automatic task pausing when the referenced template structure changes.
- Replaced the long parameter list with a six-row Forms scrollbar viewport.
- Changed parameter discovery so every parameter appears only when the selected template uses it, including `batch_date`.
- Kept `batch_date` as the occurrence date minus one day.

## Excel VBA edition — initial release

- Replaced the Java and PowerShell Outlook bridge with a pure Excel VBA workflow.
- Added editable recipient, CC, subject, and body templates.
- Added Outlook display-name address support and semicolon-separated recipient lists.
- Added manual variables and date expressions.
- Added editable-message mode and confirmation-gated direct sending.
