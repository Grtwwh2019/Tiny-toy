# Outlook Template Tool — Excel VBA Edition

This tool creates or sends one template-based message through the classic Outlook profile already signed in on the current Windows desktop. It runs entirely in Excel VBA: no Java, PowerShell, Access, Windows Task Scheduler, third-party DLL, custom EXE, or second Microsoft 365 login is required.

The package contains:

- `Outlook邮件模板工具.xlsx`: formatted workbook template.
- `vba/modOutlookTemplateTool.bas`: VBA module to import once.
- `使用说明.md`: complete Chinese installation and operating guide.
- `CHANGELOG.md`: release notes.

Features include editable To/CC/subject/body templates, Outlook display-name addresses, manual and calculated date parameters, a six-row scrollable parameter area, editable-message mode, confirmation-gated manual sending, and daily or weekly recurring sending through Excel `Application.OnTime`.

After importing the module, save the workbook as `.xlsm` and run `SetupWorkbook` once. Recurring tasks work only while Excel remains open, the Windows user remains signed in, and the computer remains awake. A task that is already past its planned time is marked missed; it is never sent late or retried automatically.

Classic Outlook for Windows is required. New Outlook does not expose the legacy COM object model used by VBA.
