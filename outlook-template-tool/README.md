# Outlook Template Tool — Excel VBA Edition

This version uses Excel VBA and the COM object model built into classic Outlook. It does not use Java, PowerShell, a third-party DLL, a custom executable, or an extra Microsoft 365 login.

The repository contains:

- `Outlook邮件模板工具.xlsx`: formatted workbook template.
- `vba/modOutlookTemplateTool.bas`: VBA module to import once.
- `使用说明.md`: installation and operating instructions.

After importing the module, save the workbook as `.xlsm` and run `SetupWorkbook` once. The workbook can then open one editable Outlook message at a time, or send it after a confirmation dialog when the optional auto-send setting is selected.

Classic Outlook for Windows is required. New Outlook does not expose the legacy COM object model used by VBA.

