# Tiny-toy

一些实用的小工具。

## Outlook 模板邮件工具

[打开项目](outlook-template-tool/) · [完整使用说明](outlook-template-tool/使用说明.md)

Excel VBA 工具，支持自定义邮件模板、手动变量、日期函数、收件人显示名称和抄送列表。

- 直接使用当前 Windows 用户已经登录的经典版 Outlook，不需要再次登录。
- 默认打开可编辑邮件；可选“确认后自动发送”。
- 内置 `batch_date`，每次运行时等于当前日期减 1 天。
- 不使用 Java、PowerShell、Access、第三方 DLL 或自定义 EXE。

### 快速开始

1. 下载并打开 `outlook-template-tool/Outlook邮件模板工具.xlsx`。
2. 另存为 `.xlsm`，导入 `vba/modOutlookTemplateTool.bas`。
3. 运行一次 `SetupWorkbook`，以后通过工作表按钮使用。

当前仅支持 Windows 经典版 Outlook。首次宏导入、公司签名或受信任位置要求见完整使用说明。

