# Tiny-toy

一些实用的小工具。

## Outlook 模板邮件工具

[打开项目](outlook-template-tool/) · [完整使用说明](outlook-template-tool/使用说明.md)

Excel VBA 工具，支持自定义收件人、抄送、主题和正文模板。所有参数只在当前模板使用时显示，参数区通过滚动条每页显示 6 项；`batch_date` 在邮件计划日期的基础上自动减 1 天。

- 直接使用当前 Windows 用户已经登录的经典版 Outlook，不需要再次登录。
- 默认打开可编辑邮件；可选“确认后自动发送”。
- 支持每天或每周循环自动发送，使用 Excel `Application.OnTime`，不使用 Windows 任务计划程序。
- 计划时间一旦过去即标记“已错过”，不会补发或自动重试，并自动计算下一次发送时间。
- 不使用 Java、PowerShell、Access、第三方 DLL 或自定义 EXE。

### 快速开始

1. 下载并打开 `outlook-template-tool/Outlook邮件模板工具.xlsx`。
2. 另存为 `.xlsm`，导入 `vba/modOutlookTemplateTool.bas`。
3. 运行一次 `SetupWorkbook`，以后通过工作表按钮使用。

循环任务要求 Excel 保持打开、Windows 用户保持登录且电脑不能休眠。当前仅支持 Windows 经典版 Outlook；首次宏导入、公司签名或受信任位置要求见完整使用说明。
