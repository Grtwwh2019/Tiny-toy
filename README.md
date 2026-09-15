# Tiny-toy

一些实用的小工具。

## Outlook 模板邮件工具

[打开项目](outlook-template-tool/) · [完整使用说明](outlook-template-tool/使用说明.md)

Windows Java 桌面工具，支持自定义邮件模板、手动变量和日期函数。可打开 Outlook 编辑邮件，也可在每次弹框确认后，通过公司 Microsoft 365 账号直接发送。

### 快速开始

1. 点击仓库右上方 **Code → Download ZIP** 并解压。
2. 进入 `outlook-template-tool`，双击 `Start.bat`（需要 Java 8 或以上）。
3. 默认模式会打开 Windows 默认邮件应用，请先将 Outlook 设为 MAILTO 默认应用。
4. 直接发送需配置租户 ID、应用 ID 并完成微软登录授权，步骤见使用说明。

项目包含可运行 JAR、源码、编译与测试脚本。72 项模拟自动化检查及发送确认流程测试已通过；Windows Outlook 唤起、DPAPI 和公司账号实际发送仍需在目标环境验证。
