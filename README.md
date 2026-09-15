# Tiny-toy

一些实用的小工具。

## Outlook 模板邮件工具

[打开项目](outlook-template-tool/) · [完整使用说明](outlook-template-tool/使用说明.md)

Windows Java 桌面工具，支持自定义邮件模板、手动变量、日期函数、收件人显示名称和抄送列表。

- 经典 Outlook：直接使用桌面 Outlook 当前账号打开或发送邮件，无需程序内再次登录。
- 新版 Outlook：通过 Windows 默认邮件应用打开编辑窗口。
- 直接发送前始终显示完整确认框。

### 快速开始

1. 点击仓库右上方 **Code → Download ZIP** 并解压。
2. 进入 `outlook-template-tool`，双击 `Start.bat`（需要 Java 8 或以上）。
3. 具体版本兼容、地址格式与升级说明见完整使用说明。

项目包含可运行 JAR、源码、编译与测试脚本。自动化检查及发送确认流程测试使用模拟 Outlook；真实 Outlook 行为需在目标 Windows 电脑验证。
