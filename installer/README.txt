AI 邮件助手（Outlook 2016）安装包
================================================

适用环境
--------
- Windows 10 / 11（x64）
- Microsoft Outlook 2016（16.0，x64；与开发机版本一致：16.0.4266.1001）
- .NET Framework 4.7.2 或更高（Windows 10/11 一般自带 4.8）
- 按用户安装：不需要管理员权限

一键安装
--------
1. 把整个文件夹解压到本机任意位置（不要直接在压缩包里双击）
2. 关闭 Outlook（重要：否则要先退出再重开才生效）
3. 双击「Setup.cmd」
4. 安装完成后会自动打开 Outlook（也可以手动打开）

首次配置 API Key
----------------
打开 Outlook → 功能区「AI 邮件助手」→「打开 AI 助手」→「设置」
  - Base URL：https://api.deepseek.com/v1
  - 模型：deepseek-chat
  - API Key：粘贴你自己的 DeepSeek Key →「测试连接」→「保存」

（如果安装时已经粘贴过 Key，这一步可以跳过。）

用法
----
- 阅读窗格 / 双击打开的邮件：选中文字 → 右键 →「AI 摘要（选中文字）」「AI 翻译（选中文字）」
- 撰写窗口：选中文字 → 右键 →「AI 润色成中文（选中文字）」「AI 润色成英文（选中文字）」
- 邮件列表：右键 →「AI 摘要」「AI 翻译」「AI 邮件助手...」
- 「写邮件」页签：选模板 + 一句话要求 → 生成邮件 → 新建邮件并填入

回滚 / 重装
-----------
- 重装：退出 Outlook 后，重新双击「Setup.cmd」即可（旧版本自动归档，不会丢配置）
- 卸载：退出 Outlook 后，双击「Uninstall.cmd」；默认保留设置（含 API Key），
  如需彻底删除配置，用命令行运行：
  powershell -NoProfile -ExecutionPolicy Bypass -File uninstall.ps1 -RemoveSettings

常见问题
--------
- 功能区没有「AI 邮件助手」：
  * 完全退出 Outlook 再打开（任务管理器确认没有 OUTLOOK.EXE）
  * 文件 → 选项 → 加载项，确认「AI 邮件助手」在列表中且已加载
  * 看日志：%APPDATA%\OutlookAIAssistant\addin.log
- AI 请求失败：检查 API Key 是否正确、公司网络能否访问 api.deepseek.com
- 插件崩溃/被禁用：文件 → 选项 → 加载项 → 管理「已禁用的项目」→ 启用

隐私
----
- API Key 使用 Windows DPAPI（当前用户）加密后保存在
  %APPDATA%\OutlookAIAssistant\settings.json，只在本机、本用户可用
- 邮件正文只会发送到你配置的 API 接口；日志不记录邮件正文和 Key
- 安装包本身不包含任何 API Key（除非管理员在 payload\apikey.txt 里放了统一 Key）

卸载
----
双击「Uninstall.cmd」（默认保留设置）；彻底删除见上文「回滚 / 重装」。
