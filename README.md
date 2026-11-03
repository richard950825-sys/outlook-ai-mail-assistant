# Outlook AI 邮件助手（Outlook 2016 原生 COM 加载项）

在 **Outlook 界面内**调用 DeepSeek（或任意 OpenAI 兼容接口），完成邮件阅读/撰写时的 AI 处理：

- **阅读**：选中文字 → 右键（**Office 原生右键菜单**）→ `AI 摘要（选中文字）` / `AI 翻译（选中文字）`
- **撰写**：选中文字 → 右键（**Office 原生右键菜单**）→ `AI 润色成中文（选中文字）` / `AI 润色成英文（选中文字）`
- **阅读窗格内嵌回复（inline reply）**：同样支持右键润色 / 替换选中文字（阅读窗格里的回复编辑器）
- **邮件列表**：右键 → AI 摘要 / AI 翻译 / AI 邮件助手
- **独立 AI 助手窗口**：摘要、翻译、润色（可一键替换选中文字）、自定义模板写邮件、模板管理、设置
- API Key 使用 Windows DPAPI（当前用户）加密保存在本机

> 技术路线：**原生 COM 加载项**（C# / .NET Framework 4.7.2），不依赖 VSTO 运行时和 Office PIA。
> 右键菜单通过 Outlook Ribbon XML 的 `<contextMenus>` 实现（`ContextMenuText` + `ContextMenuReadOnlyMail*`），是 Office 渲染的真原生菜单项，不替换、不拦截系统菜单。

## 适用环境

| 项 | 说明 |
|---|---|
| 操作系统 | Windows 10 / 11 x64 |
| Outlook | **Microsoft Outlook 2016 x64**（开发验证版本 16.0.4266.1001） |
| .NET | .NET Framework 4.7.2 或更高（Windows 10/11 一般自带 4.8） |
| 权限 | 按用户安装（HKCU），**不需要管理员权限** |
| AI 接口 | 默认 `https://api.deepseek.com/v1`、模型 `deepseek-chat`，也可填任意 OpenAI 兼容接口 |

## 一键安装（推荐给同事/普通用户）

1. 下载 [`dist/AI-Mail-Assistant-Setup-8cf60294.zip`](dist/AI-Mail-Assistant-Setup-8cf60294.zip)
2. 解压到任意目录（不要直接在压缩包里双击）
3. **完全退出 Outlook**
4. 双击 `Setup.cmd`

安装脚本会自动：

- 检查 Outlook 2016 / .NET Framework 版本
- 把插件复制到 `%LOCALAPPDATA%\OutlookAIAssistant\Addin\latest\`
- 注册 COM 组件与 Outlook 加载项（HKCU，`LoadBehavior=3`）
- 可选：把 DeepSeek API Key 用 DPAPI 加密写入本机配置（也可以装完后在「设置」里填）
- 启动 Outlook

安装后的首次使用：Outlook 功能区「AI 邮件助手」→「打开 AI 助手」→「设置」→ 填 API Key →「测试连接」→「保存」。

卸载：双击 `Uninstall.cmd`（默认保留设置；`uninstall.ps1 -RemoveSettings` 可彻底删除配置）。

> ZIP 包 SHA256：`B41CA77C0D3135F91154BE6A59B9B9EC5E3341329225EAF8D3B7F6DFF9678EFF`

## 目录结构

```
src/OutlookAIAddin/          源码（C# / net472）
  src/Connect.cs             IDTExtensibility2 + IRibbonExtensibility 入口、Ribbon 回调
  src/UI/MainForm.cs         AI 助手窗口（摘要/翻译/润色/写邮件/模板/设置）
  src/UI/WordContextMenuHook.cs 选区捕获（阅读窗格 UIA / 撰写窗口 Word Selection）
  src/Core/                  AI 客户端、提示词、结果解析
  src/Infrastructure/        设置（DPAPI）、模板存储、日志
  resources/Ribbon.xml       功能区 + 原生右键菜单定义
installer/                   一键安装包（解压版，与 dist 里的 zip 内容一致）
dist/                        一键安装包 zip
```

## 从源码构建

需要 Windows + .NET SDK（本项目用 .NET SDK 10 / VS 2026 MSBuild 构建 net472 类库，不引用 Office PIA）：

```powershell
dotnet build src/OutlookAIAddin/OutlookAIAddin.csproj -c Release
```

产物：`src/OutlookAIAddin/bin/Release/OutlookAIAddin.dll`

自己安装这一版：

1. 退出 Outlook
2. 用构建出的 DLL 覆盖 `installer/payload/OutlookAIAddin.dll`
3. 双击 `installer/Setup.cmd`

> 也可参考 `installer/install.ps1` 里的注册逻辑手动注册：
> `HKCU\Software\Classes\CLSID\{8F5D2E1A-...}`（mscoree / Assembly / CodeBase）
> 与 `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAIAddin.Connect`（LoadBehavior=3）。

## 实现要点

- **原生右键菜单**：Outlook 加载项可以在 Ribbon XML 里声明 `<contextMenus>`：
  - `ContextMenuText`：撰写窗口（可编辑 Word 正文）的右键菜单；
  - `ContextMenuReadOnlyMailText`、`ContextMenuReadOnlyMailList`、`ContextMenuReadOnlyMailTextTable`、`ContextMenuReadOnlyMailTableCell`、`ContextMenuReadOnlyMailListTable`、`ContextMenuReadOnlyMailTable`、`ContextMenuReadOnlyMailTableWhole`、`ContextMenuReadOnlyMailHyperlink`：阅读窗格 / 独立阅读窗口的只读邮件正文右键菜单。
  这样 AI 项由 Office 渲染进原生菜单，原生项（复制、同义字、翻译、查看源文件等）全部保留。
- **选区读取**：阅读窗格没有公开的 WordEditor，用 UI Automation `TextPattern` 从 `_WwG` 控件读取选中文字；独立阅读窗口和撰写窗口用 `WordEditor` 的 Word Selection（late binding，无 PIA）。
- **无管理员权限**：COM 注册全部在 HKCU（`mscoree.dll` + `Assembly` + `CodeBase`），不需要写 HKLM。

## 隐私

- API Key 用 Windows DPAPI（当前用户）加密后保存在 `%APPDATA%\OutlookAIAssistant\settings.json`，仅本机本用户可解密；
- 日志（`%APPDATA%\OutlookAIAssistant\addin.log`）不记录邮件正文和 API Key；
- 安装包不包含任何 API Key；如公司统一发 Key，可在 `installer/payload/apikey.txt` 放一行 Key（注意该文件会随包传播）。
- 邮件内容只会发送到你自己配置的 AI 接口。

## 已知限制

- 仅在 **Outlook 2016 x64（16.0.4266.1001）+ Windows 11** 验证；其他 Office 版本（2019/2021/365）未验证，理论可行但右键菜单 ID 可能需要调整。
- `deepseek-flash` 是推理模型，长文本可能把 `max_tokens` 消耗在推理上导致返回空（程序会自动加大额度重试）；建议使用 `deepseek-chat`。
- 32 位（x86）Office 未验证（本项目为 AnyCPU，理论可加载）。
- 阅读窗格内嵌回复的“替换选中文字”依赖 `Explorer.ActiveInlineResponseWordEditor`（Outlook 2013+ 提供）。

## License

[MIT](LICENSE)
