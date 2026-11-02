#requires -Version 5.1
<#
    AI 邮件助手 —— 一键安装脚本（Outlook 2016 / Windows 10-11 / 按用户安装，无需管理员）

    用法：
      双击「安装.cmd」；或
      powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 [-ApiKey "sk-..."] [-NonInteractive] [-NoStart] [-DryRun]

    说明：
      - 插件安装到 %LOCALAPPDATA%\OutlookAIAssistant\Addin\latest\，注册在 HKCU（当前用户），不写系统级注册表
      - API Key 用 Windows DPAPI（当前用户）加密后存 %APPDATA%\OutlookAIAssistant\settings.json
      - 安装包内不含任何 API Key；如公司统一发放 Key，可放 payload\apikey.txt（一行一个 Key），安装时自动读取
#>
[CmdletBinding()]
param(
    [string]$ApiKey = "",
    [switch]$NonInteractive,
    [switch]$NoStart,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$Clsid        = '{8F5D2E1A-4C6B-4A87-9E3D-1B2C3D4E5F60}'
$ProgId       = 'OutlookAIAddin.Connect'
$AsmName      = 'OutlookAIAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
$DefaultBaseUrl = 'https://api.deepseek.com/v1'
$DefaultModel   = 'deepseek-chat'

$AppDataDir  = Join-Path $env:APPDATA 'OutlookAIAssistant'
$AddinRoot   = Join-Path $env:LOCALAPPDATA 'OutlookAIAssistant\Addin'
$LatestDir   = Join-Path $AddinRoot 'latest'
$HistoryDir  = Join-Path $AddinRoot 'history'
$TargetDll   = Join-Path $LatestDir 'OutlookAIAddin.dll'
$SettingsFile = Join-Path $AppDataDir 'settings.json'

$Classes      = 'HKCU:\Software\Classes'
$AddinsKey    = 'HKCU:\Software\Microsoft\Office\Outlook\Addins'
$AddinsKeyOld = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins'

function Step($t) { Write-Host "  $t" }
function Ok($t)   { Write-Host "  $t" -ForegroundColor Green }
function Warn($t) { Write-Host "  $t" -ForegroundColor Yellow }
function Fail($t) { Write-Host "  $t" -ForegroundColor Red }

Write-Host '====================================================' -ForegroundColor Cyan
Write-Host '   AI 邮件助手（Outlook 2016）一键安装' -ForegroundColor Cyan
Write-Host '====================================================' -ForegroundColor Cyan
if ($DryRun) { Warn '（DryRun 模式：只做检查，不做任何修改）' }

# ---------------------------------------------------------------- 1) 环境检查
Write-Host ''
Write-Host '1) 检查环境' -ForegroundColor Cyan

if ([Environment]::Is64BitOperatingSystem) { Ok 'Windows x64' }
else { Warn 'Windows 为 32 位；本安装包按 Outlook 2016 x64 验证过，32 位环境未验证' }

$officeKey = 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook'
if (-not (Test-Path $officeKey)) {
    $officeKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Outlook'
}
if (Test-Path $officeKey) {
    Ok '已检测到 Outlook 2016'
    try {
        $installRoot = (Get-ItemProperty $officeKey -ErrorAction Stop).'InstallRoot'
        if ($installRoot) { Step ("Office 安装目录: " + $installRoot) }
    } catch { }
} else {
    Warn '未检测到 Outlook 2016 注册表项。如果 Outlook 版本不同，装好后可能无法加载。'
}

$ndp = 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
$release = 0
try { $release = [int](Get-ItemProperty -Path $ndp -Name Release -ErrorAction Stop).Release } catch { }
if ($release -ge 461808) {
    Ok (".NET Framework 4.7.2+ 已安装（Release={0}）" -f $release)
} else {
    Fail '.NET Framework 低于 4.7.2，无法运行本插件。Windows 10/11 一般自带 4.8；请先安装 .NET Framework 4.8 后重试。'
    exit 1
}

$srcDll = Join-Path $PSScriptRoot 'payload\OutlookAIAddin.dll'
if (-not (Test-Path $srcDll)) {
    # 自解压单文件（IExpress）会把文件平铺到临时目录，没有 payload 子目录
    $srcDll = Join-Path $PSScriptRoot 'OutlookAIAddin.dll'
}
if (-not (Test-Path $srcDll)) {
    Fail ("安装包不完整：找不到 " + $srcDll)
    exit 1
}
$newHash = (Get-FileHash $srcDll -Algorithm SHA256).Hash.Substring(0, 8).ToLower()
Step ("待安装插件版本: " + $newHash)

# ---------------------------------------------------------------- 2) Outlook 运行状态
Write-Host ''
Write-Host '2) 检查 Outlook 是否在运行' -ForegroundColor Cyan
$outlook = Get-Process OUTLOOK -ErrorAction SilentlyContinue
$outlookRunning = ($null -ne $outlook)
if ($outlookRunning) {
    Warn 'Outlook 正在运行：安装可以继续，但需要重启 Outlook 后新插件才会生效。'
} else {
    Ok 'Outlook 未运行'
}

if ($DryRun) {
    Write-Host ''
    Write-Host 'DryRun 计划执行的操作：' -ForegroundColor Cyan
    Step ("复制 " + $srcDll)
    Step ("      -> " + $TargetDll)
    Step ("注册 COM: HKCU\Software\Classes\CLSID\" + $Clsid)
    Step ("登记加载项: " + $AddinsKey + "\" + $ProgId + " (LoadBehavior=3)")
    if ($ApiKey) { Step '写入 API Key（DPAPI 加密）到 settings.json' }
    else { Step '不写入 API Key（稍后在插件“设置”里填）' }
    if (-not $NoStart) { Step '启动 Outlook' }
    Write-Host ''
    Ok 'DryRun 完成，未做任何修改。'
    exit 0
}

# ---------------------------------------------------------------- 3) 安装 DLL（latest + 旧版本归档）
Write-Host ''
Write-Host '3) 安装插件文件' -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path $LatestDir | Out-Null
New-Item -ItemType Directory -Force -Path $HistoryDir | Out-Null

$needCopy = $true
if (Test-Path $TargetDll) {
    $oldHash = (Get-FileHash $TargetDll -Algorithm SHA256).Hash.Substring(0, 8).ToLower()
    if ($oldHash -eq $newHash) {
        $needCopy = $false
        Step '相同版本已存在，跳过复制'
    } else {
        $archiveDir = Join-Path $HistoryDir ('bin-' + $oldHash)
        New-Item -ItemType Directory -Force -Path $archiveDir | Out-Null
        try {
            Copy-Item -Path $TargetDll -Destination (Join-Path $archiveDir 'OutlookAIAddin.dll') -Force
            Step ("旧版本已归档到 history\bin-" + $oldHash)
        } catch { Warn '旧版本归档失败（忽略）' }
    }
}

if ($needCopy) {
    try {
        Copy-Item -Path $srcDll -Destination $TargetDll -Force
        Ok ("已复制到 " + $TargetDll)
    } catch {
        Fail ('复制 DLL 失败：' + $_.Exception.Message)
        Fail '请先完全退出 Outlook，然后重新双击安装。'
        exit 1
    }
}

$info = @(
    ("build: " + $newHash),
    ("installed: " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
    ("source: " + $srcDll)
) -join "`r`n"
Set-Content -Path (Join-Path $LatestDir 'build-info.txt') -Value $info -Encoding UTF8

# ---------------------------------------------------------------- 4) 注册 COM + 加载项（HKCU）
Write-Host ''
Write-Host '4) 注册到当前用户（HKCU）' -ForegroundColor Cyan

New-Item -Force -Path "$Classes\$ProgId" | Out-Null
Set-ItemProperty -Path "$Classes\$ProgId" -Name '(default)' -Value 'Outlook AI Mail Assistant'
New-Item -Force -Path "$Classes\$ProgId\CLSID" | Out-Null
Set-ItemProperty -Path "$Classes\$ProgId\CLSID" -Name '(default)' -Value $Clsid
New-Item -Force -Path "$Classes\CLSID\$Clsid" | Out-Null
Set-ItemProperty -Path "$Classes\CLSID\$Clsid" -Name '(default)' -Value $ProgId

$inproc = "$Classes\CLSID\$Clsid\InprocServer32"
New-Item -Force -Path $inproc | Out-Null
Set-ItemProperty -Path $inproc -Name '(default)'      -Value 'mscoree.dll'
Set-ItemProperty -Path $inproc -Name 'ThreadingModel' -Value 'Both'
Set-ItemProperty -Path $inproc -Name 'Class'          -Value 'OutlookAIAddin.Connect'
Set-ItemProperty -Path $inproc -Name 'Assembly'       -Value $AsmName
Set-ItemProperty -Path $inproc -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty -Path $inproc -Name 'CodeBase'       -Value ('file:///' + ($TargetDll -replace '\\', '/'))
Ok '已注册 COM 组件'

New-Item -Force -Path "$AddinsKey\$ProgId" | Out-Null
Set-ItemProperty -Path "$AddinsKey\$ProgId" -Name 'FriendlyName' -Value 'AI 邮件助手（Outlook AI Assistant）'
Set-ItemProperty -Path "$AddinsKey\$ProgId" -Name 'Description'  -Value '摘要、翻译、润色选中文字、一句话写邮件（DeepSeek）'
New-ItemProperty -Path "$AddinsKey\$ProgId" -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path "$AddinsKey\$ProgId" -Name 'CommandLineSafe' -Value 0 -PropertyType DWord -Force | Out-Null
Ok '已登记 Outlook 加载项（LoadBehavior=3）'

if (Test-Path "$AddinsKeyOld\$ProgId") {
    Remove-Item -Path "$AddinsKeyOld\$ProgId" -Recurse -Force -ErrorAction SilentlyContinue
    Step '已清理旧位置的登记项'
}

# ---------------------------------------------------------------- 5) API Key 配置（可选）
Write-Host ''
Write-Host '5) 配置 API Key（可跳过）' -ForegroundColor Cyan

if (-not $ApiKey) {
    $keyFile = Join-Path $PSScriptRoot 'payload\apikey.txt'
    if (-not (Test-Path $keyFile)) { $keyFile = Join-Path $PSScriptRoot 'apikey.txt' }
    if (Test-Path $keyFile) {
        $ApiKey = (Get-Content $keyFile -Raw -Encoding UTF8).Trim()
        if ($ApiKey) { Step '已从 payload\apikey.txt 读取 API Key' }
    }
}

if (-not $ApiKey -and -not $NonInteractive) {
    Write-Host '  现在可以粘贴 DeepSeek API Key（直接回车 = 跳过，稍后在插件「设置」里填）：' -ForegroundColor Yellow
    $secure = Read-Host -AsSecureString '  API Key'
    if ($secure.Length -gt 0) {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { $ApiKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    }
}

$apiKeyEnc = ''
if ($ApiKey) {
    try {
        Add-Type -AssemblyName System.Security
        $entropy = [Text.Encoding]::UTF8.GetBytes('OutlookAIAddin.Settings.v1')
        $enc = [Security.Cryptography.ProtectedData]::Protect(
            [Text.Encoding]::UTF8.GetBytes($ApiKey.Trim()), $entropy, 'CurrentUser')
        $apiKeyEnc = [Convert]::ToBase64String($enc)
        Ok '已加密 API Key'
    } catch {
        Warn ('API Key 加密失败：' + $_.Exception.Message + '（稍后可在插件里填）')
        $apiKeyEnc = ''
    }
}

New-Item -ItemType Directory -Force -Path $AppDataDir | Out-Null
if (Test-Path $SettingsFile) {
    if ($apiKeyEnc) {
        try {
            $existing = Get-Content $SettingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $existing.apiKeyEnc = $apiKeyEnc
            ($existing | ConvertTo-Json -Compress) | Set-Content -Path $SettingsFile -Encoding UTF8
            Ok '已更新现有配置中的 API Key（其他设置不动）'
        } catch {
            Warn ('更新现有配置失败：' + $_.Exception.Message)
        }
    } else {
        Step '已有配置文件，保持不变'
    }
} else {
    $settings = [ordered]@{
        baseUrl          = $DefaultBaseUrl
        model            = $DefaultModel
        apiKeyEnc        = $apiKeyEnc
        summaryLanguage  = 'zh'
        translateTarget  = 'auto'
        temperature      = 0.3
    }
    ($settings | ConvertTo-Json -Compress) | Set-Content -Path $SettingsFile -Encoding UTF8
    if ($apiKeyEnc) { Ok '已写入默认配置 + API Key' }
    else { Step '已写入默认配置（API Key 为空，稍后在插件里填）' }
}

# ---------------------------------------------------------------- 6) 启动 Outlook / 提示
Write-Host ''
Write-Host '6) 完成' -ForegroundColor Cyan
if ($outlookRunning) {
    Warn 'Outlook 正在运行：请退出 Outlook 再重新打开，插件才会加载。'
} elseif (-not $NoStart) {
    $outlookExe = $null
    try {
        $appPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
        if (Test-Path $appPath) { $outlookExe = (Get-ItemProperty $appPath).'(default)' }
    } catch { }
    if (-not $outlookExe) { $outlookExe = 'C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE' }
    if (Test-Path $outlookExe) {
        try { Start-Process $outlookExe; Ok '已启动 Outlook' }
        catch { Warn '请手动打开 Outlook' }
    } else {
        Warn '没有自动找到 Outlook.exe，请手动打开 Outlook'
    }
}

Write-Host ''
Ok '安装完成！'
Write-Host '接下来：' -ForegroundColor Cyan
Write-Host '  1. 打开 Outlook（如果刚启动没生效，退出再打开）'
Write-Host '  2. 功能区出现「AI 邮件助手」选项卡'
Write-Host '  3. 「打开 AI 助手」→ 设置 → 填 DeepSeek API Key → 测试连接 → 保存'
Write-Host ''
Write-Host '功能：'
Write-Host '  - 阅读窗格/独立阅读窗口：选中文字 → 右键 → AI 摘要 / AI 翻译（原生菜单）'
Write-Host '  - 撰写窗口：选中文字 → 右键 → AI 润色成中文 / AI 润色成英文（原生菜单）'
Write-Host '  - 邮件列表右键：AI 摘要 / AI 翻译'
Write-Host ''
Write-Host ("配置目录: " + $AppDataDir)
Write-Host ("日志文件: " + (Join-Path $AppDataDir 'addin.log'))
Write-Host ''
if (-not $NonInteractive) {
    Read-Host '按回车键退出' | Out-Null
}
