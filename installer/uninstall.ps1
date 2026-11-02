#requires -Version 5.1
<#
    AI 邮件助手 —— 卸载脚本（按用户删除，无需管理员）
    用法：
      powershell -NoProfile -ExecutionPolicy Bypass -File uninstall.ps1 [-RemoveSettings] [-NonInteractive]
    -RemoveSettings：同时删除 %APPDATA%\OutlookAIAssistant 里的设置、模板、日志、API Key
#>
[CmdletBinding()]
param(
    [switch]$RemoveSettings,
    [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'
$Clsid = '{8F5D2E1A-4C6B-4A87-9E3D-1B2C3D4E5F60}'
$ProgId = 'OutlookAIAddin.Connect'
$AddinRoot = Join-Path $env:LOCALAPPDATA 'OutlookAIAssistant\Addin'
$AppDataDir = Join-Path $env:APPDATA 'OutlookAIAssistant'
$Classes = 'HKCU:\Software\Classes'
$AddinsKey = 'HKCU:\Software\Microsoft\Office\Outlook\Addins'
$AddinsKeyOld = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins'

Write-Host '== 卸载 AI 邮件助手 ==' -ForegroundColor Cyan
if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
    Write-Host '  提示：Outlook 正在运行，卸载后请重启 Outlook。' -ForegroundColor Yellow
}

foreach ($key in @(
        "$AddinsKey\$ProgId",
        "$AddinsKeyOld\$ProgId",
        "$Classes\$ProgId",
        "$Classes\CLSID\$Clsid")) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  已删除注册表项 $key"
    }
}

if (Test-Path $AddinRoot) {
    try {
        Remove-Item -Path $AddinRoot -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  已删除 $AddinRoot"
    } catch {
        Write-Host "  删除 $AddinRoot 失败（如果 Outlook 正在运行，请退出后重试）" -ForegroundColor Yellow
    }
}

if ($RemoveSettings) {
    if (Test-Path $AppDataDir) {
        Remove-Item -Path $AppDataDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  已删除配置目录 $AppDataDir"
    }
} else {
    Write-Host "  配置（含 API Key）保留在 $AppDataDir"
    Write-Host '  如需一并删除，请重新运行并加 -RemoveSettings'
}

Write-Host ''
Write-Host '卸载完成。请完全退出并重新打开 Outlook。' -ForegroundColor Green
if (-not $NonInteractive) { Read-Host '按回车键退出' | Out-Null }
