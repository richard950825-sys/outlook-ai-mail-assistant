@echo off
chcp 65001 >nul
title AI Mail Assistant - Setup
setlocal enabledelayedexpansion
set EXTRA=
if "%AIADDIN_DRYRUN%"=="1" set EXTRA=!EXTRA! -DryRun
if "%AIADDIN_NONINTERACTIVE%"=="1" set EXTRA=!EXTRA! -NonInteractive -NoStart
echo.
echo Installing AI Mail Assistant for Outlook 2016 ...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" !EXTRA!
set RC=%ERRORLEVEL%
echo.
if not "%RC%"=="0" (
  echo [ERROR] Installation failed, exit code = %RC%
) else (
  echo [OK] Installation finished.
)
echo.
if not "%AIADDIN_NONINTERACTIVE%"=="1" pause
exit /b %RC%
