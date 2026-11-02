@echo off
chcp 65001 >nul
title AI Mail Assistant - Uninstall
echo.
echo Uninstalling AI Mail Assistant ...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
set RC=%ERRORLEVEL%
echo.
if not "%RC%"=="0" (
  echo [ERROR] Uninstall failed, exit code = %RC%
) else (
  echo [OK] Uninstall finished.
)
echo.
if not "%AIADDIN_NONINTERACTIVE%"=="1" pause
exit /b %RC%
