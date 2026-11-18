@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Diagnose-AddIn.ps1" %*
echo.
pause
exit /b %ERRORLEVEL%