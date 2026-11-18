@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Unregister-AddIn.ps1" %*
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
    echo.
    pause
)
exit /b %EXITCODE%