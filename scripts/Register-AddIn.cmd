@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "SCRIPT_DIR=%~dp0"
set "ROOT_DIR=%SCRIPT_DIR%.."
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Register-AddIn.ps1" %*
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
    echo.
    pause
)
pause
exit /b %EXITCODE%