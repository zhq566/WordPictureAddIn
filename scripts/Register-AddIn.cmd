@echo off
setlocal EnableExtensions
set "SCRIPT_DIR=%~dp0"
set "ROOT_DIR=%SCRIPT_DIR%.."
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Register-AddIn.ps1" %*
exit /b %ERRORLEVEL%
