@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\NIT2-Windows.ps1" -Action setup-https
if errorlevel 1 goto :error
pause
exit /b 0

:error
echo.
echo Не вдалося налаштувати HTTPS.
pause
exit /b 1
