@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\NIT2-Windows.ps1" -Action start
if errorlevel 1 goto :error
echo.
echo Вікно можна закрити: NIT2 продовжить працювати у фоні.
pause
exit /b 0

:error
echo.
echo Не вдалося запустити NIT2.
pause
exit /b 1
