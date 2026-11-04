@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Спочатку запустіть "Встановити NIT Windows.bat".
  pause
  exit /b 1
)

set "NIT_HOST=0.0.0.0"
set "NIT_PORT=8080"
set "NIT_SHOW_BROWSER=0"

echo NIT2 запускається для локальної мережі школи.
echo На цьому комп'ютері: http://localhost:8080
echo На інших комп'ютерах: http://IP_ЦЬОГО_КОМП'ЮТЕРА:8080
echo.
echo IPv4-адреси цього комп'ютера:
ipconfig | findstr /R /C:"IPv4"
echo.
echo Не закривайте це вікно, поки NIT2 має працювати.
echo Для зупинки натисніть Ctrl+C.
echo.

".venv\Scripts\python.exe" nit_web.py
echo.
echo NIT2 зупинено.
pause
