@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

where py >nul 2>&1
if errorlevel 1 (
  echo Не знайдено Python Launcher.
  echo Встановіть Python 3.13 з https://www.python.org/downloads/windows/
  echo Під час встановлення увімкніть "Add python.exe to PATH".
  pause
  exit /b 1
)

py -3.13 --version >nul 2>&1
if errorlevel 1 (
  echo Не знайдено Python 3.13. Встановіть його з python.org.
  pause
  exit /b 1
)

echo Створення окремого Python-середовища...
py -3.13 -m venv .venv
if errorlevel 1 goto :error

echo Встановлення залежностей NIT2...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :error
".venv\Scripts\python.exe" -m pip install -r requirements_web.txt
if errorlevel 1 goto :error

if not exist nit_config.local.json copy /Y nit_config.example.json nit_config.local.json >nul

echo.
echo Готово. Перевірте school_id у файлі nit_config.local.json,
echo а потім запускайте "Запустити NIT Windows.bat".
echo Для постійної роботи можна окремо увімкнути автозапуск.
pause
exit /b 0

:error
echo.
echo Встановлення завершилося з помилкою.
pause
exit /b 1
