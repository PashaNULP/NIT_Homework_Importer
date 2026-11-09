@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

if not exist "NIT2-CA.cer" (
  echo Поруч із цим файлом немає NIT2-CA.cer.
  echo Скопіюйте обидва файли із серверного комп'ютера NIT2.
  pause
  exit /b 1
)

certutil.exe -user -addstore Root "NIT2-CA.cer"
if errorlevel 1 goto :error
echo.
echo Сертифікату NIT2 тепер довіряє поточний користувач Windows.
echo Закрийте й знову відкрийте браузер.
pause
exit /b 0

:error
echo.
echo Не вдалося додати сертифікат NIT2.
pause
exit /b 1
