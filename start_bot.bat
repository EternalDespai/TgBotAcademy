@echo off
chcp 65001 >nul
setlocal EnableExtensions

cd /d "%~dp0"

echo ==========================================
echo Настройка и запуск бота (строго Python 3.12)
echo ==========================================

echo Доступные Python (py -0p):
py -0p
echo.

py -3.12 --version >nul 2>&1
if errorlevel 1 goto NO_PY312

echo Создаю виртуальное окружение (Python 3.12)...
py -3.12 -m venv venv

echo Версия Python в venv:
venv\Scripts\python.exe -V

echo Устанавливаю библиотеки...
venv\Scripts\python.exe -m pip install -r requirements.txt

echo.
echo ==========================================
echo Запуск бота! (Нажмите Ctrl+C для выхода)
echo ==========================================
venv\Scripts\python.exe bot_app\main.py

pause
exit /b

:NO_PY312
echo ОШИБКА: Python 3.12 не найден через "py -3.12".
echo Установи Python 3.12 так, чтобы он появился в "py -0p" (Python312).
pause
exit /b
