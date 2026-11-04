@echo off
rem Сборка SnackReport.exe на Windows. Нужен Python 3.10+ (python.org), запуск из папки проекта.
chcp 65001 >nul
setlocal
cd /d "%~dp0"

if not exist .venv (
    python -m venv .venv || goto :error
)
call .venv\Scripts\activate.bat || goto :error
python -m pip install -q -r requirements-dev.txt || goto :error

python -m pytest -q tests || goto :error

if not exist config\template.xlsx python tools\make_template.py || goto :error

rem Шаблон зашивается внутрь exe
pyinstaller --noconfirm --clean --onefile --noconsole --name SnackReport ^
    --paths src --add-data "config\template.xlsx;." run_snack_report.py || goto :error

echo.
echo Готово: dist\SnackReport.exe
goto :eof

:error
echo Сборка не удалась.
exit /b 1
