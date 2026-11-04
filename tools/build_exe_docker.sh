#!/bin/sh
# Сборка SnackReport.exe на Linux через Docker (Wine + Windows Python).
# Результат: dist/SnackReport.exe
set -e
cd "$(dirname "$0")/.."
docker run --rm -v "$PWD":/src -w /src tobix/pywine:3.12 sh -c '
  wine python -m pip install -q -r requirements.txt pyinstaller &&
  wine pyinstaller --noconfirm --clean --onefile --noconsole --name SnackReport \
      --paths src --add-data "config/template.xlsx;." run_snack_report.py
'
ls -la dist/SnackReport.exe
