"""Точка входа для PyInstaller и запуска из исходников: python run_snack_report.py"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from snack_report.main import run  # noqa: E402

if __name__ == "__main__":
    sys.exit(run())
