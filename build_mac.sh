#!/bin/bash
# Builds ExcelPrintFormatter.app for macOS. Run this on a Mac.
set -e
cd "$(dirname "$0")"
python3 -m venv .build-venv
source .build-venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean --windowed --onefile --name ExcelPrintFormatter app.py
echo ""
echo "Done. The app is at: $(pwd)/dist/ExcelPrintFormatter.app"
