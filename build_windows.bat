@echo off
REM Builds ExcelPrintFormatter.exe for Windows. Run this on a Windows PC with Python 3.10+ installed.
cd /d "%~dp0"
python -m venv .build-venv
call .build-venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean --windowed --onefile --name ExcelPrintFormatter app.py
echo.
echo Done. The app is at: %cd%\dist\ExcelPrintFormatter.exe
pause
