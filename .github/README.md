# Excel Print Formatter

Converts counter-book Excel files (สมุดคุมหน้าเคาท์เตอร์) to landscape A4, following the
layout of the hand-edited files in the `adjust` folder. Many files can be converted in one go.

## Use the app
1. Open **ExcelPrintFormatter** (macOS: `dist/ExcelPrintFormatter.app`; Windows: `dist\ExcelPrintFormatter.exe`).
2. Click **Add files...** and select 3 to 5 (or more) `.xlsx` files, or **Add folder...** to take every `.xlsx` in a folder.
3. Check **Save to**. It defaults to a `converted` folder next to the first file. Click **Choose...** to change it.
4. Click **Convert all**. Each file gets a separate result in the log. One failed file does not stop the others.

Source files are never changed. Output files are new copies, and an existing file with the same name is not overwritten (a ` (2)` suffix is added).

## What the conversion does
- Cell values are copied unchanged. After saving, every value and merged cell is compared with the source; if anything differs the output is discarded and the file is reported as failed.
- Font TH SarabunPSK. Table data is bold and centred with wrapped text.
- Column widths follow the `adjust` files. Column M (ผู้รับรายงาน) is 25 wide.
- Record rows (with ลำดับ) are 160 pt high so there is room to sign. Continuation rows and the total row are 37.5 pt. A row grows only when its text would otherwise be cut off.
- Paper A4 landscape with a fixed scale so the full table width fits the page.
- Page breaks fall between records. Header rows 4-6 repeat on each page.

## Build the app yourself
- macOS: run `./build_mac.sh` in Terminal (needs Python 3.10+ with tkinter; the python.org installer includes it).
- Windows: double-click `build_windows.bat` (needs Python 3.10+ installed from python.org, with "tcl/tk" ticked).
PyInstaller cannot cross-compile, so build the macOS app on a Mac and the Windows app on Windows.

## Command line (optional)
    python3 converter.py raw/*.xlsx -o adjust
