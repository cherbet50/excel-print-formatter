#!/usr/bin/env python3
"""Reformat counter-book (สมุดคุมหน้าเคาท์เตอร์) Excel files for landscape A4 printing.

Layout rules were taken from the hand-edited files in the ``adjust`` folder:

* Only layout and styling change. Every cell value is compared before and after
  saving; if anything differs, the output is deleted and the file is reported as failed.
* Font TH SarabunPSK, data rows bold and centred, text wrapped.
* Column widths follow the ``adjust`` files. Column M (ผู้รับรายงาน) is 25 wide.
* Record rows (rows with ลำดับ / เลขงาน) are 160 pt high so there is room to sign.
  Continuation rows and the total row are 37.5 pt. Any row grows if its text needs more room.
* Paper A4 landscape, narrow margins, fixed scale so the table fits the page width.
* Manual page breaks fall between records, so a record is never split across pages.
  Header rows 4-6 repeat on every page.

Command line:
    python3 converter.py raw/*.xlsx -o adjust
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break, RowBreak
from openpyxl.worksheet.properties import PageSetupProperties

FONT_NAME = "TH SarabunPSK"
HEADER_LABEL = "ลำดับ"          # text in column A of the table header row
SIGN_HEADER = "ผู้รับรายงาน"      # text in column M of the table header row
TABLE_COLUMNS = 14              # A..N
LAST_COLUMN_LETTER = "N"
HIDDEN_COLUMNS = ("O",)
SOURCE_PREFIX_RE = re.compile(r"^ยังไม่แก้\s*")

# Column widths in Excel character units, learned from the adjust folder.
COLUMN_WIDTHS = {
    "A": 5.3,   # ลำดับ
    "B": 10.3,  # เลขงาน
    "C": 11.5,  # สถานที่
    "D": 15.0,  # ชื่อผู้ต้องหา
    "E": 13.5,  # เมท (item description)
    "F": 6.3,
    "G": 3.9,
    "H": 5.6,
    "I": 3.9,
    "J": 4.3,
    "K": 9.5,   # หมายเหตุ
    "L": 15.5,  # ได้รับของกลางคืนแล้ว
    "M": 25.0,  # ผู้รับรายงาน
    "N": 13.0,
}

RECORD_ROW_HEIGHT = 160.0
SUB_ROW_HEIGHT = 37.5
HEADER_MIN_HEIGHT = 21.75
TITLE_SIZE = 16
LABEL_SIZE = 14
HEADER_SIZE = 14
DATA_SIZE = 12

A4_LANDSCAPE_PT = (841.89, 595.28)
MARGIN_LR_IN = 0.25
MARGIN_TB_IN = 0.4
PT_PER_WIDTH_UNIT = 5.25        # Calibri 11 default: about 7 px per width unit, 0.75 pt per px
TH_CHAR_PT_AT_12 = 5.0          # average glyph width of TH SarabunPSK at 12 pt (slightly conservative)
LINE_FACTOR = 1.4


class LayoutError(Exception):
    """Raised when a file does not match the expected table template."""


@dataclass
class Report:
    source: str
    output: str
    records: int = 0
    last_row: int = 0
    scale: int = 100
    pages: int = 0
    values_checked: int = 0
    warnings: list = field(default_factory=list)


def _has_value(ws, row: int, col: int) -> bool:
    return ws.cell(row, col).value not in (None, "")


def _text_width_pt(letter: str) -> float:
    return COLUMN_WIDTHS.get(letter, 8.43) * PT_PER_WIDTH_UNIT


def _estimate_row_height(ws, row: int, size: float, bold: bool) -> float:
    """Height in points needed so the longest wrapped cell in the row is not cut off."""
    lines = 1
    for col in range(1, TABLE_COLUMNS + 1):
        letter = get_column_letter(col)
        value = ws.cell(row, col).value
        if value in (None, ""):
            continue
        usable = _text_width_pt(letter) - 6
        char_pt = TH_CHAR_PT_AT_12 * size / 12 * (1.08 if bold else 1.0)
        per_line = max(1, int(usable // char_pt))
        count = sum(max(1, math.ceil(len(part) / per_line)) for part in str(value).split("\n"))
        lines = max(lines, count)
    return lines * size * LINE_FACTOR + 8


def _restyle(cell, size: float, bold: bool) -> None:
    old = cell.font
    cell.font = Font(
        name=FONT_NAME,
        size=size,
        bold=bold,
        italic=old.italic,
        underline=old.underline,
        color=old.color,
    )
    cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)


def _format_sheet(ws, report: Report) -> bool:
    """Apply the adjust-folder layout to one worksheet. Returns False if it has no table."""
    header_row = None
    for r in range(1, 16):
        if ws.cell(r, 1).value == HEADER_LABEL:
            header_row = r
            break
    if header_row is None:
        report.warnings.append(f"sheet '{ws.title}': no table header found, left unchanged")
        return False
    if ws.cell(header_row, 13).value != SIGN_HEADER:
        raise LayoutError(
            f"sheet '{ws.title}': column M header is {ws.cell(header_row, 13).value!r}, "
            f"expected {SIGN_HEADER!r}"
        )

    last_row = header_row
    for r in range(header_row + 1, ws.max_row + 1):
        if any(_has_value(ws, r, c) for c in range(1, TABLE_COLUMNS + 1)):
            last_row = r
    if last_row == header_row:
        report.warnings.append(f"sheet '{ws.title}': table has no data rows")
    first_data = header_row + 1

    # Column widths
    for letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[letter].width = width
    for letter in HIDDEN_COLUMNS:
        ws.column_dimensions[letter].hidden = True

    # Fonts and alignment: title, labels, header, data
    for r in range(1, header_row):
        for c in range(1, TABLE_COLUMNS + 1):
            cell = ws.cell(r, c)
            if _has_value(ws, r, c):
                _restyle(cell, cell.font.size or LABEL_SIZE, True)
    for c in range(1, TABLE_COLUMNS + 1):
        cell = ws.cell(header_row, c)
        if _has_value(ws, header_row, c) or cell.border.left.style:
            _restyle(cell, HEADER_SIZE, True)
    for r in range(first_data, last_row + 1):
        for c in range(1, TABLE_COLUMNS + 1):
            cell = ws.cell(r, c)
            if _has_value(ws, r, c) or cell.border.left.style:
                _restyle(cell, DATA_SIZE, True)

    # Row heights
    ws.row_dimensions[header_row].height = max(
        HEADER_MIN_HEIGHT, _estimate_row_height(ws, header_row, HEADER_SIZE, True)
    )
    groups: list[list[int]] = []
    for r in range(first_data, last_row + 1):
        starts_group = _has_value(ws, r, 1) or _has_value(ws, r, 2) or not groups
        if starts_group:
            groups.append([r])
        else:
            groups[-1].append(r)
    for group in groups:
        for r in group:
            base = RECORD_ROW_HEIGHT if _has_value(ws, r, 1) else SUB_ROW_HEIGHT
            est = _estimate_row_height(ws, r, DATA_SIZE, True)
            if est > base:
                report.warnings.append(f"row {r}: grown to {est:.0f} pt so text is not cut off")
            ws.row_dimensions[r].height = max(base, est)
    report.records = sum(1 for g in groups if _has_value(ws, g[0], 1))
    report.last_row = last_row

    # Page setup: fixed scale so the table fits the A4 width
    page_w, page_h = A4_LANDSCAPE_PT
    usable_w = page_w - 2 * MARGIN_LR_IN * 72
    usable_h = page_h - 2 * MARGIN_TB_IN * 72
    table_w = sum(COLUMN_WIDTHS[get_column_letter(c)] * PT_PER_WIDTH_UNIT
                  for c in range(1, TABLE_COLUMNS + 1))
    scale = min(100, int(usable_w / table_w * 100))
    report.scale = scale
    capacity = usable_h * 100 / scale

    def height(r: int) -> float:
        return ws.row_dimensions[r].height or 15.0

    title_rows = range(header_row - 2, header_row + 1)
    used = sum(height(r) for r in range(1, header_row + 1))
    breaks: list[int] = []
    for group in groups:
        group_h = sum(height(r) for r in group)
        if used + group_h > capacity and group[0] != first_data:
            breaks.append(group[0] - 1)
            used = sum(height(r) for r in title_rows) + group_h
        else:
            used += group_h
    ws.row_breaks = RowBreak()
    for row_id in breaks:
        ws.row_breaks.append(Break(id=row_id))
    report.pages = len(breaks) + 1

    ws.print_area = f"A1:{LAST_COLUMN_LETTER}{last_row}"
    ws.print_title_rows = f"{header_row - 2}:{header_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.scale = scale
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)
    ws.page_margins = PageMargins(
        left=MARGIN_LR_IN, right=MARGIN_LR_IN,
        top=MARGIN_TB_IN, bottom=MARGIN_TB_IN, header=0.2, footer=0.2,
    )
    ws.print_options.horizontalCentered = True
    return True


def _snapshot(path: Path):
    """Every non-empty cell value and merged range, per sheet."""
    wb = openpyxl.load_workbook(path)
    data = {}
    for ws in wb.worksheets:
        values = {
            (cell.row, cell.column): cell.value
            for row in ws.iter_rows()
            for cell in row
            if cell.value not in (None, "")
        }
        merged = sorted(str(m) for m in ws.merged_cells.ranges)
        data[ws.title] = (values, merged)
    return data


def convert_file(src, dst) -> Report:
    """Convert one workbook. Writes ``dst`` only if every value matches the source."""
    src, dst = Path(src), Path(dst)
    report = Report(source=str(src), output=str(dst))
    wb = openpyxl.load_workbook(src)
    formatted = 0
    for ws in wb.worksheets:
        if _format_sheet(ws, report):
            formatted += 1
    if formatted == 0:
        raise LayoutError("no table found in any sheet")
    dst.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dst)

    before = _snapshot(src)
    after = _snapshot(dst)
    if before.keys() != after.keys():
        dst.unlink(missing_ok=True)
        raise LayoutError("sheet names changed; output discarded")
    checked = 0
    for name in before:
        rows_a, merged_a = before[name]
        rows_b, merged_b = after[name]
        if merged_a != merged_b:
            dst.unlink(missing_ok=True)
            raise LayoutError(f"merged cells changed on '{name}'; output discarded")
        if rows_a != rows_b:
            dst.unlink(missing_ok=True)
            raise LayoutError(f"cell values changed on '{name}'; output discarded")
        checked += len(rows_a)
    report.values_checked = checked
    return report


def output_name(source) -> str:
    stem = SOURCE_PREFIX_RE.sub("", Path(source).stem).strip()
    return f"{stem}.xlsx"


def unique_path(folder: Path, name: str) -> Path:
    candidate = folder / name
    counter = 2
    while candidate.exists():
        candidate = folder / f"{Path(name).stem} ({counter}){Path(name).suffix}"
        counter += 1
    return candidate


def collect_inputs(paths) -> list[Path]:
    files: list[Path] = []
    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            files += sorted(q for q in p.glob("*.xlsx") if not q.name.startswith("~$"))
        elif p.suffix.lower() == ".xlsx" and p.exists():
            files.append(p)
    return files


def convert_many(sources, out_dir, progress=None):
    """Convert several files. One failure does not stop the others.

    Returns a list of (source, Report or None, error message or None).
    """
    out_dir = Path(out_dir)
    results = []
    total = len(sources)
    for index, src in enumerate(sources, start=1):
        try:
            target = unique_path(out_dir, output_name(src))
            report = convert_file(src, target)
            results.append((Path(src), report, None))
        except Exception as exc:  # report every failure, keep going
            results.append((Path(src), None, f"{type(exc).__name__}: {exc}"))
        if progress:
            progress(index, total, Path(src).name)
    return results


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Format counter-book Excel files for A4 landscape.")
    parser.add_argument("inputs", nargs="+", help=".xlsx files or folders containing them")
    parser.add_argument("-o", "--output", help="output folder (default: 'converted' next to the first input)")
    args = parser.parse_args(argv)

    files = collect_inputs(args.inputs)
    if not files:
        print("No .xlsx files found.", file=sys.stderr)
        return 1
    out_dir = Path(args.output) if args.output else files[0].parent / "converted"
    results = convert_many(files, out_dir)
    failed = 0
    for src, report, error in results:
        if error:
            failed += 1
            print(f"FAILED  {src.name}: {error}")
        else:
            print(f"OK      {src.name} -> {Path(report.output).name}  "
                  f"records={report.records} pages={report.pages} scale={report.scale}% "
                  f"values checked={report.values_checked}")
            for w in report.warnings:
                print(f"        note: {w}")
    print(f"\nDone: {len(results) - failed} converted, {failed} failed. Output: {out_dir}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
