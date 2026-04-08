from __future__ import annotations

import argparse
from copy import copy
import json
import re
import sys
import unicodedata
import warnings
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


TARGET_DIR = Path(
    r"C:\Users\NMMIS\OneDrive - Ministry of Transportation and Communications-7280502-Taiwan Railways Administration, MOTC\文件\MMIS桌面"
)
FAULT_NOTICE_SHEET_NAME = "故障通報管理 的清單"
DATE_HEADER = "發生日期"
CAR_HEADER = "車組/車號"
NEW_HEADER = "車號"
FONT_SIZE = 12
FONT_NAME = "新細明體"
DEFAULT_FILE_TYPE = "auto"
FAULT_NOTICE_FILENAME_PATTERNS = (
    "故障通報管理",
    "未處理故障通報",
)
FAULT_NOTICE_DELETE_HEADERS = [
    "發生時間",
    "故障地點",
    "立案人員",
    "通報人員",
    "通報單位",
    "狀態",
    "配屬段別",
    "配屬段別名稱",
]

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    module="openpyxl.styles.stylesheet",
)


class FormattingError(RuntimeError):
    pass


@dataclass(frozen=True)
class FormatterStrategy:
    key: str
    filename_patterns: tuple[str, ...]
    formatter: Callable[[Path], dict[str, Any]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def mmdd_today() -> str:
    return datetime.now().strftime("%m%d")


def list_existing_files() -> list[str]:
    if not TARGET_DIR.exists():
        return []
    return sorted(item.name for item in TARGET_DIR.iterdir() if item.is_file())


def result_template(*, file_path: Path | None = None, file_type: str = DEFAULT_FILE_TYPE) -> dict[str, Any]:
    filename = file_path.name if file_path else None
    path = str(file_path) if file_path else None
    return {
        "file_found": False,
        "filename": filename,
        "file_type": file_type,
        "detected_type": None,
        "format_applied": False,
        "sorted": False,
        "column_inserted": False,
        "header_set": False,
        "value_range": None,
        "value_verification": [],
        "unparsed_count": 0,
        "saved": False,
        "reason": None,
        "existing_files": [],
        "sheet_names": [],
        "headers": [],
        "path": path,
        "used_range": None,
        "verification": [],
        "full_range_verified": False,
        "mismatch_cells": [],
        "selected_by": None,
        "deleted_headers": [],
        "header_bold_applied": False,
        "layout_applied": False,
        "autofit_applied": False,
    }


def find_used_range(sheet) -> tuple[int, int, int, int]:
    min_row = None
    min_col = None
    max_row = 0
    max_col = 0
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value not in (None, ""):
                if min_row is None or cell.row < min_row:
                    min_row = cell.row
                if min_col is None or cell.column < min_col:
                    min_col = cell.column
                if cell.row > max_row:
                    max_row = cell.row
                if cell.column > max_col:
                    max_col = cell.column
    if min_row is None or min_col is None:
        return (1, 1, sheet.max_row or 1, sheet.max_column or 1)
    return (min_row, min_col, max_row, max_col)


def apply_font_to_range(sheet, min_row: int, min_col: int, max_row: int, max_col: int) -> None:
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            new_font = copy(cell.font)
            new_font.name = FONT_NAME
            new_font.size = FONT_SIZE
            cell.font = new_font


def apply_alignment_to_range(
    sheet,
    min_row: int,
    min_col: int,
    max_row: int,
    max_col: int,
    *,
    horizontal: str,
    vertical: str,
) -> None:
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            new_alignment = copy(cell.alignment) if cell.alignment else Alignment()
            new_alignment.horizontal = horizontal
            new_alignment.vertical = vertical
            cell.alignment = new_alignment


def apply_header_row_alignment(sheet, min_col: int, max_col: int) -> None:
    apply_alignment_to_range(sheet, 1, min_col, 1, max_col, horizontal="center", vertical="center")


def display_width(value: Any) -> int:
    text = normalize_text(value)
    if not text:
        return 0
    width = 0
    for ch in text:
        width += 2 if unicodedata.east_asian_width(ch) in {"F", "W"} else 1
    return width


def autofit_columns(sheet, min_col: int, max_col: int) -> None:
    for col_idx in range(min_col, max_col + 1):
        max_width = 0
        for row_idx in range(1, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            max_width = max(max_width, display_width(value))
        adjusted_width = max(8, min(max_width + 2, 80))
        sheet.column_dimensions[get_column_letter(col_idx)].width = adjusted_width


def verify_font_rows(sheet, min_row: int, max_row: int, min_col: int) -> list[dict[str, Any]]:
    candidate_rows = [min_row, min_row + (max_row - min_row) // 2, max_row]
    ordered_rows = []
    for row_num in candidate_rows:
        if row_num not in ordered_rows:
            ordered_rows.append(row_num)
    verification: list[dict[str, Any]] = []
    for row_num in ordered_rows:
        cell = sheet.cell(row=row_num, column=min_col)
        verification.append(
            {
                "row": row_num,
                "font_name": cell.font.name,
                "font_size": cell.font.sz,
                "horizontal": cell.alignment.horizontal,
                "vertical": cell.alignment.vertical,
                "cell": cell.coordinate,
            }
        )
    return verification


def scan_range_for_font_size_mismatches(
    sheet, min_row: int, min_col: int, max_row: int, max_col: int
) -> list[dict[str, Any]]:
    mismatches: list[dict[str, Any]] = []
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if float(cell.font.sz or 0) != float(FONT_SIZE):
                mismatches.append({"cell": cell.coordinate, "font_size": cell.font.sz})
            elif normalize_text(cell.font.name) != FONT_NAME:
                mismatches.append({"cell": cell.coordinate, "font_name": cell.font.name})
    return mismatches


def ensure_car_number_column(sheet, max_row: int, result: dict[str, Any]) -> None:
    if normalize_text(sheet["C1"].value) != NEW_HEADER:
        sheet.insert_cols(3, 1)
        result["column_inserted"] = True
    else:
        result["column_inserted"] = True

    sheet["C1"] = NEW_HEADER
    result["header_set"] = normalize_text(sheet["C1"].value) == NEW_HEADER

    if max_row < 2:
        raise FormattingError("資料列為 0")

    if not any(normalize_text(sheet.cell(row=row_idx, column=2).value) for row_idx in range(2, max_row + 1)):
        raise FormattingError("找不到 B 欄")

    unparsed_count = 0
    for row_idx in range(2, max_row + 1):
        source_value = normalize_text(sheet.cell(row=row_idx, column=2).value)
        match = re.search(r"(\d+)$", source_value)
        if not match:
            sheet.cell(row=row_idx, column=3, value=None)
            unparsed_count += 1
            continue

        n = int(match.group(1))
        result_value = n // 10 if 1000 <= n < 10000 else n
        sheet.cell(row=row_idx, column=3, value=result_value)

    result["value_range"] = f"C2~C{max_row}"
    result["unparsed_count"] = unparsed_count


def apply_fault_notice_header_style(sheet, result: dict[str, Any]) -> None:
    cell = sheet["C1"]
    new_font = copy(cell.font) if cell.font else Font()
    new_font.bold = True
    cell.font = new_font
    result["header_bold_applied"] = bool(sheet["C1"].font.bold)


def delete_fault_notice_columns(sheet, result: dict[str, Any]) -> None:
    headers = [normalize_text(sheet.cell(row=1, column=col_idx).value) for col_idx in range(1, sheet.max_column + 1)]
    to_delete: list[tuple[int, str]] = []
    for header in FAULT_NOTICE_DELETE_HEADERS:
        if header in headers:
            to_delete.append((headers.index(header) + 1, header))

    deleted_headers: list[str] = []
    for column_index, header in sorted(to_delete, key=lambda item: item[0], reverse=True):
        sheet.delete_cols(column_index, 1)
        deleted_headers.append(header)

    result["deleted_headers"] = list(reversed(deleted_headers))


def verify_value_rows(sheet, min_row: int, max_row: int) -> list[dict[str, Any]]:
    if max_row < 2:
        return []
    candidate_rows = [2, 2 + (max_row - 2) // 2, max_row]
    ordered_rows = []
    for row_num in candidate_rows:
        if row_num not in ordered_rows:
            ordered_rows.append(row_num)
    verification = []
    for row_num in ordered_rows:
        source_value = normalize_text(sheet.cell(row=row_num, column=2).value)
        match = re.search(r"(\d+)$", source_value)
        parsed_n = int(match.group(1)) if match else None
        expected_value = parsed_n // 10 if parsed_n is not None and 1000 <= parsed_n < 10000 else parsed_n
        cell_value = sheet.cell(row=row_num, column=3).value
        verification.append(
            {
                "cell": f"C{row_num}",
                "source_cell": f"B{row_num}",
                "source_value": source_value,
                "parsed_n": parsed_n,
                "written_value": cell_value,
                "ok": cell_value == expected_value,
            }
        )
    return verification


def parse_date_for_sort(value: Any) -> tuple[int, str]:
    if value is None or value == "":
        return (9, "")
    if hasattr(value, "strftime"):
        return (0, value.strftime("%Y%m%d%H%M%S"))
    text = normalize_text(value)
    candidates = [
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
    ]
    for fmt in candidates:
        try:
            parsed = datetime.strptime(text, fmt)
            return (0, parsed.strftime("%Y%m%d%H%M%S"))
        except ValueError:
            continue
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) >= 8:
        return (1, digits)
    return (8, text)


def parse_car_for_sort(value: Any) -> tuple[int, list[Any], str]:
    text = normalize_text(value)
    if not text:
        return (9, [], "")
    parts: list[Any] = []
    current = ""
    current_is_digit = text[0].isdigit()
    for ch in text:
        if ch.isdigit() == current_is_digit:
            current += ch
        else:
            parts.append(int(current) if current_is_digit else current.lower())
            current = ch
            current_is_digit = ch.isdigit()
    parts.append(int(current) if current_is_digit else current.lower())
    return (0, parts, text.lower())


def format_fault_notice_excel(file_path: Path) -> dict[str, Any]:
    result = result_template(file_path=file_path, file_type="fault_notice")
    result["detected_type"] = "fault_notice"

    if not file_path.exists():
        result["reason"] = "找不到目標檔案"
        result["existing_files"] = list_existing_files()
        return result

    result["file_found"] = True

    try:
        workbook = load_workbook(file_path)
    except PermissionError:
        result["reason"] = "儲存失敗"
        return result

    result["sheet_names"] = workbook.sheetnames
    if FAULT_NOTICE_SHEET_NAME not in workbook.sheetnames:
        result["reason"] = "工作表不存在"
        workbook.close()
        return result

    sheet = workbook[FAULT_NOTICE_SHEET_NAME]
    min_row, min_col, max_row, max_col = find_used_range(sheet)
    result["used_range"] = {
        "min_row": min_row,
        "max_row": max_row,
        "min_col": min_col,
        "max_col": max_col,
    }
    header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
    headers = [normalize_text(value) for value in header_row]
    result["headers"] = headers

    try:
        date_col_idx = headers.index(DATE_HEADER) + 1
        car_col_idx = headers.index(CAR_HEADER) + 1
    except ValueError:
        result["reason"] = "欄位名稱不一致"
        workbook.close()
        return result

    apply_font_to_range(sheet, min_row, min_col, max_row, max_col)

    if max_row > min_row:
        data_rows: list[list[Any]] = []
        for row_idx in range(min_row + 1, max_row + 1):
            values = [sheet.cell(row=row_idx, column=col_idx).value for col_idx in range(min_col, max_col + 1)]
            data_rows.append(values)

        data_rows.sort(
            key=lambda row: (
                parse_date_for_sort(row[(date_col_idx - min_col)]),
                parse_car_for_sort(row[(car_col_idx - min_col)]),
            )
        )

        for row_offset, values in enumerate(data_rows, start=min_row + 1):
            for col_idx, value in enumerate(values, start=min_col):
                sheet.cell(row=row_offset, column=col_idx).value = value

    result["sorted"] = True
    ensure_car_number_column(sheet, max_row, result)
    apply_fault_notice_header_style(sheet, result)
    delete_fault_notice_columns(sheet, result)
    min_row, min_col, max_row, max_col = find_used_range(sheet)
    apply_font_to_range(sheet, min_row, min_col, max_row, max_col)
    apply_alignment_to_range(sheet, min_row, min_col, max_row, max_col, horizontal="left", vertical="top")
    apply_fault_notice_header_style(sheet, result)
    apply_header_row_alignment(sheet, min_col, max_col)
    autofit_columns(sheet, min_col, max_col)
    if "未處理故障通報" in file_path.name:
        column_width_map = {
            "車次": 5.7,
            "車組/車號": 11.6,
            "車號": 5.7,
            "發生日期": 10.8,
            "事故等級": 10.8,
            "ATP故障": 10.4,
            "故障現象": 70.7,
            "通報號": 11.6,
        }
        header_to_index = {
            normalize_text(sheet.cell(row=1, column=col_idx).value): col_idx for col_idx in range(1, sheet.max_column + 1)
        }
        for header, width in column_width_map.items():
            col_idx = header_to_index.get(header)
            if col_idx is not None:
                sheet.column_dimensions[get_column_letter(col_idx)].width = width
    result["layout_applied"] = True
    result["autofit_applied"] = True
    result["format_applied"] = True
    result["headers"] = [normalize_text(cell.value) for cell in next(sheet.iter_rows(min_row=1, max_row=1), ())]
    result["verification"] = verify_font_rows(sheet, min_row, max_row, min_col)
    result["value_verification"] = verify_value_rows(sheet, min_row, max_row)

    try:
        workbook.save(file_path)
    except PermissionError:
        result["reason"] = "儲存失敗"
        workbook.close()
        return result
    result["saved"] = True
    workbook.close()

    verify_workbook = load_workbook(file_path)
    verify_sheet = verify_workbook[FAULT_NOTICE_SHEET_NAME]
    verify_min_row, verify_min_col, verify_max_row, verify_max_col = find_used_range(verify_sheet)
    result["used_range"] = {
        "min_row": verify_min_row,
        "max_row": verify_max_row,
        "min_col": verify_min_col,
        "max_col": verify_max_col,
    }
    result["verification"] = verify_font_rows(verify_sheet, verify_min_row, verify_max_row, verify_min_col)
    result["value_verification"] = verify_value_rows(verify_sheet, verify_min_row, verify_max_row)
    result["headers"] = [
        normalize_text(cell.value) for cell in next(verify_sheet.iter_rows(min_row=1, max_row=1), ())
    ]
    result["header_bold_applied"] = bool(verify_sheet["C1"].font.bold)
    mismatches = scan_range_for_font_size_mismatches(
        verify_sheet, verify_min_row, verify_min_col, verify_max_row, verify_max_col
    )
    result["mismatch_cells"] = mismatches[:20]
    result["full_range_verified"] = len(mismatches) == 0
    verify_workbook.close()
    if mismatches:
        result["saved"] = False
        result["reason"] = "格式套用失敗"
        return result
    if any(not item["ok"] for item in result["value_verification"]):
        result["saved"] = False
        result["reason"] = "車號寫入失敗"
        return result

    return result


def detect_file_type(file_path: Path) -> str | None:
    filename = file_path.name
    if any(pattern in filename for pattern in FAULT_NOTICE_FILENAME_PATTERNS):
        return "fault_notice"
    return None


def find_latest_excel_file() -> Path | None:
    if not TARGET_DIR.exists():
        return None
    candidates = sorted(
        (
            path
            for path in TARGET_DIR.iterdir()
            if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$")
        ),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def resolve_target_file(*, file_path_arg: str | None, file_type: str) -> tuple[Path | None, str]:
    if file_path_arg:
        return Path(file_path_arg), "explicit_path"

    latest = find_latest_excel_file()
    if latest is None:
        return None, "latest_file"

    if file_type == "auto":
        return latest, "latest_file"

    # For the currently supported formatter, the newest matching Excel file is sufficient.
    if file_type == "fault_notice":
        matching = [
            path
            for path in TARGET_DIR.iterdir()
            if path.is_file()
            and path.suffix.lower() == ".xlsx"
            and not path.name.startswith("~$")
            and any(pattern in path.name for pattern in FAULT_NOTICE_FILENAME_PATTERNS)
        ]
        matching.sort(key=lambda path: path.stat().st_mtime, reverse=True)
        return (matching[0], "latest_matching_type") if matching else (latest, "latest_file")

    return latest, "latest_file"


def build_formatter_registry() -> dict[str, FormatterStrategy]:
    return {
        "fault_notice": FormatterStrategy(
            key="fault_notice",
            filename_patterns=FAULT_NOTICE_FILENAME_PATTERNS,
            formatter=format_fault_notice_excel,
        )
    }


def dispatch_formatter(file_path: Path, file_type: str) -> dict[str, Any]:
    registry = build_formatter_registry()
    detected_type = detect_file_type(file_path) if file_type == "auto" else file_type
    result = result_template(file_path=file_path, file_type=file_type)
    result["detected_type"] = detected_type

    if detected_type is None:
        result["reason"] = "無法判斷檔案類型"
        result["existing_files"] = list_existing_files()
        return result

    if detected_type not in registry:
        result["reason"] = "尚未支援的檔案類型"
        return result

    formatter_result = registry[detected_type].formatter(file_path)
    formatter_result["file_type"] = file_type
    formatter_result["detected_type"] = detected_type
    return formatter_result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Format MMIS Excel files using filename-based strategies")
    parser.add_argument("--file", dest="file_path", help="Explicit Excel file path")
    parser.add_argument(
        "--file-type",
        default=DEFAULT_FILE_TYPE,
        choices=["auto", "fault_notice"],
        help="Formatter strategy to use; default is auto-detect from filename",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    file_path, selected_by = resolve_target_file(file_path_arg=args.file_path, file_type=args.file_type)

    if file_path is None:
        result = result_template(file_type=args.file_type)
        result["reason"] = "找不到目標檔案"
        result["existing_files"] = list_existing_files()
        result["selected_by"] = selected_by
        print(json.dumps(result, ensure_ascii=False))
        return 1

    result = dispatch_formatter(file_path, args.file_type)
    result["selected_by"] = selected_by

    if not result.get("file_found"):
        result["existing_files"] = list_existing_files()
        print(json.dumps(result, ensure_ascii=False))
        return 1

    if not result.get("saved"):
        print(json.dumps(result, ensure_ascii=False))
        return 1

    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
