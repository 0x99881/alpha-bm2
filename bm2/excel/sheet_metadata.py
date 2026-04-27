from __future__ import annotations

from datetime import datetime


def parse_saved_date(date_text: str):
    return datetime.strptime(str(date_text).strip(), '%Y-%m-%d').date()


def read_sheet_meta(workbook, sheet_name: str) -> list[tuple[str, str]]:
    if sheet_name not in workbook.sheetnames:
        return []
    rows = []
    for row in workbook[sheet_name].iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            rows.append((str(row[0]), str(row[1])))
    return rows


def meta_to_date_map(workbook, sheet_name: str) -> dict[int, str]:
    mapping: dict[int, str] = {}
    for date_text, col_name in read_sheet_meta(workbook, sheet_name):
        if str(col_name).isdigit():
            mapping[int(col_name)] = str(date_text)
    return mapping


def normalize_day_code_date_text(value: str, year: int) -> str:
    if len(value) == 4 and value.isdigit():
        return f"{year:04d}-{value[:2]}-{value[2:]}"
    return value


def append_sheet_meta(workbook, sheet_name: str, date_text: str, col_name: str) -> None:
    workbook[sheet_name].append([date_text, col_name])
