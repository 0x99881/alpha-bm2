from __future__ import annotations

from typing import Any


def find_sheet_by_alias(workbook, aliases: list[str]):
    for name in aliases:
        if name in workbook.sheetnames:
            return workbook[name]
    return None


def find_column(sheet, header: str) -> int | None:
    for col in range(1, sheet.max_column + 1):
        if sheet.cell(1, col).value == header:
            return col
    return None


def parse_numeric_header(value: Any) -> int | None:
    if not isinstance(value, str):
        return None
    if not value.isdigit():
        return None
    return int(value)


def numeric_header_columns(sheet) -> list[tuple[int, int]]:
    result = []
    for col in range(1, sheet.max_column + 1):
        number = parse_numeric_header(sheet.cell(1, col).value)
        if number is not None:
            result.append((number, col))
    result.sort(key=lambda item: item[0])
    return result
