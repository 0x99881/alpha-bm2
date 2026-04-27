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
