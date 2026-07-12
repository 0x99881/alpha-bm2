from __future__ import annotations


def safe_float(value) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def sum_sheet_row_values(sheet, row: int, columns: list[int]) -> float:
    return sum(safe_float(sheet.cell(row, col).value) for col in columns)
