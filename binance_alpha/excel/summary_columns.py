from __future__ import annotations

from .header_locator import find_column


def ensure_summary_columns(sheet, total_header: str, name_header: str) -> tuple[int, int, bool]:
    changed = False
    total_col = find_column(sheet, total_header)
    name_col = find_column(sheet, name_header)
    if total_col is None:
        sheet.cell(1, sheet.max_column + 1, total_header)
        changed = True
    if name_col is None:
        sheet.cell(1, sheet.max_column + 1, name_header)
        changed = True

    total_col = find_column(sheet, total_header)
    name_col = find_column(sheet, name_header)
    assert total_col is not None and name_col is not None

    if name_col != total_col + 1:
        total_values = [sheet.cell(row, total_col).value for row in range(1, sheet.max_row + 1)]
        name_values = [sheet.cell(row, name_col).value for row in range(1, sheet.max_row + 1)]
        for col_index in sorted([total_col, name_col], reverse=True):
            sheet.delete_cols(col_index)
        insert_at = sheet.max_column + 1
        sheet.insert_cols(insert_at, 2)
        for row, value in enumerate(total_values, start=1):
            sheet.cell(row, insert_at, value)
        for row, value in enumerate(name_values, start=1):
            sheet.cell(row, insert_at + 1, value)
        changed = True

    total_col = find_column(sheet, total_header)
    name_col = find_column(sheet, name_header)
    assert total_col is not None and name_col is not None
    return total_col, name_col, changed
