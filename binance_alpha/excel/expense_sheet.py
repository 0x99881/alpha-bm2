from __future__ import annotations

from .header_locator import find_column, find_sheet_by_alias, numeric_header_columns
from .member_rows import ensure_member_rows, sort_rows_by_name
from ..constants import EXPENSE_NAME_HEADER, EXPENSE_SHEET, EXPENSE_SHEET_ALIASES


class ExpenseSheet:
    def __init__(self, store) -> None:
        self.store = store

    def sheet(self, workbook):
        sheet = find_sheet_by_alias(workbook, EXPENSE_SHEET_ALIASES)
        if sheet is not None:
            if sheet.title != EXPENSE_SHEET:
                sheet.title = EXPENSE_SHEET
            return sheet
        return workbook.create_sheet(title=EXPENSE_SHEET)

    def columns(self, sheet) -> list[tuple[int, int]]:
        return numeric_header_columns(sheet)

    def ensure_structure(self, workbook) -> bool:
        sheet = self.sheet(workbook)
        changed = False

        name_col = find_column(sheet, EXPENSE_NAME_HEADER)
        if name_col is None:
            sheet.cell(1, sheet.max_column + 1, EXPENSE_NAME_HEADER)
            changed = True
            name_col = find_column(sheet, EXPENSE_NAME_HEADER)

        assert name_col is not None
        if name_col != sheet.max_column:
            name_values = [sheet.cell(row, name_col).value for row in range(1, sheet.max_row + 1)]
            sheet.delete_cols(name_col)
            insert_at = sheet.max_column + 1
            sheet.insert_cols(insert_at, 1)
            for row, value in enumerate(name_values, start=1):
                sheet.cell(row, insert_at, value)
            changed = True
            name_col = insert_at

        changed = ensure_member_rows(
            sheet,
            members=self.store.get_members(),
            name_col=name_col,
            total_col=None,
            value_columns=[col for _, col in self.columns(sheet)],
        ) or changed
        sort_rows_by_name(sheet, name_col)
        return changed
