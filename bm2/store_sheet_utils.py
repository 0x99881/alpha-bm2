from __future__ import annotations

from typing import Any, NamedTuple

from .constants import (
    DATA_START_ROW,
    VALUE_SHEET_SPECS,
)
from .excel.header_locator import find_column, find_sheet_by_alias


class ValueSheetHelpers(NamedTuple):
    spec: dict[str, Any]
    sheet_getter: Any
    ensure_structure: Any
    columns_getter: Any
    value_formatter: Any


class StoreSheetUtilsMixin:
    def _value_sheet_spec(self, sheet_type: str) -> dict[str, Any]:
        return VALUE_SHEET_SPECS[sheet_type]

    def _value_sheet_helpers(self, sheet_type: str):
        spec = self._value_sheet_spec(sheet_type)
        return ValueSheetHelpers(
            spec=spec,
            sheet_getter=getattr(self, spec["sheet_method"]),
            ensure_structure=getattr(self, spec["ensure_method"]),
            columns_getter=getattr(self, spec["columns_method"]),
            value_formatter=getattr(self, spec["round_method"]),
        )

    def _find_member_row(self, sheet, name_col: int, member_name: str) -> int | None:
        for row in range(DATA_START_ROW, sheet.max_row + 1):
            if str(sheet.cell(row, name_col).value or "").strip() == member_name:
                return row
        return None

    def _set_member_row_hidden(self, sheet, name_col: int, member_name: str, hidden: bool) -> bool:
        row = self._find_member_row(sheet, name_col, member_name)
        if row is None:
            return False
        if bool(sheet.row_dimensions[row].hidden) == hidden:
            return False
        sheet.row_dimensions[row].hidden = hidden
        return True

    def _find_sheet_by_alias(self, workbook, aliases: list[str]):
        return find_sheet_by_alias(workbook, aliases)

    def _ensure_named_sheet(self, workbook, primary_name: str, aliases: list[str], headers: list[str], hidden: bool = False):
        sheet = self._find_sheet_by_alias(workbook, aliases)
        changed = False
        if sheet is None:
            sheet = workbook.create_sheet(title=primary_name)
            sheet.append(headers)
            changed = True
        elif sheet.title != primary_name:
            sheet.title = primary_name
            changed = True

        current_headers = [sheet.cell(1, idx).value for idx in range(1, len(headers) + 1)]
        if current_headers != headers:
            sheet.delete_rows(1, sheet.max_row)
            sheet.append(headers)
            changed = True

        if hidden and sheet.sheet_state != "hidden":
            sheet.sheet_state = "hidden"
            changed = True
        return sheet, changed

    def _wear_sheet(self, workbook):
        return self.wear_sheet.sheet(workbook)

    def _income_sheet(self, workbook):
        return self.income_sheet.sheet(workbook)

    def _expense_sheet(self, workbook):
        return self.expense_sheet.sheet(workbook)

    def _score_sheet(self, workbook):
        return self.score_sheet.sheet(workbook)

    def _wear_columns(self, sheet) -> list[tuple[int, int]]:
        return self.wear_sheet.columns(sheet)

    def _income_columns(self, sheet) -> list[tuple[int, int]]:
        return self.income_sheet.columns(sheet)

    def _expense_columns(self, sheet) -> list[tuple[int, int]]:
        return self.expense_sheet.columns(sheet)

    def _find_column(self, sheet, header: str) -> int | None:
        return find_column(sheet, header)

    def _score_date_columns(self, sheet) -> list[tuple[int, int]]:
        return self.score_sheet.date_columns(sheet)
