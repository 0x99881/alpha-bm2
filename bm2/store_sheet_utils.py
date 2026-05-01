from __future__ import annotations

from typing import Any, NamedTuple

from .constants import (
    DATA_START_ROW,
    PROFIT_SHEET,
    PROFIT_SHEET_ALIASES,
    VALUE_SHEET_SPECS,
)
from .excel.header_locator import find_column, find_sheet_by_alias, parse_numeric_header
from .excel.member_rows import build_name_row_map, ensure_member_rows_with_map, sort_named_rows, sort_rows_by_name


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

    def _set_member_row_hidden(self, sheet, name_col: int, member_name: str, hidden: bool) -> None:
        row = self._find_member_row(sheet, name_col, member_name)
        if row is not None:
            sheet.row_dimensions[row].hidden = hidden

    def _delete_member_row(self, sheet, name_col: int, member_name: str) -> None:
        row = self._find_member_row(sheet, name_col, member_name)
        if row is not None:
            sheet.delete_rows(row, 1)

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

    def _profit_sheet(self, workbook):
        sheet = self._find_sheet_by_alias(workbook, PROFIT_SHEET_ALIASES)
        if sheet is not None:
            if sheet.title != PROFIT_SHEET:
                sheet.title = PROFIT_SHEET
            return sheet
        return workbook.create_sheet(title=PROFIT_SHEET)

    def _score_sheet(self, workbook):
        return self.score_sheet.sheet(workbook)

    def _parse_value_header(self, value: Any) -> int | None:
        return parse_numeric_header(value)

    def _wear_columns(self, sheet) -> list[tuple[int, int]]:
        return self.wear_sheet.columns(sheet)

    def _income_columns(self, sheet) -> list[tuple[int, int]]:
        return self.income_sheet.columns(sheet)

    def _expense_columns(self, sheet) -> list[tuple[int, int]]:
        return self.expense_sheet.columns(sheet)

    def _find_column(self, sheet, header: str) -> int | None:
        return find_column(sheet, header)

    def _build_name_row_map(self, sheet, name_col: int) -> dict[str, int]:
        return build_name_row_map(sheet, name_col)

    def _ensure_member_rows(self, sheet, name_col: int, total_col: int | None, value_columns: list[int], extra_columns: list[int] | None = None) -> tuple[dict[str, int], bool]:
        return ensure_member_rows_with_map(
            sheet,
            members=self.get_members(),
            name_col=name_col,
            total_col=total_col,
            value_columns=value_columns,
            extra_columns=extra_columns,
        )

    def _ensure_summary_columns(self, sheet, total_header: str, name_header: str) -> tuple[int, int, bool]:
        changed = False
        total_col = self._find_column(sheet, total_header)
        name_col = self._find_column(sheet, name_header)
        if total_col is None:
            sheet.cell(1, sheet.max_column + 1, total_header)
            changed = True
        if name_col is None:
            sheet.cell(1, sheet.max_column + 1, name_header)
            changed = True

        total_col = self._find_column(sheet, total_header)
        name_col = self._find_column(sheet, name_header)
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

        total_col = self._find_column(sheet, total_header)
        name_col = self._find_column(sheet, name_header)
        assert total_col is not None and name_col is not None
        return total_col, name_col, changed

    def _ensure_score_tail_columns(self, sheet) -> tuple[int, int, int, bool]:
        return self.score_sheet.ensure_tail_columns(sheet)

    def _sort_named_rows(self, sheet, total_col: int, name_col: int) -> None:
        sort_named_rows(sheet, total_col, name_col)

    def _sort_rows_by_name(self, sheet, name_col: int) -> None:
        sort_rows_by_name(sheet, name_col)

    def _is_mmdd_header(self, value: Any) -> bool:
        return self.score_sheet.is_date_header(value)

    def _score_date_columns(self, sheet) -> list[tuple[int, int]]:
        return self.score_sheet.date_columns(sheet)

    def _d_columns(self, sheet) -> list[tuple[int, int]]:
        return self._score_date_columns(sheet)
