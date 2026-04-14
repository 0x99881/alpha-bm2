from __future__ import annotations

import re
from typing import Any

from .constants import (
    EXPENSE_SHEET,
    EXPENSE_SHEET_ALIASES,
    PROFIT_SHEET,
    PROFIT_SHEET_ALIASES,
    INCOME_SHEET,
    INCOME_SHEET_ALIASES,
    LEGACY_WEAR_SHEET,
    NAME_HEADER,
    OLD_SCORE_PREFIX,
    PROFIT_HEADER,
    SCORE_SHEET,
    SCORE_SHEET_ALIASES,
    TOTAL_HEADER,
    WEAR_SHEET,
    WEAR_SHEET_ALIASES,
)


class StoreSheetUtilsMixin:
    def _find_member_row(self, sheet, name_col: int, member_name: str) -> int | None:
        for row in range(2, sheet.max_row + 1):
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
        for name in aliases:
            if name in workbook.sheetnames:
                return workbook[name]
        return None

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
            if primary_name == WEAR_SHEET and sheet.max_row > 1:
                if LEGACY_WEAR_SHEET in workbook.sheetnames:
                    workbook.remove(workbook[LEGACY_WEAR_SHEET])
                sheet.title = LEGACY_WEAR_SHEET
                sheet = workbook.create_sheet(title=primary_name)
                sheet.append(headers)
            else:
                sheet.delete_rows(1, sheet.max_row)
                sheet.append(headers)
            changed = True

        if hidden and sheet.sheet_state != "hidden":
            sheet.sheet_state = "hidden"
            changed = True
        return sheet, changed

    def _wear_sheet(self, workbook):
        sheet = self._find_sheet_by_alias(workbook, WEAR_SHEET_ALIASES)
        if sheet is not None:
            if sheet.title != WEAR_SHEET:
                sheet.title = WEAR_SHEET
            return sheet
        return workbook.create_sheet(title=WEAR_SHEET)

    def _income_sheet(self, workbook):
        sheet = self._find_sheet_by_alias(workbook, INCOME_SHEET_ALIASES)
        if sheet is not None:
            if sheet.title != INCOME_SHEET:
                sheet.title = INCOME_SHEET
            return sheet
        return workbook.create_sheet(title=INCOME_SHEET)

    def _expense_sheet(self, workbook):
        sheet = self._find_sheet_by_alias(workbook, EXPENSE_SHEET_ALIASES)
        if sheet is not None:
            if sheet.title != EXPENSE_SHEET:
                sheet.title = EXPENSE_SHEET
            return sheet
        return workbook.create_sheet(title=EXPENSE_SHEET)


    def _profit_sheet(self, workbook):
        sheet = self._find_sheet_by_alias(workbook, PROFIT_SHEET_ALIASES)
        if sheet is not None:
            if sheet.title != PROFIT_SHEET:
                sheet.title = PROFIT_SHEET
            return sheet
        return workbook.create_sheet(title=PROFIT_SHEET)

    def _score_sheet(self, workbook):
        sheet = self._find_sheet_by_alias(workbook, SCORE_SHEET_ALIASES)
        if sheet is not None:
            if sheet.title != SCORE_SHEET:
                sheet.title = SCORE_SHEET
            return sheet
        if workbook.sheetnames:
            workbook[workbook.sheetnames[0]].title = SCORE_SHEET
            return workbook[SCORE_SHEET]
        return workbook.create_sheet(title=SCORE_SHEET)

    def _parse_value_header(self, value: Any) -> int | None:
        if not isinstance(value, str):
            return None
        if not value.isdigit():
            return None
        return int(value)

    def _wear_columns(self, sheet) -> list[tuple[int, int]]:
        result = []
        for col in range(1, sheet.max_column + 1):
            number = self._parse_value_header(sheet.cell(1, col).value)
            if number is not None:
                result.append((number, col))
        result.sort(key=lambda item: item[0])
        return result

    def _income_columns(self, sheet) -> list[tuple[int, int]]:
        return self._wear_columns(sheet)

    def _expense_columns(self, sheet) -> list[tuple[int, int]]:
        return self._wear_columns(sheet)

    def _find_column(self, sheet, header: str) -> int | None:
        for col in range(1, sheet.max_column + 1):
            if sheet.cell(1, col).value == header:
                return col
        return None

    def _build_name_row_map(self, sheet, name_col: int) -> dict[str, int]:
        mapping: dict[str, int] = {}
        for row in range(2, sheet.max_row + 1):
            name = sheet.cell(row, name_col).value
            if name:
                mapping[str(name).strip()] = row
        return mapping

    def _is_meaningful_value(self, value: Any) -> bool:
        return value not in (None, '')

    def _should_replace_cell(self, current: Any, candidate: Any) -> bool:
        if not self._is_meaningful_value(candidate):
            return False
        if not self._is_meaningful_value(current):
            return True
        if isinstance(current, (int, float)) and current == 0 and isinstance(candidate, (int, float)) and candidate != 0:
            return True
        return False

    def _drop_columns_after(self, sheet, keep_col: int) -> bool:
        if keep_col >= sheet.max_column:
            return False
        sheet.delete_cols(keep_col + 1, sheet.max_column - keep_col)
        return True

    def _dedupe_member_rows(self, sheet, name_col: int, protected_cols: list[int]) -> bool:
        first_rows: dict[str, int] = {}
        duplicate_rows: list[int] = []
        changed = False
        for row in range(2, sheet.max_row + 1):
            raw_name = sheet.cell(row, name_col).value
            if not raw_name:
                continue
            name = str(raw_name).strip()
            if name not in first_rows:
                first_rows[name] = row
                continue
            keeper_row = first_rows[name]
            for col in protected_cols:
                current = sheet.cell(keeper_row, col).value
                candidate = sheet.cell(row, col).value
                if self._should_replace_cell(current, candidate):
                    sheet.cell(keeper_row, col, candidate)
                    changed = True
            keeper_hidden = bool(sheet.row_dimensions[keeper_row].hidden)
            row_hidden = bool(sheet.row_dimensions[row].hidden)
            if keeper_hidden and not row_hidden:
                sheet.row_dimensions[keeper_row].hidden = False
                changed = True
            duplicate_rows.append(row)
        for row in reversed(duplicate_rows):
            sheet.delete_rows(row, 1)
            changed = True
        return changed

    def _ensure_member_rows(self, sheet, name_col: int, total_col: int | None, value_columns: list[int], extra_columns: list[int] | None = None) -> tuple[dict[str, int], bool]:
        changed = False
        changed = self._drop_columns_after(sheet, name_col) or changed
        extra_columns = extra_columns or []
        protected_cols = sorted({*value_columns, *extra_columns, *([total_col] if total_col is not None else []), name_col})
        changed = self._dedupe_member_rows(sheet, name_col, protected_cols) or changed
        row_map = self._build_name_row_map(sheet, name_col)
        for member in self.get_members():
            name = member["name"]
            if name in row_map:
                continue
            row = sheet.max_row + 1
            for col in value_columns:
                sheet.cell(row, col, 0)
            if total_col is not None:
                sheet.cell(row, total_col, 0)
            for col in extra_columns:
                sheet.cell(row, col, 0)
            sheet.cell(row, name_col, name)
            row_map[name] = row
            changed = True
        return row_map, changed

    def _meta_to_date_map(self, workbook, sheet_name: str) -> dict[int, str]:
        mapping: dict[int, str] = {}
        for date_text, col_name in self._read_sheet_meta(workbook, sheet_name):
            if str(col_name).isdigit():
                mapping[int(col_name)] = str(date_text)
        return mapping

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
        total_col = self._find_column(sheet, TOTAL_HEADER)
        profit_col = self._find_column(sheet, PROFIT_HEADER)
        name_col = self._find_column(sheet, NAME_HEADER)
        changed = False

        if total_col is not None and profit_col is not None and name_col is not None and profit_col == total_col + 1 and name_col == profit_col + 1:
            sheet.cell(1, total_col, TOTAL_HEADER)
            sheet.cell(1, profit_col, PROFIT_HEADER)
            sheet.cell(1, name_col, NAME_HEADER)
            return total_col, profit_col, name_col, changed

        total_values = [sheet.cell(row, total_col).value for row in range(1, sheet.max_row + 1)] if total_col is not None else [None] * sheet.max_row
        profit_values = [sheet.cell(row, profit_col).value for row in range(1, sheet.max_row + 1)] if profit_col is not None else [None] * sheet.max_row
        name_values = [sheet.cell(row, name_col).value for row in range(1, sheet.max_row + 1)] if name_col is not None else [None] * sheet.max_row

        existing_cols = [col for col in [total_col, profit_col, name_col] if col is not None]
        for col_index in sorted(existing_cols, reverse=True):
            sheet.delete_cols(col_index)
        insert_at = sheet.max_column + 1
        sheet.insert_cols(insert_at, 3)
        for row, value in enumerate(total_values, start=1):
            sheet.cell(row, insert_at, value)
        for row, value in enumerate(profit_values, start=1):
            sheet.cell(row, insert_at + 1, value)
        for row, value in enumerate(name_values, start=1):
            sheet.cell(row, insert_at + 2, value)
        sheet.cell(1, insert_at, TOTAL_HEADER)
        sheet.cell(1, insert_at + 1, PROFIT_HEADER)
        sheet.cell(1, insert_at + 2, NAME_HEADER)
        changed = True
        return insert_at, insert_at + 1, insert_at + 2, changed

    def _sort_named_rows(self, sheet, total_col: int, name_col: int) -> None:
        rows = []
        for row in range(2, sheet.max_row + 1):
            values = [sheet.cell(row, col).value for col in range(1, sheet.max_column + 1)]
            if values[name_col - 1]:
                rows.append(values)
        rows.sort(
            key=lambda row: (
                -(float(row[total_col - 1]) if isinstance(row[total_col - 1], (int, float)) else 0.0),
                str(row[name_col - 1]),
            )
        )
        for row_index, values in enumerate(rows, start=2):
            for col_index, value in enumerate(values, start=1):
                sheet.cell(row_index, col_index, value)

    def _sort_rows_by_name(self, sheet, name_col: int) -> None:
        rows = []
        for row in range(2, sheet.max_row + 1):
            values = [sheet.cell(row, col).value for col in range(1, sheet.max_column + 1)]
            if values[name_col - 1]:
                rows.append(values)
        rows.sort(key=lambda row: str(row[name_col - 1]))
        for row_index, values in enumerate(rows, start=2):
            for col_index, value in enumerate(values, start=1):
                sheet.cell(row_index, col_index, value)

    def _parse_d_header(self, value: Any) -> int | None:
        if not isinstance(value, str):
            return None
        normalized = value.replace(f"{OLD_SCORE_PREFIX}-", "").replace(OLD_SCORE_PREFIX, "")
        if not normalized.startswith("D"):
            return None
        suffix = normalized[1:]
        return int(suffix) if suffix.isdigit() else None

    def _is_mmdd_header(self, value: Any) -> bool:
        return isinstance(value, str) and re.fullmatch(r"\d{2}-\d{2}", value) is not None

    def _d_columns(self, sheet) -> list[tuple[int, int]]:
        total_col = self._find_column(sheet, TOTAL_HEADER)
        if total_col is not None:
            result = []
            for number, col in enumerate(range(1, total_col), start=1):
                if sheet.cell(1, col).value is not None:
                    result.append((number, col))
            return result

        result = []
        for col in range(1, sheet.max_column + 1):
            number = self._parse_d_header(sheet.cell(1, col).value)
            if number is not None:
                result.append((number, col))
        result.sort(key=lambda item: item[0])
        return result
