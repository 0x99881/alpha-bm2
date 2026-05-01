from __future__ import annotations

from openpyxl import Workbook, load_workbook

from ..constants import (
    EXPENSE_META_SHEET_ALIASES,
    EXPENSE_NAME_HEADER,
    DISABLED,
    INCOME_META_SHEET_ALIASES,
    INCOME_NAME_HEADER,
    META_HEADERS,
    META_SHEET,
    META_SHEET_ALIASES,
    NAME_HEADER,
    PROFIT_HEADERS,
    PROFIT_HEADER,
    PROFIT_SHEET,
    PROFIT_SHEET_ALIASES,
    WEAR_META_SHEET_ALIASES,
    WEAR_NAME_HEADER,
    WINDOW_SIZE,
)
from .number_utils import safe_float, sum_sheet_row_values
from .profit_recalculator import recalculate_score_profits


class StoreStructureMixin:
    def _ensure_sheet_member_rows(self, sheet, *, name_col: int, total_col: int | None, value_columns: list[int], extra_columns: list[int] | None = None) -> bool:
        _, changed = self._ensure_member_rows(
            sheet,
            name_col=name_col,
            total_col=total_col,
            value_columns=value_columns,
            extra_columns=extra_columns,
        )
        return changed

    def _create_initial_score_sheet(self, workbook) -> None:
        self.score_sheet.create_initial_sheet(workbook)

    def _ensure_meta_sheets(self, workbook) -> bool:
        changed = False
        _, meta_changed = self._ensure_named_sheet(workbook, META_SHEET, META_SHEET_ALIASES, META_HEADERS, hidden=True)
        changed = meta_changed or changed
        for sheet_type in ('wear', 'income', 'expense'):
            _, value_meta_changed = self._ensure_named_sheet(
                workbook,
                self._value_sheet_spec(sheet_type)['meta_sheet'],
                self._value_sheet_spec(sheet_type)['meta_aliases'],
                self._value_sheet_spec(sheet_type)['meta_headers'],
                hidden=True,
            )
            changed = value_meta_changed or changed
        return changed

    def _ensure_score_sheet_structure(self, workbook) -> bool:
        return self.score_sheet.ensure_structure(workbook)

    def _recalculate_score_profits(self, workbook, score_sheet, profit_col: int) -> None:
        recalculate_score_profits(workbook, score_sheet, profit_col, self.income_sheet, self.wear_sheet)

    def _recalculate_wear_totals(self, sheet, total_col: int) -> None:
        self.wear_sheet.recalculate_totals(sheet, total_col)

    def _format_wear_sheet(self, sheet, total_col: int) -> None:
        self.wear_sheet.format_sheet(sheet, total_col)

    def _ensure_wear_sheet_structure(self, workbook) -> bool:
        return self.wear_sheet.ensure_structure(workbook)

    def _ensure_income_sheet_structure(self, workbook) -> bool:
        return self.income_sheet.ensure_structure(workbook)

    def _ensure_expense_sheet_structure(self, workbook) -> bool:
        return self.expense_sheet.ensure_structure(workbook)

    def _sum_sheet_row_values(self, sheet, row: int, columns: list[int]) -> float:
        return sum_sheet_row_values(sheet, row, columns)

    def _sort_profit_rows(self, sheet, *, name_col: int, profit_col: int) -> None:
        rows = []
        for row in range(2, sheet.max_row + 1):
            values = [sheet.cell(row, col).value for col in range(1, sheet.max_column + 1)]
            if values[name_col - 1]:
                rows.append(values)
        rows.sort(key=lambda row: (-safe_float(row[profit_col - 1]), str(row[name_col - 1])))
        for row_index, values in enumerate(rows, start=2):
            for col_index, value in enumerate(values, start=1):
                sheet.cell(row_index, col_index, value)

    def _recalculate_profit_sheet(self, profit_sheet, income_sheet, wear_sheet) -> None:
        profit_name_col = self._find_column(profit_sheet, NAME_HEADER)
        income_name_col = self._find_column(income_sheet, INCOME_NAME_HEADER)
        wear_name_col = self._find_column(wear_sheet, WEAR_NAME_HEADER)
        if profit_name_col is None or income_name_col is None or wear_name_col is None:
            return

        income_row_map = self._build_name_row_map(income_sheet, income_name_col)
        wear_row_map = self._build_name_row_map(wear_sheet, wear_name_col)
        income_columns = [col for _, col in self._income_columns(income_sheet)]
        wear_columns = [col for _, col in self._wear_columns(wear_sheet)]

        for row in range(2, profit_sheet.max_row + 1):
            member_name = str(profit_sheet.cell(row, profit_name_col).value or '').strip()
            if not member_name:
                continue
            income_total = self._round_income(
                self._sum_sheet_row_values(income_sheet, income_row_map[member_name], income_columns)
                if member_name in income_row_map else 0
            )
            wear_total = self._round_wear(
                self._sum_sheet_row_values(wear_sheet, wear_row_map[member_name], wear_columns)
                if member_name in wear_row_map else 0
            )
            profit_sheet.cell(row, 2, income_total)
            profit_sheet.cell(row, 3, wear_total)
            profit_sheet.cell(row, 4, self._round_income(income_total - wear_total))

    def _ensure_profit_sheet_structure(self, workbook) -> bool:
        sheet, changed = self._ensure_named_sheet(workbook, PROFIT_SHEET, PROFIT_SHEET_ALIASES, PROFIT_HEADERS)
        name_col = self._find_column(sheet, NAME_HEADER)
        if name_col is None:
            return changed
        changed = self._ensure_sheet_member_rows(sheet, name_col=name_col, total_col=None, value_columns=[2, 3, 4]) or changed
        self._recalculate_profit_sheet(sheet, self._income_sheet(workbook), self._wear_sheet(workbook))
        self._sort_profit_rows(sheet, name_col=name_col, profit_col=4)
        return changed

    def _ensure_workbook(self) -> None:
        changed = False
        if self.workbook_path.exists():
            workbook = load_workbook(self.workbook_path)
        else:
            if getattr(self, 'read_only', False):
                raise FileNotFoundError(f"Workbook not found: {self.workbook_path}")
            workbook = Workbook()
            self._create_initial_score_sheet(workbook)
            changed = True

        try:
            changed = self._ensure_meta_sheets(workbook) or changed
            changed = self._ensure_score_sheet_structure(workbook) or changed
            changed = self._ensure_wear_sheet_structure(workbook) or changed
            changed = self._ensure_income_sheet_structure(workbook) or changed
            changed = self._ensure_expense_sheet_structure(workbook) or changed
            if PROFIT_SHEET in workbook.sheetnames:
                workbook.remove(workbook[PROFIT_SHEET])
                changed = True
            if changed and not getattr(self, 'read_only', False):
                self._sync_member_visibility_in_workbook(workbook)
                self.workbook_repository.save(workbook)
        finally:
            workbook.close()

    def _sync_member_visibility_in_workbook(self, workbook) -> None:
        self._ensure_score_sheet_structure(workbook)
        self._ensure_wear_sheet_structure(workbook)
        self._ensure_income_sheet_structure(workbook)
        self._ensure_expense_sheet_structure(workbook)

        score_sheet = self._score_sheet(workbook)
        wear_sheet = self._wear_sheet(workbook)
        income_sheet = self._income_sheet(workbook)
        expense_sheet = self._expense_sheet(workbook)
        score_name_col = self._find_column(score_sheet, NAME_HEADER)
        wear_name_col = self._find_column(wear_sheet, WEAR_NAME_HEADER)
        income_name_col = self._find_column(income_sheet, INCOME_NAME_HEADER)
        expense_name_col = self._find_column(expense_sheet, EXPENSE_NAME_HEADER)

        for member in self.get_members():
            should_hide = member['status'] == DISABLED
            if score_name_col is not None:
                self._set_member_row_hidden(score_sheet, score_name_col, member['name'], should_hide)
            if wear_name_col is not None:
                self._set_member_row_hidden(wear_sheet, wear_name_col, member['name'], should_hide)
            if income_name_col is not None:
                self._set_member_row_hidden(income_sheet, income_name_col, member['name'], should_hide)
            if expense_name_col is not None:
                self._set_member_row_hidden(expense_sheet, expense_name_col, member['name'], should_hide)

    def _append_meta(self, workbook, date_text: str, col_name: str) -> None:
        workbook[META_SHEET].append([date_text, col_name])

    def _append_value_meta(self, workbook, sheet_type: str, date_text: str, col_name: str) -> None:
        workbook[self._value_sheet_spec(sheet_type)['meta_sheet']].append([date_text, col_name])

    def _recalculate_totals(self, sheet, total_col: int) -> None:
        self.score_sheet.recalculate_totals(sheet, total_col)

    def _format_score_sheet(self, sheet, recent_numbers: list[int], total_col: int) -> None:
        self.score_sheet.format_sheet(sheet, recent_numbers, total_col)
