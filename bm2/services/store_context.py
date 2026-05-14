from __future__ import annotations

from ..excel.workbook_structure import StoreStructureMixin
from ..store_sheet_utils import StoreSheetUtilsMixin


class StoreContext(StoreSheetUtilsMixin, StoreStructureMixin):
    def ensure_workbook(self) -> None:
        return self._ensure_workbook()

    def score_sheet_for(self, workbook):
        return self._score_sheet(workbook)

    def score_date_columns_for(self, sheet):
        return self._score_date_columns(sheet)

    def wear_sheet_for(self, workbook):
        return self._wear_sheet(workbook)

    def income_sheet_for(self, workbook):
        return self._income_sheet(workbook)

    def expense_sheet_for(self, workbook):
        return self._expense_sheet(workbook)

    def wear_columns_for(self, sheet):
        return self._wear_columns(sheet)

    def income_columns_for(self, sheet):
        return self._income_columns(sheet)

    def expense_columns_for(self, sheet):
        return self._expense_columns(sheet)

    def value_sheet_helpers_for(self, sheet_type: str):
        return self._value_sheet_helpers(sheet_type)

    def ensure_score_sheet_structure(self, workbook) -> bool:
        return self._ensure_score_sheet_structure(workbook)

    def ensure_wear_sheet_structure(self, workbook) -> bool:
        return self._ensure_wear_sheet_structure(workbook)

    def ensure_income_sheet_structure(self, workbook) -> bool:
        return self._ensure_income_sheet_structure(workbook)

    def ensure_expense_sheet_structure(self, workbook) -> bool:
        return self._ensure_expense_sheet_structure(workbook)

    def sync_member_visibility_in_workbook(self, workbook) -> None:
        return self._sync_member_visibility_in_workbook(workbook)

    def recalculate_score_totals(self, sheet, total_col: int) -> None:
        return self._recalculate_totals(sheet, total_col)

    def recalculate_score_profits(self, workbook, score_sheet, profit_col: int) -> None:
        return self._recalculate_score_profits(workbook, score_sheet, profit_col)

    def format_score_sheet(self, sheet, recent_numbers: list[int], total_col: int) -> None:
        return self._format_score_sheet(sheet, recent_numbers, total_col)

    def recalculate_wear_totals(self, sheet, total_col: int) -> None:
        return self._recalculate_wear_totals(sheet, total_col)

    def get_members(self):
        return self.query_service.get_members()

    def get_active_members(self):
        return self.query_service.get_active_members()

    def get_score_summary_data(self):
        return self.query_service.get_score_summary_data()

    def get_score_sheet_snapshot(self):
        return self.query_service.get_score_sheet_snapshot()

    def get_wear_sheet_snapshot(self):
        return self.query_service.get_wear_sheet_snapshot()

    def get_value_sheet_snapshot(self, sheet_type: str):
        return self.query_service.get_value_sheet_snapshot(sheet_type)

    def get_member_calendar_dataset(self, name: str, year: int, month: int, workbook=None):
        return self.query_service.get_member_calendar_dataset(name, year, month, workbook=workbook)

    def get_all_members_calendar_dataset(self, year: int, month: int):
        return self.query_service.get_all_members_calendar_dataset(year, month)

    def get_wear_abnormal_threshold(self) -> float:
        return self.query_service.get_wear_abnormal_threshold()

    def format_wear_value(self, value):
        return self.workbook_writer.format_wear_value(value)

    def format_income_value(self, value):
        return self.workbook_writer.format_income_value(value)

    def format_expense_value(self, value):
        return self.workbook_writer.format_expense_value(value)
