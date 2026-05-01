from __future__ import annotations

from ..excel.workbook_structure import StoreStructureMixin
from ..store_sheet_utils import StoreSheetUtilsMixin


class StoreContext(StoreSheetUtilsMixin, StoreStructureMixin):
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

    def get_member_calendar_dataset(self, name: str, year: int, month: int, workbook=None):
        return self.query_service.get_member_calendar_dataset(name, year, month, workbook=workbook)

    def get_all_members_calendar_dataset(self, year: int, month: int):
        return self.query_service.get_all_members_calendar_dataset(year, month)

    def get_wear_abnormal_threshold(self) -> float:
        return self.query_service.get_wear_abnormal_threshold()

    def _round_wear(self, value):
        return self._writer._round_wear(value)

    def _round_income(self, value):
        return self._writer._round_income(value)

    def _round_expense(self, value):
        return self._writer._round_expense(value)
