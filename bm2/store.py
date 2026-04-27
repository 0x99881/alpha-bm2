from __future__ import annotations

# COMPATIBILITY LAYER

from .constants import DISABLED, ENABLED
from .excel import ExpenseSheet, IncomeSheet, ScoreSheet, WearSheet
from .presenters import ProfitCalendarPresenter, ScorePresenter, WearPresenter
from .services import DailyEntryService, MemberService
from .store_base import StoreBaseMixin
from .store_sheet_utils import StoreSheetUtilsMixin
from .store_structure import StoreStructureMixin
from .store_reader_facade import StoreReaderFacade
from .store_writer_facade import StoreWriterFacade


class ExcelStore(
    StoreBaseMixin,
    StoreSheetUtilsMixin,
    StoreStructureMixin,
):
    """External facade.

    Keep web-facing calls on this facade while legacy mixins are being split
    into smaller services and rules.
    """

    def __init__(self, base_dir):
        self.score_sheet = ScoreSheet(self)
        self.wear_sheet = WearSheet(self)
        self.income_sheet = IncomeSheet(self)
        self.expense_sheet = ExpenseSheet(self)
        self._reader = StoreReaderFacade(self)
        self._writer = StoreWriterFacade(self)
        super().__init__(base_dir)
        self.daily_entry_service = DailyEntryService(self._writer)
        self.member_service = MemberService(
            self.get_members,
            self.config_repository,
            self.sync_members_to_workbook,
            self.delete_member_from_workbook,
        )
        self.score_presenter = ScorePresenter(self)
        self.wear_presenter = WearPresenter(self)
        self.profit_calendar_presenter = ProfitCalendarPresenter(self)

    def _round_wear(self, value):
        return self._writer._round_wear(value)

    def _round_income(self, value):
        return self._writer._round_income(value)

    def _round_expense(self, value):
        return self._writer._round_expense(value)

    def save_scores_and_wear(self, date_text: str, entries: list[dict[str, str]]):
        return self._writer.save_scores_and_wear(date_text, entries)

    def get_score_rankings(self, limit: int = 999, workbook=None):
        return self._reader.get_score_rankings(limit=limit, workbook=workbook)

    def get_score_latest_column(self) -> str:
        return self._reader.get_score_latest_column()

    def get_score_column_count(self) -> int:
        return self._reader.get_score_column_count()

    def get_active_member_profit_map(self) -> dict[str, float]:
        return self._reader.get_active_member_profit_map()

    def get_score_summary_data(self):
        return self._reader.get_score_summary_data()

    def get_member_wear_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self._reader.get_member_wear_records(name, year_hint=year_hint, workbook=workbook)

    def get_member_income_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self._reader.get_member_income_records(name, year_hint=year_hint, workbook=workbook)

    def get_wear_sheet_snapshot(self):
        return self._reader.get_wear_sheet_snapshot()

    def get_member_calendar_dataset(self, name: str, year: int, month: int, workbook=None):
        return self._reader.get_member_calendar_dataset(name, year, month, workbook=workbook)

    def get_all_members_calendar_dataset(self, year: int, month: int):
        return self._reader.get_all_members_calendar_dataset(year, month)

    def get_next_score_date(self) -> str:
        return self._reader.get_next_score_date()

    def get_score_summary(self):
        return self.score_presenter.build_score_summary()

    def get_wear_sheet_view(self):
        return self.wear_presenter.build_wear_sheet_view()

    def get_member_profit_calendar(self, name: str, year: int, month: int):
        return self.profit_calendar_presenter.build_member_profit_calendar(name=name, year=year, month=month)

    def get_active_members(self) -> list[dict[str, str]]:
        return self.member_service.get_active_members()

    def add_member(self, name: str, note: str = '') -> None:
        self.member_service.add_member(name, note)

    def update_member(self, name: str, note: str, status: str | None = None) -> None:
        self.member_service.update_member(name, note, status=status)

    def delete_member(self, name: str) -> None:
        self.member_service.delete_member(name)

    def reorder_active_members(self, ordered_names: list[str]) -> None:
        self.member_service.reorder_active_members(ordered_names)


__all__ = ["ExcelStore", "ENABLED", "DISABLED"]
