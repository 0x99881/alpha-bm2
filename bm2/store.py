from __future__ import annotations

from .constants import DISABLED, ENABLED
from .presenters import ProfitCalendarPresenter, ScorePresenter, WearPresenter
from .record_repository import RecordRepository
from .store_base import StoreBaseMixin
from .store_read import StoreReadMixin
from .store_sheet_utils import StoreSheetUtilsMixin
from .store_structure import StoreStructureMixin
from .store_write import StoreWriteMixin


class ExcelStore(
    StoreBaseMixin,
    StoreSheetUtilsMixin,
    StoreStructureMixin,
    StoreWriteMixin,
    StoreReadMixin,
):
    """External facade.

    New read/write and presentation entry points should go through the
    repository, presenters, workbook manager, or services instead of adding
    fresh behavior into the legacy mixins.
    """

    def __init__(self, base_dir):
        super().__init__(base_dir)
        self.record_repository = RecordRepository(self)
        self.score_presenter = ScorePresenter(self.record_repository)
        self.wear_presenter = WearPresenter(self.record_repository)
        self.profit_calendar_presenter = ProfitCalendarPresenter(self.record_repository)

    def get_score_rankings(self, limit: int = 999, workbook=None):
        return self.record_repository.get_score_rankings(limit=limit, workbook=workbook)

    def get_score_summary(self):
        return self.score_presenter.build_score_summary()

    def get_active_member_profit_map(self):
        return self.record_repository.get_active_member_profit_map()

    def get_member_wear_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self.record_repository.get_member_wear_records(name=name, year_hint=year_hint, workbook=workbook)

    def get_member_income_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self.record_repository.get_member_income_records(name=name, year_hint=year_hint, workbook=workbook)

    def get_wear_sheet_view(self):
        return self.wear_presenter.build_wear_sheet_view()

    def get_member_profit_calendar(self, name: str, year: int, month: int):
        return self.profit_calendar_presenter.build_member_profit_calendar(name=name, year=year, month=month)

    def get_next_score_date(self) -> str:
        return self.record_repository.get_next_score_date()

    def save_scores_and_wear(self, date_text: str, entries):
        return self.record_repository.save_scores_and_wear(date_text, entries)


__all__ = ["ExcelStore", "ENABLED", "DISABLED"]
