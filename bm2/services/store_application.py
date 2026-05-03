from __future__ import annotations

from .store_bootstrap_service import StoreBootstrapService


class StoreApplication:
    def __init__(self, base_dir, read_only: bool = False):
        self._context = StoreBootstrapService().create_context(base_dir, read_only=read_only)

    @property
    def workbook_path(self):
        return self._context.workbook_path

    @workbook_path.setter
    def workbook_path(self, value) -> None:
        self._context.workbook_path = value

    @property
    def daily_entry_service(self):
        return self._context.daily_entry_service

    @property
    def member_service(self):
        return self._context.member_service

    @property
    def score_presenter(self):
        return self._context.score_presenter

    def get_quick_scores(self):
        return self._context.query_service.get_quick_scores()

    def get_members(self):
        return self._context.query_service.get_members()

    def get_active_members(self):
        return self._context.query_service.get_active_members()

    def get_member_wear_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self._context.query_service.get_member_wear_records(name, year_hint=year_hint, workbook=workbook)

    def get_score_summary(self):
        return self._context.query_service.get_score_summary()

    def get_score_sheet_view(self):
        return self._context.query_service.get_score_sheet_view()

    def get_wear_sheet_view(self):
        return self._context.query_service.get_wear_sheet_view()

    def get_member_profit_calendar(self, name: str, year: int, month: int):
        return self._context.query_service.get_member_profit_calendar(name, year, month)

    def get_next_score_date(self):
        return self._context.query_service.get_next_score_date()

    def get_score_rows_for_date(self, score_date: str):
        return self._context.query_service.get_score_rows_for_date(score_date)

    def get_online_active_members(self):
        return self._context.query_service.get_online_active_members()

    def get_online_score_summary(self):
        return self._context.query_service.get_online_score_summary()

    def get_online_next_score_date(self):
        return self._context.query_service.get_online_next_score_date()

    def get_mobile_overview(self):
        return self._context.query_service.get_mobile_overview()

    def set_wear_abnormal_threshold(self, value_text: str) -> float:
        return self._context.command_service.set_wear_abnormal_threshold(value_text)

    def refresh_local_database(self) -> dict[str, object]:
        return self._context.command_service.refresh_local_database()

    def export_to_excel(self, date_text: str | None = None) -> int:
        return self._context.export_service.export_to_excel(date_text)

    def create_new_cycle(self, start_date_text: str) -> str:
        return self._context.command_service.create_new_cycle(start_date_text)

    def delete_score_date(self, date_text: str) -> dict[str, int]:
        return self._context.command_service.delete_score_date(date_text)

    def is_supabase_configured(self) -> bool:
        return self._context.command_service.is_supabase_configured()

    def supabase_push(self, *, force_full: bool = True) -> dict[str, int]:
        return self._context.command_service.supabase_push(force_full=force_full)

    def supabase_pull(self) -> dict[str, int]:
        return self._context.command_service.supabase_pull()


__all__ = ["StoreApplication"]
