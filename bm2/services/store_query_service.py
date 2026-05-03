from __future__ import annotations

from datetime import datetime

from ..constants import DEFAULT_QUICK_SCORES, ENABLED, WEAR_ABNORMAL_THRESHOLD, normalize_status


class StoreQueryService:
    def __init__(self, context) -> None:
        self._context = context

    def _load_config(self) -> dict:
        return self._context.config_repository.load()

    def config_members(self) -> list[dict[str, str]]:
        self._context.config_repository.ensure_member_config(self._context.bootstrap_service.timestamp)
        return self._context.config_repository.load_members(normalize_status)

    def get_member(self, name: str) -> dict[str, str] | None:
        for member in self.get_members():
            if member["name"] == name:
                return member
        return None

    def get_quick_scores(self) -> list[int]:
        config = self._load_config()
        return [int(value) for value in config.get("quick_scores", DEFAULT_QUICK_SCORES)]

    def get_default_score_date(self) -> str:
        config = self._load_config()
        default_date = str(config.get("default_score_date") or "").strip()
        if default_date:
            try:
                return datetime.strptime(default_date, "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:
                return datetime.now().strftime("%Y-%m-%d")
        return datetime.now().strftime("%Y-%m-%d")

    def get_wear_abnormal_threshold(self) -> float:
        override_value = getattr(self._context, "_wear_abnormal_threshold_override", None)
        if override_value is not None:
            return float(override_value)
        config = self._load_config()
        raw_value = config.get("wear_abnormal_threshold", WEAR_ABNORMAL_THRESHOLD)
        try:
            return float(raw_value)
        except (TypeError, ValueError):
            return float(WEAR_ABNORMAL_THRESHOLD)

    def get_score_rankings(self, limit: int = 999, workbook=None):
        if self._context.read_only:
            return self._context._reader.get_score_rankings(limit=limit, workbook=workbook)
        rankings = self._context.local_db.get_score_summary_data()["rankings"]
        return rankings[:limit]

    def get_score_latest_column(self) -> str:
        if self._context.read_only:
            return self._context._reader.get_score_latest_column()
        return self._context.local_db.get_score_summary_data()["latest_column"]

    def get_score_column_count(self) -> int:
        return self._context._reader.get_score_column_count()

    def get_active_member_profit_map(self) -> dict[str, float]:
        return self._context._reader.get_active_member_profit_map()

    def get_score_summary_data(self):
        if self._context.read_only:
            return self._context._reader.get_score_summary_data()
        return self._context.local_db.get_score_summary_data()

    def get_score_sheet_snapshot(self):
        if self._context.read_only:
            return self._context._reader.get_score_sheet_snapshot()
        return self._context.local_db.get_score_sheet_snapshot(self.get_active_member_profit_map())

    def get_member_wear_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self._context._reader.get_member_wear_records(name, year_hint=year_hint, workbook=workbook)

    def get_member_income_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self._context._reader.get_member_income_records(name, year_hint=year_hint, workbook=workbook)

    def get_wear_sheet_snapshot(self):
        return self._context._reader.get_wear_sheet_snapshot()

    def get_member_calendar_dataset(self, name: str, year: int, month: int, workbook=None):
        return self._context._reader.get_member_calendar_dataset(name, year, month, workbook=workbook)

    def get_all_members_calendar_dataset(self, year: int, month: int):
        return self._context._reader.get_all_members_calendar_dataset(year, month)

    def get_next_score_date(self) -> str:
        if self._context.read_only:
            return self._context._reader.get_next_score_date()
        return self._context.local_db.get_next_score_date(self.get_default_score_date())

    def has_existing_score_date(self, score_date: str) -> bool:
        if self._context.read_only:
            return bool(self.online_score_rows_for_date(score_date))
        return bool(self._context.local_db.get_score_rows_for_date(score_date))

    def get_score_rows_for_date(self, score_date: str) -> list[dict]:
        if self._context.read_only:
            return self.online_score_rows_for_date(score_date)
        return self._context.local_db.get_score_rows_for_date(score_date)

    def get_score_summary(self):
        return self._context.score_presenter.build_score_summary()

    def get_score_sheet_view(self):
        return self._context.score_presenter.build_score_sheet_view()

    def get_wear_sheet_view(self):
        return self._context.wear_presenter.build_wear_sheet_view()

    def get_member_profit_calendar(self, name: str, year: int, month: int):
        return self._context.profit_calendar_presenter.build_member_profit_calendar(name=name, year=year, month=month)

    def get_active_members(self) -> list[dict[str, str]]:
        if self._context.read_only:
            return self._context.member_service.get_active_members()
        return [self._member_row_to_dict(item) for item in self._context.local_db.get_active_member_rows()]

    def get_members(self) -> list[dict[str, str]]:
        if self._context.read_only or not hasattr(self._context, "local_db"):
            return self.config_members()
        return [self._member_row_to_dict(item) for item in self._context.local_db.get_member_rows()]

    def _member_row_to_dict(self, item) -> dict[str, str]:
        return {
            "name": str(item["name"]),
            "status": str(item["status"]),
            "note": str(item["note"]),
            "created_at": str(item["created_at"]),
            "disabled_at": str(item["disabled_at"]),
            "sort_order": int(item["sort_order"]),
        }

    def _application_service(self):
        if getattr(self._context.application_service, "_supabase_client", None) is not self._context.supabase:
            self._context.application_service = self._context.bootstrap_service.create_online_application_service(self._context)
        return self._context.application_service

    def online_score_rows_for_date(self, score_date: str) -> list[dict]:
        return self._application_service().online_score_rows_for_date(score_date)

    def get_online_active_members(self) -> list[dict[str, str]]:
        return self._application_service().online_active_members()

    def get_online_score_summary_data(self) -> dict:
        return self._application_service().online_score_summary_data()

    def get_online_score_summary(self):
        return self._application_service().online_score_summary()

    def get_online_score_sheet_view(self):
        return self._application_service().online_score_sheet_view()

    def get_online_next_score_date(self) -> str:
        return self._application_service().online_next_score_date()

    def get_mobile_overview(self):
        return self._application_service().mobile_overview()
