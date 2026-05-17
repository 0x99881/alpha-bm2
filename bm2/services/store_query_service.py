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
        config = self._load_config()
        raw_value = config.get("wear_abnormal_threshold", WEAR_ABNORMAL_THRESHOLD)
        try:
            return float(raw_value)
        except (TypeError, ValueError):
            return float(WEAR_ABNORMAL_THRESHOLD)

    def _local_score_summary_data(self) -> dict:
        return self._context.local_db.get_score_summary_data()

    def _local_score_sheet_snapshot(self) -> dict:
        return self._context.local_db.get_score_sheet_snapshot(self.get_active_member_profit_map())

    def _local_score_rows_for_date(self, score_date: str) -> list[dict]:
        return self._context.local_db.get_score_rows_for_date(score_date)

    def _local_score_date_notes(self, score_date: str) -> dict[str, str]:
        return self._context.local_db.get_score_date_notes(score_date)

    def _local_next_score_date(self) -> str:
        return self._context.local_db.get_next_score_date(self.get_default_score_date())

    def _local_member_rows(self, *, active_only: bool) -> list:
        if active_only:
            return self._context.local_db.get_active_member_rows()
        return self._context.local_db.get_member_rows()

    def get_active_member_profit_map(self) -> dict[str, float]:
        return self._context.workbook_reader.get_active_member_profit_map()

    def get_score_summary_data(self):
        if self._context.read_only:
            return self._context.workbook_reader.get_score_summary_data()
        return self._local_score_summary_data()

    def get_score_sheet_snapshot(self):
        if self._context.read_only:
            return self._context.workbook_reader.get_score_sheet_snapshot()
        return self._local_score_sheet_snapshot()

    def get_member_wear_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self._context.workbook_reader.get_member_wear_records(name, year_hint=year_hint, workbook=workbook)

    def get_member_income_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self._context.workbook_reader.get_member_income_records(name, year_hint=year_hint, workbook=workbook)

    def get_wear_sheet_snapshot(self):
        return self._context.workbook_reader.get_wear_sheet_snapshot()

    def get_value_sheet_snapshot(self, sheet_type: str):
        return self._context.workbook_reader.get_value_sheet_snapshot(sheet_type)

    def get_member_calendar_dataset(self, name: str, year: int, month: int, workbook=None):
        return self._context.workbook_reader.get_member_calendar_dataset(name, year, month, workbook=workbook)

    def get_all_members_calendar_dataset(self, year: int, month: int):
        return self._context.workbook_reader.get_all_members_calendar_dataset(year, month)

    def get_next_score_date(self) -> str:
        if self._context.read_only:
            return self._context.workbook_reader.get_next_score_date()
        return self._local_next_score_date()

    def get_score_rows_for_date(self, score_date: str) -> list[dict]:
        if self._context.read_only:
            return self.online_score_rows_for_date(score_date)
        return self._local_score_rows_for_date(score_date)

    def get_score_date_notes(self, score_date: str) -> dict[str, str]:
        if self._context.read_only:
            return {"note1": "", "note2": "", "note3": ""}
        return self._local_score_date_notes(score_date)

    def get_score_summary(self):
        return self._context.score_presenter.build_score_summary()

    def get_score_sheet_view(self):
        return self._context.score_presenter.build_score_sheet_view()

    def get_wear_sheet_view(self):
        return self._context.wear_presenter.build_wear_sheet_view()

    def get_value_sheet_view(self, sheet_type: str, *, cycle_window=None):
        return self._context.value_sheet_presenter.build_sheet_view(
            sheet_type, cycle_window=cycle_window
        )

    def get_member_profit_calendar(self, name: str, year: int, month: int):
        return self._context.profit_calendar_presenter.build_member_profit_calendar(name=name, year=year, month=month)

    def get_member_profit_calendar_for_cycle(self, name: str, cycle_window: dict, cycles: list):
        return self._context.profit_calendar_presenter.build_member_profit_calendar_for_cycle(
            name=name, cycle_window=cycle_window, cycles=cycles,
        )

    def get_all_cycles(self) -> list[dict]:
        return self._context.local_db.get_settlement_cycles()

    def get_cycle_profit_view(self, cycle_id: str | None = None):
        return self._context.cycle_profit_presenter.build_cycle_profit_view(cycle_id)

    def get_active_members(self) -> list[dict[str, str]]:
        if self._context.read_only:
            return self._context.member_service.get_active_members()
        rows = self._local_member_rows(active_only=False)
        if not rows:
            return [item for item in self.config_members() if item["status"] == ENABLED]
        return [self._member_row_to_dict(item) for item in rows if str(item["status"]) == ENABLED]

    def get_members(self) -> list[dict[str, str]]:
        if self._context.read_only or not hasattr(self._context, "local_db"):
            return self.config_members()
        rows = self._local_member_rows(active_only=False)
        if not rows:
            return self.config_members()
        return [self._member_row_to_dict(item) for item in rows]

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
        if not self._context.application_service.uses_supabase_client(self._context.supabase):
            self._context.application_service = self._context.bootstrap_service.create_online_application_service(self._context)
        return self._context.application_service

    def online_score_rows_for_date(self, score_date: str) -> list[dict]:
        return self._application_service().online_score_rows_for_date(score_date)

    def get_online_active_members(self) -> list[dict[str, str]]:
        return self._application_service().online_active_members()

    def get_online_score_summary(self):
        return self._application_service().online_score_summary()

    def get_online_wear_sheet_view(self):
        return self._application_service().online_wear_sheet_view()

    def get_online_next_score_date(self) -> str:
        return self._application_service().online_next_score_date()

    def get_mobile_overview(self):
        return self._application_service().mobile_overview()
