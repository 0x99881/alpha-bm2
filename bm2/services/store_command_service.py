from __future__ import annotations
class StoreCommandService:
    def __init__(self, context) -> None:
        self._context = context

    def _load_config(self) -> dict:
        return self._context.config_repository.load()

    def _save_config(self, config: dict) -> None:
        self._context.config_repository.save(config)

    def set_wear_abnormal_threshold(self, value_text: str) -> float:
        threshold = float(value_text.strip())
        self._context._wear_abnormal_threshold_override = threshold
        try:
            self._context.workbook_repository.update_workbook(self._context._ensure_wear_sheet_structure)
        finally:
            if hasattr(self._context, "_wear_abnormal_threshold_override"):
                delattr(self._context, "_wear_abnormal_threshold_override")

        config = self._load_config()
        config["wear_abnormal_threshold"] = threshold
        self._save_config(config)
        return threshold

    def refresh_local_database(self) -> dict[str, object]:
        before_members = self._context.local_db._get_all_member_rows_for_push()
        before_scores = self._context.local_db._get_all_score_rows_for_push()
        self._context.excel_import_service.refresh_from_excel()
        after_members = self._context.local_db._get_all_member_rows_for_push()
        after_scores = self._context.local_db._get_all_score_rows_for_push()
        return {
            "changed": before_members != after_members or before_scores != after_scores,
            "members": len(after_members),
            "score_entries": len(after_scores),
        }

    def create_new_cycle(self, start_date_text: str) -> str:
        filename = self._context.workbook_repository.create_new_cycle(self._context, start_date_text)
        self._context.workbook_path = self._context.workbook_repository.workbook_path
        self.refresh_local_database()
        return filename

    def is_supabase_configured(self) -> bool:
        return self._context.sync_service.is_configured()

    def supabase_push(self, *, force_full: bool = True) -> dict[str, int]:
        score_profit_map = self._context.query_service.get_active_member_profit_map()
        return self._context.sync_service.push(force_full=force_full, score_profit_map=score_profit_map)

    def supabase_pull(self) -> dict[str, int]:
        return self._context.sync_service.pull()
