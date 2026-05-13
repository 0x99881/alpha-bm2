from __future__ import annotations

from ..repositories.supabase_client import SupabaseSchemaError
from ..ui_text import MESSAGES


class StoreCommandService:
    def __init__(self, context) -> None:
        self._context = context

    def _load_config(self) -> dict:
        return self._context.config_repository.load()

    def _save_config(self, config: dict) -> None:
        self._context.config_repository.save(config)

    def _change_detection_snapshot(self) -> dict[str, object]:
        return self._context.local_db.get_change_detection_snapshot()

    def _delete_score_date_for_export(self, date_text: str) -> dict:
        return self._context.local_db.delete_score_entries_for_export(date_text, source="local")

    def _restore_score_date_export_delete(self, snapshot: dict) -> None:
        self._context.local_db.restore_score_date_export_delete(snapshot, source="local")

    def set_wear_abnormal_threshold(self, value_text: str) -> float:
        threshold = float(value_text.strip())
        config = self._load_config()
        previous_value = config.get("wear_abnormal_threshold")
        had_previous_value = "wear_abnormal_threshold" in config
        config["wear_abnormal_threshold"] = threshold
        self._save_config(config)
        try:
            self._context.workbook_repository.update_workbook(self._context.ensure_wear_sheet_structure)
        except (OSError, ValueError):
            rollback_config = self._load_config()
            if had_previous_value:
                rollback_config["wear_abnormal_threshold"] = previous_value
            else:
                rollback_config.pop("wear_abnormal_threshold", None)
            self._save_config(rollback_config)
            raise
        return threshold

    def refresh_local_database(self) -> dict[str, object]:
        before_snapshot = self._change_detection_snapshot()
        self._context.excel_import_service.refresh_from_excel()
        after_snapshot = self._change_detection_snapshot()
        return {
            "changed": before_snapshot != after_snapshot,
            "members": len(after_snapshot["members"]),
            "score_entries": len(after_snapshot["score_entries"]),
        }

    def save_daily_entry(
        self,
        form_data,
        selected_date: str,
        *,
        existing_notes: dict[str, str] | None = None,
        export_to_excel: bool = True,
    ) -> dict:
        if self._context.read_only:
            active_members = self._context.query_service.get_online_active_members()
            existing_notes = existing_notes or {"note1": "", "note2": "", "note3": ""}
        else:
            active_members = self._context.query_service.get_active_members()
            if existing_notes is None:
                existing_notes = self._context.query_service.get_score_date_notes(selected_date)
        submission = self._context.daily_entry_service.process_submission(
            active_members,
            form_data,
            selected_date,
            existing_notes=existing_notes,
        )
        submission["export_error"] = None
        if submission["ok"] and export_to_excel and not self._context.read_only:
            try:
                self._context.export_service.export_to_excel(submission["result"]["saved_date"])
            except (OSError, ValueError) as exc:
                submission["export_error"] = exc
        return submission

    def create_new_cycle(self, start_date_text: str) -> str:
        filename = self._context.workbook_repository.create_new_cycle(self._context, start_date_text)
        self._context.workbook_path = self._context.workbook_repository.workbook_path
        self.refresh_local_database()
        return filename

    def delete_score_date(self, date_text: str) -> dict[str, int]:
        delete_snapshot = self._delete_score_date_for_export(date_text)
        try:
            exported_dates = self._context.export_service.export_to_excel(date_text)
        except (OSError, ValueError):
            self._restore_score_date_export_delete(delete_snapshot)
            raise
        return {
            "deleted_rows": len(delete_snapshot["entry_ids"]),
            "deleted_notes": int(
                delete_snapshot["date_note"] is not None
                and int(delete_snapshot["date_note"].get("deleted", 0) or 0) == 0
            ),
            "exported_dates": exported_dates,
        }

    def is_supabase_configured(self) -> bool:
        return self._context.sync_service.is_configured()

    def supabase_push(self, *, force_full: bool = True) -> dict[str, int]:
        score_profit_map = self._context.query_service.get_active_member_profit_map()
        try:
            return self._context.sync_service.push(force_full=force_full, score_profit_map=score_profit_map)
        except SupabaseSchemaError as exc:
            raise ValueError(MESSAGES["supabase_schema_missing_columns"]) from exc

    def supabase_pull(self) -> dict[str, int]:
        return self._context.sync_service.pull()
