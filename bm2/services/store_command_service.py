from __future__ import annotations

import logging
from datetime import datetime


LOGGER = logging.getLogger(__name__)


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

    def set_default_score_date(self, date_text: str) -> str:
        saved_date = datetime.strptime(date_text.strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
        config = self._load_config()
        config["default_score_date"] = saved_date
        self._save_config(config)
        return saved_date

    def save_scores_and_wear(self, date_text: str, entries: list[dict[str, str]]):
        if self._context._sqlite_writer is not None:
            return self._context._sqlite_writer.save_scores_and_wear(date_text, entries)
        raise RuntimeError("save_scores_and_wear called in read-only mode")

    def add_member(self, name: str, note: str = "") -> None:
        self._context.member_service.add_member(name, note)

    def update_member(self, name: str, note: str, status: str | None = None) -> None:
        self._context.member_service.update_member(name, note, status=status)

    def delete_member(self, name: str) -> None:
        self._context.member_service.delete_member(name)

    def reorder_active_members(self, ordered_names: list[str]) -> None:
        self._context.member_service.reorder_active_members(ordered_names)

    def refresh_local_database(self) -> dict[str, object]:
        before_members = self._context.local_db._get_all_member_rows_for_push()
        before_scores = self._context.local_db._get_all_score_rows_for_push()
        self._context.migration_service.refresh_from_legacy_store()
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

    def after_local_score_submission(self, result: dict) -> dict[str, object]:
        saved_date = str(result.get("saved_date") or "").strip()
        if self._context.read_only or not saved_date:
            return {"ok": False, "configured": False, "push": {"members": 0, "score_entries": 0}}
        if not self._context.sync_service.is_configured():
            return {"ok": False, "configured": False, "push": {"members": 0, "score_entries": 0}}
        try:
            push = self._context.sync_service.push()
        except Exception:
            LOGGER.exception("Supabase auto-push failed after local submission")
            return {"ok": False, "configured": True, "push": {"members": 0, "score_entries": 0}}
        return {"ok": True, "configured": True, "push": push}

    def is_supabase_configured(self) -> bool:
        return self._context.sync_service.is_configured()

    def supabase_push(self) -> dict[str, int]:
        return self._context.sync_service.push()

    def supabase_pull(self) -> dict[str, int]:
        return self._context.sync_service.pull()

    def supabase_sync(self) -> dict[str, dict[str, int]]:
        return self._context.sync_service.sync()
