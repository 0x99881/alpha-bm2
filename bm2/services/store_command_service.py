from __future__ import annotations

import logging
from datetime import datetime

from ..ui_text import MESSAGES


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
            result = self._context._sqlite_writer.save_scores_and_wear(date_text, entries)
            result["excel"] = self._write_entries_to_excel_safely(date_text, entries)
            return result
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
        excel = result.get("excel") or {"ok": True, "message": None}
        if self._context.read_only or not saved_date:
            return {"ok": False, "configured": False, "push": {"members": 0, "score_entries": 0}, "excel": excel}
        if not self._context.sync_service.is_configured():
            return {"ok": False, "configured": False, "push": {"members": 0, "score_entries": 0}, "excel": excel}
        return {"ok": True, "configured": True, "push": {"members": 0, "score_entries": 0}, "excel": excel}

    def write_entries_to_excel(self, date_text: str, entries: list[dict[str, str]]) -> dict[str, object]:
        return self._write_entries_to_excel_safely(date_text, entries)

    def _write_entries_to_excel_safely(self, date_text: str, entries: list[dict[str, str]]) -> dict[str, object]:
        if self._context.read_only:
            return {"ok": True, "exported": 0, "message": None}
        try:
            excel_result = self._context._writer.save_scores_and_wear(date_text, entries)
        except (ValueError, PermissionError, OSError) as exc:
            LOGGER.exception("Excel export failed after local submission")
            return {"ok": False, "exported": 0, "message": str(exc)}
        except Exception:
            LOGGER.exception("Excel export failed after local submission")
            workbook_name = getattr(self._context.workbook_path, "name", "")
            return {"ok": False, "exported": 0, "message": MESSAGES["excel_save_failed"].format(filename=workbook_name)}
        return {"ok": True, "exported": 1, "message": None, "result": excel_result}

    def is_supabase_configured(self) -> bool:
        return self._context.sync_service.is_configured()

    def supabase_push(self, *, force_full: bool = True) -> dict[str, int]:
        score_profit_map = self._context.query_service.get_active_member_profit_map()
        return self._context.sync_service.push(force_full=force_full, score_profit_map=score_profit_map)

    def supabase_pull(self) -> dict[str, int]:
        return self._context.sync_service.pull()

    def supabase_sync(self) -> dict[str, dict[str, int]]:
        return self._context.sync_service.sync()
