from __future__ import annotations

from datetime import datetime
from typing import Any

from ..constants import WINDOW_SIZE


class SupabaseEntryWriter:
    """Entry writer used by the deployed/mobile app.

    Online submissions go directly to Supabase; the legacy JSON blob path is
    intentionally bypassed.
    """

    def __init__(self, supabase_sync) -> None:
        self._supabase = supabase_sync

    @staticmethod
    def _member_id(name: str) -> str:
        return f"member:{name.strip()}"

    @staticmethod
    def _score_entry_id(member_name: str, score_date: str) -> str:
        return f"score:{score_date}:{member_name.strip()}"

    @staticmethod
    def _now_text() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def _score_value(entry: dict[str, str]) -> int:
        raw_value = str(entry.get("score", "") or "0").strip()
        return int(raw_value or 0)

    @staticmethod
    def _count_wear_entries(entries: list[dict[str, str]]) -> int:
        count = 0
        for entry in entries:
            has_manual = str(entry.get("manual_wear", "")).strip()
            has_before = str(entry.get("before_balance", "")).strip()
            has_after = str(entry.get("after_balance", "")).strip()
            if has_manual or (has_before and has_after):
                count += 1
        return count

    def save_scores_and_wear(
        self,
        date_text: str,
        entries: list[dict[str, str]],
    ) -> dict[str, Any]:
        if not self._supabase.is_configured():
            raise ValueError("线上数据库未配置，暂时不能保存。")

        saved_date = datetime.strptime(date_text.strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
        now_text = self._now_text()
        ids = [
            self._score_entry_id(str(entry.get("name", "")).strip(), saved_date)
            for entry in entries
            if str(entry.get("name", "")).strip()
        ]
        try:
            existing = {
                str(row["id"]): row
                for row in self._supabase.fetch_score_entries_by_ids(ids)
            }
        except Exception as exc:
            raise ValueError("线上数据库暂时连接失败，请稍后再试。") from exc

        rows: list[dict[str, Any]] = []
        for entry in entries:
            member_name = str(entry.get("name", "")).strip()
            if not member_name:
                continue
            row_id = self._score_entry_id(member_name, saved_date)
            previous = existing.get(row_id) or {}
            rows.append(
                {
                    "id": row_id,
                    "member_id": self._member_id(member_name),
                    "member_name": member_name,
                    "score_date": saved_date,
                    "score": self._score_value(entry),
                    "updated_at": now_text,
                    "version": int(previous.get("version", 0) or 0) + 1,
                    "deleted": 0,
                    "source": "online",
                }
            )

        try:
            self._supabase.push_score_entries(rows)
        except Exception as exc:
            raise ValueError("线上数据库暂时连接失败，请稍后再试。") from exc
        return {
            "target_column": saved_date[5:],
            "wear_column": saved_date[5:].replace("-", ""),
            "wear_rows_added": self._count_wear_entries(entries),
            "window_size": WINDOW_SIZE,
            "saved_date": saved_date,
        }
