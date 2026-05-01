from __future__ import annotations

from datetime import datetime
from typing import Any

from ..constants import WINDOW_SIZE
from ..entry_helpers import count_wear_entries, member_id, score_entry_id
from ..source_metadata import parse_source_profit, with_source_profit


class SupabaseEntryWriter:
    """Entry writer used by the deployed/mobile app.

    Online submissions go directly to Supabase.
    """

    def __init__(self, supabase_sync) -> None:
        self._supabase = supabase_sync

    @staticmethod
    def _now_text() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def _score_value(entry: dict[str, str]) -> int:
        raw_value = str(entry.get("score", "") or "0").strip()
        return int(raw_value or 0)

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
            score_entry_id(str(entry.get("name", "")).strip(), saved_date)
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
            row_id = score_entry_id(member_name, saved_date)
            previous = existing.get(row_id) or {}
            previous_profit = parse_source_profit(previous.get("source", ""))
            source = "online" if previous_profit is None else with_source_profit("online", previous_profit)
            rows.append(
                {
                    "id": row_id,
                    "member_id": member_id(member_name),
                    "member_name": member_name,
                    "score_date": saved_date,
                    "score": self._score_value(entry),
                    "updated_at": now_text,
                    "version": int(previous.get("version", 0) or 0) + 1,
                    "deleted": 0,
                    "source": source,
                }
            )

        try:
            self._supabase.push_score_entries(rows)
        except Exception as exc:
            raise ValueError("线上数据库暂时连接失败，请稍后再试。") from exc
        return {
            "target_column": saved_date[5:],
            "wear_column": saved_date[5:].replace("-", ""),
            "wear_rows_added": count_wear_entries(entries),
            "window_size": WINDOW_SIZE,
            "saved_date": saved_date,
        }
