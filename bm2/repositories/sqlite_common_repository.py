from __future__ import annotations

from datetime import datetime


class SQLiteCommonRepositoryMixin:
    def _now_text(self) -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _member_id(self, name: str) -> str:
        return f"member:{name.strip()}"

    def _score_entry_id(self, member_name: str, score_date: str) -> str:
        return f"score:{score_date}:{member_name.strip()}"
