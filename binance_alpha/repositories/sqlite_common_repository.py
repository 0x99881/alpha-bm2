from __future__ import annotations

from datetime import datetime

from ..entry_helpers import member_id, score_entry_id


class SQLiteCommonRepositoryMixin:
    def _now_text(self) -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _member_id(self, name: str) -> str:
        return member_id(name)

    def _score_entry_id(self, member_name: str, score_date: str) -> str:
        return score_entry_id(member_name, score_date)
