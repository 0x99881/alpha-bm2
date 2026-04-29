from __future__ import annotations

from typing import Any


class SQLiteScoreEntryRepositoryMixin:
    def replace_score_entries_from_snapshot(self, rows: list[dict[str, Any]], *, source: str = "local") -> None:
        if self.read_only:
            return
        connection = self._connect()
        try:
            updated_at = self._now_text()
            seen_keys: set[tuple[str, str]] = set()

            for item in rows:
                member_name = str(item.get("member_name", "")).strip()
                score_date = str(item.get("score_date", "")).strip()
                if not member_name:
                    continue
                if not score_date:
                    continue
                member_id = self._member_id(member_name)
                try:
                    score_value = int(item.get("score", 0) or 0)
                except (TypeError, ValueError):
                    score_value = 0
                seen_keys.add((member_name, score_date))
                connection.execute(
                    """
                    INSERT INTO score_entries (
                        id, member_id, member_name, score_date, score, updated_at, version, deleted, source
                    ) VALUES (?, ?, ?, ?, ?, ?, 1, 0, ?)
                    ON CONFLICT(member_name, score_date) DO UPDATE SET
                        score=excluded.score,
                        member_id=excluded.member_id,
                        updated_at=excluded.updated_at,
                        version=score_entries.version + 1,
                        deleted=0,
                        source=excluded.source
                    """,
                    (
                        self._score_entry_id(member_name, score_date),
                        member_id,
                        member_name,
                        score_date,
                        score_value,
                        updated_at,
                        source,
                    ),
                )

            if seen_keys:
                connection.execute("UPDATE score_entries SET deleted = 1, updated_at = ?, version = version + 1 WHERE deleted = 0", (updated_at,))
                for member_name, score_date in seen_keys:
                    connection.execute(
                        """
                        UPDATE score_entries
                        SET deleted = 0
                        WHERE member_name = ? AND score_date = ?
                        """,
                        (member_name, score_date),
                    )
            connection.commit()
        finally:
            connection.close()

    def record_score_entries(self, saved_date: str, entries: list[dict[str, str]], *, source: str = "local") -> None:
        if self.read_only:
            return
        updated_at = self._now_text()
        connection = self._connect()
        try:
            for entry in entries:
                member_name = str(entry.get("name", "")).strip()
                if not member_name:
                    continue
                try:
                    score_value = int(str(entry.get("score", "") or "0").strip() or 0)
                except ValueError:
                    score_value = 0
                connection.execute(
                    """
                    INSERT INTO score_entries (
                        id, member_id, member_name, score_date, score, updated_at, version, deleted, source
                    ) VALUES (?, ?, ?, ?, ?, ?, 1, 0, ?)
                    ON CONFLICT(member_name, score_date) DO UPDATE SET
                        score=excluded.score,
                        member_id=excluded.member_id,
                        updated_at=excluded.updated_at,
                        version=score_entries.version + 1,
                        deleted=0,
                        source=excluded.source
                    """,
                    (
                        self._score_entry_id(member_name, saved_date),
                        self._member_id(member_name),
                        member_name,
                        saved_date,
                        score_value,
                        updated_at,
                        source,
                    ),
                )
            connection.commit()
        finally:
            connection.close()

    def get_score_rows(self, *, include_deleted: bool = False) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            query = """
                SELECT member_name, score_date, score, updated_at, version, deleted, source
                FROM score_entries
            """
            if not include_deleted:
                query += " WHERE deleted = 0"
            query += " ORDER BY score_date ASC, member_name ASC"
            rows = connection.execute(query).fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

    def get_score_rows_for_date(
        self,
        score_date: str,
        *,
        include_deleted: bool = False,
    ) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            query = """
                SELECT member_name, score_date, score, updated_at, version, deleted, source
                FROM score_entries
                WHERE score_date = ?
            """
            params: list[Any] = [score_date]
            if not include_deleted:
                query += " AND deleted = 0"
            query += " ORDER BY member_name ASC"
            rows = connection.execute(query, params).fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

