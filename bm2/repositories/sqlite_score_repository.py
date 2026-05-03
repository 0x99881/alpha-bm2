from __future__ import annotations

from typing import Any


class SQLiteScoreEntryRepositoryMixin:
    @staticmethod
    def _entry_text(entry: dict[str, Any], key: str) -> str:
        return str(entry.get(key, "") or "").strip()

    def _default_zero_entry_text(self, entry: dict[str, Any], key: str) -> str:
        value = self._entry_text(entry, key)
        return value if value else "0"

    def _default_zero_wear_text(self, entry: dict[str, Any]) -> str:
        manual_wear = self._entry_text(entry, "manual_wear")
        if manual_wear:
            return manual_wear
        before_balance = self._entry_text(entry, "before_balance")
        after_balance = self._entry_text(entry, "after_balance")
        if before_balance or after_balance:
            return ""
        return "0"

    def replace_score_entries_from_snapshot(self, rows: list[dict[str, Any]], *, source: str = "local") -> None:
        if self.read_only:
            return
        connection = self._connect()
        try:
            updated_at = self._now_text()
            snapshot_keys: set[tuple[str, str]] = set()
            snapshot_dates: set[str] = set()

            for item in rows:
                member_name = str(item.get("member_name", "")).strip()
                score_date = str(item.get("score_date", "")).strip()
                if not member_name:
                    continue
                if not score_date:
                    continue
                snapshot_keys.add((member_name, score_date))
                snapshot_dates.add(score_date)
                member_id = self._member_id(member_name)
                try:
                    score_value = int(item.get("score", 0) or 0)
                except (TypeError, ValueError):
                    score_value = 0
                connection.execute(
                    """
                    INSERT INTO score_entries (
                        id, member_id, member_name, score_date, score,
                        before_balance, after_balance, manual_wear, income, other_expense,
                        updated_at, version, deleted, source
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 0, ?)
                    ON CONFLICT(member_name, score_date) DO UPDATE SET
                        score=excluded.score,
                        member_id=excluded.member_id,
                        before_balance=excluded.before_balance,
                        after_balance=excluded.after_balance,
                        manual_wear=excluded.manual_wear,
                        income=excluded.income,
                        other_expense=excluded.other_expense,
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
                        self._entry_text(item, "before_balance"),
                        self._entry_text(item, "after_balance"),
                        self._default_zero_wear_text(item),
                        self._default_zero_entry_text(item, "income"),
                        self._default_zero_entry_text(item, "other_expense"),
                        updated_at,
                        source,
                    ),
                )

            if snapshot_keys:
                min_snapshot_date = min(snapshot_dates)
                existing_rows = connection.execute(
                    """
                    SELECT member_name, score_date
                    FROM score_entries
                    WHERE deleted = 0 AND score_date >= ?
                    """,
                    (min_snapshot_date,),
                ).fetchall()
                for existing in existing_rows:
                    key = (str(existing["member_name"]).strip(), str(existing["score_date"]).strip())
                    if key in snapshot_keys:
                        continue
                    connection.execute(
                        """
                        UPDATE score_entries
                        SET deleted = 1,
                            updated_at = ?,
                            version = version + 1,
                            source = ?
                        WHERE member_name = ? AND score_date = ?
                        """,
                        (updated_at, source, key[0], key[1]),
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
                        id, member_id, member_name, score_date, score,
                        before_balance, after_balance, manual_wear, income, other_expense,
                        updated_at, version, deleted, source
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 0, ?)
                    ON CONFLICT(member_name, score_date) DO UPDATE SET
                        score=excluded.score,
                        member_id=excluded.member_id,
                        before_balance=excluded.before_balance,
                        after_balance=excluded.after_balance,
                        manual_wear=excluded.manual_wear,
                        income=excluded.income,
                        other_expense=excluded.other_expense,
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
                        self._entry_text(entry, "before_balance"),
                        self._entry_text(entry, "after_balance"),
                        self._default_zero_wear_text(entry),
                        self._default_zero_entry_text(entry, "income"),
                        self._default_zero_entry_text(entry, "other_expense"),
                        updated_at,
                        source,
                    ),
                )
            connection.commit()
        finally:
            connection.close()

    def delete_score_entries_for_date(self, score_date: str, *, source: str = "local") -> list[str]:
        if self.read_only:
            return []
        cleaned_date = str(score_date or "").strip()
        if not cleaned_date:
            return []
        updated_at = self._now_text()
        connection = self._connect()
        try:
            rows = connection.execute(
                """
                SELECT id
                FROM score_entries
                WHERE score_date = ? AND deleted = 0
                """,
                (cleaned_date,),
            ).fetchall()
            changed_ids = [str(row["id"]) for row in rows]
            if changed_ids:
                connection.execute(
                    """
                    UPDATE score_entries
                    SET deleted = 1,
                        updated_at = ?,
                        version = version + 1,
                        source = ?
                    WHERE score_date = ? AND deleted = 0
                    """,
                    (updated_at, source, cleaned_date),
                )
            connection.commit()
            return changed_ids
        finally:
            connection.close()

    def restore_score_entries_by_ids(self, entry_ids: list[str], *, source: str = "local") -> None:
        if self.read_only or not entry_ids:
            return
        updated_at = self._now_text()
        connection = self._connect()
        try:
            placeholders = ",".join("?" for _ in entry_ids)
            connection.execute(
                f"""
                UPDATE score_entries
                SET deleted = 0,
                    updated_at = ?,
                    version = version + 1,
                    source = ?
                WHERE id IN ({placeholders})
                """,
                [updated_at, source, *entry_ids],
            )
            connection.commit()
        finally:
            connection.close()

    def get_score_rows(self, *, include_deleted: bool = False) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            query = """
                SELECT
                    member_name, score_date, score,
                    before_balance, after_balance, manual_wear, income, other_expense,
                    updated_at, version, deleted, source
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
                SELECT
                    member_name, score_date, score,
                    before_balance, after_balance, manual_wear, income, other_expense,
                    updated_at, version, deleted, source
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

