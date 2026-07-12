from __future__ import annotations

from decimal import Decimal, InvalidOperation
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

    @staticmethod
    def _numeric_equal(left: str, right: str) -> bool:
        try:
            return Decimal(str(left or "0")) == Decimal(str(right or "0"))
        except (InvalidOperation, ValueError):
            return False

    def _preserve_numeric_text(self, existing_value: Any, incoming_value: str) -> str:
        existing_text = self._entry_text({"value": existing_value}, "value")
        if existing_text and self._numeric_equal(existing_text, incoming_value):
            return existing_text
        return incoming_value

    def _snapshot_detail_values(
        self,
        item: dict[str, Any],
        existing: dict[str, Any] | None,
    ) -> dict[str, str]:
        before_balance = (
            self._entry_text(item, "before_balance")
            if "before_balance" in item
            else self._entry_text(existing or {}, "before_balance")
        )
        after_balance = (
            self._entry_text(item, "after_balance")
            if "after_balance" in item
            else self._entry_text(existing or {}, "after_balance")
        )
        existing_has_balance = bool(
            self._entry_text(existing or {}, "before_balance")
            or self._entry_text(existing or {}, "after_balance")
        )
        snapshot_has_balance = bool(
            self._entry_text(item, "before_balance")
            or self._entry_text(item, "after_balance")
        )
        if existing_has_balance and not snapshot_has_balance:
            manual_wear = self._entry_text(existing or {}, "manual_wear")
        elif "manual_wear" in item:
            manual_wear = self._preserve_numeric_text(
                (existing or {}).get("manual_wear", ""),
                self._default_zero_wear_text(item),
            )
        else:
            # Reader didn't surface this date's wear (e.g. it lives in a
            # historical cycle block that the active-block-only reader
            # doesn't scan). Preserve whatever the DB already has so refresh
            # doesn't silently zero out historical entries.
            manual_wear = self._entry_text(existing or {}, "manual_wear")

        if "income" in item:
            income = self._preserve_numeric_text(
                (existing or {}).get("income", ""),
                self._default_zero_entry_text(item, "income"),
            )
        else:
            income = self._entry_text(existing or {}, "income")

        if "other_expense" in item:
            other_expense = self._preserve_numeric_text(
                (existing or {}).get("other_expense", ""),
                self._default_zero_entry_text(item, "other_expense"),
            )
        else:
            other_expense = self._entry_text(existing or {}, "other_expense")

        return {
            "before_balance": before_balance,
            "after_balance": after_balance,
            "manual_wear": manual_wear,
            "income": income,
            "other_expense": other_expense,
            "profit": (
                self._entry_text(item, "profit")
                if "profit" in item
                else self._entry_text(existing or {}, "profit") or "0"
            ),
        }

    def _record_detail_values(
        self,
        entry: dict[str, Any],
        existing: dict[str, Any] | None,
    ) -> dict[str, str]:
        before_balance = self._entry_text(entry, "before_balance")
        after_balance = self._entry_text(entry, "after_balance")
        manual_wear_input = self._entry_text(entry, "manual_wear")

        if manual_wear_input:
            manual_wear = self._preserve_numeric_text((existing or {}).get("manual_wear", ""), manual_wear_input)
        elif before_balance or after_balance:
            manual_wear = ""
        else:
            manual_wear = ""

        def _detail_text(key: str) -> str:
            incoming_value = self._entry_text(entry, key)
            if incoming_value:
                return self._preserve_numeric_text((existing or {}).get(key, ""), incoming_value)
            return ""

        return {
            "before_balance": before_balance,
            "after_balance": after_balance,
            "manual_wear": manual_wear,
            "income": _detail_text("income"),
            "other_expense": _detail_text("other_expense"),
            "profit": self._entry_text(existing or {}, "profit") or "0",
        }

    def _score_entry_changed(self, existing: dict[str, Any], incoming: dict[str, Any]) -> bool:
        if int(existing.get("score", 0) or 0) != int(incoming["score"]):
            return True
        if int(existing.get("deleted", 0) or 0) != 0:
            return True
        for key in (
            "member_id",
            "before_balance",
            "after_balance",
            "manual_wear",
            "income",
            "other_expense",
            "profit",
        ):
            if self._entry_text(existing, key) != self._entry_text(incoming, key):
                return True
        return False

    def replace_score_entries_from_snapshot(self, rows: list[dict[str, Any]], *, source: str = "local") -> None:
        if self.read_only:
            return
        connection = self._connect()
        try:
            updated_at = self._now_text()
            existing_rows = {
                (str(row["member_name"]).strip(), str(row["score_date"]).strip()): dict(row)
                for row in connection.execute(
                    """
                    SELECT
                        id, member_id, member_name, score_date, score,
                        before_balance, after_balance, manual_wear, income, other_expense, profit,
                        updated_at, version, deleted, source
                    FROM score_entries
                    """
                ).fetchall()
            }

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
                existing = existing_rows.get((member_name, score_date))
                incoming = {
                    "id": self._score_entry_id(member_name, score_date),
                    "member_id": member_id,
                    "member_name": member_name,
                    "score_date": score_date,
                    "score": score_value,
                    **self._snapshot_detail_values(item, existing),
                    "updated_at": updated_at,
                    "source": source,
                }
                if existing is None:
                    connection.execute(
                        """
                        INSERT INTO score_entries (
                            id, member_id, member_name, score_date, score,
                            before_balance, after_balance, manual_wear, income, other_expense, profit,
                            updated_at, version, deleted, source
                        ) VALUES (
                            :id, :member_id, :member_name, :score_date, :score,
                            :before_balance, :after_balance, :manual_wear, :income, :other_expense, :profit,
                            :updated_at, 1, 0, :source
                        )
                        """,
                        incoming,
                    )
                    continue
                if not self._score_entry_changed(existing, incoming):
                    continue
                connection.execute(
                    """
                    UPDATE score_entries
                    SET member_id = :member_id,
                        score = :score,
                        before_balance = :before_balance,
                        after_balance = :after_balance,
                        manual_wear = :manual_wear,
                        income = :income,
                        other_expense = :other_expense,
                        profit = :profit,
                        updated_at = :updated_at,
                        version = version + 1,
                        deleted = 0,
                        source = :source
                    WHERE member_name = :member_name AND score_date = :score_date
                    """,
                    incoming,
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
            existing_rows = {
                str(row["member_name"]).strip(): dict(row)
                for row in connection.execute(
                    """
                    SELECT
                        id, member_id, member_name, score_date, score,
                        before_balance, after_balance, manual_wear, income, other_expense, profit,
                        updated_at, version, deleted, source
                    FROM score_entries
                    WHERE score_date = ?
                    """,
                    (saved_date,),
                ).fetchall()
            }
            for entry in entries:
                member_name = str(entry.get("name", "")).strip()
                if not member_name:
                    continue
                try:
                    score_value = int(str(entry.get("score", "") or "0").strip() or 0)
                except ValueError:
                    score_value = 0
                existing = existing_rows.get(member_name)
                incoming = {
                    "id": self._score_entry_id(member_name, saved_date),
                    "member_id": self._member_id(member_name),
                    "member_name": member_name,
                    "score_date": saved_date,
                    "score": score_value,
                    **self._record_detail_values(entry, existing),
                    "updated_at": updated_at,
                    "source": source,
                }
                if existing is not None and not self._score_entry_changed(existing, incoming):
                    continue
                if existing is None:
                    connection.execute(
                        """
                        INSERT INTO score_entries (
                            id, member_id, member_name, score_date, score,
                            before_balance, after_balance, manual_wear, income, other_expense, profit,
                            updated_at, version, deleted, source
                        ) VALUES (
                            :id, :member_id, :member_name, :score_date, :score,
                            :before_balance, :after_balance, :manual_wear, :income, :other_expense, :profit,
                            :updated_at, 1, 0, :source
                        )
                        """,
                        incoming,
                    )
                    continue
                connection.execute(
                    """
                    UPDATE score_entries
                    SET
                        score = :score,
                        member_id = :member_id,
                        before_balance = :before_balance,
                        after_balance = :after_balance,
                        manual_wear = :manual_wear,
                        income = :income,
                        other_expense = :other_expense,
                        profit = :profit,
                        updated_at = :updated_at,
                        version = version + 1,
                        deleted=0,
                        source = :source
                    WHERE member_name = :member_name AND score_date = :score_date
                    """,
                    incoming,
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

    def delete_score_entries_for_export(self, score_date: str, *, source: str = "local") -> dict[str, Any]:
        return {
            "entry_ids": self.delete_score_entries_for_date(score_date, source=source),
            "date_note": self.delete_score_date_notes(score_date, source=source),
        }

    def restore_score_date_export_delete(self, snapshot: dict[str, Any], *, source: str = "local") -> None:
        self.restore_score_entries_by_ids(list(snapshot.get("entry_ids", [])), source=source)
        self.restore_score_date_note(snapshot.get("date_note"), source=source)

    def get_score_rows(self, *, include_deleted: bool = False) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            query = """
                SELECT
                    member_name, score_date, score,
                    before_balance, after_balance, manual_wear, income, other_expense, profit,
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
                    before_balance, after_balance, manual_wear, income, other_expense, profit,
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

