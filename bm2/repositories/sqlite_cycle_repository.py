from __future__ import annotations

import uuid
from typing import Any


class SQLiteCycleRepositoryMixin:
    """Settlement cycle storage for the 周期盈亏情况 feature.

    A cycle has a single start_date and (after settling) a single settle_date.
    Per-member rows hold only manual start/end balances; wear/red-packet/income
    are summed live from score_entries within the cycle date range.
    """

    @staticmethod
    def _cycle_text(value: Any) -> str:
        return str(value or "").strip()

    @staticmethod
    def _short_cycle_name(start_date_text: str) -> str:
        mmdd = start_date_text[5:] if len(start_date_text) >= 10 else start_date_text
        return f"周期{mmdd}"

    def create_settlement_cycle(self, start_date: str) -> str:
        cleaned_start = self._cycle_text(start_date)
        cycle_id = uuid.uuid4().hex
        if self.read_only:
            return cycle_id
        now = self._now_text()
        connection = self._connect()
        try:
            connection.execute(
                """
                INSERT INTO settlement_cycles (
                    id, name, start_date, settle_date, settled,
                    created_at, updated_at, version, deleted, source
                ) VALUES (?, ?, ?, '', 0, ?, ?, 1, 0, 'local')
                """,
                (cycle_id, self._short_cycle_name(cleaned_start), cleaned_start, now, now),
            )
            connection.commit()
        finally:
            connection.close()
        return cycle_id

    def delete_settlement_cycle(self, cycle_id: str) -> None:
        if self.read_only:
            return
        cleaned_cycle = self._cycle_text(cycle_id)
        if not cleaned_cycle:
            return
        now = self._now_text()
        connection = self._connect()
        try:
            connection.execute(
                "UPDATE settlement_cycles "
                "SET deleted = 1, updated_at = ?, version = version + 1 "
                "WHERE id = ?",
                (now, cleaned_cycle),
            )
            connection.commit()
        finally:
            connection.close()

    def settle_settlement_cycle(self, cycle_id: str, settle_date: str) -> None:
        if self.read_only:
            return
        cleaned_cycle = self._cycle_text(cycle_id)
        cleaned_settle = self._cycle_text(settle_date)
        if not cleaned_cycle:
            return
        now = self._now_text()
        connection = self._connect()
        try:
            connection.execute(
                """
                UPDATE settlement_cycles
                SET settle_date = ?, settled = 1, updated_at = ?,
                    version = version + 1
                WHERE id = ? AND deleted = 0
                """,
                (cleaned_settle, now, cleaned_cycle),
            )
            connection.commit()
        finally:
            connection.close()

    def unsettle_settlement_cycle(self, cycle_id: str) -> None:
        """Revert a settled cycle back to unsettled state. Used for rollback."""
        if self.read_only:
            return
        cleaned_cycle = self._cycle_text(cycle_id)
        if not cleaned_cycle:
            return
        now = self._now_text()
        connection = self._connect()
        try:
            connection.execute(
                """
                UPDATE settlement_cycles
                SET settle_date = '', settled = 0, updated_at = ?,
                    version = version + 1
                WHERE id = ? AND deleted = 0
                """,
                (now, cleaned_cycle),
            )
            connection.commit()
        finally:
            connection.close()

    def get_settlement_cycles(self) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            rows = connection.execute(
                """
                SELECT id, name, start_date, settle_date, settled, created_at, updated_at
                FROM settlement_cycles
                WHERE deleted = 0
                ORDER BY created_at DESC, name DESC
                """
            ).fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

    def get_settlement_cycle(self, cycle_id: str) -> dict[str, Any] | None:
        connection = self._connect()
        try:
            row = connection.execute(
                """
                SELECT id, name, start_date, settle_date, settled, created_at, updated_at
                FROM settlement_cycles
                WHERE id = ? AND deleted = 0
                """,
                (self._cycle_text(cycle_id),),
            ).fetchone()
            return dict(row) if row else None
        finally:
            connection.close()

    def get_settlement_entries(self, cycle_id: str) -> dict[str, dict[str, Any]]:
        connection = self._connect()
        try:
            rows = connection.execute(
                """
                SELECT member_name, start_balance, end_balance, is_extra, sort_order
                FROM settlement_entries
                WHERE cycle_id = ? AND deleted = 0
                """,
                (self._cycle_text(cycle_id),),
            ).fetchall()
            return {str(row["member_name"]).strip(): dict(row) for row in rows}
        finally:
            connection.close()

    def save_settlement_entries(self, cycle_id: str, entries: list[dict[str, str]]) -> None:
        if self.read_only:
            return
        cleaned_cycle = self._cycle_text(cycle_id)
        if not cleaned_cycle:
            return
        updated_at = self._now_text()
        connection = self._connect()
        try:
            for entry in entries:
                member_name = self._cycle_text(entry.get("member_name"))
                if not member_name:
                    continue
                connection.execute(
                    """
                INSERT INTO settlement_entries (
                    cycle_id, member_name, start_date, start_balance,
                    settle_date, end_balance, updated_at, version, deleted, source,
                    is_extra
                ) VALUES (?, ?, '', ?, '', ?, ?, 1, 0, 'local', ?)
                ON CONFLICT(cycle_id, member_name) DO UPDATE SET
                        start_balance=excluded.start_balance,
                        end_balance=excluded.end_balance,
                        is_extra=excluded.is_extra,
                        updated_at=excluded.updated_at,
                        version=settlement_entries.version + 1,
                        deleted=0,
                        source=excluded.source
                    """,
                    (
                        cleaned_cycle,
                        member_name,
                        self._cycle_text(entry.get("start_balance")),
                        self._cycle_text(entry.get("end_balance")),
                        updated_at,
                        int(entry.get("is_extra", 0) or 0),
                    ),
                )
            connection.commit()
        finally:
            connection.close()

    def add_cycle_extra_member(self, cycle_id: str, member_name: str) -> bool:
        if self.read_only:
            return False
        cleaned_cycle = self._cycle_text(cycle_id)
        cleaned_name = self._cycle_text(member_name)
        if not cleaned_cycle or not cleaned_name:
            return False
        now = self._now_text()
        connection = self._connect()
        try:
            existing = connection.execute(
                "SELECT is_extra, deleted FROM settlement_entries WHERE cycle_id = ? AND member_name = ?",
                (cleaned_cycle, cleaned_name),
            ).fetchone()
            if existing and not int(existing["deleted"] or 0):
                return False
            max_row = connection.execute(
                "SELECT COALESCE(MAX(sort_order), 0) AS max_order FROM settlement_entries WHERE cycle_id = ?",
                (cleaned_cycle,),
            ).fetchone()
            next_order = int(max_row["max_order"] or 0) + 1
            connection.execute(
                """
                INSERT INTO settlement_entries (
                    cycle_id, member_name, start_date, start_balance,
                    settle_date, end_balance, updated_at, version, deleted, source,
                    is_extra, sort_order
                ) VALUES (?, ?, '', '', '', '', ?, 1, 0, 'local', 1, ?)
                ON CONFLICT(cycle_id, member_name) DO UPDATE SET
                    deleted=0, is_extra=1, sort_order=excluded.sort_order,
                    updated_at=excluded.updated_at,
                    version=settlement_entries.version + 1
                """,
                (cleaned_cycle, cleaned_name, now, next_order),
            )
            connection.commit()
            return True
        finally:
            connection.close()

    def remove_cycle_extra_member(self, cycle_id: str, member_name: str) -> bool:
        if self.read_only:
            return False
        cleaned_cycle = self._cycle_text(cycle_id)
        cleaned_name = self._cycle_text(member_name)
        if not cleaned_cycle or not cleaned_name:
            return False
        now = self._now_text()
        connection = self._connect()
        try:
            cursor = connection.execute(
                """
                UPDATE settlement_entries
                SET deleted = 1, updated_at = ?, version = version + 1
                WHERE cycle_id = ? AND member_name = ? AND is_extra = 1 AND deleted = 0
                """,
                (now, cleaned_cycle, cleaned_name),
            )
            connection.commit()
            return cursor.rowcount > 0
        finally:
            connection.close()

    def set_cycle_member_order(self, cycle_id: str, ordered_names: list[str]) -> None:
        if self.read_only:
            return
        cleaned_cycle = self._cycle_text(cycle_id)
        if not cleaned_cycle:
            return
        now = self._now_text()
        connection = self._connect()
        try:
            existing_names = {
                str(row["member_name"]).strip()
                for row in connection.execute(
                    "SELECT member_name FROM settlement_entries WHERE cycle_id = ? AND deleted = 0",
                    (cleaned_cycle,),
                ).fetchall()
            }
            for index, name in enumerate(ordered_names, start=1):
                cleaned_name = self._cycle_text(name)
                if not cleaned_name:
                    continue
                if cleaned_name in existing_names:
                    connection.execute(
                        """
                        UPDATE settlement_entries
                        SET sort_order = ?, updated_at = ?, version = version + 1
                        WHERE cycle_id = ? AND member_name = ?
                        """,
                        (index, now, cleaned_cycle, cleaned_name),
                    )
                else:
                    connection.execute(
                        """
                        INSERT INTO settlement_entries (
                            cycle_id, member_name, start_date, start_balance,
                            settle_date, end_balance, updated_at, version, deleted,
                            source, is_extra, sort_order
                        ) VALUES (?, ?, '', '', '', '', ?, 1, 0, 'local', 0, ?)
                        """,
                        (cleaned_cycle, cleaned_name, now, index),
                    )
            connection.commit()
        finally:
            connection.close()
