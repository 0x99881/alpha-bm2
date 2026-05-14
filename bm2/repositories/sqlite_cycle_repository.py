from __future__ import annotations

import uuid
from typing import Any


class SQLiteCycleRepositoryMixin:
    """Settlement cycle storage for the 周期盈亏情况 feature.

    A settlement cycle is a named period kept as history. Each member has one
    manual settlement row per cycle: start/settle dates and balances. Wear and
    red-packet amounts are not stored here; they are summed live from
    score_entries within each member's date range.
    """

    @staticmethod
    def _cycle_text(value: Any) -> str:
        return str(value or "").strip()

    def create_settlement_cycle(self, name: str) -> str:
        cycle_id = uuid.uuid4().hex
        if self.read_only:
            return cycle_id
        now = self._now_text()
        connection = self._connect()
        try:
            connection.execute(
                """
                INSERT INTO settlement_cycles (id, name, created_at, updated_at, deleted)
                VALUES (?, ?, ?, ?, 0)
                """,
                (cycle_id, self._cycle_text(name), now, now),
            )
            connection.commit()
        finally:
            connection.close()
        return cycle_id

    def get_settlement_cycles(self) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            rows = connection.execute(
                """
                SELECT id, name, created_at, updated_at
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
                SELECT id, name, created_at, updated_at
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
                SELECT member_name, start_date, start_balance, settle_date, end_balance
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
                        settle_date, end_balance, updated_at, version, deleted, source
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, 1, 0, 'local')
                    ON CONFLICT(cycle_id, member_name) DO UPDATE SET
                        start_date=excluded.start_date,
                        start_balance=excluded.start_balance,
                        settle_date=excluded.settle_date,
                        end_balance=excluded.end_balance,
                        updated_at=excluded.updated_at,
                        version=settlement_entries.version + 1,
                        deleted=0,
                        source=excluded.source
                    """,
                    (
                        cleaned_cycle,
                        member_name,
                        self._cycle_text(entry.get("start_date")),
                        self._cycle_text(entry.get("start_balance")),
                        self._cycle_text(entry.get("settle_date")),
                        self._cycle_text(entry.get("end_balance")),
                        updated_at,
                    ),
                )
            connection.commit()
        finally:
            connection.close()
