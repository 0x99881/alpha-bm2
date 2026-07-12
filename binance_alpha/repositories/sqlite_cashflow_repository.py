from __future__ import annotations

import uuid
from typing import Any


class SQLiteCashFlowRepositoryMixin:
    """Personal cash-flow ledger (出金/入金 记录).

    A flat append-only ledger: every entry is one row keyed by a random id, so
    two identical entries on the same day (e.g. 两笔 200 红包) never collide.
    Deletes are soft (``deleted = 1``) so a mis-tap never destroys history —
    the same caution the rest of the project uses for real financial data.

    This table intentionally has no version/source columns: the ledger is
    local-only and does not participate in Supabase sync.
    """

    CASH_FLOW_COLUMNS = (
        "id", "entry_date", "direction", "amount", "category", "note",
        "created_at", "updated_at", "deleted",
    )

    @staticmethod
    def _cash_flow_text(value: Any) -> str:
        return str(value or "").strip()

    def add_cash_flow(
        self,
        *,
        entry_date: str,
        direction: str,
        amount: str,
        category: str,
        note: str,
    ) -> str:
        if self.read_only:
            return ""
        flow_id = uuid.uuid4().hex
        now = self._now_text()
        connection = self._connect()
        try:
            connection.execute(
                """
                INSERT INTO cash_flows (
                    id, entry_date, direction, amount, category, note,
                    created_at, updated_at, deleted
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0)
                """,
                (
                    flow_id,
                    self._cash_flow_text(entry_date),
                    self._cash_flow_text(direction) or "out",
                    self._cash_flow_text(amount),
                    self._cash_flow_text(category),
                    self._cash_flow_text(note),
                    now,
                    now,
                ),
            )
            connection.commit()
        finally:
            connection.close()
        return flow_id

    def list_cash_flows(self, *, include_deleted: bool = False) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            where = "" if include_deleted else "WHERE deleted = 0"
            rows = connection.execute(
                f"""
                SELECT {', '.join(self.CASH_FLOW_COLUMNS)}
                FROM cash_flows
                {where}
                ORDER BY entry_date DESC, created_at DESC, id DESC
                """
            ).fetchall()
        finally:
            connection.close()
        return [dict(row) for row in rows]

    def soft_delete_cash_flow(self, flow_id: str) -> bool:
        """Mark a ledger entry deleted. Returns True if a live row was hit."""
        if self.read_only:
            return False
        cleaned = self._cash_flow_text(flow_id)
        if not cleaned:
            return False
        now = self._now_text()
        connection = self._connect()
        try:
            cursor = connection.execute(
                "UPDATE cash_flows SET deleted = 1, updated_at = ? "
                "WHERE id = ? AND deleted = 0",
                (now, cleaned),
            )
            connection.commit()
            return cursor.rowcount > 0
        finally:
            connection.close()
