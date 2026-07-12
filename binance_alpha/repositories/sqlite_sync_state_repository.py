from __future__ import annotations


class SQLiteSyncStateRepositoryMixin:
    def get_sync_state(self, key: str) -> str:
        connection = self._connect()
        try:
            row = connection.execute("SELECT value FROM sync_state WHERE key = ?", (key,)).fetchone()
            return str(row["value"]) if row else ""
        finally:
            connection.close()

    def set_sync_state(self, key: str, value: str) -> None:
        if self.read_only:
            return
        connection = self._connect()
        try:
            connection.execute(
                """
                INSERT INTO sync_state(key, value)
                VALUES (?, ?)
                ON CONFLICT(key) DO UPDATE SET value=excluded.value
                """,
                (key, value),
            )
            connection.commit()
        finally:
            connection.close()

    def _get_sync_state(self, key: str) -> str:
        return self.get_sync_state(key)

    def _set_sync_state(self, key: str, value: str) -> None:
        self.set_sync_state(key, value)
