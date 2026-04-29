from __future__ import annotations


class SQLiteSchemaMixin:
    def _ensure_schema(self) -> None:
        connection = self._connect()
        try:
            connection.executescript(
                """
                PRAGMA foreign_keys = ON;

                CREATE TABLE IF NOT EXISTS members (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL UNIQUE,
                    status TEXT NOT NULL,
                    note TEXT NOT NULL DEFAULT '',
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT '',
                    disabled_at TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL,
                    version INTEGER NOT NULL DEFAULT 1,
                    deleted INTEGER NOT NULL DEFAULT 0
                );

                CREATE TABLE IF NOT EXISTS score_entries (
                    id TEXT PRIMARY KEY,
                    member_id TEXT NOT NULL,
                    member_name TEXT NOT NULL,
                    score_date TEXT NOT NULL,
                    score INTEGER NOT NULL,
                    updated_at TEXT NOT NULL,
                    version INTEGER NOT NULL DEFAULT 1,
                    deleted INTEGER NOT NULL DEFAULT 0,
                    source TEXT NOT NULL DEFAULT 'local',
                    UNIQUE(member_name, score_date)
                );

                CREATE TABLE IF NOT EXISTS sync_state (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL DEFAULT ''
                );

                CREATE INDEX IF NOT EXISTS idx_score_entries_date
                ON score_entries(score_date);
                """
            )
            connection.commit()
        finally:
            connection.close()
