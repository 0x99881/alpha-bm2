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
                    before_balance TEXT NOT NULL DEFAULT '',
                    after_balance TEXT NOT NULL DEFAULT '',
                    manual_wear TEXT NOT NULL DEFAULT '',
                    income TEXT NOT NULL DEFAULT '',
                    other_expense TEXT NOT NULL DEFAULT '',
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
            existing_columns = {
                row["name"]
                for row in connection.execute("PRAGMA table_info(score_entries)").fetchall()
            }
            for column_name in (
                "before_balance",
                "after_balance",
                "manual_wear",
                "income",
                "other_expense",
            ):
                if column_name not in existing_columns:
                    connection.execute(
                        f"ALTER TABLE score_entries ADD COLUMN {column_name} TEXT NOT NULL DEFAULT ''"
                    )
            connection.commit()
        finally:
            connection.close()
