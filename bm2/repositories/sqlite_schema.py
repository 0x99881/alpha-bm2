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
                    profit TEXT NOT NULL DEFAULT '0',
                    updated_at TEXT NOT NULL,
                    version INTEGER NOT NULL DEFAULT 1,
                    deleted INTEGER NOT NULL DEFAULT 0,
                    source TEXT NOT NULL DEFAULT 'local',
                    UNIQUE(member_name, score_date)
                );

                CREATE TABLE IF NOT EXISTS score_date_notes (
                    score_date TEXT PRIMARY KEY,
                    note1 TEXT NOT NULL DEFAULT '',
                    note2 TEXT NOT NULL DEFAULT '',
                    note3 TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL,
                    version INTEGER NOT NULL DEFAULT 1,
                    deleted INTEGER NOT NULL DEFAULT 0,
                    source TEXT NOT NULL DEFAULT 'local'
                );

                CREATE TABLE IF NOT EXISTS sync_state (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS settlement_cycles (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    start_date TEXT NOT NULL DEFAULT '',
                    settle_date TEXT NOT NULL DEFAULT '',
                    settled INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    version INTEGER NOT NULL DEFAULT 1,
                    deleted INTEGER NOT NULL DEFAULT 0,
                    source TEXT NOT NULL DEFAULT 'local'
                );

                CREATE TABLE IF NOT EXISTS settlement_entries (
                    cycle_id TEXT NOT NULL,
                    member_name TEXT NOT NULL,
                    start_date TEXT NOT NULL DEFAULT '',
                    start_balance TEXT NOT NULL DEFAULT '',
                    settle_date TEXT NOT NULL DEFAULT '',
                    end_balance TEXT NOT NULL DEFAULT '',
                    is_extra INTEGER NOT NULL DEFAULT 0,
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    updated_at TEXT NOT NULL,
                    version INTEGER NOT NULL DEFAULT 1,
                    deleted INTEGER NOT NULL DEFAULT 0,
                    source TEXT NOT NULL DEFAULT 'local',
                    PRIMARY KEY (cycle_id, member_name)
                );

                CREATE INDEX IF NOT EXISTS idx_score_entries_date
                ON score_entries(score_date);

                CREATE INDEX IF NOT EXISTS idx_settlement_entries_cycle
                ON settlement_entries(cycle_id);
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
                "profit",
            ):
                if column_name not in existing_columns:
                    connection.execute(
                        f"ALTER TABLE score_entries ADD COLUMN {column_name} TEXT NOT NULL DEFAULT ''"
                    )
            existing_note_columns = {
                row["name"]
                for row in connection.execute("PRAGMA table_info(score_date_notes)").fetchall()
            }
            for column_name in ("note3",):
                if column_name not in existing_note_columns:
                    connection.execute(
                        f"ALTER TABLE score_date_notes ADD COLUMN {column_name} TEXT NOT NULL DEFAULT ''"
                    )
            existing_cycle_columns = {
                row["name"]
                for row in connection.execute("PRAGMA table_info(settlement_cycles)").fetchall()
            }
            for column_name, default_clause in (
                ("start_date", "TEXT NOT NULL DEFAULT ''"),
                ("settle_date", "TEXT NOT NULL DEFAULT ''"),
                ("settled", "INTEGER NOT NULL DEFAULT 0"),
                ("version", "INTEGER NOT NULL DEFAULT 1"),
                ("source", "TEXT NOT NULL DEFAULT 'local'"),
            ):
                if column_name not in existing_cycle_columns:
                    connection.execute(
                        f"ALTER TABLE settlement_cycles ADD COLUMN {column_name} {default_clause}"
                    )
            existing_entry_columns = {
                row["name"]
                for row in connection.execute("PRAGMA table_info(settlement_entries)").fetchall()
            }
            for column_name, default_clause in (
                ("is_extra", "INTEGER NOT NULL DEFAULT 0"),
                ("sort_order", "INTEGER NOT NULL DEFAULT 0"),
            ):
                if column_name not in existing_entry_columns:
                    connection.execute(
                        f"ALTER TABLE settlement_entries ADD COLUMN {column_name} {default_clause}"
                    )
            connection.commit()
        finally:
            connection.close()
