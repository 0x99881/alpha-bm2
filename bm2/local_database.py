from __future__ import annotations

# This file is only the SQLite database composition entry.
# It owns the connection and assembles repository mixins; it must not contain
# business logic, sync orchestration, Excel migration logic, or CRUD details.

import sqlite3
from pathlib import Path

from .repositories.sqlite_common_repository import SQLiteCommonRepositoryMixin
from .repositories.sqlite_member_repository import SQLiteMemberRepositoryMixin
from .repositories.sqlite_report_repository import SQLiteReportRepositoryMixin
from .repositories.sqlite_schema import SQLiteSchemaMixin
from .repositories.sqlite_score_repository import SQLiteScoreEntryRepositoryMixin
from .repositories.sqlite_sync_merge_repository import SQLiteSyncMergeRepositoryMixin
from .repositories.sqlite_sync_state_repository import SQLiteSyncStateRepositoryMixin


class LocalDatabase(
    SQLiteSchemaMixin,
    SQLiteCommonRepositoryMixin,
    SQLiteSyncStateRepositoryMixin,
    SQLiteMemberRepositoryMixin,
    SQLiteScoreEntryRepositoryMixin,
    SQLiteReportRepositoryMixin,
    SQLiteSyncMergeRepositoryMixin,
):
    def __init__(self, base_dir: Path, *, read_only: bool = False) -> None:
        self.base_dir = Path(base_dir)
        self.read_only = read_only
        self.db_path = self.base_dir / "bm2_local.db"
        self._ensure_schema()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        return connection
