from __future__ import annotations

import logging
from datetime import datetime
from typing import Any

from ..constants import WINDOW_SIZE
from ..entry_helpers import count_wear_entries

LOGGER = logging.getLogger(__name__)


class SQLiteEntryWriter:
    """Primary local entry writer.

    Score submissions land in SQLite first. Repeating the same submission for
    the same day should update that day's rows instead of silently shifting the
    write to later dates.
    """

    def __init__(self, local_db) -> None:
        self._local_db = local_db

    def _resolve_save_date(self, date_text: str) -> str:
        try:
            return datetime.strptime(date_text.strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            return date_text

    def save_scores_and_wear(
        self, date_text: str, entries: list[dict[str, str]]
    ) -> dict[str, Any]:
        saved_date = self._resolve_save_date(date_text)
        self._local_db.record_score_entries(saved_date, entries, source="local")

        return {
            "target_column": saved_date[5:],
            "wear_column": saved_date[5:].replace("-", ""),
            "wear_rows_added": count_wear_entries(entries),
            "window_size": min(WINDOW_SIZE, self._local_db.get_score_date_count()),
            "saved_date": saved_date,
        }
