from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from ..constants import NAME_HEADER, PROFIT_HEADER, TOTAL_HEADER


class SQLiteReportRepositoryMixin:
    def _cycle_start_date(self) -> datetime.date | None:
        """Always None now — the score sheet is cycle-agnostic.

        The old per-cycle-file workflow stored ``score_cycle_start_date`` in
        ``sync_state`` to clip out previous-cycle dates. With the cycle-block
        layout the score sheet keeps rolling across cycles and the 15-day
        red-highlight is enforced by ``recent_numbers`` formatting, not by
        clipping the data. Returning ``None`` here makes ``get_filtered_*``
        return all active entries.
        """
        return None

    def _filtered_score_rows(self) -> list[dict[str, Any]]:
        return self.get_filtered_score_rows()

    def get_filtered_score_rows(self) -> list[dict[str, Any]]:
        """Return all active score rows (cycle-agnostic)."""
        rows = []
        for row in self.get_score_rows():
            try:
                datetime.strptime(str(row["score_date"]).strip(), "%Y-%m-%d").date()
            except ValueError:
                continue
            rows.append(row)
        return rows

    def get_filtered_score_date_notes(self) -> dict[str, list[str]]:
        return self.get_score_date_notes_map()

    def get_next_score_date(self, default_date: str) -> str:
        rows = self._filtered_score_rows()
        if rows:
            latest = max(datetime.strptime(str(row["score_date"]).strip(), "%Y-%m-%d").date() for row in rows)
            return (latest + timedelta(days=1)).strftime("%Y-%m-%d")
        return default_date

    def get_score_date_count(self) -> int:
        connection = self._connect()
        try:
            cycle_start = self._cycle_start_date()
            if cycle_start is None:
                row = connection.execute(
                    "SELECT COUNT(DISTINCT score_date) AS count FROM score_entries WHERE deleted = 0"
                ).fetchone()
            else:
                row = connection.execute(
                    """
                    SELECT COUNT(DISTINCT score_date) AS count
                    FROM score_entries
                    WHERE deleted = 0 AND score_date >= ?
                    """,
                    (cycle_start.strftime("%Y-%m-%d"),),
                ).fetchone()
            return int(row["count"] or 0)
        finally:
            connection.close()

    def _window_dates(self) -> list[str]:
        rows = self._filtered_score_rows()
        cycle_start = self._cycle_start_date()
        if rows:
            latest = max(datetime.strptime(str(row["score_date"]).strip(), "%Y-%m-%d").date() for row in rows)
        elif cycle_start is not None:
            latest = cycle_start
        else:
            latest = datetime.now().date()
        return [
            (latest - timedelta(days=offset)).strftime("%Y-%m-%d")
            for offset in range(14, -1, -1)
        ]

    def get_score_summary_data(self) -> dict[str, Any]:
        window_dates = self._window_dates()
        rows = self._filtered_score_rows()
        score_map = {
            (str(row["member_name"]).strip(), str(row["score_date"]).strip()): int(row["score"] or 0)
            for row in rows
        }
        rankings = []
        for member in self.get_active_member_rows():
            member_name = str(member["name"]).strip()
            total = sum(score_map.get((member_name, score_date), 0) for score_date in window_dates)
            rankings.append({"name": member_name, "total": total})
        rankings.sort(key=lambda item: (-int(item["total"]), str(item["name"])))
        return {
            "latest_column": window_dates[-1][5:],
            "column_count": len(window_dates),
            "rankings": rankings,
        }

    def get_score_sheet_snapshot(self, profit_map: dict[str, float] | None = None) -> dict[str, Any]:
        profit_values = profit_map or {}
        window_dates = self._window_dates()
        rows = self._filtered_score_rows()
        notes_map = self.get_filtered_score_date_notes()
        score_map = {
            (str(row["member_name"]).strip(), str(row["score_date"]).strip()): int(row["score"] or 0)
            for row in rows
        }
        headers = [score_date[5:] for score_date in window_dates]
        headers.extend([TOTAL_HEADER, PROFIT_HEADER, NAME_HEADER])
        raw_rows: list[list[Any]] = []
        for member in self.get_active_member_rows():
            member_name = str(member["name"]).strip()
            date_scores = [score_map.get((member_name, score_date), 0) for score_date in window_dates]
            total = sum(date_scores)
            profit = float(profit_values.get(member_name, 0.0))
            raw_rows.append([*date_scores, total, profit, member_name])
        visible_notes = {score_date: notes_map.get(score_date, []) for score_date in window_dates}
        if any(visible_notes.values()):
            for note_index in range(3):
                note_row = [
                    visible_notes.get(score_date, [])[note_index]
                    if len(visible_notes.get(score_date, [])) > note_index
                    else ""
                    for score_date in window_dates
                ]
                raw_rows.append([*note_row, "", "", ""])
        return {
            "headers": headers,
            "raw_rows": raw_rows,
        }

    # ------------------------------------------------------------------
    # Supabase push / pull
    # ------------------------------------------------------------------
