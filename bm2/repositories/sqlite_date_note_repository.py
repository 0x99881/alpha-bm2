from __future__ import annotations

from typing import Any


class SQLiteDateNoteRepositoryMixin:
    NOTE_KEYS = ("note1", "note2", "note3")

    @staticmethod
    def _note_text(value: Any) -> str:
        return str(value or "").strip()

    def record_score_date_notes(
        self,
        score_date: str,
        notes: dict[str, str],
        *,
        source: str = "local",
    ) -> None:
        if self.read_only:
            return
        cleaned_date = self._note_text(score_date)
        if not cleaned_date:
            return
        note1 = self._note_text(notes.get("note1"))
        note2 = self._note_text(notes.get("note2"))
        note3 = self._note_text(notes.get("note3"))
        deleted = 0 if note1 or note2 or note3 else 1
        updated_at = self._now_text()
        connection = self._connect()
        try:
            connection.execute(
                """
                INSERT INTO score_date_notes (
                    score_date, note1, note2, note3, updated_at, version, deleted, source
                ) VALUES (?, ?, ?, ?, ?, 1, ?, ?)
                ON CONFLICT(score_date) DO UPDATE SET
                    note1=excluded.note1,
                    note2=excluded.note2,
                    note3=excluded.note3,
                    updated_at=excluded.updated_at,
                    version=score_date_notes.version + 1,
                    deleted=excluded.deleted,
                    source=excluded.source
                WHERE score_date_notes.note1 != excluded.note1
                    OR score_date_notes.note2 != excluded.note2
                    OR score_date_notes.note3 != excluded.note3
                    OR score_date_notes.deleted != excluded.deleted
                    OR score_date_notes.source != excluded.source
                """,
                (cleaned_date, note1, note2, note3, updated_at, deleted, source),
            )
            connection.commit()
        finally:
            connection.close()

    def get_score_date_note_row(self, score_date: str) -> dict[str, Any] | None:
        connection = self._connect()
        try:
            row = connection.execute(
                """
                SELECT score_date, note1, note2, note3, updated_at, version, deleted, source
                FROM score_date_notes
                WHERE score_date = ?
                """,
                (self._note_text(score_date),),
            ).fetchone()
            return dict(row) if row else None
        finally:
            connection.close()

    def get_score_date_notes(self, score_date: str) -> dict[str, str]:
        row = self.get_score_date_note_row(score_date)
        if not row or int(row.get("deleted", 0) or 0):
            return {"note1": "", "note2": "", "note3": ""}
        return {
            "note1": self._note_text(row.get("note1")),
            "note2": self._note_text(row.get("note2")),
            "note3": self._note_text(row.get("note3")),
        }

    def get_score_date_notes_map(self) -> dict[str, list[str]]:
        connection = self._connect()
        try:
            rows = connection.execute(
                """
                SELECT score_date, note1, note2, note3
                FROM score_date_notes
                WHERE deleted = 0
                ORDER BY score_date ASC
                """
            ).fetchall()
            result: dict[str, list[str]] = {}
            for row in rows:
                notes = [self._note_text(row[key]) for key in self.NOTE_KEYS]
                compacted = [note for note in notes if note]
                if compacted:
                    result[str(row["score_date"]).strip()] = compacted[:3]
            return result
        finally:
            connection.close()

    def delete_score_date_notes(self, score_date: str, *, source: str = "local") -> dict[str, Any] | None:
        if self.read_only:
            return None
        cleaned_date = self._note_text(score_date)
        if not cleaned_date:
            return None
        previous = self.get_score_date_note_row(cleaned_date)
        if previous is None or int(previous.get("deleted", 0) or 0):
            return previous
        updated_at = self._now_text()
        connection = self._connect()
        try:
            connection.execute(
                """
                UPDATE score_date_notes
                SET note1 = '',
                    note2 = '',
                    note3 = '',
                    updated_at = ?,
                    version = version + 1,
                    deleted = 1,
                    source = ?
                WHERE score_date = ?
                """,
                (updated_at, source, cleaned_date),
            )
            connection.commit()
            return previous
        finally:
            connection.close()

    def restore_score_date_note(self, previous: dict[str, Any] | None, *, source: str = "local") -> None:
        if self.read_only or previous is None:
            return
        score_date = self._note_text(previous.get("score_date"))
        if not score_date:
            return
        connection = self._connect()
        try:
            connection.execute(
                """
                INSERT INTO score_date_notes (
                    score_date, note1, note2, note3, updated_at, version, deleted, source
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(score_date) DO UPDATE SET
                    note1=excluded.note1,
                    note2=excluded.note2,
                    note3=excluded.note3,
                    updated_at=excluded.updated_at,
                    version=excluded.version,
                    deleted=excluded.deleted,
                    source=excluded.source
                """,
                (
                    score_date,
                    self._note_text(previous.get("note1")),
                    self._note_text(previous.get("note2")),
                    self._note_text(previous.get("note3")),
                    self._note_text(previous.get("updated_at")) or self._now_text(),
                    int(previous.get("version", 1) or 1),
                    int(previous.get("deleted", 0) or 0),
                    source,
                ),
            )
            connection.commit()
        finally:
            connection.close()
