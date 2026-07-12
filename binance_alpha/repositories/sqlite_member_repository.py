from __future__ import annotations

from typing import Any

from ..constants import DISABLED, ENABLED


class SQLiteMemberRepositoryMixin:
    def sync_members(self, members: list[dict[str, Any]]) -> None:
        if self.read_only:
            return
        updated_at = self._now_text()
        current_names = {str(item["name"]).strip() for item in members}
        connection = self._connect()
        try:
            for member in members:
                name = str(member["name"]).strip()
                connection.execute(
                    """
                    INSERT INTO members (
                        id, name, status, note, sort_order, created_at, disabled_at, updated_at, version, deleted
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, 0)
                    ON CONFLICT(name) DO UPDATE SET
                        status=excluded.status,
                        note=excluded.note,
                        sort_order=excluded.sort_order,
                        created_at=CASE
                            WHEN members.created_at = '' THEN excluded.created_at
                            ELSE members.created_at
                        END,
                        disabled_at=excluded.disabled_at,
                        updated_at=excluded.updated_at,
                        version=members.version + 1,
                        deleted=0
                    """,
                    (
                        self._member_id(name),
                        name,
                        str(member.get("status", "")).strip(),
                        str(member.get("note", "")).strip(),
                        int(member.get("sort_order", 0) or 0),
                        str(member.get("created_at", "")).strip(),
                        str(member.get("disabled_at", "")).strip(),
                        updated_at,
                    ),
                )

            if current_names:
                placeholders = ",".join("?" for _ in current_names)
                connection.execute(
                    f"""
                    UPDATE members
                    SET deleted = 1,
                        updated_at = ?,
                        version = version + 1
                    WHERE name NOT IN ({placeholders})
                    """,
                    (updated_at, *sorted(current_names)),
                )
            connection.commit()
        finally:
            connection.close()

    def get_member_rows(self, *, include_deleted: bool = False) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            query = """
                SELECT name, status, note, sort_order, created_at, disabled_at, updated_at, version, deleted
                FROM members
            """
            params: list[Any] = []
            if not include_deleted:
                query += " WHERE deleted = 0"
            query += " ORDER BY deleted ASC, sort_order ASC, name ASC"
            rows = connection.execute(query, params).fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

    def get_active_member_rows(self) -> list[dict[str, Any]]:
        return [row for row in self.get_member_rows() if row["status"] == ENABLED and int(row["deleted"] or 0) == 0]

    def add_member(self, name: str, note: str = "") -> None:
        if self.read_only:
            return
        cleaned_name = name.strip()
        updated_at = self._now_text()
        existing = self.get_member_rows(include_deleted=True)
        next_sort_order = max((int(item.get("sort_order", 0) or 0) for item in existing), default=0) + 1
        connection = self._connect()
        try:
            connection.execute(
                """
                INSERT INTO members (
                    id, name, status, note, sort_order, created_at, disabled_at, updated_at, version, deleted
                ) VALUES (?, ?, ?, ?, ?, ?, '', ?, 1, 0)
                ON CONFLICT(name) DO UPDATE SET
                    status=excluded.status,
                    note=excluded.note,
                    sort_order=excluded.sort_order,
                    disabled_at='',
                    updated_at=excluded.updated_at,
                    version=members.version + 1,
                    deleted=0
                """,
                (
                    self._member_id(cleaned_name),
                    cleaned_name,
                    ENABLED,
                    note.strip(),
                    next_sort_order,
                    updated_at,
                    updated_at,
                ),
            )
            connection.commit()
        finally:
            connection.close()

    def update_member(self, name: str, note: str, status: str | None = None) -> bool:
        if self.read_only:
            return False
        cleaned_name = name.strip()
        rows = self.get_member_rows(include_deleted=True)
        existing = next((row for row in rows if str(row.get("name", "")).strip() == cleaned_name and int(row.get("deleted", 0) or 0) == 0), None)
        if existing is None:
            return False
        normalized_status = status if status in {ENABLED, DISABLED} else str(existing.get("status") or ENABLED)
        disabled_at = str(existing.get("disabled_at") or "")
        if normalized_status == DISABLED and str(existing.get("status")) != DISABLED:
            disabled_at = self._now_text()
        if normalized_status == ENABLED:
            disabled_at = ""
        updated_at = self._now_text()
        connection = self._connect()
        try:
            connection.execute(
                """
                UPDATE members
                SET status = ?,
                    note = ?,
                    disabled_at = ?,
                    updated_at = ?,
                    version = version + 1
                WHERE name = ? AND deleted = 0
                """,
                (normalized_status, note.strip(), disabled_at, updated_at, cleaned_name),
            )
            connection.commit()
            return True
        finally:
            connection.close()

    def delete_member(self, name: str) -> bool:
        if self.read_only:
            return False
        cleaned_name = name.strip()
        updated_at = self._now_text()
        connection = self._connect()
        try:
            cursor = connection.execute(
                """
                UPDATE members
                SET deleted = 1,
                    updated_at = ?,
                    version = version + 1
                WHERE name = ? AND deleted = 0
                """,
                (updated_at, cleaned_name),
            )
            connection.commit()
            return cursor.rowcount > 0
        finally:
            connection.close()

    def reorder_active_members(self, ordered_names: list[str]) -> None:
        if self.read_only:
            return
        current_members = self.get_member_rows()
        active_members = [item for item in current_members if item["status"] == ENABLED]
        disabled_members = [item for item in current_members if item["status"] == DISABLED]
        active_names = [str(item["name"]) for item in active_members]
        unique_requested: list[str] = []
        for name in ordered_names:
            if name in active_names and name not in unique_requested:
                unique_requested.append(name)
        final_names = unique_requested + [name for name in active_names if name not in unique_requested]
        final_names.extend(str(item["name"]) for item in disabled_members)
        updated_at = self._now_text()
        connection = self._connect()
        try:
            for index, member_name in enumerate(final_names, start=1):
                connection.execute(
                    """
                    UPDATE members
                    SET sort_order = ?,
                        updated_at = ?,
                        version = version + 1
                    WHERE name = ? AND deleted = 0
                    """,
                    (index, updated_at, member_name),
                )
            connection.commit()
        finally:
            connection.close()
