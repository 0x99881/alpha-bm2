from __future__ import annotations

import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Any

LOGGER = logging.getLogger(__name__)


class SQLiteSyncMergeRepositoryMixin:
    # _LWW_TABLES (declared lower in the class) lists the column tuple
    # for every synced table; both push and pull queries derive their
    # column list from it, so adding a synced column means editing one
    # place.
    def _select_all_for_push(self, table: str) -> list[dict[str, Any]]:
        spec = self._LWW_TABLES[table]
        columns = ", ".join(spec["columns"])
        connection = self._connect()
        try:
            rows = connection.execute(f"SELECT {columns} FROM {table}").fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

    def _select_unsynced_for_push(self, table: str, since_ts: str) -> list[dict[str, Any]]:
        spec = self._LWW_TABLES[table]
        columns = ", ".join(spec["columns"])
        connection = self._connect()
        try:
            rows = connection.execute(
                f"SELECT {columns} FROM {table} WHERE updated_at > ?",
                (since_ts,),
            ).fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

    def _get_all_member_rows_for_push(self) -> list[dict[str, Any]]:
        return self._select_all_for_push("members")

    def _get_all_score_rows_for_push(self) -> list[dict[str, Any]]:
        return self._select_all_for_push("score_entries")

    def get_change_detection_snapshot(self) -> dict[str, object]:
        return {
            "members": self._get_all_member_rows_for_push(),
            "score_entries": self._get_all_score_rows_for_push(),
            "date_notes": self.get_score_date_notes_map(),
        }

    def _get_unsynced_member_rows_for_push(self, since_ts: str) -> list[dict[str, Any]]:
        return self._select_unsynced_for_push("members", since_ts)

    def _get_unsynced_score_rows_for_push(self, since_ts: str) -> list[dict[str, Any]]:
        return self._select_unsynced_for_push("score_entries", since_ts)

    @staticmethod
    def _score_row_with_detail_defaults(row: dict[str, Any]) -> dict[str, Any]:
        normalized = dict(row)
        for key in ("before_balance", "after_balance", "manual_wear", "income", "other_expense", "profit"):
            normalized[key] = str(normalized.get(key, "") or "")
        return normalized

    def _deleted_remote_score_rows(self, remote_rows: list[dict[str, Any]], local_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        local_ids = {str(row.get("id", "")).strip() for row in local_rows if str(row.get("id", "")).strip()}
        now_text = self._now_text()
        deleted_rows: list[dict[str, Any]] = []
        for row in remote_rows:
            row_id = str(row.get("id", "")).strip()
            if not row_id or row_id in local_ids:
                continue
            if int(row.get("deleted", 0) or 0) != 0:
                continue
            deleted = dict(row)
            deleted["deleted"] = 1
            deleted["updated_at"] = now_text
            deleted["version"] = int(row.get("version", 0) or 0) + 1
            deleted_rows.append(deleted)
        return deleted_rows

    @staticmethod
    def _attach_score_profit_values(
        rows: list[dict[str, Any]],
        score_profit_map: dict[str, float] | None,
    ) -> list[dict[str, Any]]:
        if not score_profit_map:
            return rows
        enriched_rows: list[dict[str, Any]] = []
        for row in rows:
            member_name = str(row.get("member_name", "")).strip()
            if member_name not in score_profit_map:
                enriched_rows.append(row)
                continue
            enriched = dict(row)
            enriched["profit"] = str(round(float(score_profit_map[member_name]), 1))
            enriched_rows.append(enriched)
        return enriched_rows

    @staticmethod
    def _push_members_and_scores(supabase_sync, members: list[dict[str, Any]], scores: list[dict[str, Any]]) -> tuple[int, int]:
        if members and scores:
            with ThreadPoolExecutor(max_workers=2) as executor:
                member_future = executor.submit(supabase_sync.push_members, members)
                score_future = executor.submit(supabase_sync.push_score_entries, scores)
                return member_future.result(), score_future.result()
        return supabase_sync.push_members(members), supabase_sync.push_score_entries(scores)

    @staticmethod
    def _push_cycles_and_entries(
        supabase_sync,
        cycles: list[dict[str, Any]],
        entries: list[dict[str, Any]],
    ) -> tuple[int, int]:
        # Settlement cycles and their entries are independent at the
        # Supabase row level (no FK enforced) so they can ship in
        # parallel. Cycle counts are tiny but parallelism keeps push
        # latency dominated by the worst-case single call rather than
        # their sum.
        if cycles and entries:
            with ThreadPoolExecutor(max_workers=2) as executor:
                cycle_future = executor.submit(supabase_sync.push_settlement_cycles, cycles)
                entry_future = executor.submit(supabase_sync.push_settlement_entries, entries)
                return cycle_future.result(), entry_future.result()
        return (
            supabase_sync.push_settlement_cycles(cycles),
            supabase_sync.push_settlement_entries(entries),
        )

    def push_to_supabase(
        self,
        supabase_sync,
        *,
        force_full: bool = False,
        score_profit_map: dict[str, float] | None = None,
    ) -> dict[str, int]:
        """Push local rows to Supabase (incremental by updated_at > last_push).
        Manual upload treats the local database as the source of truth.
        First push is a full push. Upserts by id on Supabase side.

        Tables pushed: members, score_entries, settlement_cycles,
        settlement_entries. Cycle data is included so the online/mobile
        side knows which date window counts as 'this cycle'."""
        if self.read_only:
            return self._empty_sync_result()
        since = None if force_full else self._get_sync_state("supabase_last_push") or None
        mode = "incremental" if since else "full"
        if since:
            members = self._select_unsynced_for_push("members", since)
            scores = self._select_unsynced_for_push("score_entries", since)
            cycles = self._select_unsynced_for_push("settlement_cycles", since)
            entries = self._select_unsynced_for_push("settlement_entries", since)
        else:
            members = self._select_all_for_push("members")
            scores = self._select_all_for_push("score_entries")
            cycles = self._select_all_for_push("settlement_cycles")
            entries = self._select_all_for_push("settlement_entries")
        scores = self._attach_score_profit_values(scores, score_profit_map)
        if scores:
            supabase_sync.ensure_score_entries_schema()
        LOGGER.info(
            "Supabase push (%s): %d members, %d score_entries, %d cycles, %d cycle_entries queued",
            mode, len(members), len(scores), len(cycles), len(entries),
        )
        m_count, s_count = self._push_members_and_scores(supabase_sync, members, scores)
        c_count, e_count = self._push_cycles_and_entries(supabase_sync, cycles, entries)
        deleted_count = 0
        if force_full:
            remote_scores = supabase_sync.pull_score_entries()
            deleted_scores = self._deleted_remote_score_rows(remote_scores, scores)
            deleted_count = supabase_sync.mark_score_entries_deleted(deleted_scores)
        self._set_sync_state("supabase_last_push", self._now_text())
        LOGGER.info(
            "Supabase push done: sent %d members, %d score_entries, %d cycles, %d cycle_entries, "
            "deleted %d remote extras",
            m_count, s_count, c_count, e_count, deleted_count,
        )
        return {
            "members": m_count,
            "score_entries": s_count,
            "settlement_cycles": c_count,
            "settlement_entries": e_count,
            "score_entries_deleted": deleted_count,
        }

    def pull_from_supabase(self, supabase_sync) -> dict[str, int]:
        """Pull rows from Supabase (since last pull), merge into local SQLite
        with last-writer-wins on (version, updated_at). Delete tombstones
        are sticky: once a row is deleted on either side, the deletion
        state survives the merge.

        Tables pulled: members, score_entries, settlement_cycles,
        settlement_entries. Cycle pulls are best-effort; if the online
        cycle tables don't exist yet (older Supabase project) the call
        returns no rows and we keep going."""
        if self.read_only:
            return self._empty_sync_result()
        since = self._get_sync_state("supabase_last_pull") or None
        remote_members = supabase_sync.pull_members(since_updated_at=since)
        remote_scores = supabase_sync.pull_score_entries(since_updated_at=since)
        remote_cycles = supabase_sync.pull_settlement_cycles(since_updated_at=since)
        remote_entries = supabase_sync.pull_settlement_entries(since_updated_at=since)
        LOGGER.info(
            "Supabase pull: received %d members, %d score_entries, %d cycles, %d cycle_entries "
            "(since=%s)",
            len(remote_members), len(remote_scores), len(remote_cycles), len(remote_entries),
            since or "beginning",
        )

        m = self._upsert_members_from_remote(remote_members)
        s = self._upsert_scores_from_remote(remote_scores)
        c = self._upsert_cycles_from_remote(remote_cycles)
        e = self._upsert_cycle_entries_from_remote(remote_entries)
        for label, stats in (("members", m), ("scores", s), ("cycles", c), ("cycle_entries", e)):
            LOGGER.info(
                "Supabase pull %s: total=%d inserted=%d updated=%d skipped=%d",
                label, stats["total"], stats["inserted"], stats["updated"], stats["skipped"],
            )
        self._set_sync_state("supabase_last_pull", self._now_text())
        return {
            "members": m["inserted"] + m["updated"],
            "score_entries": s["inserted"] + s["updated"],
            "settlement_cycles": c["inserted"] + c["updated"],
            "settlement_entries": e["inserted"] + e["updated"],
            "members_skipped": m["skipped"],
            "scores_skipped": s["skipped"],
        }

    @staticmethod
    def _empty_sync_result() -> dict[str, int]:
        return {
            "members": 0,
            "score_entries": 0,
            "settlement_cycles": 0,
            "settlement_entries": 0,
        }

    @staticmethod
    def _remote_wins(remote: dict, local: dict) -> bool:
        rd = int(remote.get("deleted", 0) or 0)
        ld = int(local.get("deleted", 0) or 0)
        if rd != ld:
            return rd > ld
        rv, ra = int(remote["version"]), str(remote["updated_at"])
        lv, la = int(local["version"]), str(local["updated_at"])
        return rv > lv or (rv == lv and ra > la)

    # ------------------------------------------------------------------
    # Generic LWW upsert
    #
    # All tables that participate in Supabase sync share the same merge
    # contract: each row has version + updated_at + deleted, and the
    # remote-wins rule lives in ``_remote_wins``. The helper below builds
    # the INSERT / UPDATE SQL from a column list and a primary-key tuple,
    # so adding a new synced table is a one-call change (see the
    # _upsert_*_from_remote wrappers further down).
    #
    # Composite PKs are supported (e.g. settlement_entries keyed by
    # (cycle_id, member_name)); single-column PKs fall through the same
    # path.
    # ------------------------------------------------------------------
    _LWW_TABLES = {
        "members": {
            "pk": ("id",),
            "columns": (
                "id", "name", "status", "note", "sort_order",
                "created_at", "disabled_at", "updated_at", "version", "deleted",
            ),
        },
        "score_entries": {
            "pk": ("id",),
            "columns": (
                "id", "member_id", "member_name", "score_date", "score",
                "before_balance", "after_balance", "manual_wear", "income",
                "other_expense", "profit",
                "updated_at", "version", "deleted", "source",
            ),
        },
        "settlement_cycles": {
            "pk": ("id",),
            "columns": (
                "id", "name", "start_date", "settle_date", "settled",
                "created_at", "updated_at", "version", "deleted", "source",
            ),
        },
        "settlement_entries": {
            "pk": ("cycle_id", "member_name"),
            "columns": (
                "cycle_id", "member_name", "start_date", "start_balance",
                "settle_date", "end_balance", "is_extra", "sort_order",
                "updated_at", "version", "deleted", "source",
            ),
        },
    }

    def _lww_upsert_rows(
        self,
        *,
        table: str,
        rows: list[dict[str, Any]],
        row_normalizer=None,
    ) -> dict[str, int]:
        """Merge ``rows`` into ``table`` using last-writer-wins.

        Adds an INSERT or UPDATE per row depending on whether the local
        side has a record under the same primary key, and whether the
        remote row wins under ``_remote_wins``.

        ``row_normalizer`` is an optional per-row hook to fill in default
        values that may be missing from the wire format (used by
        score_entries to surface NULL detail columns as empty strings).
        """
        if not rows:
            return {"total": 0, "inserted": 0, "updated": 0, "skipped": 0}
        spec = self._LWW_TABLES[table]
        pk_columns: tuple[str, ...] = spec["pk"]
        all_columns: tuple[str, ...] = spec["columns"]
        non_pk_columns = tuple(c for c in all_columns if c not in pk_columns)

        if row_normalizer is not None:
            rows = [row_normalizer(row) for row in rows]

        connection = self._connect()
        try:
            existing = self._fetch_existing_lww_state(connection, table, pk_columns, rows)
            insert_sql = (
                f"INSERT INTO {table} ({','.join(all_columns)}) "
                f"VALUES ({','.join(':' + c for c in all_columns)})"
            )
            update_sql = (
                f"UPDATE {table} SET "
                + ", ".join(f"{c}=:{c}" for c in non_pk_columns)
                + " WHERE "
                + " AND ".join(f"{c}=:{c}" for c in pk_columns)
            )

            inserted = updated = skipped = 0
            for row in rows:
                key = tuple(row[c] for c in pk_columns)
                local = existing.get(key)
                if local is None:
                    connection.execute(insert_sql, row)
                    inserted += 1
                elif self._remote_wins(row, local):
                    connection.execute(update_sql, row)
                    updated += 1
                else:
                    skipped += 1
            connection.commit()
            return {
                "total": len(rows),
                "inserted": inserted,
                "updated": updated,
                "skipped": skipped,
            }
        finally:
            connection.close()

    @staticmethod
    def _fetch_existing_lww_state(
        connection,
        table: str,
        pk_columns: tuple[str, ...],
        rows: list[dict[str, Any]],
    ) -> dict[tuple, dict[str, Any]]:
        """Return ``{pk_tuple: {version, updated_at, deleted}}`` for the
        subset of ``rows`` whose PK is already present in ``table``.
        Single-column and composite PKs share the same shape; callers
        compose the lookup key as a tuple regardless of arity.
        """
        if len(pk_columns) == 1:
            pk = pk_columns[0]
            placeholders = ",".join("?" * len(rows))
            where_clause = f"{pk} IN ({placeholders})"
            params: list[Any] = [row[pk] for row in rows]
        else:
            tuple_clause = ",".join(f"({','.join('?' * len(pk_columns))})" for _ in rows)
            where_clause = f"({','.join(pk_columns)}) IN ({tuple_clause})"
            params = []
            for row in rows:
                params.extend(row[c] for c in pk_columns)

        select_cols = ",".join((*pk_columns, "version", "updated_at", "deleted"))
        query = f"SELECT {select_cols} FROM {table} WHERE {where_clause}"
        existing: dict[tuple, dict[str, Any]] = {}
        for row in connection.execute(query, params).fetchall():
            key = tuple(row[c] for c in pk_columns)
            existing[key] = {
                "version": row["version"],
                "updated_at": row["updated_at"],
                "deleted": row["deleted"],
            }
        return existing

    def _upsert_members_from_remote(self, rows: list[dict[str, Any]]) -> dict[str, int]:
        return self._lww_upsert_rows(table="members", rows=rows)

    def _upsert_scores_from_remote(self, rows: list[dict[str, Any]]) -> dict[str, int]:
        return self._lww_upsert_rows(
            table="score_entries",
            rows=rows,
            row_normalizer=self._score_row_with_detail_defaults,
        )

    def _upsert_cycles_from_remote(self, rows: list[dict[str, Any]]) -> dict[str, int]:
        return self._lww_upsert_rows(table="settlement_cycles", rows=rows)

    def _upsert_cycle_entries_from_remote(self, rows: list[dict[str, Any]]) -> dict[str, int]:
        return self._lww_upsert_rows(table="settlement_entries", rows=rows)
