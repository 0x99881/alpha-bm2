from __future__ import annotations

import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Any

LOGGER = logging.getLogger(__name__)


class SQLiteSyncMergeRepositoryMixin:
    def _get_all_member_rows_for_push(self) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            rows = connection.execute(
                "SELECT id, name, status, note, sort_order, created_at, disabled_at, updated_at, version, deleted FROM members"
            ).fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

    def _get_all_score_rows_for_push(self) -> list[dict[str, Any]]:
        connection = self._connect()
        try:
            rows = connection.execute(
                "SELECT id, member_id, member_name, score_date, score, "
                "before_balance, after_balance, manual_wear, income, other_expense, profit, "
                "updated_at, version, deleted, source FROM score_entries"
            ).fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

    def get_change_detection_snapshot(self) -> dict[str, object]:
        return {
            "members": self._get_all_member_rows_for_push(),
            "score_entries": self._get_all_score_rows_for_push(),
            "date_notes": self.get_score_date_notes_map(),
        }

    def _get_unsynced_member_rows_for_push(self, since_ts: str) -> list[dict[str, Any]]:
        """Return member rows updated after *since_ts* (incremental push)."""
        connection = self._connect()
        try:
            rows = connection.execute(
                "SELECT id, name, status, note, sort_order, created_at, disabled_at, updated_at, version, deleted "
                "FROM members WHERE updated_at > ?",
                (since_ts,),
            ).fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

    def _get_unsynced_score_rows_for_push(self, since_ts: str) -> list[dict[str, Any]]:
        """Return score_entries rows updated after *since_ts* (incremental push)."""
        connection = self._connect()
        try:
            rows = connection.execute(
                "SELECT id, member_id, member_name, score_date, score, "
                "before_balance, after_balance, manual_wear, income, other_expense, profit, "
                "updated_at, version, deleted, source "
                "FROM score_entries WHERE updated_at > ?",
                (since_ts,),
            ).fetchall()
            return [dict(row) for row in rows]
        finally:
            connection.close()

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

    def push_to_supabase(
        self,
        supabase_sync,
        *,
        force_full: bool = False,
        score_profit_map: dict[str, float] | None = None,
    ) -> dict[str, int]:
        """Push local rows to Supabase (incremental by updated_at > last_push).
        Manual upload treats the local database as the source of truth.
        First push is a full push. Upserts by id on Supabase side."""
        if self.read_only:
            return {"members": 0, "score_entries": 0}
        since = None if force_full else self._get_sync_state("supabase_last_push") or None
        mode = "incremental" if since else "full"
        if since:
            members = self._get_unsynced_member_rows_for_push(since)
            scores = self._get_unsynced_score_rows_for_push(since)
        else:
            members = self._get_all_member_rows_for_push()
            scores = self._get_all_score_rows_for_push()
        scores = self._attach_score_profit_values(scores, score_profit_map)
        if scores:
            supabase_sync.ensure_score_entries_schema()
        LOGGER.info(
            "Supabase push (%s): %d members, %d score_entries queued",
            mode, len(members), len(scores),
        )
        m_count, s_count = self._push_members_and_scores(supabase_sync, members, scores)
        deleted_count = 0
        if force_full:
            remote_scores = supabase_sync.pull_score_entries()
            deleted_scores = self._deleted_remote_score_rows(remote_scores, scores)
            deleted_count = supabase_sync.mark_score_entries_deleted(deleted_scores)
        self._set_sync_state("supabase_last_push", self._now_text())
        LOGGER.info(
            "Supabase push done: sent %d members, %d score_entries, deleted %d remote extras",
            m_count, s_count, deleted_count,
        )
        return {"members": m_count, "score_entries": s_count, "score_entries_deleted": deleted_count}

    def pull_from_supabase(self, supabase_sync) -> dict[str, int]:
        """Pull rows from Supabase (since last pull), merge into local SQLite with LWW.
        Conflict rule: higher version wins; tie → higher updated_at wins."""
        if self.read_only:
            return {"members": 0, "score_entries": 0}
        since = self._get_sync_state("supabase_last_pull") or None
        remote_members = supabase_sync.pull_members(since_updated_at=since)
        remote_scores = supabase_sync.pull_score_entries(since_updated_at=since)
        LOGGER.info(
            "Supabase pull: received %d members, %d score_entries from remote (since=%s)",
            len(remote_members), len(remote_scores), since or "beginning",
        )

        m = self._upsert_members_from_remote(remote_members)
        s = self._upsert_scores_from_remote(remote_scores)
        LOGGER.info(
            "Supabase pull members: total=%d inserted=%d updated=%d skipped=%d",
            m["total"], m["inserted"], m["updated"], m["skipped"],
        )
        LOGGER.info(
            "Supabase pull scores:  total=%d inserted=%d updated=%d skipped=%d",
            s["total"], s["inserted"], s["updated"], s["skipped"],
        )
        self._set_sync_state("supabase_last_pull", self._now_text())
        return {
            "members": m["inserted"] + m["updated"],
            "score_entries": s["inserted"] + s["updated"],
            "members_skipped": m["skipped"],
            "scores_skipped": s["skipped"],
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

    def _upsert_members_from_remote(self, rows: list[dict[str, Any]]) -> dict[str, int]:
        if not rows:
            return {"total": 0, "inserted": 0, "updated": 0, "skipped": 0}
        connection = self._connect()
        try:
            ids = [r["id"] for r in rows]
            ph = ",".join("?" * len(ids))
            existing = {
                row["id"]: {
                    "version": row["version"],
                    "updated_at": row["updated_at"],
                    "deleted": row["deleted"],
                }
                for row in connection.execute(
                    f"SELECT id, version, updated_at, deleted FROM members WHERE id IN ({ph})", ids
                ).fetchall()
            }
            inserted = updated = skipped = 0
            for row in rows:
                local = existing.get(row["id"])
                if local is None:
                    connection.execute(
                        "INSERT INTO members "
                        "(id, name, status, note, sort_order, created_at, disabled_at, updated_at, version, deleted) "
                        "VALUES (:id, :name, :status, :note, :sort_order, :created_at, :disabled_at, :updated_at, :version, :deleted)",
                        row,
                    )
                    inserted += 1
                elif self._remote_wins(row, local):
                    connection.execute(
                        "UPDATE members SET name=:name, status=:status, note=:note, sort_order=:sort_order, "
                        "created_at=:created_at, disabled_at=:disabled_at, updated_at=:updated_at, "
                        "version=:version, deleted=:deleted WHERE id=:id",
                        row,
                    )
                    updated += 1
                else:
                    skipped += 1
            connection.commit()
            return {"total": len(rows), "inserted": inserted, "updated": updated, "skipped": skipped}
        finally:
            connection.close()

    def _upsert_scores_from_remote(self, rows: list[dict[str, Any]]) -> dict[str, int]:
        if not rows:
            return {"total": 0, "inserted": 0, "updated": 0, "skipped": 0}
        connection = self._connect()
        try:
            ids = [r["id"] for r in rows]
            ph = ",".join("?" * len(ids))
            existing = {
                row["id"]: {
                    "version": row["version"],
                    "updated_at": row["updated_at"],
                    "deleted": row["deleted"],
                }
                for row in connection.execute(
                    f"SELECT id, version, updated_at, deleted FROM score_entries WHERE id IN ({ph})", ids
                ).fetchall()
            }
            inserted = updated = skipped = 0
            for row in rows:
                row = self._score_row_with_detail_defaults(row)
                local = existing.get(row["id"])
                if local is None:
                    connection.execute(
                        "INSERT INTO score_entries "
                        "(id, member_id, member_name, score_date, score, "
                        "before_balance, after_balance, manual_wear, income, other_expense, profit, "
                        "updated_at, version, deleted, source) "
                        "VALUES (:id, :member_id, :member_name, :score_date, :score, "
                        ":before_balance, :after_balance, :manual_wear, :income, :other_expense, :profit, "
                        ":updated_at, :version, :deleted, :source)",
                        row,
                    )
                    inserted += 1
                elif self._remote_wins(row, local):
                    connection.execute(
                        "UPDATE score_entries SET score=:score, member_id=:member_id, member_name=:member_name, "
                        "score_date=:score_date, before_balance=:before_balance, after_balance=:after_balance, "
                        "manual_wear=:manual_wear, income=:income, other_expense=:other_expense, profit=:profit, "
                        "updated_at=:updated_at, version=:version, "
                        "deleted=:deleted, source=:source WHERE id=:id",
                        row,
                    )
                    updated += 1
                else:
                    skipped += 1
            connection.commit()
            return {"total": len(rows), "inserted": inserted, "updated": updated, "skipped": skipped}
        finally:
            connection.close()
