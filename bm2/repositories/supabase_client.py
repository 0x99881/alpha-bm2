from __future__ import annotations

import logging
import os
import time
from pathlib import Path
from typing import Any, Callable, TypeVar

import httpx
from postgrest.exceptions import APIError

LOGGER = logging.getLogger(__name__)

_T = TypeVar("_T")
_RETRY_ATTEMPTS = 3
_RETRY_DELAY_SECONDS = 1.5


def _retry_on_transient(call: Callable[[], _T], *, label: str) -> _T:
    last_exc: Exception | None = None
    for attempt in range(1, _RETRY_ATTEMPTS + 1):
        try:
            return call()
        except (httpx.TimeoutException, httpx.TransportError) as exc:
            last_exc = exc
            LOGGER.warning(
                "Supabase %s transient failure (attempt %d/%d): %s: %s",
                label,
                attempt,
                _RETRY_ATTEMPTS,
                type(exc).__name__,
                exc,
            )
            if attempt < _RETRY_ATTEMPTS:
                time.sleep(_RETRY_DELAY_SECONDS)
    assert last_exc is not None
    raise last_exc

_SUPABASE_URL_KEY = "SUPABASE_URL"
_SUPABASE_KEY_KEY = "SUPABASE_SERVICE_ROLE_KEY"
_SCORE_ENTRY_UPLOAD_COLUMNS = (
    "id",
    "member_id",
    "member_name",
    "score_date",
    "score",
    "before_balance",
    "after_balance",
    "manual_wear",
    "income",
    "other_expense",
    "profit",
    "updated_at",
    "version",
    "deleted",
    "source",
)
_SETTLEMENT_CYCLE_COLUMNS = (
    "id",
    "name",
    "start_date",
    "settle_date",
    "settled",
    "created_at",
    "updated_at",
    "version",
    "deleted",
    "source",
)
_SETTLEMENT_ENTRY_COLUMNS = (
    "cycle_id",
    "member_name",
    "start_date",
    "start_balance",
    "settle_date",
    "end_balance",
    "is_extra",
    "sort_order",
    "updated_at",
    "version",
    "deleted",
    "source",
)


class SupabaseSchemaError(RuntimeError):
    """Raised when the online score table is missing required columns."""


SUPABASE_REQUEST_ERRORS = (APIError, httpx.TimeoutException, httpx.TransportError)


class SupabaseClient:
    def __init__(self, base_dir: Path) -> None:
        self.base_dir = Path(base_dir)
        self._env_cache: dict[str, str] | None = None
        self._client_cache = None
        self._schema_ok = False

    def _load_env_file(self) -> dict[str, str]:
        if self._env_cache is not None:
            return self._env_cache
        env: dict[str, str] = {}
        path = self.base_dir / ".env.local"
        if path.exists():
            for line in path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                env[key.strip()] = value.strip().strip('"').strip("'")
        self._env_cache = env
        return env

    def _env(self, key: str) -> str:
        return os.environ.get(key) or self._load_env_file().get(key, "")

    def _url(self) -> str:
        return self._env(_SUPABASE_URL_KEY).rstrip("/")

    def _key(self) -> str:
        return self._env(_SUPABASE_KEY_KEY)

    def is_configured(self) -> bool:
        return bool(self._url() and self._key())

    def _client(self):
        from supabase import create_client  # type: ignore

        if self._client_cache is None:
            self._client_cache = create_client(self._url(), self._key())
        return self._client_cache

    def ensure_score_entries_schema(self) -> None:
        if self._schema_ok:
            return

        def _run() -> None:
            self._client().table("score_entries").select(",".join(_SCORE_ENTRY_UPLOAD_COLUMNS)).limit(1).execute()

        try:
            _retry_on_transient(_run, label="ensure_score_entries_schema")
            self._schema_ok = True
        except APIError as exc:
            message = str(exc)
            if "score_entries." in message and "does not exist" in message:
                raise SupabaseSchemaError(message) from exc
            raise

    def push_members(self, rows: list[dict[str, Any]]) -> int:
        if not rows:
            return 0
        LOGGER.info("Supabase push_members: upserting %d rows (on_conflict=id)", len(rows))
        _retry_on_transient(
            lambda: self._client().table("members").upsert(rows, on_conflict="id").execute(),
            label="push_members",
        )
        return len(rows)

    def push_score_entries(self, rows: list[dict[str, Any]]) -> int:
        if not rows:
            return 0
        LOGGER.info("Supabase push_score_entries: upserting %d rows (on_conflict=id)", len(rows))
        _retry_on_transient(
            lambda: self._client().table("score_entries").upsert(rows, on_conflict="id").execute(),
            label="push_score_entries",
        )
        return len(rows)

    def fetch_members_by_ids(self, ids: list[str]) -> list[dict[str, Any]]:
        if not ids:
            return []
        result = _retry_on_transient(
            lambda: self._client().table("members").select("*").in_("id", ids).execute(),
            label="fetch_members_by_ids",
        )
        return result.data or []

    def fetch_score_entries_by_ids(self, ids: list[str]) -> list[dict[str, Any]]:
        if not ids:
            return []
        result = _retry_on_transient(
            lambda: self._client().table("score_entries").select("*").in_("id", ids).execute(),
            label="fetch_score_entries_by_ids",
        )
        return result.data or []

    def mark_score_entries_deleted(self, rows: list[dict[str, Any]]) -> int:
        if not rows:
            return 0
        LOGGER.info("Supabase mark_score_entries_deleted: upserting %d rows", len(rows))
        _retry_on_transient(
            lambda: self._client().table("score_entries").upsert(rows, on_conflict="id").execute(),
            label="mark_score_entries_deleted",
        )
        return len(rows)

    def pull_members(self, since_updated_at: str | None = None) -> list[dict[str, Any]]:
        LOGGER.info("Supabase pull_members: since=%s", since_updated_at or "beginning")

        def _run():
            query = self._client().table("members").select("*")
            if since_updated_at:
                query = query.gt("updated_at", since_updated_at)
            return query.order("updated_at").execute()

        result = _retry_on_transient(_run, label="pull_members")
        return result.data or []

    def pull_score_entries(self, since_updated_at: str | None = None) -> list[dict[str, Any]]:
        LOGGER.info("Supabase pull_score_entries: since=%s", since_updated_at or "beginning")

        def _run():
            query = self._client().table("score_entries").select("*")
            if since_updated_at:
                query = query.gt("updated_at", since_updated_at)
            return query.order("updated_at").execute()

        result = _retry_on_transient(_run, label="pull_score_entries")
        return result.data or []

    # ------------------------------------------------------------------
    # Settlement cycles + entries
    #
    # These tables are optional on the Supabase side: an older project
    # may not have them yet. Missing-table errors are caught and
    # surfaced as an empty result so the pull path stays robust during
    # the rollout window where local code knows about cycles but the
    # remote schema hasn't been migrated.
    # ------------------------------------------------------------------
    @staticmethod
    def _is_missing_table_error(exc: APIError) -> bool:
        message = str(exc)
        return "does not exist" in message or "could not find" in message.lower()

    def push_settlement_cycles(self, rows: list[dict[str, Any]]) -> int:
        if not rows:
            return 0
        LOGGER.info("Supabase push_settlement_cycles: upserting %d rows", len(rows))
        try:
            _retry_on_transient(
                lambda: self._client().table("settlement_cycles").upsert(rows, on_conflict="id").execute(),
                label="push_settlement_cycles",
            )
        except APIError as exc:
            if self._is_missing_table_error(exc):
                LOGGER.warning(
                    "Supabase settlement_cycles table missing; skipping push. "
                    "Run docs/supabase_schema.sql in the Supabase SQL Editor."
                )
                return 0
            raise
        return len(rows)

    def push_settlement_entries(self, rows: list[dict[str, Any]]) -> int:
        if not rows:
            return 0
        LOGGER.info("Supabase push_settlement_entries: upserting %d rows", len(rows))
        try:
            _retry_on_transient(
                lambda: self._client()
                    .table("settlement_entries")
                    .upsert(rows, on_conflict="cycle_id,member_name")
                    .execute(),
                label="push_settlement_entries",
            )
        except APIError as exc:
            if self._is_missing_table_error(exc):
                LOGGER.warning(
                    "Supabase settlement_entries table missing; skipping push. "
                    "Run docs/supabase_schema.sql in the Supabase SQL Editor."
                )
                return 0
            raise
        return len(rows)

    def pull_settlement_cycles(self, since_updated_at: str | None = None) -> list[dict[str, Any]]:
        LOGGER.info("Supabase pull_settlement_cycles: since=%s", since_updated_at or "beginning")

        def _run():
            query = self._client().table("settlement_cycles").select("*")
            if since_updated_at:
                query = query.gt("updated_at", since_updated_at)
            return query.order("updated_at").execute()

        try:
            result = _retry_on_transient(_run, label="pull_settlement_cycles")
        except APIError as exc:
            if self._is_missing_table_error(exc):
                LOGGER.warning(
                    "Supabase settlement_cycles table missing; treating pull as empty."
                )
                return []
            raise
        return result.data or []

    def pull_settlement_entries(self, since_updated_at: str | None = None) -> list[dict[str, Any]]:
        LOGGER.info("Supabase pull_settlement_entries: since=%s", since_updated_at or "beginning")

        def _run():
            query = self._client().table("settlement_entries").select("*")
            if since_updated_at:
                query = query.gt("updated_at", since_updated_at)
            return query.order("updated_at").execute()

        try:
            result = _retry_on_transient(_run, label="pull_settlement_entries")
        except APIError as exc:
            if self._is_missing_table_error(exc):
                LOGGER.warning(
                    "Supabase settlement_entries table missing; treating pull as empty."
                )
                return []
            raise
        return result.data or []
