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
# PostgREST caps a single response at 1000 rows by default. Past this
# threshold the pull silently truncates the tail of an ``order(...)``
# result, which is what made the online /score-overview lag behind the
# local one once score_entries grew past 1000.
_PAGE_SIZE = 1000


# Substrings that mark an APIError as a transient gateway/DB blip worth
# retrying. Anything NOT matching (schema errors, permission, PGRST logic
# codes) is a permanent failure and re-raised immediately — retrying those
# would just burn 3×1.5s before failing anyway.
_TRANSIENT_API_HINTS = (
    "timeout", "timed out", "temporarily unavailable", "unavailable",
    "connection", "reset", "gateway", "overloaded", "too many",
    "502", "503", "504",
)


def _is_transient_api_error(exc: APIError) -> bool:
    """True for server-side/gateway blips (Supabase 5xx, connection resets,
    DB restarts) that a retry can plausibly recover from."""
    code = str(getattr(exc, "code", "") or "")
    if code[:1] == "5":  # 5xx-style gateway/server error
        return True
    blob = " ".join(
        str(getattr(exc, attr, "") or "")
        for attr in ("message", "details", "hint", "code")
    ).lower()
    return any(hint in blob for hint in _TRANSIENT_API_HINTS)


def _retry_on_transient(call: Callable[[], _T], *, label: str) -> _T:
    last_exc: Exception | None = None
    for attempt in range(1, _RETRY_ATTEMPTS + 1):
        try:
            return call()
        except (httpx.TimeoutException, httpx.TransportError) as exc:
            last_exc = exc
        except APIError as exc:
            # Permanent API errors (schema/permission/logic) fail fast; only
            # transient server-side blips are worth another attempt.
            if not _is_transient_api_error(exc):
                raise
            last_exc = exc
        LOGGER.warning(
            "Supabase %s transient failure (attempt %d/%d): %s: %s",
            label,
            attempt,
            _RETRY_ATTEMPTS,
            type(last_exc).__name__,
            last_exc,
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
        self._env_cache_mtime: float | None = None
        self._client_cache = None
        self._schema_ok = False

    def _load_env_file(self) -> dict[str, str]:
        path = self.base_dir / ".env.local"
        env_mtime = path.stat().st_mtime if path.exists() else None
        if self._env_cache is not None and self._env_cache_mtime == env_mtime:
            return self._env_cache
        env: dict[str, str] = {}
        if path.exists():
            for line in path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                env[key.strip()] = value.strip().strip('"').strip("'")
        self._env_cache = env
        self._env_cache_mtime = env_mtime
        self._client_cache = None
        self._schema_ok = False
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

    def _paginated_pull(self, build_query, *, label: str) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        offset = 0
        while True:
            page_label = f"{label}[offset={offset}]"
            result = _retry_on_transient(
                lambda: build_query().range(offset, offset + _PAGE_SIZE - 1).execute(),
                label=page_label,
            )
            data = list(result.data or [])
            rows.extend(data)
            if len(data) < _PAGE_SIZE:
                return rows
            offset += _PAGE_SIZE

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

        def _build():
            query = self._client().table("members").select("*")
            if since_updated_at:
                query = query.gt("updated_at", since_updated_at)
            return query.order("updated_at")

        return self._paginated_pull(_build, label="pull_members")

    def pull_score_entries(self, since_updated_at: str | None = None) -> list[dict[str, Any]]:
        LOGGER.info("Supabase pull_score_entries: since=%s", since_updated_at or "beginning")

        def _build():
            query = self._client().table("score_entries").select("*")
            if since_updated_at:
                query = query.gt("updated_at", since_updated_at)
            return query.order("updated_at")

        return self._paginated_pull(_build, label="pull_score_entries")

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

        def _build():
            query = self._client().table("settlement_cycles").select("*")
            if since_updated_at:
                query = query.gt("updated_at", since_updated_at)
            return query.order("updated_at")

        try:
            return self._paginated_pull(_build, label="pull_settlement_cycles")
        except APIError as exc:
            if self._is_missing_table_error(exc):
                LOGGER.warning(
                    "Supabase settlement_cycles table missing; treating pull as empty."
                )
                return []
            raise

    def pull_settlement_entries(self, since_updated_at: str | None = None) -> list[dict[str, Any]]:
        LOGGER.info("Supabase pull_settlement_entries: since=%s", since_updated_at or "beginning")

        def _build():
            query = self._client().table("settlement_entries").select("*")
            if since_updated_at:
                query = query.gt("updated_at", since_updated_at)
            return query.order("updated_at")

        try:
            return self._paginated_pull(_build, label="pull_settlement_entries")
        except APIError as exc:
            if self._is_missing_table_error(exc):
                LOGGER.warning(
                    "Supabase settlement_entries table missing; treating pull as empty."
                )
                return []
            raise
