from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any

LOGGER = logging.getLogger(__name__)

_SUPABASE_URL_KEY = "SUPABASE_URL"
_SUPABASE_KEY_KEY = "SUPABASE_SERVICE_ROLE_KEY"


class SupabaseClient:
    def __init__(self, base_dir: Path) -> None:
        self.base_dir = Path(base_dir)
        self._env_cache: dict[str, str] | None = None

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

        return create_client(self._url(), self._key())

    def push_members(self, rows: list[dict[str, Any]]) -> int:
        if not rows:
            return 0
        LOGGER.info("Supabase push_members: upserting %d rows (on_conflict=id)", len(rows))
        self._client().table("members").upsert(rows, on_conflict="id").execute()
        return len(rows)

    def push_score_entries(self, rows: list[dict[str, Any]]) -> int:
        if not rows:
            return 0
        LOGGER.info("Supabase push_score_entries: upserting %d rows (on_conflict=id)", len(rows))
        self._client().table("score_entries").upsert(rows, on_conflict="id").execute()
        return len(rows)

    def fetch_members_by_ids(self, ids: list[str]) -> list[dict[str, Any]]:
        if not ids:
            return []
        result = self._client().table("members").select("*").in_("id", ids).execute()
        return result.data or []

    def fetch_score_entries_by_ids(self, ids: list[str]) -> list[dict[str, Any]]:
        if not ids:
            return []
        result = self._client().table("score_entries").select("*").in_("id", ids).execute()
        return result.data or []

    def pull_members(self, since_updated_at: str | None = None) -> list[dict[str, Any]]:
        LOGGER.info("Supabase pull_members: since=%s", since_updated_at or "beginning")
        query = self._client().table("members").select("*")
        if since_updated_at:
            query = query.gt("updated_at", since_updated_at)
        result = query.order("updated_at").execute()
        return result.data or []

    def pull_score_entries(self, since_updated_at: str | None = None) -> list[dict[str, Any]]:
        LOGGER.info("Supabase pull_score_entries: since=%s", since_updated_at or "beginning")
        query = self._client().table("score_entries").select("*")
        if since_updated_at:
            query = query.gt("updated_at", since_updated_at)
        result = query.order("updated_at").execute()
        return result.data or []
