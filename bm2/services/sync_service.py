from __future__ import annotations

from typing import Callable


class SyncService:
    def __init__(self, local_database, supabase_client, after_pull: Callable[[], None] | None = None) -> None:
        self._local_database = local_database
        self._supabase_client = supabase_client
        self._after_pull = after_pull

    def is_configured(self) -> bool:
        return self._supabase_client.is_configured()

    def push(self) -> dict[str, int]:
        if not self.is_configured():
            return {"members": 0, "score_entries": 0}
        return self._local_database.push_to_supabase(self._supabase_client)

    def pull(self) -> dict[str, int]:
        if not self.is_configured():
            return {"members": 0, "score_entries": 0}
        result = self._local_database.pull_from_supabase(self._supabase_client)
        changed = result.get("score_entries", 0) > 0 or result.get("members", 0) > 0
        if changed and self._after_pull is not None:
            self._after_pull()
        return result

    def sync(self) -> dict[str, dict[str, int]]:
        if not self.is_configured():
            return {
                "pull": {"members": 0, "score_entries": 0},
                "push": {"members": 0, "score_entries": 0},
            }
        push = self.push()
        pull = self.pull()
        return {"pull": pull, "push": push}
