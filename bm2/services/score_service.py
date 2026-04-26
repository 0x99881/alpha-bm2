from __future__ import annotations

# COMPATIBILITY LAYER

from typing import Any


def build_score_entries(store: Any, active_members: list[dict[str, Any]], form_data: Any) -> list[dict[str, str]]:
    return store.daily_entry_service.build_entries(active_members, form_data)


def validate_score_entries(store: Any, entries: list[dict[str, str]]) -> str | None:
    return store.daily_entry_service.validate_entries(entries)


def process_score_submission(store: Any, active_members: list[dict[str, Any]], form_data: Any, selected_date: str) -> dict[str, Any]:
    return store.daily_entry_service.process_submission(active_members, form_data, selected_date)


__all__ = ['build_score_entries', 'process_score_submission', 'validate_score_entries']
