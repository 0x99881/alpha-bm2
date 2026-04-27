from __future__ import annotations

from typing import Any

from ..domain.rules.daily_entry import find_invalid_score_member
from ..ui_text import MESSAGES


class DailyEntryService:
    def __init__(self, entry_writer) -> None:
        self._entry_writer = entry_writer

    def build_entries(self, active_members: list[dict[str, Any]], form_data: Any) -> list[dict[str, str]]:
        entries: list[dict[str, str]] = []
        for member in active_members:
            name = member['name']
            entries.append(
                {
                    'name': name,
                    'score': form_data.get(f'score_{name}', '').strip(),
                    'before_balance': form_data.get(f'before_{name}', '').strip(),
                    'after_balance': form_data.get(f'after_{name}', '').strip(),
                    'manual_wear': form_data.get(f'manual_wear_{name}', '').strip(),
                    'income': form_data.get(f'income_{name}', '').strip(),
                    'other_expense': form_data.get(f'other_expense_{name}', '').strip(),
                }
            )
        return entries

    def validate_entries(self, entries: list[dict[str, str]]) -> str | None:
        invalid_member = find_invalid_score_member(entries)
        if invalid_member is not None:
            return MESSAGES['score_must_integer'].format(name=invalid_member)
        return None

    def process_submission(self, active_members: list[dict[str, Any]], form_data: Any, selected_date: str) -> dict[str, Any]:
        entries = self.build_entries(active_members, form_data)
        validation_error = self.validate_entries(entries)
        if validation_error is not None:
            return {
                'ok': False,
                'entries': entries,
                'error': validation_error,
            }

        try:
            result = self._entry_writer.save_scores_and_wear(selected_date, entries)
        except ValueError as exc:
            return {
                'ok': False,
                'entries': entries,
                'error': str(exc),
            }

        return {
            'ok': True,
            'entries': entries,
            'result': result,
        }
