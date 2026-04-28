from __future__ import annotations

from datetime import datetime
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

    def validate_selected_date(self, selected_date: str) -> str | None:
        raw_date = str(selected_date or "").strip()
        if not raw_date:
            return MESSAGES["score_date_required"]
        try:
            normalized = datetime.strptime(raw_date, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            return MESSAGES["score_date_invalid"]
        if normalized != raw_date:
            return MESSAGES["score_date_invalid"]
        return None

    def validate_member_fields(
        self,
        active_members: list[dict[str, Any]],
        form_data: Any,
    ) -> str | None:
        active_names = {str(member.get("name", "")).strip() for member in active_members}
        for key in form_data.keys():
            if not str(key).startswith("score_"):
                continue
            member_name = str(key)[6:].strip()
            raw_value = str(form_data.get(key, "") or "").strip()
            if not raw_value:
                continue
            if member_name not in active_names:
                return MESSAGES["score_member_inactive"].format(name=member_name)
        return None

    def validate_entries(self, entries: list[dict[str, str]]) -> str | None:
        invalid_member = find_invalid_score_member(entries)
        if invalid_member is not None:
            return MESSAGES['score_must_integer'].format(name=invalid_member)
        return None

    def validate_submission(
        self,
        active_members: list[dict[str, Any]],
        form_data: Any,
        selected_date: str,
        entries: list[dict[str, str]],
    ) -> str | None:
        date_error = self.validate_selected_date(selected_date)
        if date_error is not None:
            return date_error
        member_error = self.validate_member_fields(active_members, form_data)
        if member_error is not None:
            return member_error
        return self.validate_entries(entries)

    def process_submission(self, active_members: list[dict[str, Any]], form_data: Any, selected_date: str) -> dict[str, Any]:
        entries = self.build_entries(active_members, form_data)
        validation_error = self.validate_submission(active_members, form_data, selected_date, entries)
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
