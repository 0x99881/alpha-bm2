from __future__ import annotations

from typing import Any

from ..ui_text import MESSAGES


def build_score_entries(active_members: list[dict[str, Any]], form_data: Any) -> list[dict[str, str]]:
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


def validate_score_entries(entries: list[dict[str, str]]) -> str | None:
    for entry in entries:
        score_text = entry['score']
        if not score_text:
            continue
        try:
            int(score_text)
        except ValueError:
            return MESSAGES['score_must_integer'].format(name=entry['name'])
    return None


def process_score_submission(store: Any, active_members: list[dict[str, Any]], form_data: Any, selected_date: str) -> dict[str, Any]:
    entries = build_score_entries(active_members, form_data)
    validation_error = validate_score_entries(entries)
    if validation_error is not None:
        return {
            'ok': False,
            'entries': entries,
            'error': validation_error,
        }

    try:
        result = store.save_scores_and_wear(selected_date, entries)
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
