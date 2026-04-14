from __future__ import annotations

from calendar import Calendar
from typing import Any


MONTH_YEAR_SEPARATOR = '\u5e74'
MONTH_SUFFIX = '\u6708'


def build_month_label(year: int, month: int) -> str:
    return f"{year}{MONTH_YEAR_SEPARATOR}{month:02d}{MONTH_SUFFIX}"


def build_month_neighbors(year: int, month: int) -> tuple[int, int, int, int]:
    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    next_year, next_month = (year + 1, 1) if month == 12 else (year, month + 1)
    return prev_year, prev_month, next_year, next_month


def empty_member_day_record(date_text: str) -> dict[str, Any]:
    return {'date': date_text, 'wear': 0.0, 'income': 0.0, 'note': ''}


def build_calendar_weeks(*, year: int, month: int, record_map: dict[str, dict[str, Any]], max_abs_wear: float, note_text: str = '') -> list[list[dict[str, Any]]]:
    calendar_view = Calendar(firstweekday=0)
    weeks: list[list[dict[str, Any]]] = []
    for week in calendar_view.monthdatescalendar(year, month):
        cells = []
        for day in week:
            date_text = day.strftime('%Y-%m-%d')
            record = record_map.get(date_text)
            wear = None if record is None else float(record.get('wear', 0) or 0)
            income = None if record is None else float(record.get('income', 0) or 0)
            intensity = 0.0
            if wear is not None and max_abs_wear > 0:
                intensity = min(abs(wear) / max_abs_wear, 1.0)
            cells.append(
                {
                    'date': date_text,
                    'day': day.day,
                    'in_month': day.month == month,
                    'is_today': date_text == note_text,
                    'record': record,
                    'wear': wear,
                    'income': income,
                    'note': '' if record is None else str(record.get('note') or ''),
                    'wear_account_count': 0 if record is None else int(record.get('wear_account_count', 0) or 0),
                    'income_account_count': 0 if record is None else int(record.get('income_account_count', 0) or 0),
                    'avg_wear': None if record is None else record.get('avg_wear'),
                    'avg_income': None if record is None else record.get('avg_income'),
                    'breakdown': [] if record is None else list(record.get('breakdown', [])),
                    'intensity': intensity,
                }
            )
        weeks.append(cells)
    return weeks


def sort_breakdown_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows.sort(key=lambda item: (-abs(float(item.get('wear', 0) or 0)), item.get('name', '')))
    return rows
