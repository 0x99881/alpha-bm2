from __future__ import annotations

from datetime import datetime

from .header_locator import find_column
from .sheet_metadata import meta_to_date_map, normalize_day_code_date_text
from .value_normalizer import normalize_income, normalize_wear
from .value_sheet_spec import get_value_sheet_spec


_SHEET_HANDLER_ATTRS = {
    'wear': 'wear_sheet',
    'income': 'income_sheet',
    'expense': 'expense_sheet',
}


def read_member_value_records(owner, workbook, *, name: str, year_hint: int | None, sheet_type: str) -> list[dict[str, object]]:
    spec = get_value_sheet_spec(sheet_type)
    handler = _get_sheet_handler(owner, sheet_type)
    handler.ensure_structure(workbook)
    sheet = handler.sheet(workbook)
    date_by_number = meta_to_date_map(workbook, spec['meta_sheet'])
    name_col = find_column(sheet, spec['name_header'])
    if year_hint is None:
        year_hint = datetime.now().year
    if name_col is None:
        return []

    target_row = None
    for row in range(2, sheet.max_row + 1):
        if str(sheet.cell(row, name_col).value or '') == name:
            target_row = row
            break
    if target_row is None:
        return []

    normalizer = _get_value_normalizer(sheet_type)
    rows = []
    for number, col in handler.columns(sheet):
        raw_date = date_by_number.get(number, str(sheet.cell(1, col).value))
        rows.append(
            {
                'date': normalize_day_code_date_text(raw_date, year_hint),
                spec['value_key']: normalizer(sheet.cell(target_row, col).value or 0),
            }
        )
    rows.sort(key=lambda item: item['date'])
    return rows


def get_value_sheet_columns(owner, workbook, sheet_type: str) -> list[tuple[int, int]]:
    handler = _get_sheet_handler(owner, sheet_type)
    sheet = handler.sheet(workbook)
    return handler.columns(sheet)


def _get_sheet_handler(owner, sheet_type: str):
    return getattr(owner, _SHEET_HANDLER_ATTRS[sheet_type])


def _get_value_normalizer(sheet_type: str):
    if sheet_type == 'wear':
        return normalize_wear
    return normalize_income
