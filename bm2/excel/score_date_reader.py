from __future__ import annotations

from datetime import datetime

from .sheet_metadata import parse_saved_date, read_sheet_meta
from .value_sheet_reader import get_value_sheet_columns
from ..constants import EXPENSE_META_SHEET, INCOME_META_SHEET, WEAR_META_SHEET


def score_date_columns(score_sheet_handler, sheet) -> list[tuple[int, int]]:
    return score_sheet_handler.date_columns(sheet)


def latest_header_date(workbook, score_sheet_handler) -> object | None:
    score_latest_date = score_sheet_handler.latest_used_date(workbook)
    latest_meta_dates = [
        parse_saved_date(saved_date)
        for sheet_name in (WEAR_META_SHEET, INCOME_META_SHEET, EXPENSE_META_SHEET)
        for saved_date, _ in read_sheet_meta(workbook, sheet_name)
    ]
    if score_latest_date is not None:
        latest_meta_dates.append(score_latest_date)
    year_hint = max(latest_meta_dates).year if latest_meta_dates else datetime.now().year
    latest_dates = list(latest_meta_dates)

    for sheet_type in ('wear', 'income', 'expense'):
        columns = get_value_sheet_columns(score_sheet_handler.store, workbook, sheet_type)
        if not columns:
            continue
        latest_header = f'{columns[-1][0]:04d}'
        latest_dates.append(datetime.strptime(f'{year_hint}-{latest_header[:2]}-{latest_header[2:]}', '%Y-%m-%d').date())

    return max(latest_dates) if latest_dates else None
