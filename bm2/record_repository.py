from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from .constants import EXPENSE_META_SHEET, INCOME_META_SHEET, META_SHEET, VALUE_SHEET_SPECS, WEAR_META_SHEET
from .store_read import StoreReadMixin
from .store_write import StoreWriteMixin


class RecordRepository:
    def __init__(self, store) -> None:
        self.store = store

    def get_score_rankings(self, limit: int = 999, workbook=None):
        return StoreReadMixin.get_score_rankings(self.store, limit=limit, workbook=workbook)

    def get_score_latest_column(self):
        return StoreReadMixin.get_score_latest_column(self.store)

    def get_score_column_count(self):
        return StoreReadMixin.get_score_column_count(self.store)

    def get_active_member_profit_map(self):
        return StoreReadMixin.get_active_member_profit_map(self.store)

    def get_active_members(self):
        return self.store.get_active_members()

    def _value_sheet_spec(self, sheet_type: str) -> dict[str, Any]:
        return VALUE_SHEET_SPECS[sheet_type]

    def _get_member_value_records(self, name: str, year_hint: int | None, sheet_type: str, workbook=None):
        spec = self._value_sheet_spec(sheet_type)
        own_workbook = workbook is None
        if own_workbook:
            workbook = self.store.workbook_manager.open_workbook()
        try:
            ensure_structure = getattr(self.store, spec['ensure_method'])
            sheet_getter = getattr(self.store, spec['sheet_method'])
            columns_getter = getattr(self.store, spec['columns_method'])
            value_formatter = getattr(self.store, spec['round_method'])

            ensure_structure(workbook)
            sheet = sheet_getter(workbook)
            date_by_number = self.store._meta_to_date_map(workbook, spec['meta_sheet'])
            name_col = self.store._find_column(sheet, spec['name_header'])
            rows = []
            effective_year = year_hint or datetime.now().year
            if name_col is not None:
                target_row = None
                for row in range(2, sheet.max_row + 1):
                    if str(sheet.cell(row, name_col).value or '') == name:
                        target_row = row
                        break
                if target_row is not None:
                    for number, col in columns_getter(sheet):
                        raw_date = date_by_number.get(number, str(sheet.cell(1, col).value))
                        rows.append(
                            {
                                'date': self.store._normalize_wear_date_text(raw_date, effective_year),
                                spec['value_key']: value_formatter(sheet.cell(target_row, col).value or 0),
                            }
                        )
            rows.sort(key=lambda item: item['date'])
            return rows
        finally:
            if own_workbook:
                workbook.close()

    def get_member_wear_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self._get_member_value_records(name=name, year_hint=year_hint, sheet_type='wear', workbook=workbook)

    def get_member_income_records(self, name: str, year_hint: int | None = None, workbook=None):
        return self._get_member_value_records(name=name, year_hint=year_hint, sheet_type='income', workbook=workbook)

    def get_wear_sheet_snapshot(self):
        return StoreReadMixin.get_wear_sheet_snapshot(self.store)

    def get_member_calendar_dataset(self, name: str, year: int, month: int, workbook=None):
        return StoreReadMixin.get_member_calendar_dataset(self.store, name=name, year=year, month=month, workbook=workbook)

    def get_all_members_calendar_dataset(self, year: int, month: int):
        return StoreReadMixin.get_all_members_calendar_dataset(self.store, year=year, month=month)

    def _latest_header_date(self, workbook):
        latest_meta_dates = [
            self.store._parse_saved_date(saved_date)
            for sheet_name in (META_SHEET, WEAR_META_SHEET, INCOME_META_SHEET, EXPENSE_META_SHEET)
            for saved_date, _ in self.store._read_sheet_meta(workbook, sheet_name)
        ]
        year_hint = max(latest_meta_dates).year if latest_meta_dates else datetime.now().year
        latest_dates = list(latest_meta_dates)

        score_sheet = self.store._score_sheet(workbook)
        score_columns = self.store._d_columns(score_sheet)
        if score_columns:
            latest_score_header = str(score_sheet.cell(1, score_columns[-1][1]).value or '').strip()
            if self.store._is_mmdd_header(latest_score_header):
                latest_dates.append(datetime.strptime(f'{year_hint}-{latest_score_header}', '%Y-%m-%d').date())

        for sheet_type in ('wear', 'income', 'expense'):
            spec = self._value_sheet_spec(sheet_type)
            sheet_getter = getattr(self.store, spec['sheet_method'])
            columns_getter = getattr(self.store, spec['columns_method'])
            sheet = sheet_getter(workbook)
            columns = columns_getter(sheet)
            if not columns:
                continue
            latest_header = f'{columns[-1][0]:04d}'
            latest_dates.append(datetime.strptime(f'{year_hint}-{latest_header[:2]}-{latest_header[2:]}', '%Y-%m-%d').date())

        return max(latest_dates) if latest_dates else None

    def get_next_score_date(self) -> str:
        workbook = self.store.workbook_manager.open_workbook()
        try:
            latest_date = self._latest_header_date(workbook)
            if latest_date is None:
                return datetime.now().strftime('%Y-%m-%d')
            return (latest_date + timedelta(days=1)).strftime('%Y-%m-%d')
        finally:
            workbook.close()

    def save_scores_and_wear(self, date_text: str, entries):
        return StoreWriteMixin.save_scores_and_wear(self.store, date_text, entries)
