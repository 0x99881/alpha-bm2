from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from ..profit_calendar_utils import empty_member_day_record, sort_breakdown_rows
from .score_date_reader import latest_header_date, score_date_columns
from .value_sheet_reader import read_member_value_records
from .header_locator import find_column
from .value_sheet_spec import get_value_sheet_spec
from .value_normalizer import normalize_expense, normalize_income, normalize_wear


class ExcelWorkbookReader:
    def __init__(self, store) -> None:
        self._store = store

    def _read_score_rankings_from_sheet(self, sheet) -> list[dict[str, int | str]]:
        return self._store.score_sheet.read_rankings(sheet)

    def get_active_member_profit_map(self) -> dict[str, float]:
        workbook = self._store.workbook_repository.open()
        try:
            self._store.ensure_score_sheet_structure(workbook)
            sheet = self._store.score_sheet_for(workbook)
            return self._store.score_sheet.active_member_profit_map(sheet, self._store.get_active_members())
        finally:
            workbook.close()

    def get_score_summary_data(self) -> dict[str, Any]:
        workbook = self._store.workbook_repository.open()
        try:
            self._store.ensure_score_sheet_structure(workbook)
            sheet = self._store.score_sheet_for(workbook)
            score_columns = score_date_columns(self._store.score_sheet, sheet)
            latest_column = self._store.score_sheet.latest_column(sheet)
            rankings = self._read_score_rankings_from_sheet(sheet)
            rankings.sort(key=lambda item: (-int(item['total']), str(item['name'])))
            return {
                'latest_column': latest_column,
                'column_count': len(score_columns),
                'rankings': rankings,
            }
        finally:
            workbook.close()

    def get_score_sheet_snapshot(self) -> dict[str, Any]:
        workbook = self._store.workbook_repository.open()
        try:
            return self._store.score_sheet.snapshot(workbook)
        finally:
            workbook.close()

    def _get_member_value_records(
        self,
        *,
        name: str,
        year_hint: int | None,
        sheet_type: str,
        workbook=None,
    ) -> list[dict[str, Any]]:
        own_workbook = workbook is None
        if own_workbook:
            workbook = self._store.workbook_repository.open()
        try:
            return read_member_value_records(self._store, workbook, name=name, year_hint=year_hint, sheet_type=sheet_type)
        finally:
            if own_workbook:
                workbook.close()

    def get_member_wear_records(self, name: str, year_hint: int | None = None, workbook=None) -> list[dict[str, Any]]:
        return self._get_member_value_records(name=name, year_hint=year_hint, sheet_type='wear', workbook=workbook)

    def get_member_income_records(self, name: str, year_hint: int | None = None, workbook=None) -> list[dict[str, Any]]:
        return self._get_member_value_records(name=name, year_hint=year_hint, sheet_type='income', workbook=workbook)

    def get_member_expense_records(self, name: str, year_hint: int | None = None, workbook=None) -> list[dict[str, Any]]:
        return self._get_member_value_records(name=name, year_hint=year_hint, sheet_type='expense', workbook=workbook)

    def get_wear_sheet_snapshot(self) -> dict[str, Any]:
        workbook = self._store.workbook_repository.open()
        try:
            return self._store.wear_sheet.snapshot(workbook)
        finally:
            workbook.close()

    def get_value_sheet_snapshot(self, sheet_type: str) -> dict[str, Any]:
        spec = get_value_sheet_spec(sheet_type)
        handler = getattr(self._store, f"{sheet_type}_sheet")
        workbook = self._store.workbook_repository.open()
        try:
            sheet = handler.sheet(workbook)
            headers = [sheet.cell(1, col).value for col in range(1, sheet.max_column + 1)]
            name_col = find_column(sheet, spec["name_header"])
            raw_rows = []
            # The cycle-block layout puts the active block on rows 2..M and any
            # historical cycles below as self-contained sub-tables (preceded
            # by banner rows, with their own header / 周期合计 / 周期平均
            # markers in name_col). Only the active block belongs in the
            # chart/preview snapshot — historical blocks reference different
            # date columns that wouldn't align with row 1's headers.
            from .cycle_block_sheets import CYCLE_AVG_LABEL, CYCLE_BANNER_PREFIX, CYCLE_TOTAL_LABEL, WEAR_DAILY_AVG_LABEL
            block_terminators = {CYCLE_TOTAL_LABEL, CYCLE_AVG_LABEL, WEAR_DAILY_AVG_LABEL}
            for row in range(2, sheet.max_row + 1):
                if bool(sheet.row_dimensions[row].hidden):
                    continue
                first_cell = str(sheet.cell(row, 1).value or "").strip()
                if first_cell.startswith(CYCLE_BANNER_PREFIX):
                    break  # entered historical territory
                values = [sheet.cell(row, col).value for col in range(1, sheet.max_column + 1)]
                if name_col is not None:
                    name_value = str(values[name_col - 1] or "").strip()
                    if not name_value:
                        continue
                    if name_value in block_terminators:
                        break
                raw_rows.append(values)
            value_col_indices = [col - 1 for _, col in handler.columns(sheet)]
            return {
                "headers": headers,
                "raw_rows": raw_rows,
                "value_col_indices": value_col_indices,
            }
        finally:
            workbook.close()

    def _build_member_record_map(self, name: str, year: int, month: int, workbook=None) -> dict[str, dict[str, Any]]:
        """Per-month wear / income / expense map straight from SQLite.

        Previously this aggregated Excel sheets. With the cycle-block sheet
        layout, a member appears in multiple rows (one per cycle), so reading
        from Excel cleanly is awkward. The DB is the source of truth and has
        every entry indexed by date, so we ask it directly.
        """
        local_db = getattr(self._store, 'local_db', None)
        month_prefix = f'{year:04d}-{month:02d}-'
        if local_db is None:
            return {}

        from ..value_utils import to_float_or_none
        from ..domain.rules.daily_entry import resolve_wear_value, IncompleteBalanceInput
        from .value_normalizer import normalize_expense, normalize_income, normalize_wear, parse_decimal

        merged_map: dict[str, dict[str, Any]] = {}
        for row in local_db.get_score_rows():
            row_name = str(row.get('member_name', '') or '').strip()
            if row_name != name:
                continue
            date_text = str(row.get('score_date', '') or '').strip()
            if not date_text.startswith(month_prefix):
                continue
            entry = merged_map.setdefault(date_text, empty_member_day_record(date_text))
            income = to_float_or_none(str(row.get('income', '') or '').strip())
            expense = to_float_or_none(str(row.get('other_expense', '') or '').strip())
            try:
                wear = resolve_wear_value(
                    before_text=str(row.get('before_balance', '') or '').strip(),
                    after_text=str(row.get('after_balance', '') or '').strip(),
                    manual_wear_text=str(row.get('manual_wear', '') or '').strip(),
                    parse_decimal=parse_decimal,
                    manual_field='manual_wear',
                    before_field='before',
                    after_field='after',
                )
            except (IncompleteBalanceInput, ValueError):
                wear = None
            if income is not None:
                entry['income'] = normalize_income(income)
            if expense is not None:
                entry['expense'] = normalize_expense(expense)
            if wear is not None:
                entry['wear'] = normalize_wear(wear)
        return merged_map

    def get_member_calendar_dataset(self, name: str, year: int, month: int, workbook=None) -> dict[str, Any]:
        month_record_map = self._build_member_record_map(name, year, month, workbook=workbook)
        month_records = self._finalize_member_month_records(month_record_map)
        max_abs_wear = max((abs(float(item['wear'])) for item in month_records), default=0.0)
        return {
            'year': year,
            'month': month,
            'month_records': month_records,
            'record_map': month_record_map,
            'max_abs_wear': max_abs_wear,
            'note_text': '',
        }

    def _finalize_member_month_records(self, month_record_map: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
        month_records = sorted(month_record_map.values(), key=lambda item: item['date'], reverse=True)
        for item in month_records:
            wear_value = float(item.get('wear', 0) or 0)
            income_value = float(item.get('income', 0) or 0)
            item['wear_account_count'] = 1 if wear_value != 0 else 0
            item['income_account_count'] = 1 if income_value != 0 else 0
            item['avg_wear'] = wear_value
            item['avg_income'] = income_value
            item['breakdown'] = []
        return month_records

    def get_all_members_calendar_dataset(self, year: int, month: int) -> dict[str, Any]:
        workbook = self._store.workbook_repository.open()
        try:
            merged_map: dict[str, dict[str, Any]] = {}
            for member in self._store.get_active_members():
                member_name = member['name']
                member_calendar = self.get_member_calendar_dataset(member_name, year, month, workbook=workbook)
                for item in member_calendar['month_records']:
                    self._merge_all_member_day_record(merged_map, member_name=member_name, item=item)
        finally:
            workbook.close()

        month_records = self._finalize_all_members_month_records(merged_map)
        max_abs_wear = max((abs(float(item['wear'])) for item in month_records), default=0.0)
        return {
            'year': year,
            'month': month,
            'month_records': month_records,
            'record_map': {item['date']: item for item in month_records},
            'max_abs_wear': max_abs_wear,
            'note_text': '閸欏苯鍤仦鏇炵磻閸氬嫬褰块弰搴ｇ矎',
        }

    def _merge_all_member_day_record(self, merged_map: dict[str, dict[str, Any]], *, member_name: str, item: dict[str, Any]) -> None:
        day_record = merged_map.setdefault(
            item['date'],
            {
                'date': item['date'],
                'wear': 0.0,
                'income': 0.0,
                'expense': 0.0,
                'note': '',
                'wear_account_count': 0,
                'income_account_count': 0,
                'breakdown': [],
            },
        )
        wear_value = float(item.get('wear', 0) or 0)
        income_value = float(item.get('income', 0) or 0)
        expense_value = float(item.get('expense', 0) or 0)
        day_record['wear'] = normalize_wear(float(day_record['wear']) + wear_value)
        day_record['income'] = normalize_income(float(day_record['income']) + income_value)
        day_record['expense'] = normalize_expense(float(day_record['expense']) + expense_value)
        if wear_value != 0:
            day_record['wear_account_count'] = int(day_record['wear_account_count']) + 1
        if income_value != 0:
            day_record['income_account_count'] = int(day_record['income_account_count']) + 1
        if wear_value != 0 or income_value != 0 or expense_value != 0:
            day_record['breakdown'].append(
                {
                    'name': member_name,
                    'wear': normalize_wear(wear_value),
                    'income': normalize_income(income_value),
                    'expense': normalize_expense(expense_value),
                }
            )

    def _finalize_all_members_month_records(self, merged_map: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
        month_records = sorted(merged_map.values(), key=lambda item: item['date'], reverse=True)
        for item in month_records:
            wear_count = int(item.get('wear_account_count', 0) or 0)
            income_count = int(item.get('income_account_count', 0) or 0)
            item['avg_wear'] = normalize_wear(float(item['wear']) / wear_count) if wear_count else 0.0
            item['avg_income'] = normalize_income(float(item['income']) / income_count) if income_count else 0.0
            sort_breakdown_rows(item['breakdown'])
        return month_records

    def get_next_score_date(self) -> str:
        workbook = self._store.workbook_repository.open()
        try:
            latest_date = latest_header_date(workbook, self._store.score_sheet)
        finally:
            workbook.close()
        if latest_date is not None:
            return (latest_date + timedelta(days=1)).strftime('%Y-%m-%d')
        configured = str(self._store.config_repository.load().get('default_score_date') or '').strip()
        if configured:
            try:
                return datetime.strptime(configured, '%Y-%m-%d').strftime('%Y-%m-%d')
            except ValueError:
                return datetime.now().strftime('%Y-%m-%d')
        return datetime.now().strftime('%Y-%m-%d')

