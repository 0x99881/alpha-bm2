from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from ..profit_calendar_utils import empty_member_day_record, sort_breakdown_rows
from ..value_utils import build_threshold_abnormal_flags, to_float_or_none
from .score_date_reader import latest_header_date, score_date_columns
from .value_sheet_reader import read_member_value_records
from .value_normalizer import normalize_income, normalize_wear


class ExcelWorkbookReader:
    def __init__(self, store) -> None:
        self._store = store

    def get_score_rankings(self, limit: int = 999, workbook=None) -> list[dict[str, int | str]]:
        own_workbook = workbook is None
        if own_workbook:
            workbook = self._store.workbook_repository.open()
        rows = []
        try:
            self._store._ensure_score_sheet_structure(workbook)
            sheet = self._store._score_sheet(workbook)
            rows = self._read_score_rankings_from_sheet(sheet)
        finally:
            if own_workbook:
                workbook.close()
        rows.sort(key=lambda item: (-int(item['total']), str(item['name'])))
        return rows[:limit]

    def _read_score_rankings_from_sheet(self, sheet) -> list[dict[str, int | str]]:
        return self._store.score_sheet.read_rankings(sheet)

    def get_score_latest_column(self) -> str:
        workbook = self._store.workbook_repository.open()
        try:
            self._store._ensure_score_sheet_structure(workbook)
            sheet = self._store._score_sheet(workbook)
            return self._store.score_sheet.latest_column(sheet)
        finally:
            workbook.close()

    def get_score_column_count(self) -> int:
        workbook = self._store.workbook_repository.open()
        try:
            self._store._ensure_score_sheet_structure(workbook)
            sheet = self._store._score_sheet(workbook)
            return self._store.score_sheet.column_count(sheet)
        finally:
            workbook.close()

    def get_active_member_profit_map(self) -> dict[str, float]:
        workbook = self._store.workbook_repository.open()
        try:
            self._store._ensure_score_sheet_structure(workbook)
            sheet = self._store._score_sheet(workbook)
            return self._store.score_sheet.active_member_profit_map(sheet, self._store.get_active_members())
        finally:
            workbook.close()

    def get_score_summary_data(self) -> dict[str, Any]:
        workbook = self._store.workbook_repository.open()
        try:
            self._store._ensure_score_sheet_structure(workbook)
            sheet = self._store._score_sheet(workbook)
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

    def _to_float_or_none(self, value: Any) -> float | None:
        return to_float_or_none(value)

    def _median(self, sorted_values: list[float]) -> float:
        if not sorted_values:
            return 0.0
        middle = len(sorted_values) // 2
        if len(sorted_values) % 2:
            return sorted_values[middle]
        return (sorted_values[middle - 1] + sorted_values[middle]) / 2

    def _build_abnormal_flags(self, rows: list[list[Any]], target_cols: list[int], threshold: float) -> dict[tuple[int, int], bool]:
        return build_threshold_abnormal_flags(rows, target_cols, threshold)

    def get_wear_sheet_snapshot(self) -> dict[str, Any]:
        workbook = self._store.workbook_repository.open()
        try:
            return self._store.wear_sheet.snapshot(workbook)
        finally:
            workbook.close()

    def _build_member_record_map(self, name: str, year: int, month: int, workbook=None) -> dict[str, dict[str, Any]]:
        wear_records = self.get_member_wear_records(name, year_hint=year, workbook=workbook)
        income_records = self.get_member_income_records(name, year_hint=year, workbook=workbook)
        merged_map: dict[str, dict[str, Any]] = {}
        for item in wear_records:
            merged_map.setdefault(item['date'], empty_member_day_record(item['date']))
            merged_map[item['date']]['wear'] = item['wear']
        for item in income_records:
            merged_map.setdefault(item['date'], empty_member_day_record(item['date']))
            merged_map[item['date']]['income'] = item['income']
        month_prefix = f'{year:04d}-{month:02d}-'
        return {key: value for key, value in merged_map.items() if key.startswith(month_prefix)}

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
                'note': '',
                'wear_account_count': 0,
                'income_account_count': 0,
                'breakdown': [],
            },
        )
        wear_value = float(item.get('wear', 0) or 0)
        income_value = float(item.get('income', 0) or 0)
        day_record['wear'] = normalize_wear(float(day_record['wear']) + wear_value)
        day_record['income'] = normalize_income(float(day_record['income']) + income_value)
        if wear_value != 0:
            day_record['wear_account_count'] = int(day_record['wear_account_count']) + 1
        if income_value != 0:
            day_record['income_account_count'] = int(day_record['income_account_count']) + 1
        if wear_value != 0 or income_value != 0:
            day_record['breakdown'].append(
                {
                    'name': member_name,
                    'wear': normalize_wear(wear_value),
                    'income': normalize_income(income_value),
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

