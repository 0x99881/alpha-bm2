from __future__ import annotations

from typing import Any

from .constants import NAME_HEADER, PROFIT_HEADER, TOTAL_HEADER, VALUE_SHEET_SPECS, WEAR_NAME_HEADER
from .profit_calendar_utils import empty_member_day_record, sort_breakdown_rows
from .value_utils import build_threshold_abnormal_flags, to_float_or_none


class StoreReadMixin:
    def _value_sheet_spec(self, sheet_type: str) -> dict[str, Any]:
        return VALUE_SHEET_SPECS[sheet_type]

    def _value_sheet_read_helpers(self, sheet_type: str):
        spec = self._value_sheet_spec(sheet_type)
        return (
            spec,
            getattr(self, spec['sheet_method']),
            getattr(self, spec['ensure_method']),
            getattr(self, spec['columns_method']),
            getattr(self, spec['round_method']),
        )

    def get_score_rankings(self, limit: int = 999, workbook=None) -> list[dict[str, int | str]]:
        own_workbook = workbook is None
        if own_workbook:
            workbook = self._open_workbook()
        self._ensure_score_sheet_structure(workbook)
        sheet = self._score_sheet(workbook)
        total_col = self._find_column(sheet, TOTAL_HEADER)
        name_col = self._find_column(sheet, NAME_HEADER)
        rows = []
        if total_col and name_col:
            for row in range(2, sheet.max_row + 1):
                name = sheet.cell(row, name_col).value
                total = sheet.cell(row, total_col).value
                if name:
                    rows.append({'name': str(name), 'total': int(total or 0)})
        if own_workbook:
            workbook.close()
        rows.sort(key=lambda item: (-int(item['total']), str(item['name'])))
        return rows[:limit]

    def get_score_latest_column(self) -> str:
        workbook = self._open_workbook()
        self._ensure_score_sheet_structure(workbook)
        sheet = self._score_sheet(workbook)
        d_cols = self._d_columns(sheet)
        latest_column = str(sheet.cell(1, d_cols[-1][1]).value) if d_cols else '-'
        workbook.close()
        return latest_column

    def get_score_column_count(self) -> int:
        workbook = self._open_workbook()
        self._ensure_score_sheet_structure(workbook)
        sheet = self._score_sheet(workbook)
        count = len(self._d_columns(sheet))
        workbook.close()
        return count

    def get_active_member_profit_map(self) -> dict[str, float]:
        workbook = self._open_workbook()
        self._ensure_score_sheet_structure(workbook)
        sheet = self._score_sheet(workbook)
        name_col = self._find_column(sheet, NAME_HEADER)
        profit_col = self._find_column(sheet, PROFIT_HEADER)
        profit_map = {member['name']: 0.0 for member in self.get_active_members()}
        if name_col is not None and profit_col is not None:
            for row in range(2, sheet.max_row + 1):
                member_name = str(sheet.cell(row, name_col).value or '').strip()
                if member_name not in profit_map:
                    continue
                profit_map[member_name] = self._round_income(sheet.cell(row, profit_col).value or 0)
        workbook.close()
        return profit_map

    def _get_member_value_records(
        self,
        *,
        name: str,
        year_hint: int | None,
        sheet_type: str,
        workbook=None,
    ) -> list[dict[str, Any]]:
        spec, sheet_getter, ensure_structure, columns_getter, value_formatter = self._value_sheet_read_helpers(sheet_type)
        own_workbook = workbook is None
        if own_workbook:
            workbook = self._open_workbook()
        ensure_structure(workbook)
        sheet = sheet_getter(workbook)
        date_by_number = self._meta_to_date_map(workbook, spec['meta_sheet'])
        name_col = self._find_column(sheet, spec['name_header'])
        rows = []
        if year_hint is None:
            year_hint = datetime.now().year
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
                            'date': self._normalize_wear_date_text(raw_date, year_hint),
                            spec['value_key']: value_formatter(sheet.cell(target_row, col).value or 0),
                        }
                    )
        if own_workbook:
            workbook.close()
        rows.sort(key=lambda item: item['date'])
        return rows

    def get_member_wear_records(self, name: str, year_hint: int | None = None, workbook=None) -> list[dict[str, Any]]:
        return self._get_member_value_records(
            name=name,
            year_hint=year_hint,
            sheet_type='wear',
            workbook=workbook,
        )

    def get_member_income_records(self, name: str, year_hint: int | None = None, workbook=None) -> list[dict[str, Any]]:
        return self._get_member_value_records(
            name=name,
            year_hint=year_hint,
            sheet_type='income',
            workbook=workbook,
        )

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
        workbook = self._open_workbook()
        self._ensure_wear_sheet_structure(workbook)
        sheet = self._wear_sheet(workbook)
        headers = [sheet.cell(1, col).value for col in range(1, sheet.max_column + 1)]
        raw_rows = []
        name_col = self._find_column(sheet, WEAR_NAME_HEADER)
        for row in range(2, sheet.max_row + 1):
            values = [sheet.cell(row, col).value for col in range(1, sheet.max_column + 1)]
            if name_col is not None and not values[name_col - 1]:
                continue
            raw_rows.append(values)
        wear_col_indices = [col - 1 for _, col in self._wear_columns(sheet)]
        threshold = self.get_wear_abnormal_threshold()
        workbook.close()
        return {
            'headers': headers,
            'raw_rows': raw_rows,
            'wear_col_indices': wear_col_indices,
            'threshold': threshold,
        }

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
        workbook = self._open_workbook()
        merged_map: dict[str, dict[str, Any]] = {}
        for member in self.get_active_members():
            member_name = member['name']
            member_calendar = self.get_member_calendar_dataset(member_name, year, month, workbook=workbook)
            for item in member_calendar['month_records']:
                self._merge_all_member_day_record(merged_map, member_name=member_name, item=item)
        workbook.close()

        month_records = self._finalize_all_members_month_records(merged_map)
        max_abs_wear = max((abs(float(item['wear'])) for item in month_records), default=0.0)
        return {
            'year': year,
            'month': month,
            'month_records': month_records,
            'record_map': {item['date']: item for item in month_records},
            'max_abs_wear': max_abs_wear,
            'note_text': '双击展开各号明细',
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
        day_record['wear'] = self._round_wear(float(day_record['wear']) + wear_value)
        day_record['income'] = self._round_income(float(day_record['income']) + income_value)
        if wear_value != 0:
            day_record['wear_account_count'] = int(day_record['wear_account_count']) + 1
        if income_value != 0:
            day_record['income_account_count'] = int(day_record['income_account_count']) + 1
        if wear_value != 0 or income_value != 0:
            day_record['breakdown'].append(
                {
                    'name': member_name,
                    'wear': self._round_wear(wear_value),
                    'income': self._round_income(income_value),
                }
            )

    def _finalize_all_members_month_records(self, merged_map: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
        month_records = sorted(merged_map.values(), key=lambda item: item['date'], reverse=True)
        for item in month_records:
            wear_count = int(item.get('wear_account_count', 0) or 0)
            income_count = int(item.get('income_account_count', 0) or 0)
            item['avg_wear'] = self._round_wear(float(item['wear']) / wear_count) if wear_count else 0.0
            item['avg_income'] = self._round_income(float(item['income']) / income_count) if income_count else 0.0
            sort_breakdown_rows(item['breakdown'])
        return month_records
