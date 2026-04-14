from __future__ import annotations

from datetime import datetime
from typing import Any

from .constants import INCOME_META_SHEET, INCOME_NAME_HEADER, NAME_HEADER, PROFIT_HEADER, TOTAL_HEADER, WEAR_META_SHEET, WEAR_NAME_HEADER, WINDOW_SIZE
from .profit_calendar_utils import (
    build_calendar_weeks,
    build_month_label,
    build_month_neighbors,
    empty_member_day_record,
    sort_breakdown_rows,
)
from .ui_text import UI_TEXT


class StoreReadMixin:
    def _sum_profit_calendar_totals(self, month_records: list[dict[str, Any]], stats_allowed: bool) -> tuple[float, float]:
        """Return month income and wear totals using the existing calendar records."""
        if not stats_allowed:
            return 0.0, 0.0
        month_income = self._round_income(sum(float(item.get('income', 0) or 0) for item in month_records))
        month_wear = self._round_wear(sum(float(item.get('wear', 0) or 0) for item in month_records))
        return month_income, month_wear

    def _count_profit_stats_members(self, *, name: str, stats_allowed: bool, active_members: list[dict[str, Any]]) -> int:
        if name == 'all':
            return len(active_members)
        return 1 if stats_allowed else 0

    def _build_profit_average_stats(
        self,
        *,
        month_income: float,
        month_wear: float,
        active_member_count: int,
        data_day_count: int,
    ) -> dict[str, float | int]:
        avg_income = self._round_income(month_income / active_member_count) if active_member_count else 0.0
        avg_wear = self._round_wear(month_wear / active_member_count) if active_member_count else 0.0
        avg_per_member_per_day = round(month_wear / (active_member_count * data_day_count), 2) if active_member_count and data_day_count else 0.0
        avg_member_daily_wear = round(month_wear / data_day_count, 2) if data_day_count else 0.0
        return {
            'month_avg_income': avg_income,
            'month_avg_wear': avg_wear,
            'month_income_account_count': active_member_count,
            'month_wear_account_count': active_member_count,
            'month_day_count': data_day_count,
            'avg_per_member_per_day': avg_per_member_per_day,
            'avg_member_daily_wear': avg_member_daily_wear,
        }

    def _build_profit_calendar_payload(
        self,
        *,
        calendar_data: dict[str, Any],
        month_income: float,
        month_wear: float,
        average_stats: dict[str, float | int],
        board_rows: list[dict[str, Any]],
        profit_positive_list: list[dict[str, Any]],
        profit_negative_list: list[dict[str, Any]],
        stats_allowed: bool,
        is_all_members: bool,
    ) -> dict[str, Any]:
        return {
            **calendar_data,
            'month_income_total': month_income,
            'month_wear_total': month_wear,
            'month_profit_total': self._round_income(month_income - month_wear),
            **average_stats,
            'profit_board_rows': board_rows,
            'profit_positive_list': profit_positive_list,
            'profit_negative_list': profit_negative_list,
            'today': datetime.now().strftime('%Y-%m-%d'),
            'is_all_members': is_all_members,
            'stats_allowed': stats_allowed,
        }

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

    def get_score_summary(self) -> dict[str, Any]:
        workbook = self._open_workbook()
        self._ensure_score_sheet_structure(workbook)
        sheet = self._score_sheet(workbook)
        d_cols = self._d_columns(sheet)
        latest_column = str(sheet.cell(1, d_cols[-1][1]).value) if d_cols else '-'
        rankings = self.get_score_rankings(limit=999, workbook=workbook)
        workbook.close()
        return {
            'latest_column': latest_column,
            'window_size': min(WINDOW_SIZE, len(d_cols)),
            'rankings': rankings,
        }

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
        sheet_getter,
        ensure_structure,
        meta_sheet_name: str,
        name_header: str,
        columns_getter,
        value_key: str,
        value_formatter,
        workbook=None,
    ) -> list[dict[str, Any]]:
        own_workbook = workbook is None
        if own_workbook:
            workbook = self._open_workbook()
        ensure_structure(workbook)
        sheet = sheet_getter(workbook)
        date_by_number = self._meta_to_date_map(workbook, meta_sheet_name)
        name_col = self._find_column(sheet, name_header)
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
                            value_key: value_formatter(sheet.cell(target_row, col).value or 0),
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
            sheet_getter=self._wear_sheet,
            ensure_structure=self._ensure_wear_sheet_structure,
            meta_sheet_name=WEAR_META_SHEET,
            name_header=WEAR_NAME_HEADER,
            columns_getter=self._wear_columns,
            value_key='wear',
            value_formatter=self._round_wear,
            workbook=workbook,
        )

    def get_member_income_records(self, name: str, year_hint: int | None = None, workbook=None) -> list[dict[str, Any]]:
        return self._get_member_value_records(
            name=name,
            year_hint=year_hint,
            sheet_getter=self._income_sheet,
            ensure_structure=self._ensure_income_sheet_structure,
            meta_sheet_name=INCOME_META_SHEET,
            name_header=INCOME_NAME_HEADER,
            columns_getter=self._income_columns,
            value_key='income',
            value_formatter=self._round_income,
            workbook=workbook,
        )

    def _to_float_or_none(self, value: Any) -> float | None:
        if value in (None, ''):
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    def _median(self, sorted_values: list[float]) -> float:
        if not sorted_values:
            return 0.0
        middle = len(sorted_values) // 2
        if len(sorted_values) % 2:
            return sorted_values[middle]
        return (sorted_values[middle - 1] + sorted_values[middle]) / 2

    def _build_abnormal_flags(self, rows: list[list[Any]], target_cols: list[int], threshold: float) -> dict[tuple[int, int], bool]:
        abnormal_flags: dict[tuple[int, int], bool] = {}
        for col_index in target_cols:
            for row_index, row in enumerate(rows):
                numeric_value = self._to_float_or_none(row[col_index])
                if numeric_value is None:
                    continue
                abnormal_flags[(row_index, col_index)] = numeric_value > threshold
        return abnormal_flags

    def get_wear_sheet_view(self) -> dict[str, Any]:
        workbook = self._open_workbook()
        self._ensure_wear_sheet_structure(workbook)
        sheet = self._wear_sheet(workbook)
        headers = [sheet.cell(1, col).value for col in range(1, sheet.max_column + 1)]
        headers.append(UI_TEXT['wear_member_avg'])
        raw_rows = []
        name_col = self._find_column(sheet, WEAR_NAME_HEADER)
        for row in range(2, sheet.max_row + 1):
            values = [sheet.cell(row, col).value for col in range(1, sheet.max_column + 1)]
            if name_col is not None and not values[name_col - 1]:
                continue
            raw_rows.append(values)
        wear_col_indices = [col - 1 for _, col in self._wear_columns(sheet)]
        threshold = self.get_wear_abnormal_threshold()
        abnormal_flags = self._build_abnormal_flags(raw_rows, wear_col_indices, threshold)
        wear_values = [
            numeric_value
            for row in raw_rows
            for col_index in wear_col_indices
            if (numeric_value := self._to_float_or_none(row[col_index])) is not None
        ]
        abnormal_count = sum(1 for value in wear_values if value > threshold)
        rows = []
        for row_index, row in enumerate(raw_rows):
            row_cells = [
                {
                    'value': value,
                    'is_abnormal': abnormal_flags.get((row_index, col_index), False),
                }
                for col_index, value in enumerate(row)
            ]
            row_wear_values = [
                numeric_value
                for col_index in wear_col_indices
                if (numeric_value := self._to_float_or_none(row[col_index])) is not None
            ]
            row_cells.append(
                {
                    'value': self._round_wear(sum(row_wear_values) / len(row_wear_values)) if row_wear_values else 0.0,
                    'is_abnormal': False,
                }
            )
            rows.append(row_cells)
        workbook.close()
        return {
            'headers': headers,
            'rows': rows,
            'row_count': len(rows),
            'column_count': len(headers),
            'avg_daily_wear': self._round_wear(sum(wear_values) / len(wear_values)) if wear_values else 0.0,
            'abnormal_threshold': threshold,
            'abnormal_count': abnormal_count,
        }

    def _build_calendar_payload(
        self,
        *,
        year: int,
        month: int,
        month_records: list[dict[str, Any]],
        record_map: dict[str, dict[str, Any]],
        max_abs_wear: float,
        note_text: str = '',
    ) -> dict[str, Any]:
        weeks = build_calendar_weeks(
            year=year,
            month=month,
            record_map=record_map,
            max_abs_wear=max_abs_wear,
            note_text=note_text,
        )
        prev_year, prev_month, next_year, next_month = build_month_neighbors(year, month)
        return {
            'year': year,
            'month': month,
            'month_label': build_month_label(year, month),
            'weeks': weeks,
            'records': month_records,
            'prev_year': prev_year,
            'prev_month': prev_month,
            'next_year': next_year,
            'next_month': next_month,
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

    def get_member_calendar(self, name: str, year: int, month: int, workbook=None) -> dict[str, Any]:
        month_record_map = self._build_member_record_map(name, year, month, workbook=workbook)
        month_records = self._finalize_member_month_records(month_record_map)
        max_abs_wear = max((abs(float(item['wear'])) for item in month_records), default=0.0)
        return self._build_calendar_payload(
            year=year,
            month=month,
            month_records=month_records,
            record_map=month_record_map,
            max_abs_wear=max_abs_wear,
        )

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

    def _get_all_members_calendar(self, year: int, month: int) -> dict[str, Any]:
        workbook = self._open_workbook()
        merged_map: dict[str, dict[str, Any]] = {}
        for member in self.get_active_members():
            member_name = member['name']
            member_calendar = self.get_member_calendar(member_name, year, month, workbook=workbook)
            for item in member_calendar['records']:
                self._merge_all_member_day_record(merged_map, member_name=member_name, item=item)
        workbook.close()

        month_records = self._finalize_all_members_month_records(merged_map)
        record_map = {item['date']: item for item in month_records}
        max_abs_wear = max((abs(float(item['wear'])) for item in month_records), default=0.0)
        return self._build_calendar_payload(
            year=year,
            month=month,
            month_records=month_records,
            record_map=record_map,
            max_abs_wear=max_abs_wear,
            note_text=UI_TEXT['double_click_breakdown_hint'],
        )

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

    def _accumulate_profit_board_totals(self, month_records: list[dict[str, Any]], active_members: list[dict[str, Any]]) -> dict[str, dict[str, float | str]]:
        row_map = {item['name']: {'name': item['name'], 'income': 0.0, 'wear': 0.0, 'profit': 0.0} for item in active_members}
        for item in month_records:
            for row in item.get('breakdown', []):
                member_name = row.get('name')
                if member_name not in row_map:
                    continue
                row_map[member_name]['income'] += float(row.get('income', 0) or 0)
                row_map[member_name]['wear'] += float(row.get('wear', 0) or 0)
        return row_map

    def _build_profit_board_rows(self, row_map: dict[str, dict[str, float | str]]) -> list[dict[str, Any]]:
        board_rows = []
        for row in row_map.values():
            income_value = round(float(row['income']), 1)
            wear_value = round(float(row['wear']), 1)
            profit_value = round(income_value - wear_value, 1)
            board_rows.append({
                'name': str(row['name']),
                'income': income_value,
                'wear': wear_value,
                'profit': profit_value,
            })
        board_rows.sort(key=lambda item: (-item['profit'], -item['income'], item['name']))
        return board_rows

    def _split_profit_board_rows(self, board_rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        profit_positive_list = [item for item in board_rows if item['profit'] >= 0]
        profit_negative_list = sorted((item for item in board_rows if item['profit'] < 0), key=lambda item: (item['profit'], item['name']))
        return profit_positive_list, profit_negative_list

    def _build_profit_board_lists(self, month_records: list[dict[str, Any]], active_members: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
        row_map = self._accumulate_profit_board_totals(month_records, active_members)
        board_rows = self._build_profit_board_rows(row_map)
        profit_positive_list, profit_negative_list = self._split_profit_board_rows(board_rows)
        return board_rows, profit_positive_list, profit_negative_list

    def get_member_profit_calendar(self, name: str, year: int, month: int) -> dict[str, Any]:
        active_members = self.get_active_members()
        active_names = {item['name'] for item in active_members}
        stats_allowed = name == 'all' or name in active_names
        calendar_data = self._get_all_members_calendar(year, month) if name == 'all' else self.get_member_calendar(name, year, month)
        month_records = calendar_data['records']
        month_income, month_wear = self._sum_profit_calendar_totals(month_records, stats_allowed)
        active_member_count = self._count_profit_stats_members(name=name, stats_allowed=stats_allowed, active_members=active_members)
        data_day_count = len(month_records)
        average_stats = self._build_profit_average_stats(
            month_income=month_income,
            month_wear=month_wear,
            active_member_count=active_member_count,
            data_day_count=data_day_count,
        )
        board_rows, profit_positive_list, profit_negative_list = self._build_profit_board_lists(month_records, active_members) if name == 'all' else ([], [], [])
        return self._build_profit_calendar_payload(
            calendar_data=calendar_data,
            month_income=month_income,
            month_wear=month_wear,
            average_stats=average_stats,
            board_rows=board_rows,
            profit_positive_list=profit_positive_list,
            profit_negative_list=profit_negative_list,
            stats_allowed=stats_allowed,
            is_all_members=name == 'all',
        )
