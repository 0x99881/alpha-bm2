from __future__ import annotations

from datetime import datetime
from typing import Any

from ..excel.value_normalizer import normalize_expense, normalize_income, normalize_wear
from ..profit_calendar_utils import build_calendar_weeks, build_month_label, build_month_neighbors


class ProfitCalendarPresenter:
    def __init__(self, repository) -> None:
        self.repository = repository

    def _build_calendar_payload(self, *, dataset: dict[str, Any]) -> dict[str, Any]:
        year = dataset['year']
        month = dataset['month']
        weeks = build_calendar_weeks(
            year=year,
            month=month,
            record_map=dataset['record_map'],
            max_abs_wear=dataset['max_abs_wear'],
            note_text=dataset.get('note_text', ''),
        )
        prev_year, prev_month, next_year, next_month = build_month_neighbors(year, month)
        return {
            'year': year,
            'month': month,
            'month_label': build_month_label(year, month),
            'weeks': weeks,
            'records': dataset['month_records'],
            'prev_year': prev_year,
            'prev_month': prev_month,
            'next_year': next_year,
            'next_month': next_month,
        }

    def _sum_profit_calendar_totals(self, month_records: list[dict[str, Any]], stats_allowed: bool) -> tuple[float, float, float]:
        if not stats_allowed:
            return 0.0, 0.0, 0.0
        month_income = normalize_income(sum(float(item.get('income', 0) or 0) for item in month_records))
        month_wear = normalize_wear(sum(float(item.get('wear', 0) or 0) for item in month_records))
        month_expense = normalize_expense(sum(float(item.get('expense', 0) or 0) for item in month_records))
        return month_income, month_wear, month_expense

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
        avg_income = normalize_income(month_income / active_member_count) if active_member_count else 0.0
        avg_wear = normalize_wear(month_wear / active_member_count) if active_member_count else 0.0
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

    def _accumulate_profit_board_totals(self, month_records: list[dict[str, Any]], active_members: list[dict[str, Any]]) -> dict[str, dict[str, float | str]]:
        row_map = {item['name']: {'name': item['name'], 'income': 0.0, 'wear': 0.0, 'expense': 0.0, 'profit': 0.0} for item in active_members}
        for item in month_records:
            for row in item.get('breakdown', []):
                member_name = row.get('name')
                if member_name not in row_map:
                    continue
                row_map[member_name]['income'] += float(row.get('income', 0) or 0)
                row_map[member_name]['wear'] += float(row.get('wear', 0) or 0)
                row_map[member_name]['expense'] += float(row.get('expense', 0) or 0)
        return row_map

    def _build_profit_board_rows(self, row_map: dict[str, dict[str, float | str]]) -> list[dict[str, Any]]:
        board_rows = []
        for row in row_map.values():
            income_value = round(float(row['income']), 1)
            wear_value = round(float(row['wear']), 1)
            expense_value = round(float(row['expense']), 1)
            profit_value = round(income_value - wear_value - expense_value, 1)
            board_rows.append({
                'name': str(row['name']),
                'income': income_value,
                'wear': wear_value,
                'expense': expense_value,
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

    def build_member_profit_calendar(self, name: str, year: int, month: int) -> dict[str, Any]:
        active_members = self.repository.get_active_members()
        active_names = {item['name'] for item in active_members}
        stats_allowed = name == 'all' or name in active_names
        dataset = self.repository.get_all_members_calendar_dataset(year, month) if name == 'all' else self.repository.get_member_calendar_dataset(name, year, month)
        calendar_data = self._build_calendar_payload(dataset=dataset)
        month_records = calendar_data['records']
        month_income, month_wear, month_expense = self._sum_profit_calendar_totals(month_records, stats_allowed)
        active_member_count = self._count_profit_stats_members(name=name, stats_allowed=stats_allowed, active_members=active_members)
        data_day_count = len(month_records)
        average_stats = self._build_profit_average_stats(
            month_income=month_income,
            month_wear=month_wear,
            active_member_count=active_member_count,
            data_day_count=data_day_count,
        )
        board_rows, profit_positive_list, profit_negative_list = self._build_profit_board_lists(month_records, active_members) if name == 'all' else ([], [], [])
        return {
            **calendar_data,
            'month_income_total': month_income,
            'month_wear_total': month_wear,
            'month_expense_total': month_expense,
            'month_profit_total': normalize_income(month_income - month_wear - month_expense),
            **average_stats,
            'profit_board_rows': board_rows,
            'profit_positive_list': profit_positive_list,
            'profit_negative_list': profit_negative_list,
            'today': datetime.now().strftime('%Y-%m-%d'),
            'is_all_members': name == 'all',
            'stats_allowed': stats_allowed,
        }
