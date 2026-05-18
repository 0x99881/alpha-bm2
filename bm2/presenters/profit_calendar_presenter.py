from __future__ import annotations

import hashlib
from datetime import date, datetime, timedelta
from typing import Any

from ..excel.value_normalizer import normalize_expense, normalize_income, normalize_wear
from ..profit_calendar_utils import build_calendar_weeks, build_month_label, build_month_neighbors


def _cycle_color(cycle_id: str) -> str:
    """Stable HSL color from cycle id — same cycle always picks the same hue.

    Saturation/lightness are fixed so the calendar legend stays calm; the
    contrast between cycles comes from the hue distance alone.
    """
    digest = hashlib.md5((cycle_id or '').encode('utf-8')).hexdigest()
    hue = int(digest[:6], 16) % 360
    return f"hsl({hue}, 55%, 55%)"


def _build_cycle_palette(cycles: list[dict[str, Any]]) -> list[dict[str, Any]]:
    today_iso = date.today().strftime('%Y-%m-%d')
    palette: list[dict[str, Any]] = []
    for c in cycles or []:
        cid = str(c.get('id') or '')
        if not cid:
            continue
        start = str(c.get('start_date') or '').strip()
        settle = str(c.get('settle_date') or '').strip()
        settled = bool(int(c.get('settled', 0) or 0))
        end = settle if (settled and settle) else today_iso
        palette.append({
            'id': cid,
            'name': str(c.get('name') or ''),
            'settled': settled,
            'start_iso': start,
            'end_iso': end,
            'color': _cycle_color(cid),
        })
    return palette


def _cycle_for_date(date_iso: str, palette: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not date_iso:
        return None
    for entry in palette:
        if not entry['start_iso']:
            continue
        if entry['start_iso'] <= date_iso <= entry['end_iso']:
            return entry
    return None


def _iter_year_months(start_iso: str, end_iso: str):
    """Yield (year, month) tuples for every month spanned by [start, end].

    Yields nothing when either bound is empty or unparseable so callers
    can pass cycle-window dicts without pre-checking; this matters for
    the empty-cycle fallback path where ``start_date`` may be ``""``.
    """
    try:
        start = datetime.strptime(start_iso, "%Y-%m-%d").date()
        end = datetime.strptime(end_iso, "%Y-%m-%d").date()
    except ValueError:
        return
    if end < start:
        return
    cursor = start.replace(day=1)
    end_anchor = end.replace(day=1)
    while cursor <= end_anchor:
        yield cursor.year, cursor.month
        if cursor.month == 12:
            cursor = cursor.replace(year=cursor.year + 1, month=1)
        else:
            cursor = cursor.replace(month=cursor.month + 1)


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

    def build_calendar_combo(
        self,
        *,
        name: str,
        year: int,
        month: int,
        cycle_window: dict[str, Any] | None,
        all_cycles: list[dict[str, Any]],
    ) -> dict[str, Any]:
        """Single-month grid + cycle-scoped stats — the unified payload that
        replaces the old "month mode" and "cycle mode" branches.

        - The grid always shows the requested ``year/month`` (prev/next arrows
          page through months freely).
        - Each day cell carries ``cycle_id`` + ``cycle_color`` so the template
          can render a left-edge stripe colored by which cycle owns that day.
        - Stats / 盈亏榜 use ``cycle_window`` (the "current cycle" the user
          picked from the dropdown). ``cycle_window=None`` ⇒ 全部历史 — every
          day contributes and every day renders fully opaque.
        """
        active_members = self.repository.get_active_members()
        active_names = {item['name'] for item in active_members}
        stats_allowed = name == 'all' or name in active_names

        dataset = (
            self.repository.get_all_members_calendar_dataset(year, month)
            if name == 'all'
            else self.repository.get_member_calendar_dataset(name, year, month)
        )
        calendar_data = self._build_calendar_payload(dataset=dataset)
        month_records = calendar_data['records']

        palette = _build_cycle_palette(all_cycles)
        palette_by_id = {entry['id']: entry for entry in palette}
        selected_cycle_id = (cycle_window or {}).get('cycle_id', '') if cycle_window else 'all'
        view_all = (cycle_window is None) or selected_cycle_id == 'all'

        # ---- Stats / board scope: cycle window if focused, else month ----
        if view_all:
            scope_records = month_records
        else:
            start = cycle_window.get('start_date', '')
            end = cycle_window.get('end_date', '')
            scope_records = [
                r for r in month_records
                if start <= str(r.get('date') or '') <= end
            ]
        # When the user is focused on a specific cycle, the stats card should
        # aggregate THAT cycle's full window — not just the days that happen
        # to land in the currently-viewed month. So we widen the source to
        # cover every month inside the cycle window.
        if not view_all and cycle_window and cycle_window.get('has_cycle'):
            scope_records = self._collect_cycle_records(
                name=name,
                cycle_window=cycle_window,
                fallback_month_records=month_records,
                current_year=year,
                current_month=month,
            )

        scope_income, scope_wear, scope_expense = self._sum_profit_calendar_totals(scope_records, stats_allowed)
        active_member_count = self._count_profit_stats_members(name=name, stats_allowed=stats_allowed, active_members=active_members)
        data_day_count = len(scope_records)
        average_stats = self._build_profit_average_stats(
            month_income=scope_income,
            month_wear=scope_wear,
            active_member_count=active_member_count,
            data_day_count=data_day_count,
        )
        board_rows, profit_positive_list, profit_negative_list = (
            self._build_profit_board_lists(scope_records, active_members)
            if name == 'all'
            else ([], [], [])
        )

        # ---- Annotate each day cell with cycle ownership + focus state ----
        for week in calendar_data['weeks']:
            for cell in week:
                date_iso = cell.get('date') or ''
                entry = _cycle_for_date(date_iso, palette)
                cell['cycle_id'] = entry['id'] if entry else ''
                cell['cycle_color'] = entry['color'] if entry else ''
                cell['cycle_name'] = entry['name'] if entry else ''
                if view_all:
                    cell['cycle_focus'] = 'all'  # everything fully opaque
                elif entry and entry['id'] == selected_cycle_id:
                    cell['cycle_focus'] = 'in'   # bright
                elif entry:
                    cell['cycle_focus'] = 'other'  # dimmed-but-visible
                else:
                    cell['cycle_focus'] = 'none'   # darkest grey, no record

        prev_year, prev_month, next_year, next_month = build_month_neighbors(year, month)

        # window descriptor for stat-card title text
        if view_all:
            window_start, window_end = '', ''
        else:
            window_start = (cycle_window or {}).get('start_date', '')
            window_end = (cycle_window or {}).get('end_date', '')

        return {
            'mode': 'combo',
            'year': year,
            'month': month,
            'month_label': build_month_label(year, month),
            'weeks': calendar_data['weeks'],
            'month_blocks': [{
                'year': year, 'month': month,
                'month_label': build_month_label(year, month),
                'weeks': calendar_data['weeks'],
            }],
            'prev_year': prev_year, 'prev_month': prev_month,
            'next_year': next_year, 'next_month': next_month,
            'records': scope_records,
            'cycle_palette': palette,
            'cycle_id': selected_cycle_id,
            'cycles': all_cycles,
            'cycle_start_date': window_start,
            'cycle_end_date': window_end,
            'is_settled': bool((cycle_window or {}).get('is_settled')),
            'month_income_total': scope_income,
            'month_wear_total': scope_wear,
            'month_expense_total': scope_expense,
            'month_profit_total': normalize_income(scope_income - scope_wear - scope_expense),
            **average_stats,
            'profit_board_rows': board_rows,
            'profit_positive_list': profit_positive_list,
            'profit_negative_list': profit_negative_list,
            'today': datetime.now().strftime('%Y-%m-%d'),
            'is_all_members': name == 'all',
            'stats_allowed': stats_allowed,
        }

    def _collect_cycle_records(self, *, name, cycle_window, fallback_month_records, current_year, current_month):
        """Records across every month in the cycle window (so cycle stats are
        full-window, not clipped to the current viewing month)."""
        start_iso = cycle_window.get('start_date', '')
        end_iso = cycle_window.get('end_date', '')
        if not start_iso or not end_iso:
            return fallback_month_records
        out: list[dict[str, Any]] = []
        for year, month in _iter_year_months(start_iso, end_iso):
            if year == current_year and month == current_month:
                # Already have it via the visible month's dataset.
                out.extend([
                    r for r in fallback_month_records
                    if start_iso <= str(r.get('date') or '') <= end_iso
                ])
                continue
            dataset = (
                self.repository.get_all_members_calendar_dataset(year, month)
                if name == 'all'
                else self.repository.get_member_calendar_dataset(name, year, month)
            )
            for item in dataset['month_records']:
                date_text = str(item.get('date') or '')
                if start_iso <= date_text <= end_iso:
                    out.append(item)
        return out

    def build_member_profit_calendar_for_cycle(
        self,
        name: str,
        cycle_window: dict[str, Any],
        cycles: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        """Calendar payload scoped to a single cycle window.

        Composes per-month datasets across every calendar month touched by the
        cycle's [start_date, end_date] window, sums cycle-scope totals, and
        returns a payload shaped compatibly with the legacy month payload so
        the template can branch cleanly.
        """
        active_members = self.repository.get_active_members()
        active_names = {item['name'] for item in active_members}
        stats_allowed = name == 'all' or name in active_names

        start_iso = cycle_window.get('start_date') or ''
        end_iso = cycle_window.get('end_date') or datetime.now().strftime('%Y-%m-%d')

        # Collect month blocks and aggregate records across the cycle window.
        month_blocks: list[dict[str, Any]] = []
        cycle_records: list[dict[str, Any]] = []
        for year, month in _iter_year_months(start_iso, end_iso):
            dataset = (
                self.repository.get_all_members_calendar_dataset(year, month)
                if name == 'all'
                else self.repository.get_member_calendar_dataset(name, year, month)
            )
            weeks = build_calendar_weeks(
                year=year,
                month=month,
                record_map=dataset['record_map'],
                max_abs_wear=dataset['max_abs_wear'],
                note_text=dataset.get('note_text', ''),
            )
            month_blocks.append({
                'year': year,
                'month': month,
                'month_label': build_month_label(year, month),
                'weeks': weeks,
            })
            for item in dataset['month_records']:
                item_date = str(item.get('date') or '')
                if start_iso <= item_date <= end_iso:
                    cycle_records.append(item)

        # Cycle-scope totals identical in shape to month totals so the template
        # can keep the same card structure.
        cycle_income, cycle_wear, cycle_expense = self._sum_profit_calendar_totals(cycle_records, stats_allowed)
        active_member_count = self._count_profit_stats_members(
            name=name, stats_allowed=stats_allowed, active_members=active_members,
        )
        data_day_count = len(cycle_records)
        average_stats = self._build_profit_average_stats(
            month_income=cycle_income,
            month_wear=cycle_wear,
            active_member_count=active_member_count,
            data_day_count=data_day_count,
        )
        board_rows, profit_positive_list, profit_negative_list = (
            self._build_profit_board_lists(cycle_records, active_members)
            if name == 'all'
            else ([], [], [])
        )

        # Anchor year/month is the first month in the span — kept for backward
        # compatibility with parts of the template that still reference them
        # (e.g. modal date prefill).
        anchor_year = month_blocks[0]['year'] if month_blocks else datetime.now().year
        anchor_month = month_blocks[0]['month'] if month_blocks else datetime.now().month

        return {
            'mode': 'cycle',
            'year': anchor_year,
            'month': anchor_month,
            'month_label': month_blocks[0]['month_label'] if month_blocks else '',
            'month_blocks': month_blocks,
            'weeks': month_blocks[0]['weeks'] if month_blocks else [],
            'records': cycle_records,
            'cycle_start_date': start_iso,
            'cycle_end_date': end_iso,
            'cycle_id': cycle_window.get('cycle_id', ''),
            'is_settled': bool(cycle_window.get('is_settled')),
            'cycles': list(cycles or []),
            'month_income_total': cycle_income,
            'month_wear_total': cycle_wear,
            'month_expense_total': cycle_expense,
            'month_profit_total': normalize_income(cycle_income - cycle_wear - cycle_expense),
            **average_stats,
            'profit_board_rows': board_rows,
            'profit_positive_list': profit_positive_list,
            'profit_negative_list': profit_negative_list,
            'today': datetime.now().strftime('%Y-%m-%d'),
            'is_all_members': name == 'all',
            'stats_allowed': stats_allowed,
        }

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
        # Legacy month mode also exposes a single-element ``month_blocks`` list
        # so the template can render with one unified loop regardless of mode.
        month_blocks = [{
            'year': calendar_data['year'],
            'month': calendar_data['month'],
            'month_label': calendar_data['month_label'],
            'weeks': calendar_data['weeks'],
        }]
        return {
            'mode': 'month',
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
            'cycles': [],
            'cycle_id': '',
            'month_blocks': month_blocks,
            'cycle_start_date': '',
            'cycle_end_date': '',
        }
