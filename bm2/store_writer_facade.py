from __future__ import annotations

# COMPATIBILITY LAYER

from datetime import datetime, timedelta
from decimal import Decimal
from typing import Any

from .constants import META_SHEET, WINDOW_SIZE, UNIQUE_DATE_SEARCH_LIMIT
from .domain.rules.daily_entry import IncompleteBalanceInput, resolve_wear_value
from .excel.daily_target import DailyTarget
from .excel.header_locator import find_column
from .excel.member_rows import ensure_member_rows_with_map, sort_named_rows, sort_rows_by_name
from .excel.sheet_metadata import append_sheet_meta
from .excel.value_normalizer import normalize_expense, normalize_income, normalize_wear, parse_decimal
from .excel.value_sheet_spec import get_value_sheet_spec
from .ui_text import MESSAGES


class StoreWriterFacade:
    def __init__(self, store) -> None:
        self._store = store

    def __getattr__(self, name):
        return getattr(self._store, name)

    def _daily_target_class(self):
        return DailyTarget

    def _prepare_score_target(self, score_sheet, effective_date_text: str) -> DailyTarget:
        return self.score_sheet.prepare_target(score_sheet, effective_date_text)

    def _prepare_value_target(self, sheet, *, total_header: str | None, name_header: str, header_value: str, invalid_message: str) -> DailyTarget:
        total_col = find_column(sheet, total_header) if total_header else None
        name_col = find_column(sheet, name_header)
        if name_col is None or (total_header and total_col is None):
            raise ValueError(invalid_message)

        insert_col = total_col if total_col is not None else name_col
        sheet.insert_cols(insert_col, 1)
        sheet.cell(1, insert_col, header_value)
        return DailyTarget(
            sheet=sheet,
            target_col=insert_col,
            total_col=insert_col + 1 if total_col is not None else None,
            name_col=insert_col + 2 if total_col is not None else insert_col + 1,
            header=header_value,
        )

    def _prepare_daily_targets(self, workbook, effective_date_text: str) -> dict[str, DailyTarget]:
        day_code = effective_date_text[5:].replace('-', '')
        score_sheet = self._score_sheet(workbook)
        targets = {'score': self._prepare_score_target(score_sheet, effective_date_text)}
        for sheet_type in ('wear', 'income', 'expense'):
            helpers = self._value_sheet_helpers(sheet_type)
            targets[sheet_type] = self._prepare_value_target(
                helpers.sheet_getter(workbook),
                total_header=helpers.spec['total_header'],
                name_header=helpers.spec['name_header'],
                header_value=day_code,
                invalid_message=MESSAGES[helpers.spec['invalid_message_key']],
            )
        return targets

    def _prepare_daily_row_maps(self, targets: dict[str, DailyTarget]) -> dict[str, dict[str, int]]:
        score_target = targets['score']
        row_maps = {
            'score': ensure_member_rows_with_map(
                score_target.sheet,
                members=self.get_members(),
                name_col=score_target.name_col,
                total_col=score_target.total_col,
                value_columns=[col for _, col in self._score_date_columns(score_target.sheet)],
                extra_columns=[score_target.profit_col] if score_target.profit_col is not None else [],
            )[0],
        }
        for sheet_type in ('wear', 'income', 'expense'):
            helpers = self._value_sheet_helpers(sheet_type)
            target = targets[sheet_type]
            row_maps[sheet_type] = ensure_member_rows_with_map(
                target.sheet,
                members=self.get_members(),
                name_col=target.name_col,
                total_col=target.total_col,
                value_columns=[col for _, col in helpers.columns_getter(target.sheet)],
            )[0]
        return row_maps

    def _finalize_daily_save(self, workbook, *, effective_date_text: str, recent_numbers: list[int], targets: dict[str, DailyTarget]) -> None:
        score_target = targets['score']
        wear_target = targets['wear']
        self._recalculate_totals(score_target.sheet, score_target.total_col)
        self._recalculate_score_profits(workbook, score_target.sheet, score_target.profit_col)
        sort_named_rows(score_target.sheet, score_target.total_col, score_target.name_col)
        self._format_score_sheet(score_target.sheet, recent_numbers, score_target.total_col)
        self._recalculate_wear_totals(wear_target.sheet, wear_target.total_col)
        sort_named_rows(wear_target.sheet, wear_target.total_col, wear_target.name_col)
        for sheet_type in ('income', 'expense'):
            target = targets[sheet_type]
            sort_rows_by_name(target.sheet, target.name_col)
        append_sheet_meta(workbook, META_SHEET, effective_date_text, f"D{score_target.next_number}")
        for sheet_type in ('wear', 'income', 'expense'):
            append_sheet_meta(workbook, get_value_sheet_spec(sheet_type)['meta_sheet'], effective_date_text, targets[sheet_type].header)
        self._sync_member_visibility_in_workbook(workbook)

    def _normalize_entry_payload(self, entry: dict[str, str]) -> dict[str, str]:
        return {
            'name': entry['name'],
            'score_text': entry.get('score', '').strip(),
            'before_text': entry.get('before_balance', '').strip(),
            'after_text': entry.get('after_balance', '').strip(),
            'manual_wear_text': entry.get('manual_wear', '').strip(),
            'income_text': entry.get('income', '').strip(),
            'other_expense_text': entry.get('other_expense', '').strip(),
        }

    def _write_entry_score_income_expense(self, payload: dict[str, str], *, targets: dict[str, DailyTarget], row_maps: dict[str, dict[str, int]]) -> None:
        name = payload['name']
        score_value = int(payload['score_text']) if payload['score_text'] else 0
        targets['score'].sheet.cell(row_maps['score'][name], targets['score'].target_col, score_value)
        targets['wear'].sheet.cell(row_maps['wear'][name], targets['wear'].target_col, 0)

        income_value = normalize_income(
            parse_decimal(payload['income_text'], MESSAGES['income_field'].format(name=name))
            if payload['income_text'] else 0
        )
        expense_value = normalize_expense(
            parse_decimal(payload['other_expense_text'], MESSAGES['expense_field'].format(name=name))
            if payload['other_expense_text'] else 0
        )
        targets['income'].sheet.cell(row_maps['income'][name], targets['income'].target_col, income_value)
        targets['expense'].sheet.cell(row_maps['expense'][name], targets['expense'].target_col, expense_value)

    def _resolve_entry_wear_value(self, payload: dict[str, str]) -> Decimal | None:
        name = payload['name']
        try:
            return resolve_wear_value(
                before_text=payload['before_text'],
                after_text=payload['after_text'],
                manual_wear_text=payload['manual_wear_text'],
                parse_decimal=parse_decimal,
                manual_field=MESSAGES['manual_wear_field'].format(name=name),
                before_field=MESSAGES['before_balance_field'].format(name=name),
                after_field=MESSAGES['after_balance_field'].format(name=name),
            )
        except IncompleteBalanceInput as exc:
            raise ValueError(MESSAGES['balance_pair_required'].format(name=name)) from exc

    def _write_entry_wear(self, name: str, wear_value: Decimal | None, *, targets: dict[str, DailyTarget], row_maps: dict[str, dict[str, int]]) -> bool:
        if wear_value is None:
            return False
        targets['wear'].sheet.cell(row_maps['wear'][name], targets['wear'].target_col, normalize_wear(wear_value))
        return True

    def _parse_decimal(self, value: str, field_name: str) -> Decimal:
        return parse_decimal(value, field_name)

    def _round_wear(self, value: Decimal | float | int) -> float:
        return normalize_wear(value)

    def _round_income(self, value: Decimal | float | int) -> float:
        return normalize_income(value)

    def _round_expense(self, value: Decimal | float | int) -> float:
        return normalize_expense(value)

    def _sheet_has_day_code(self, sheet, columns_getter, day_code: str) -> bool:
        for number, _ in columns_getter(sheet):
            if f'{number:04d}' == day_code:
                return True
        return False

    def _score_sheet_has_date(self, sheet, mmdd_text: str) -> bool:
        return self.score_sheet.has_date(sheet, mmdd_text)

    def _parse_saved_date(self, date_text: str):
        return datetime.strptime(str(date_text).strip(), '%Y-%m-%d').date()

    def _resolve_unique_daily_date(self, workbook, date_text: str) -> str:
        current = self._parse_saved_date(date_text)
        score_sheet = self._score_sheet(workbook)
        for _ in range(UNIQUE_DATE_SEARCH_LIMIT):
            score_header = current.strftime('%m-%d')
            day_code = current.strftime('%m%d')
            exists = self._score_sheet_has_date(score_sheet, score_header)
            for sheet_type in ('wear', 'income', 'expense'):
                helpers = self._value_sheet_helpers(sheet_type)
                if self._sheet_has_day_code(helpers.sheet_getter(workbook), helpers.columns_getter, day_code):
                    exists = True
                    break
            if not exists:
                return current.strftime('%Y-%m-%d')
            current += timedelta(days=1)
        raise ValueError(f'无法在 {date_text} 之后 {UNIQUE_DATE_SEARCH_LIMIT} 天内找到可用日期')

    def save_scores_and_wear(self, date_text: str, entries: list[dict[str, str]]) -> dict[str, Any]:
        workbook = self._open_workbook()
        try:
            self._ensure_score_sheet_structure(workbook)
            self._ensure_wear_sheet_structure(workbook)
            self._ensure_income_sheet_structure(workbook)
            self._ensure_expense_sheet_structure(workbook)
            effective_date_text = self._resolve_unique_daily_date(workbook, date_text)
            targets = self._prepare_daily_targets(workbook, effective_date_text)
            row_maps = self._prepare_daily_row_maps(targets)

            wear_rows_added = 0
            for entry in entries:
                payload = self._normalize_entry_payload(entry)
                self._write_entry_score_income_expense(payload, targets=targets, row_maps=row_maps)
                wear_value = self._resolve_entry_wear_value(payload)
                wear_rows_added += int(self._write_entry_wear(payload['name'], wear_value, targets=targets, row_maps=row_maps))

            recent_numbers = [number for number, _ in self._score_date_columns(targets['score'].sheet)][-WINDOW_SIZE:]
            self._finalize_daily_save(workbook, effective_date_text=effective_date_text, recent_numbers=recent_numbers, targets=targets)
            self._save_workbook(workbook)
            return {
                'target_column': targets['score'].header,
                'wear_column': targets['wear'].header,
                'wear_rows_added': wear_rows_added,
                'window_size': min(WINDOW_SIZE, len(recent_numbers)),
                'saved_date': effective_date_text,
            }
        finally:
            workbook.close()
