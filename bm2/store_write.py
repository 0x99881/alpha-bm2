from __future__ import annotations

from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any

from .constants import (
    EXPENSE_META_SHEET,
    EXPENSE_NAME_HEADER,
    INCOME_META_SHEET,
    INCOME_NAME_HEADER,
    NAME_HEADER,
    TOTAL_HEADER,
    WEAR_META_SHEET,
    WEAR_NAME_HEADER,
    WEAR_TOTAL_HEADER,
    WINDOW_SIZE,
)
from .ui_text import MESSAGES


class StoreWriteMixin:
    def _prepare_score_target(self, score_sheet, effective_date_text: str) -> dict[str, Any]:
        total_col = self._find_column(score_sheet, TOTAL_HEADER)
        name_col = self._find_column(score_sheet, NAME_HEADER)
        if total_col is None or name_col is None:
            raise ValueError(MESSAGES['score_sheet_invalid'])

        d_cols = self._d_columns(score_sheet)
        next_number = d_cols[-1][0] + 1 if d_cols else 1
        score_sheet.insert_cols(total_col, 1)
        target_col = total_col
        score_header = effective_date_text[5:]
        score_sheet.cell(1, target_col, score_header)
        return {
            'sheet': score_sheet,
            'target_col': target_col,
            'total_col': total_col + 1,
            'profit_col': total_col + 2,
            'name_col': total_col + 3,
            'header': score_header,
            'next_number': next_number,
        }

    def _prepare_value_target(self, sheet, *, total_header: str | None, name_header: str, header_value: str, invalid_message: str) -> dict[str, Any]:
        total_col = self._find_column(sheet, total_header) if total_header else None
        name_col = self._find_column(sheet, name_header)
        if name_col is None or (total_header and total_col is None):
            raise ValueError(invalid_message)

        insert_col = total_col if total_col is not None else name_col
        sheet.insert_cols(insert_col, 1)
        sheet.cell(1, insert_col, header_value)
        return {
            'sheet': sheet,
            'target_col': insert_col,
            'total_col': insert_col + 1 if total_col is not None else None,
            'name_col': insert_col + 2 if total_col is not None else insert_col + 1,
            'header': header_value,
        }

    def _prepare_daily_targets(self, workbook, effective_date_text: str) -> dict[str, dict[str, Any]]:
        day_code = effective_date_text[5:].replace('-', '')
        score_sheet = self._score_sheet(workbook)
        wear_sheet = self._wear_sheet(workbook)
        income_sheet = self._income_sheet(workbook)
        expense_sheet = self._expense_sheet(workbook)
        return {
            'score': self._prepare_score_target(score_sheet, effective_date_text),
            'wear': self._prepare_value_target(
                wear_sheet,
                total_header=WEAR_TOTAL_HEADER,
                name_header=WEAR_NAME_HEADER,
                header_value=day_code,
                invalid_message=MESSAGES['wear_sheet_invalid'],
            ),
            'income': self._prepare_value_target(
                income_sheet,
                total_header=None,
                name_header=INCOME_NAME_HEADER,
                header_value=day_code,
                invalid_message=MESSAGES['income_sheet_invalid'],
            ),
            'expense': self._prepare_value_target(
                expense_sheet,
                total_header=None,
                name_header=EXPENSE_NAME_HEADER,
                header_value=day_code,
                invalid_message=MESSAGES['expense_sheet_invalid'],
            ),
        }

    def _prepare_daily_row_maps(self, targets: dict[str, dict[str, Any]]) -> dict[str, dict[str, int]]:
        return {
            'score': self._ensure_member_rows(
                targets['score']['sheet'],
                name_col=targets['score']['name_col'],
                total_col=targets['score']['total_col'],
                value_columns=[col for _, col in self._d_columns(targets['score']['sheet'])],
                extra_columns=[targets['score']['profit_col']],
            )[0],
            'wear': self._ensure_member_rows(
                targets['wear']['sheet'],
                name_col=targets['wear']['name_col'],
                total_col=targets['wear']['total_col'],
                value_columns=[col for _, col in self._wear_columns(targets['wear']['sheet'])],
            )[0],
            'income': self._ensure_member_rows(
                targets['income']['sheet'],
                name_col=targets['income']['name_col'],
                total_col=None,
                value_columns=[col for _, col in self._income_columns(targets['income']['sheet'])],
            )[0],
            'expense': self._ensure_member_rows(
                targets['expense']['sheet'],
                name_col=targets['expense']['name_col'],
                total_col=None,
                value_columns=[col for _, col in self._expense_columns(targets['expense']['sheet'])],
            )[0],
        }

    def _finalize_daily_save(self, workbook, *, effective_date_text: str, recent_numbers: list[int], targets: dict[str, dict[str, Any]]) -> None:
        self._recalculate_totals(targets['score']['sheet'], targets['score']['total_col'])
        self._recalculate_score_profits(workbook, targets['score']['sheet'], targets['score']['profit_col'])
        self._sort_named_rows(targets['score']['sheet'], targets['score']['total_col'], targets['score']['name_col'])
        self._format_score_sheet(targets['score']['sheet'], recent_numbers, targets['score']['total_col'])
        self._recalculate_wear_totals(targets['wear']['sheet'], targets['wear']['total_col'])
        self._sort_named_rows(targets['wear']['sheet'], targets['wear']['total_col'], targets['wear']['name_col'])
        self._sort_rows_by_name(targets['income']['sheet'], targets['income']['name_col'])
        self._sort_rows_by_name(targets['expense']['sheet'], targets['expense']['name_col'])
        self._append_meta(workbook, effective_date_text, f"D{targets['score']['next_number']}")
        self._append_wear_meta(workbook, effective_date_text, targets['wear']['header'])
        self._append_income_meta(workbook, effective_date_text, targets['income']['header'])
        self._append_expense_meta(workbook, effective_date_text, targets['expense']['header'])
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

    def _write_entry_score_income_expense(self, payload: dict[str, str], *, targets: dict[str, dict[str, Any]], row_maps: dict[str, dict[str, int]]) -> None:
        name = payload['name']
        score_value = int(payload['score_text']) if payload['score_text'] else 0
        targets['score']['sheet'].cell(row_maps['score'][name], targets['score']['target_col'], score_value)
        targets['wear']['sheet'].cell(row_maps['wear'][name], targets['wear']['target_col'], 0)

        income_value = self._round_income(
            self._parse_decimal(payload['income_text'], MESSAGES['income_field'].format(name=name))
            if payload['income_text'] else 0
        )
        expense_value = self._round_expense(
            self._parse_decimal(payload['other_expense_text'], MESSAGES['expense_field'].format(name=name))
            if payload['other_expense_text'] else 0
        )
        targets['income']['sheet'].cell(row_maps['income'][name], targets['income']['target_col'], income_value)
        targets['expense']['sheet'].cell(row_maps['expense'][name], targets['expense']['target_col'], expense_value)

    def _resolve_entry_wear_value(self, payload: dict[str, str]) -> Decimal | None:
        name = payload['name']
        has_balance_input = bool(payload['before_text'] or payload['after_text'])
        has_manual_input = bool(payload['manual_wear_text'])
        if not (has_balance_input or has_manual_input):
            return None

        if has_manual_input:
            return self._parse_decimal(payload['manual_wear_text'], MESSAGES['manual_wear_field'].format(name=name))

        if not payload['before_text'] or not payload['after_text']:
            raise ValueError(MESSAGES['balance_pair_required'].format(name=name))
        before_value = self._parse_decimal(payload['before_text'], MESSAGES['before_balance_field'].format(name=name))
        after_value = self._parse_decimal(payload['after_text'], MESSAGES['after_balance_field'].format(name=name))
        return before_value - after_value

    def _write_entry_wear(self, name: str, wear_value: Decimal | None, *, targets: dict[str, dict[str, Any]], row_maps: dict[str, dict[str, int]]) -> bool:
        if wear_value is None:
            return False
        targets['wear']['sheet'].cell(row_maps['wear'][name], targets['wear']['target_col'], self._round_wear(wear_value))
        return True

    def _parse_decimal(self, value: str, field_name: str) -> Decimal:
        try:
            return Decimal(value)
        except InvalidOperation as exc:
            raise ValueError(MESSAGES['must_be_number'].format(field_name=field_name)) from exc

    def _round_wear(self, value: Decimal | float | int) -> float:
        return round(float(value), 1)

    def _round_income(self, value: Decimal | float | int) -> float:
        return round(float(value), 1)

    def _round_expense(self, value: Decimal | float | int) -> float:
        return round(float(value), 1)

    def _sheet_has_day_code(self, sheet, columns_getter, day_code: str) -> bool:
        for number, _ in columns_getter(sheet):
            if f'{number:04d}' == day_code:
                return True
        return False

    def _resolve_unique_daily_date(self, workbook, date_text: str) -> str:
        current = datetime.strptime(date_text, '%Y-%m-%d')
        wear_sheet = self._wear_sheet(workbook)
        income_sheet = self._income_sheet(workbook)
        expense_sheet = self._expense_sheet(workbook)
        for _ in range(400):
            day_code = current.strftime('%m%d')
            exists = (
                self._sheet_has_day_code(wear_sheet, self._wear_columns, day_code)
                or self._sheet_has_day_code(income_sheet, self._income_columns, day_code)
                or self._sheet_has_day_code(expense_sheet, self._expense_columns, day_code)
            )
            if not exists:
                return current.strftime('%Y-%m-%d')
            current += timedelta(days=1)
        raise ValueError(f'\u65e0\u6cd5\u5728 {date_text} \u4e4b\u540e 400 \u5929\u5185\u627e\u5230\u53ef\u7528\u65e5\u671f')

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

            recent_numbers = [number for number, _ in self._d_columns(targets['score']['sheet'])][-WINDOW_SIZE:]
            self._finalize_daily_save(workbook, effective_date_text=effective_date_text, recent_numbers=recent_numbers, targets=targets)
            self._save_workbook(workbook)
            return {
                'target_column': targets['score']['header'],
                'wear_column': targets['wear']['header'],
                'wear_rows_added': wear_rows_added,
                'window_size': min(WINDOW_SIZE, len(recent_numbers)),
                'saved_date': effective_date_text,
            }
        finally:
            workbook.close()
