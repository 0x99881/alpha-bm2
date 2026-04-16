from __future__ import annotations

from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any

from .constants import (
    EXPENSE_META_SHEET,
    INCOME_META_SHEET,
    META_SHEET,
    NAME_HEADER,
    TOTAL_HEADER,
    VALUE_SHEET_SPECS,
    WEAR_META_SHEET,
    WINDOW_SIZE,
)
from .ui_text import MESSAGES
from .value_utils import round_expense, round_income, round_value_sheet_number, round_wear


class StoreWriteMixin:
    def _value_sheet_spec(self, sheet_type: str) -> dict[str, Any]:
        return VALUE_SHEET_SPECS[sheet_type]

    def _value_sheet_write_helpers(self, sheet_type: str):
        spec = self._value_sheet_spec(sheet_type)
        return (
            spec,
            getattr(self, spec['sheet_method']),
            getattr(self, spec['columns_method']),
            getattr(self, spec['round_method']),
        )

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
        targets = {'score': self._prepare_score_target(score_sheet, effective_date_text)}
        for sheet_type in ('wear', 'income', 'expense'):
            spec, sheet_getter, _, _ = self._value_sheet_write_helpers(sheet_type)
            targets[sheet_type] = self._prepare_value_target(
                sheet_getter(workbook),
                total_header=spec['total_header'],
                name_header=spec['name_header'],
                header_value=day_code,
                invalid_message=MESSAGES[spec['invalid_message_key']],
            )
        return targets

    def _prepare_daily_row_maps(self, targets: dict[str, dict[str, Any]]) -> dict[str, dict[str, int]]:
        row_maps = {
            'score': self._ensure_member_rows(
                targets['score']['sheet'],
                name_col=targets['score']['name_col'],
                total_col=targets['score']['total_col'],
                value_columns=[col for _, col in self._d_columns(targets['score']['sheet'])],
                extra_columns=[targets['score']['profit_col']],
            )[0],
        }
        for sheet_type in ('wear', 'income', 'expense'):
            spec, _, columns_getter, _ = self._value_sheet_write_helpers(sheet_type)
            row_maps[sheet_type] = self._ensure_member_rows(
                targets[sheet_type]['sheet'],
                name_col=targets[sheet_type]['name_col'],
                total_col=targets[sheet_type]['total_col'],
                value_columns=[col for _, col in columns_getter(targets[sheet_type]['sheet'])],
            )[0]
        return row_maps

    def _finalize_daily_save(self, workbook, *, effective_date_text: str, recent_numbers: list[int], targets: dict[str, dict[str, Any]]) -> None:
        self._recalculate_totals(targets['score']['sheet'], targets['score']['total_col'])
        self._recalculate_score_profits(workbook, targets['score']['sheet'], targets['score']['profit_col'])
        self._sort_named_rows(targets['score']['sheet'], targets['score']['total_col'], targets['score']['name_col'])
        self._format_score_sheet(targets['score']['sheet'], recent_numbers, targets['score']['total_col'])
        self._recalculate_wear_totals(targets['wear']['sheet'], targets['wear']['total_col'])
        self._sort_named_rows(targets['wear']['sheet'], targets['wear']['total_col'], targets['wear']['name_col'])
        for sheet_type in ('income', 'expense'):
            self._sort_rows_by_name(targets[sheet_type]['sheet'], targets[sheet_type]['name_col'])
        self._append_meta(workbook, effective_date_text, f"D{targets['score']['next_number']}")
        for sheet_type in ('wear', 'income', 'expense'):
            self._append_value_meta(workbook, sheet_type, effective_date_text, targets[sheet_type]['header'])
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

    def _round_value_sheet_number(self, value: Decimal | float | int) -> float:
        return round_value_sheet_number(value)

    def _round_wear(self, value: Decimal | float | int) -> float:
        return round_wear(value)

    def _round_income(self, value: Decimal | float | int) -> float:
        return round_income(value)

    def _round_expense(self, value: Decimal | float | int) -> float:
        return round_expense(value)

    def _sheet_has_day_code(self, sheet, columns_getter, day_code: str) -> bool:
        for number, _ in columns_getter(sheet):
            if f'{number:04d}' == day_code:
                return True
        return False

    def _score_sheet_has_date(self, sheet, mmdd_text: str) -> bool:
        total_col = self._find_column(sheet, TOTAL_HEADER)
        if total_col is None:
            return False
        for col in range(1, total_col):
            if str(sheet.cell(1, col).value or '').strip() == mmdd_text:
                return True
        return False

    def _parse_saved_date(self, date_text: str):
        return datetime.strptime(str(date_text).strip(), '%Y-%m-%d').date()

    def _latest_header_date(self, workbook):
        latest_meta_dates = [
            self._parse_saved_date(saved_date)
            for sheet_name in (META_SHEET, WEAR_META_SHEET, INCOME_META_SHEET, EXPENSE_META_SHEET)
            for saved_date, _ in self._read_sheet_meta(workbook, sheet_name)
        ]
        year_hint = max(latest_meta_dates).year if latest_meta_dates else datetime.now().year
        latest_dates = list(latest_meta_dates)

        score_sheet = self._score_sheet(workbook)
        score_columns = self._d_columns(score_sheet)
        if score_columns:
            latest_score_header = str(score_sheet.cell(1, score_columns[-1][1]).value or '').strip()
            if self._is_mmdd_header(latest_score_header):
                latest_dates.append(datetime.strptime(f'{year_hint}-{latest_score_header}', '%Y-%m-%d').date())

        for sheet_type in ('wear', 'income', 'expense'):
            _, sheet_getter, columns_getter, _ = self._value_sheet_write_helpers(sheet_type)
            sheet = sheet_getter(workbook)
            columns = columns_getter(sheet)
            if not columns:
                continue
            latest_header = f'{columns[-1][0]:04d}'
            latest_dates.append(datetime.strptime(f'{year_hint}-{latest_header[:2]}-{latest_header[2:]}', '%Y-%m-%d').date())

        return max(latest_dates) if latest_dates else None

    def get_next_score_date(self) -> str:
        workbook = self._open_workbook()
        try:
            latest_date = self._latest_header_date(workbook)
            if latest_date is None:
                return datetime.now().strftime('%Y-%m-%d')
            return (latest_date + timedelta(days=1)).strftime('%Y-%m-%d')
        finally:
            workbook.close()

    def _resolve_unique_daily_date(self, workbook, date_text: str) -> str:
        current = self._parse_saved_date(date_text)
        score_sheet = self._score_sheet(workbook)
        for _ in range(400):
            score_header = current.strftime('%m-%d')
            day_code = current.strftime('%m%d')
            exists = self._score_sheet_has_date(score_sheet, score_header)
            for sheet_type in ('wear', 'income', 'expense'):
                _, sheet_getter, columns_getter, _ = self._value_sheet_write_helpers(sheet_type)
                if self._sheet_has_day_code(sheet_getter(workbook), columns_getter, day_code):
                    exists = True
                    break
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
