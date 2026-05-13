from __future__ import annotations

from typing import Any

from openpyxl.styles import Font

from .header_locator import find_column, find_sheet_by_alias, numeric_header_columns
from .member_rows import ensure_member_rows, sort_named_rows
from .summary_columns import ensure_summary_columns
from .value_normalizer import normalize_wear
from ..constants import (
    DATA_START_ROW,
    WEAR_ABNORMAL_FONT_COLOR,
    WEAR_NAME_HEADER,
    WEAR_SHEET,
    WEAR_SHEET_ALIASES,
    WEAR_TOTAL_HEADER,
)


class WearSheet:
    def __init__(self, store) -> None:
        self.store = store

    def sheet(self, workbook):
        sheet = find_sheet_by_alias(workbook, WEAR_SHEET_ALIASES)
        if sheet is not None:
            if sheet.title != WEAR_SHEET:
                sheet.title = WEAR_SHEET
            return sheet
        return workbook.create_sheet(title=WEAR_SHEET)

    def columns(self, sheet) -> list[tuple[int, int]]:
        return numeric_header_columns(sheet)

    def recalculate_totals(self, sheet, total_col: int) -> None:
        name_col = find_column(sheet, WEAR_NAME_HEADER)
        if name_col is None:
            return
        for row in range(DATA_START_ROW, sheet.max_row + 1):
            if not sheet.cell(row, name_col).value:
                continue
            total = 0.0
            for _, col in self.columns(sheet):
                value = sheet.cell(row, col).value
                numeric = normalize_wear(value or 0)
                sheet.cell(row, col, numeric)
                total += numeric
            sheet.cell(row, total_col, normalize_wear(total))

    def format_sheet(self, sheet, total_col: int) -> None:
        threshold = self.store.get_wear_abnormal_threshold()
        normal_font = Font(color='000000', bold=False)
        abnormal_font = Font(color=WEAR_ABNORMAL_FONT_COLOR, bold=True)
        for _, wear_col in self.columns(sheet):
            for row in range(DATA_START_ROW, sheet.max_row + 1):
                cell = sheet.cell(row, wear_col)
                value = cell.value
                if value in (None, ''):
                    cell.font = normal_font
                    continue
                try:
                    numeric_value = float(value)
                except (TypeError, ValueError):
                    cell.font = normal_font
                    continue
                cell.font = abnormal_font if numeric_value > threshold else normal_font

        for row in range(DATA_START_ROW, sheet.max_row + 1):
            sheet.cell(row, total_col).font = normal_font

    def ensure_structure(self, workbook) -> bool:
        sheet = self.sheet(workbook)
        changed = False

        total_col, name_col, tail_changed = ensure_summary_columns(sheet, WEAR_TOTAL_HEADER, WEAR_NAME_HEADER)
        changed = changed or tail_changed

        changed = ensure_member_rows(
            sheet,
            members=self.store.get_members(),
            name_col=name_col,
            total_col=total_col,
            value_columns=[col for _, col in self.columns(sheet)],
        ) or changed

        self.recalculate_totals(sheet, total_col)
        sort_named_rows(sheet, total_col, name_col)
        self.format_sheet(sheet, total_col)
        return changed

    def snapshot(self, workbook) -> dict[str, Any]:
        self.ensure_structure(workbook)
        sheet = self.sheet(workbook)
        headers = [sheet.cell(1, col).value for col in range(1, sheet.max_column + 1)]
        raw_rows = []
        name_col = find_column(sheet, WEAR_NAME_HEADER)
        for row in range(DATA_START_ROW, sheet.max_row + 1):
            if bool(sheet.row_dimensions[row].hidden):
                continue
            values = [sheet.cell(row, col).value for col in range(1, sheet.max_column + 1)]
            if name_col is not None and not values[name_col - 1]:
                continue
            raw_rows.append(values)
        wear_col_indices = [col - 1 for _, col in self.columns(sheet)]
        return {
            'headers': headers,
            'raw_rows': raw_rows,
            'wear_col_indices': wear_col_indices,
            'threshold': self.store.get_wear_abnormal_threshold(),
        }
