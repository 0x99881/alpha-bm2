from __future__ import annotations

from .header_locator import find_column
from .member_rows import build_name_row_map
from .number_utils import sum_sheet_row_values
from .value_normalizer import normalize_expense, normalize_income, normalize_wear
from ..constants import EXPENSE_NAME_HEADER, INCOME_NAME_HEADER, NAME_HEADER, WEAR_NAME_HEADER


def recalculate_score_profits(workbook, score_sheet, profit_col: int, income_sheet_handler, wear_sheet_handler, expense_sheet_handler) -> None:
    score_name_col = find_column(score_sheet, NAME_HEADER)
    income_sheet = income_sheet_handler.sheet(workbook)
    wear_sheet = wear_sheet_handler.sheet(workbook)
    expense_sheet = expense_sheet_handler.sheet(workbook)
    income_name_col = find_column(income_sheet, INCOME_NAME_HEADER)
    wear_name_col = find_column(wear_sheet, WEAR_NAME_HEADER)
    expense_name_col = find_column(expense_sheet, EXPENSE_NAME_HEADER)
    if score_name_col is None or income_name_col is None or wear_name_col is None:
        return

    income_row_map = build_name_row_map(income_sheet, income_name_col)
    wear_row_map = build_name_row_map(wear_sheet, wear_name_col)
    expense_row_map = build_name_row_map(expense_sheet, expense_name_col) if expense_name_col is not None else {}
    income_columns = [col for _, col in income_sheet_handler.columns(income_sheet)]
    wear_columns = [col for _, col in wear_sheet_handler.columns(wear_sheet)]
    expense_columns = [col for _, col in expense_sheet_handler.columns(expense_sheet)] if expense_name_col is not None else []

    for row in range(2, score_sheet.max_row + 1):
        member_name = str(score_sheet.cell(row, score_name_col).value or '').strip()
        if not member_name:
            continue
        income_total = normalize_income(
            sum_sheet_row_values(income_sheet, income_row_map[member_name], income_columns)
            if member_name in income_row_map else 0
        )
        wear_total = normalize_wear(
            sum_sheet_row_values(wear_sheet, wear_row_map[member_name], wear_columns)
            if member_name in wear_row_map else 0
        )
        expense_total = normalize_expense(
            sum_sheet_row_values(expense_sheet, expense_row_map[member_name], expense_columns)
            if member_name in expense_row_map else 0
        )
        score_sheet.cell(row, profit_col, normalize_income(income_total - wear_total - expense_total))
