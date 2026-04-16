from __future__ import annotations

from decimal import Decimal
from typing import Any


def to_float_or_none(value: Any) -> float | None:
    if value in (None, ''):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def round_value_sheet_number(value: Decimal | float | int) -> float:
    return round(float(value), 1)


def round_wear(value: Decimal | float | int) -> float:
    return round_value_sheet_number(value)


def round_income(value: Decimal | float | int) -> float:
    return round_value_sheet_number(value)


def round_expense(value: Decimal | float | int) -> float:
    return round_value_sheet_number(value)


def build_threshold_abnormal_flags(
    rows: list[list[Any]],
    target_cols: list[int],
    threshold: float,
) -> dict[tuple[int, int], bool]:
    abnormal_flags: dict[tuple[int, int], bool] = {}
    for col_index in target_cols:
        for row_index, row in enumerate(rows):
            numeric_value = to_float_or_none(row[col_index])
            if numeric_value is None:
                continue
            abnormal_flags[(row_index, col_index)] = numeric_value > threshold
    return abnormal_flags
