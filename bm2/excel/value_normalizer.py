from __future__ import annotations

from decimal import Decimal, InvalidOperation

from ..ui_text import MESSAGES
from ..value_utils import round_value_sheet_number


def normalize_wear(value: Decimal | float | int) -> float:
    return round_value_sheet_number(value)


def normalize_income(value: Decimal | float | int) -> float:
    return round_value_sheet_number(value)


def normalize_expense(value: Decimal | float | int) -> float:
    return round_value_sheet_number(value)


def parse_decimal(value: str, field_name: str) -> Decimal:
    try:
        return Decimal(value)
    except InvalidOperation as exc:
        raise ValueError(MESSAGES['must_be_number'].format(field_name=field_name)) from exc
