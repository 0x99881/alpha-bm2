from __future__ import annotations

from decimal import Decimal
from typing import Callable


class IncompleteBalanceInput(ValueError):
    """Raised when only one side of a balance pair is provided."""


DecimalParser = Callable[[str, str], Decimal]


def resolve_wear_value(
    *,
    before_text: str,
    after_text: str,
    manual_wear_text: str,
    parse_decimal: DecimalParser,
    manual_field: str,
    before_field: str,
    after_field: str,
) -> Decimal | None:
    has_balance_input = bool(before_text or after_text)
    has_manual_input = bool(manual_wear_text)
    if not (has_balance_input or has_manual_input):
        return None

    if has_manual_input:
        return parse_decimal(manual_wear_text, manual_field)

    if not before_text or not after_text:
        raise IncompleteBalanceInput

    before_value = parse_decimal(before_text, before_field)
    after_value = parse_decimal(after_text, after_field)
    return before_value - after_value


def find_invalid_score_member(entries: list[dict[str, str]]) -> str | None:
    for entry in entries:
        score_text = entry['score']
        if not score_text:
            continue
        try:
            int(score_text)
        except ValueError:
            return entry['name']
    return None
