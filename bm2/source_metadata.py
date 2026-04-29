from __future__ import annotations

import re


_PROFIT_PATTERN = re.compile(r"(?:^|\|)profit=(-?\d+(?:\.\d+)?)")


def parse_source_profit(source: object) -> float | None:
    match = _PROFIT_PATTERN.search(str(source or ""))
    if match is None:
        return None
    try:
        return round(float(match.group(1)), 1)
    except ValueError:
        return None


def with_source_profit(source: object, profit: float) -> str:
    base = str(source or "local").split("|profit=", 1)[0] or "local"
    return f"{base}|profit={round(float(profit), 1)}"
