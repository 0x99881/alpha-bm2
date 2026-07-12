from __future__ import annotations

from typing import Any, NamedTuple


class DailyTarget(NamedTuple):
    sheet: Any
    target_col: int
    total_col: int | None
    name_col: int
    header: str
    profit_col: int | None = None
    next_number: int | None = None
