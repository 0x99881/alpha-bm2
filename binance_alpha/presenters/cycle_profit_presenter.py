from __future__ import annotations

from typing import Any


class CycleProfitPresenter:
    def __init__(self, cycle_service) -> None:
        self._cycle_service = cycle_service

    def build_cycle_profit_view(self, cycle_id: str | None = None) -> dict[str, Any]:
        return self._cycle_service.get_cycle_profit_data(cycle_id)
