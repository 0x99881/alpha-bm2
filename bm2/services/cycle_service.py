from __future__ import annotations

from typing import Any, Callable

from ..excel.value_normalizer import normalize_expense, normalize_income, normalize_wear
from ..value_utils import to_float_or_none

ENTRY_FIELDS = ("start_date", "start_balance", "settle_date", "end_balance")


class CycleService:
    """Business logic for the 周期盈亏情况 (cycle settlement) feature.

    Manual inputs (start/settle dates and balances) live in settlement tables.
    Wear and red-packet amounts are summed live from score_entries within each
    member's [start_date, settle_date] range.

    Profit = end_balance - start_balance - red-packet total. Wear is shown as a
    reference total only; it is already reflected in the balance change.
    """

    def __init__(self, local_db, get_active_members: Callable[[], list[dict[str, Any]]]) -> None:
        self._local_db = local_db
        self._get_active_members = get_active_members

    @staticmethod
    def _text(value: Any) -> str:
        return str(value or "").strip()

    def _resolve_wear(self, row: dict[str, Any]) -> float:
        manual = to_float_or_none(self._text(row.get("manual_wear")))
        if manual is not None:
            return manual
        before = to_float_or_none(self._text(row.get("before_balance")))
        after = to_float_or_none(self._text(row.get("after_balance")))
        if before is None or after is None:
            return 0.0
        return before - after

    def create_cycle(self, name: str) -> str:
        cleaned = self._text(name)
        if not cleaned:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_name_required"])
        return self._local_db.create_settlement_cycle(cleaned)

    def save_settlement(self, cycle_id: str, form_data: Any) -> None:
        cleaned_cycle = self._text(cycle_id)
        if not cleaned_cycle or self._local_db.get_settlement_cycle(cleaned_cycle) is None:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_not_found"])
        entries = []
        for member in self._get_active_members():
            name = self._text(member["name"])
            entries.append(
                {
                    "member_name": name,
                    "start_date": self._text(form_data.get(f"start_date_{name}", "")),
                    "start_balance": self._text(form_data.get(f"start_balance_{name}", "")),
                    "settle_date": self._text(form_data.get(f"settle_date_{name}", "")),
                    "end_balance": self._text(form_data.get(f"end_balance_{name}", "")),
                }
            )
        self._local_db.save_settlement_entries(cleaned_cycle, entries)

    def _member_range_rows(
        self,
        score_rows_by_member: dict[str, list[dict[str, Any]]],
        member_name: str,
        start_date: str,
        settle_date: str,
    ) -> list[dict[str, Any]]:
        if not start_date or not settle_date:
            return []
        low, high = sorted((start_date, settle_date))
        return [
            row
            for row in score_rows_by_member.get(member_name, [])
            if low <= self._text(row.get("score_date")) <= high
        ]

    def _build_member_row(
        self,
        member_name: str,
        settlement: dict[str, Any],
        range_rows: list[dict[str, Any]],
    ) -> dict[str, Any]:
        start_balance_text = self._text(settlement.get("start_balance"))
        end_balance_text = self._text(settlement.get("end_balance"))
        start_balance = to_float_or_none(start_balance_text)
        end_balance = to_float_or_none(end_balance_text)

        wear_total = normalize_wear(sum(self._resolve_wear(row) for row in range_rows))
        redpacket_items = []
        redpacket_sum = 0.0
        for row in sorted(range_rows, key=lambda item: self._text(item.get("score_date"))):
            amount = to_float_or_none(self._text(row.get("other_expense")))
            if amount is None or amount == 0:
                continue
            redpacket_sum += amount
            redpacket_items.append(
                {"date": self._text(row.get("score_date")), "amount": normalize_expense(amount)}
            )
        redpacket_total = normalize_expense(redpacket_sum)

        has_balances = start_balance is not None and end_balance is not None
        profit = (
            normalize_income(end_balance - start_balance - redpacket_total)
            if has_balances
            else None
        )
        return {
            "member_name": member_name,
            "start_date": self._text(settlement.get("start_date")),
            "start_balance": start_balance_text,
            "settle_date": self._text(settlement.get("settle_date")),
            "end_balance": end_balance_text,
            "wear_total": wear_total,
            "redpacket_total": redpacket_total,
            "redpacket_items": redpacket_items,
            "profit": profit,
            "has_balances": has_balances,
        }

    def get_cycle_profit_data(self, cycle_id: str | None = None) -> dict[str, Any]:
        cycles = self._local_db.get_settlement_cycles()
        selected_cycle = None
        if cycle_id:
            selected_cycle = self._local_db.get_settlement_cycle(self._text(cycle_id))
        if selected_cycle is None and cycles:
            selected_cycle = cycles[0]

        if selected_cycle is None:
            return {
                "cycles": cycles,
                "selected_cycle": None,
                "has_cycle": False,
                "rows": [],
                "totals": {"profit": 0.0, "wear": 0.0, "redpacket": 0.0},
            }

        settlements = self._local_db.get_settlement_entries(selected_cycle["id"])
        score_rows_by_member: dict[str, list[dict[str, Any]]] = {}
        for row in self._local_db.get_score_rows():
            score_rows_by_member.setdefault(self._text(row.get("member_name")), []).append(row)

        rows = []
        total_profit = 0.0
        total_wear = 0.0
        total_redpacket = 0.0
        for member in self._get_active_members():
            member_name = self._text(member["name"])
            settlement = settlements.get(member_name, {})
            range_rows = self._member_range_rows(
                score_rows_by_member,
                member_name,
                self._text(settlement.get("start_date")),
                self._text(settlement.get("settle_date")),
            )
            member_row = self._build_member_row(member_name, settlement, range_rows)
            rows.append(member_row)
            total_wear += member_row["wear_total"]
            total_redpacket += member_row["redpacket_total"]
            if member_row["profit"] is not None:
                total_profit += member_row["profit"]

        return {
            "cycles": cycles,
            "selected_cycle": selected_cycle,
            "has_cycle": True,
            "rows": rows,
            "totals": {
                "profit": normalize_income(total_profit),
                "wear": normalize_wear(total_wear),
                "redpacket": normalize_expense(total_redpacket),
            },
        }
