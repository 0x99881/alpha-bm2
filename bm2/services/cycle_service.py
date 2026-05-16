from __future__ import annotations

from datetime import datetime
from typing import Any, Callable

from ..excel.value_normalizer import normalize_expense, normalize_income, normalize_wear
from ..value_utils import to_float_or_none


class CycleService:
    """Business logic for 周期盈亏情况.

    A cycle has a single start_date and (after settling) a single settle_date.
    Manual inputs per member: start_balance, end_balance. Wear, income and
    red-packet amounts are summed live from score_entries within the cycle
    date range.

    Two profit views (only after settle):
      - profit       = end_balance - start_balance - redpacket_total (余额法)
      - profit_flow  = income_total - wear_total - redpacket_total   (流水法)
      - profit_delta = profit - profit_flow (records day-entry vs balance gap)
    """

    def __init__(
        self,
        local_db,
        get_active_members: Callable[[], list[dict[str, Any]]],
        on_overview_regen=None,
    ) -> None:
        self._local_db = local_db
        self._get_active_members = get_active_members
        # Excel layer hook: receives the full cycles-data list and rewrites the
        # 周期盈亏记录 overview tab; wired by bootstrap.
        self._on_overview_regen = on_overview_regen

    @staticmethod
    def _text(value: Any) -> str:
        return str(value or "").strip()

    @staticmethod
    def _is_iso_date(text: str) -> bool:
        if not text:
            return False
        try:
            return datetime.strptime(text, "%Y-%m-%d").strftime("%Y-%m-%d") == text
        except ValueError:
            return False

    def _resolve_wear(self, row: dict[str, Any]) -> float:
        manual = to_float_or_none(self._text(row.get("manual_wear")))
        if manual is not None:
            return manual
        before = to_float_or_none(self._text(row.get("before_balance")))
        after = to_float_or_none(self._text(row.get("after_balance")))
        if before is None or after is None:
            return 0.0
        return before - after

    # ---- mutations ----------------------------------------------------------

    def create_cycle(self, start_date: str) -> str:
        cleaned = self._text(start_date)
        if not self._is_iso_date(cleaned):
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_start_date_required"])
        cycle_id = self._local_db.create_settlement_cycle(cleaned)
        self._notify_change(cycle_id)
        return cycle_id

    def save_settlement(self, cycle_id: str, form_data: Any) -> None:
        """Partial update: only persist fields present in ``form_data``."""
        cleaned_cycle = self._text(cycle_id)
        if not cleaned_cycle or self._local_db.get_settlement_cycle(cleaned_cycle) is None:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_not_found"])
        existing = self._local_db.get_settlement_entries(cleaned_cycle)
        active_names = [self._text(m["name"]) for m in self._get_active_members()]
        names_to_save = list(active_names) + [
            name for name in existing.keys() if name not in active_names
        ]
        entries = []
        for name in names_to_save:
            prior = existing.get(name, {})
            entry = {
                "member_name": name,
                "start_balance": self._text(prior.get("start_balance", "")),
                "end_balance": self._text(prior.get("end_balance", "")),
            }
            start_key = f"start_balance_{name}"
            end_key = f"end_balance_{name}"
            if start_key in form_data:
                entry["start_balance"] = self._text(form_data.get(start_key, ""))
            if end_key in form_data:
                entry["end_balance"] = self._text(form_data.get(end_key, ""))
            entries.append(entry)
        self._local_db.save_settlement_entries(cleaned_cycle, entries)
        self._notify_change(cleaned_cycle)

    def settle_cycle(self, cycle_id: str, settle_date: str) -> None:
        cleaned_cycle = self._text(cycle_id)
        cleaned_settle = self._text(settle_date)
        cycle = self._local_db.get_settlement_cycle(cleaned_cycle)
        if cycle is None:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_not_found"])
        if not self._is_iso_date(cleaned_settle):
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_settle_date_required"])
        self._local_db.settle_settlement_cycle(cleaned_cycle, cleaned_settle)
        self._notify_change(cleaned_cycle)

    def delete_cycle(self, cycle_id: str) -> None:
        # Rewrite Excel first with the post-delete cycle list. If Excel is locked
        # the regen raises and DB stays intact so the user can retry cleanly.
        cleaned_cycle = self._text(cycle_id)
        cycle = self._local_db.get_settlement_cycle(cleaned_cycle)
        if cycle is None:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_not_found"])
        if self._on_overview_regen is not None:
            remaining = self._build_cycles_data(exclude_cycle_id=cleaned_cycle)
            self._on_overview_regen(remaining)
        self._local_db.delete_settlement_cycle(cleaned_cycle)

    def add_extra_member(self, cycle_id: str, member_name: str) -> None:
        cleaned_cycle = self._text(cycle_id)
        cleaned_name = self._text(member_name)
        if not cleaned_name:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_member_name_required"])
        cycle = self._local_db.get_settlement_cycle(cleaned_cycle)
        if cycle is None:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_not_found"])
        active_names = {self._text(m["name"]) for m in self._get_active_members()}
        if cleaned_name in active_names:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_member_active_conflict"].format(name=cleaned_name))
        added = self._local_db.add_cycle_extra_member(cleaned_cycle, cleaned_name)
        if not added:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_member_already_exists"].format(name=cleaned_name))
        self._notify_change(cleaned_cycle)

    def remove_extra_member(self, cycle_id: str, member_name: str) -> None:
        cleaned_cycle = self._text(cycle_id)
        cleaned_name = self._text(member_name)
        cycle = self._local_db.get_settlement_cycle(cleaned_cycle)
        if cycle is None:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_not_found"])
        removed = self._local_db.remove_cycle_extra_member(cleaned_cycle, cleaned_name)
        if not removed:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_member_not_removable"].format(name=cleaned_name))
        self._notify_change(cleaned_cycle)

    def reorder_members(self, cycle_id: str, ordered_names: list[str]) -> None:
        cleaned_cycle = self._text(cycle_id)
        cycle = self._local_db.get_settlement_cycle(cleaned_cycle)
        if cycle is None:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_not_found"])
        clean = [self._text(name) for name in ordered_names if self._text(name)]
        self._local_db.set_cycle_member_order(cleaned_cycle, clean)
        self._notify_change(cleaned_cycle)

    def _notify_change(self, cycle_id: str) -> None:
        if self._on_overview_regen is None:
            return
        self._on_overview_regen(self._build_cycles_data())

    def _build_cycles_data(self, *, exclude_cycle_id: str | None = None) -> list[dict[str, Any]]:
        """All non-deleted cycles in chronological order, with full view data."""
        cycles = self._local_db.get_settlement_cycles()

        def sort_key(cycle: dict[str, Any]) -> str:
            return self._text(cycle.get("start_date")) or self._text(cycle.get("created_at"))

        cycles_sorted = sorted(cycles, key=sort_key)
        return [
            self.get_cycle_profit_data(cycle["id"])
            for cycle in cycles_sorted
            if cycle["id"] != exclude_cycle_id
        ]

    def get_all_cycles_data(self) -> list[dict[str, Any]]:
        return self._build_cycles_data()

    # ---- read ---------------------------------------------------------------

    def _range_rows_for_cycle(
        self,
        score_rows_by_member: dict[str, list[dict[str, Any]]],
        member_name: str,
        start_date: str,
        settle_date: str,
    ) -> list[dict[str, Any]]:
        if not start_date:
            return []
        high = settle_date if settle_date else "9999-12-31"
        low, hi = sorted((start_date, high))
        return [
            row
            for row in score_rows_by_member.get(member_name, [])
            if low <= self._text(row.get("score_date")) <= hi
        ]

    def _build_member_row(
        self,
        member_name: str,
        settlement: dict[str, Any],
        range_rows: list[dict[str, Any]],
        is_settled: bool,
        is_extra: bool = False,
    ) -> dict[str, Any]:
        start_balance_text = self._text(settlement.get("start_balance"))
        end_balance_text = self._text(settlement.get("end_balance"))
        start_balance = to_float_or_none(start_balance_text)
        end_balance = to_float_or_none(end_balance_text)

        wear_total = normalize_wear(sum(self._resolve_wear(row) for row in range_rows))
        income_sum = 0.0
        for row in range_rows:
            amount = to_float_or_none(self._text(row.get("income")))
            if amount is not None:
                income_sum += amount
        income_total = normalize_income(income_sum)
        redpacket_by_date: dict[str, float] = {}
        for row in range_rows:
            amount = to_float_or_none(self._text(row.get("other_expense")))
            if amount is None or amount == 0:
                continue
            date_text = self._text(row.get("score_date"))
            redpacket_by_date[date_text] = redpacket_by_date.get(date_text, 0.0) + amount
        redpacket_total = normalize_expense(sum(redpacket_by_date.values()))

        has_balances = start_balance is not None and end_balance is not None
        # 目前盈亏 (流水法) is computable any time daily-entry data exists — show
        # it during the cycle as a running estimate, not only after settle.
        profit_flow = normalize_income(income_total - wear_total - redpacket_total)
        profit = None
        profit_delta = None
        if is_settled and has_balances:
            profit = normalize_income(end_balance - start_balance - redpacket_total)
            profit_delta = normalize_income(profit - profit_flow)

        return {
            "member_name": member_name,
            "start_balance": start_balance_text,
            "end_balance": end_balance_text,
            "wear_total": wear_total,
            "income_total": income_total,
            "redpacket_by_date": redpacket_by_date,
            "redpacket_total": redpacket_total,
            "profit": profit,
            "profit_flow": profit_flow,
            "profit_delta": profit_delta,
            "has_balances": has_balances,
            "is_extra": is_extra,
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
                "is_settled": False,
                "start_date": "",
                "settle_date": "",
                "redpacket_dates": [],
                "rows": [],
                "totals": {
                    "profit": 0.0, "profit_flow": 0.0, "profit_delta": 0.0,
                    "wear": 0.0, "income": 0.0, "redpacket": 0.0,
                },
            }

        start_date = self._text(selected_cycle.get("start_date"))
        settle_date = self._text(selected_cycle.get("settle_date"))
        is_settled = bool(int(selected_cycle.get("settled", 0) or 0))

        settlements = self._local_db.get_settlement_entries(selected_cycle["id"])
        score_rows_by_member: dict[str, list[dict[str, Any]]] = {}
        for row in self._local_db.get_score_rows():
            score_rows_by_member.setdefault(self._text(row.get("member_name")), []).append(row)

        active_member_names = [self._text(m["name"]) for m in self._get_active_members()]
        active_name_set = set(active_member_names)
        extra_names = [
            name for name, entry in settlements.items()
            if int(entry.get("is_extra", 0) or 0) == 1 and name not in active_name_set
        ]
        all_names = active_member_names + extra_names

        def sort_key(name: str) -> tuple[int, int]:
            order = int((settlements.get(name) or {}).get("sort_order", 0) or 0)
            default_idx = all_names.index(name)
            return (0, order) if order > 0 else (1, default_idx)

        ordered_names = sorted(all_names, key=sort_key)

        rows = []
        all_dates: set[str] = set()
        total_profit = 0.0
        total_profit_flow = 0.0
        total_profit_delta = 0.0
        total_wear = 0.0
        total_income = 0.0
        total_redpacket = 0.0
        for member_name in ordered_names:
            settlement = settlements.get(member_name, {})
            is_extra = member_name not in active_name_set
            range_rows = self._range_rows_for_cycle(
                score_rows_by_member, member_name, start_date, settle_date,
            )
            member_row = self._build_member_row(
                member_name, settlement, range_rows, is_settled, is_extra=is_extra,
            )
            rows.append(member_row)
            all_dates.update(member_row["redpacket_by_date"].keys())
            total_wear += member_row["wear_total"]
            total_income += member_row["income_total"]
            total_redpacket += member_row["redpacket_total"]
            if member_row["profit"] is not None:
                total_profit += member_row["profit"]
            if member_row["profit_flow"] is not None:
                total_profit_flow += member_row["profit_flow"]
            if member_row["profit_delta"] is not None:
                total_profit_delta += member_row["profit_delta"]

        redpacket_dates = sorted(all_dates)
        return {
            "cycles": cycles,
            "selected_cycle": selected_cycle,
            "has_cycle": True,
            "is_settled": is_settled,
            "start_date": start_date,
            "settle_date": settle_date,
            "redpacket_dates": redpacket_dates,
            "rows": rows,
            "totals": {
                "profit": normalize_income(total_profit) if is_settled else 0.0,
                "profit_flow": normalize_income(total_profit_flow),
                "profit_delta": normalize_income(total_profit_delta) if is_settled else 0.0,
                "wear": normalize_wear(total_wear),
                "income": normalize_income(total_income),
                "redpacket": normalize_expense(total_redpacket),
            },
        }
