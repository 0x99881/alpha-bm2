from __future__ import annotations

import logging
from datetime import date, datetime, timedelta
from typing import Any, Callable

from ..constants import RESERVED_MEMBER_NAMES
from ..domain.rules.daily_entry import IncompleteBalanceInput, resolve_wear_value
from ..excel.value_normalizer import normalize_expense, normalize_income, normalize_wear, parse_decimal
from ..value_utils import to_float_or_none

LOGGER = logging.getLogger(__name__)


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

    @staticmethod
    def _iso_next_day(text: str) -> str:
        parsed = datetime.strptime(text, "%Y-%m-%d").date()
        return (parsed + timedelta(days=1)).strftime("%Y-%m-%d")

    @staticmethod
    def _today_iso() -> str:
        return date.today().strftime("%Y-%m-%d")

    def _resolve_wear(self, row: dict[str, Any]) -> float | None:
        try:
            value = resolve_wear_value(
                before_text=self._text(row.get("before_balance")),
                after_text=self._text(row.get("after_balance")),
                manual_wear_text=self._text(row.get("manual_wear")),
                parse_decimal=parse_decimal,
                manual_field="manual_wear",
                before_field="before_balance",
                after_field="after_balance",
            )
        except (IncompleteBalanceInput, ValueError):
            return None
        return None if value is None else normalize_wear(value)

    def _has_incomplete_balance(self, row: dict[str, Any]) -> bool:
        before = self._text(row.get("before_balance"))
        after = self._text(row.get("after_balance"))
        return bool(before) != bool(after)

    # ---- mutations ----------------------------------------------------------

    def create_cycle(self, start_date: str) -> str:
        from ..ui_text import MESSAGES

        cleaned = self._text(start_date)
        if not self._is_iso_date(cleaned):
            raise ValueError(MESSAGES["cycle_start_date_required"])

        # Only one open cycle at a time. Without this guard a stray double-
        # click or stale browser tab can quietly create a duplicate cycle that
        # shares a start date with the existing one, which then makes every
        # downstream view (charts, calendar, settle dialog) ambiguous.
        existing_unsettled = [
            c for c in self._local_db.get_settlement_cycles()
            if not int(c.get("settled", 0) or 0)
        ]
        if existing_unsettled:
            raise ValueError(MESSAGES["cycle_unsettled_exists"])

        # Chronological ordering: the new cycle must start strictly after the
        # most recently settled cycle. Anything else creates overlapping
        # windows that produce nonsense per-cycle aggregations.
        all_cycles = self._local_db.get_settlement_cycles()
        if all_cycles:
            latest_end = max(
                (self._text(c.get("settle_date")) or self._text(c.get("start_date")))
                for c in all_cycles
            )
            if cleaned <= latest_end:
                raise ValueError(
                    MESSAGES["cycle_start_must_be_after"].format(date=latest_end)
                )

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
                "is_extra": "0" if name in active_names else self._text(prior.get("is_extra", "0")),
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
        try:
            self._local_db.delete_settlement_cycle(cleaned_cycle)
        except Exception as exc:
            LOGGER.exception(
                "Cycle delete failed after Excel regeneration; cycle_id=%s",
                cleaned_cycle,
            )
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_delete_db_failed_after_excel"]) from exc

    # ---- current-cycle window helpers --------------------------------------

    def get_current_cycle(self) -> dict[str, Any] | None:
        """Latest un-settled cycle; if every cycle is settled, the most recent."""
        cycles = self._local_db.get_settlement_cycles()
        if not cycles:
            return None
        ordered = sorted(
            cycles,
            key=lambda c: self._text(c.get("start_date")) or self._text(c.get("created_at")),
        )
        for cycle in reversed(ordered):
            if not int(cycle.get("settled", 0) or 0):
                return cycle
        return ordered[-1]

    def get_window(self, cycle_id: str | None = None) -> dict[str, Any]:
        """Return the date window for a cycle (or the current cycle).

        Returns dict with: cycle_id, start_date, end_date, is_settled, has_cycle.
        end_date is settle_date if settled, else today (ISO).
        """
        cycle = None
        if cycle_id:
            cycle = self._local_db.get_settlement_cycle(self._text(cycle_id))
        if cycle is None:
            cycle = self.get_current_cycle()
        if cycle is None:
            return {
                "has_cycle": False,
                "cycle_id": "",
                "start_date": "",
                "end_date": "",
                "is_settled": False,
            }
        start = self._text(cycle.get("start_date"))
        is_settled = bool(int(cycle.get("settled", 0) or 0))
        settle = self._text(cycle.get("settle_date"))
        end = settle if (is_settled and settle) else self._today_iso()
        return {
            "has_cycle": True,
            "cycle_id": self._text(cycle.get("id")),
            "start_date": start,
            "end_date": end,
            "is_settled": is_settled,
        }

    def get_current_cycle_wear_summary(self, cycle_id: str | None = None) -> dict[str, Any]:
        """Wear total / per-member-avg for a cycle window.

        Used by the 磨损 page summary cards. Counts only members who have any
        wear entry within the window when computing the average, so an inactive
        member with no rows doesn't dilute the figure.

        ``cycle_id`` selects a specific cycle (any cycle the user picked from
        the dropdown); ``None`` falls back to the current/latest cycle.
        """
        window = self.get_window(cycle_id)
        if not window["has_cycle"]:
            return {
                "has_cycle": False,
                "cycle_name": "",
                "start_date": "",
                "end_date": "",
                "total_wear": 0.0,
                "member_count": 0,
                "avg_wear_per_member": 0.0,
            }
        start = window["start_date"]
        end = window["end_date"]
        cycle = self._local_db.get_settlement_cycle(window["cycle_id"]) or {}
        cycle_name = self._text(cycle.get("name"))

        wear_by_member: dict[str, float] = {}
        for row in self._local_db.get_score_rows():
            date_text = self._text(row.get("score_date"))
            if not (start <= date_text <= end):
                continue
            name = self._text(row.get("member_name"))
            if not name:
                continue
            wear_value = self._resolve_wear(row)
            if wear_value is None:
                continue
            wear_by_member[name] = wear_by_member.get(name, 0.0) + wear_value

        total = normalize_wear(sum(wear_by_member.values()))
        contributing = sum(1 for value in wear_by_member.values() if value)
        avg = normalize_wear(total / contributing) if contributing else 0.0
        return {
            "has_cycle": True,
            "cycle_name": cycle_name,
            "start_date": start,
            "end_date": end,
            "total_wear": total,
            "member_count": contributing,
            "avg_wear_per_member": avg,
        }

    # ---- atomic settle + new-cycle composite ------------------------------

    def settle_and_create_next(
        self,
        cycle_id: str,
        settle_date: str,
        end_balances: dict[str, str],
    ) -> str:
        """Settle the given cycle and create the next cycle starting settle_date+1.

        Atomic: if the Excel overview regen fails, all DB mutations are rolled
        back so the user can retry once Excel is closed.

        ``end_balances`` is a mapping of member_name -> end_balance string. Only
        present members will be persisted; missing names keep their prior value.
        """
        from ..ui_text import MESSAGES

        cleaned_cycle = self._text(cycle_id)
        cleaned_settle = self._text(settle_date)
        cycle = self._local_db.get_settlement_cycle(cleaned_cycle)
        if cycle is None:
            raise ValueError(MESSAGES["cycle_not_found"])
        if int(cycle.get("settled", 0) or 0):
            raise ValueError(MESSAGES["cycle_already_settled"])
        if not self._is_iso_date(cleaned_settle):
            raise ValueError(MESSAGES["cycle_settle_date_required"])
        start_date = self._text(cycle.get("start_date"))
        if start_date and cleaned_settle < start_date:
            raise ValueError(MESSAGES["cycle_settle_before_start"])
        next_start = self._iso_next_day(cleaned_settle)

        # Snapshot end_balances for rollback (start_balance is not touched).
        prior_entries = self._local_db.get_settlement_entries(cleaned_cycle)
        prior_end_balances = {
            name: self._text(entry.get("end_balance"))
            for name, entry in prior_entries.items()
        }

        # Reuse save_settlement by synthesising a form_data dict with only the
        # end_balance keys present. This respects the partial-update semantics
        # already implemented in save_settlement.
        synthetic_form: dict[str, str] = {}
        for name, value in (end_balances or {}).items():
            clean_name = self._text(name)
            if not clean_name:
                continue
            synthetic_form[f"end_balance_{clean_name}"] = self._text(value)

        # ---- apply DB mutations -----
        applied_end_balances = bool(synthetic_form)
        new_cycle_id: str | None = None
        stage = "prepare"
        try:
            if applied_end_balances:
                # save_settlement also calls _notify_change which regens Excel;
                # suppress that here so we can do a single regen at the end.
                stage = "save_end_balances"
                self._save_settlement_no_notify(cleaned_cycle, synthetic_form)
            stage = "settle_cycle"
            self._local_db.settle_settlement_cycle(cleaned_cycle, cleaned_settle)
            stage = "create_next_cycle"
            new_cycle_id = self._local_db.create_settlement_cycle(next_start)
            # ---- Excel regen (the only operation that can plausibly fail) ----
            if self._on_overview_regen is not None:
                stage = "regenerate_excel"
                self._on_overview_regen(self._build_cycles_data())
        except Exception:
            LOGGER.exception(
                "Settle-and-create-next failed; cycle_id=%s settle_date=%s stage=%s new_cycle_id=%s",
                cleaned_cycle,
                cleaned_settle,
                stage,
                new_cycle_id or "",
            )
            # Roll back in reverse order. Each repo call is best-effort.
            if new_cycle_id is not None:
                self._local_db.delete_settlement_cycle(new_cycle_id)
            self._local_db.unsettle_settlement_cycle(cleaned_cycle)
            if applied_end_balances:
                self._restore_end_balances(cleaned_cycle, prior_end_balances)
            raise
        return new_cycle_id or ""

    def _save_settlement_no_notify(self, cycle_id: str, form_data: Any) -> None:
        """Same as save_settlement but without the overview regen side effect."""
        existing = self._local_db.get_settlement_entries(cycle_id)
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
                "is_extra": "0" if name in active_names else self._text(prior.get("is_extra", "0")),
            }
            start_key = f"start_balance_{name}"
            end_key = f"end_balance_{name}"
            if start_key in form_data:
                entry["start_balance"] = self._text(form_data.get(start_key, ""))
            if end_key in form_data:
                entry["end_balance"] = self._text(form_data.get(end_key, ""))
            entries.append(entry)
        self._local_db.save_settlement_entries(cycle_id, entries)

    def _restore_end_balances(
        self, cycle_id: str, prior_end_balances: dict[str, str]
    ) -> None:
        synthetic: dict[str, str] = {}
        for name, value in prior_end_balances.items():
            synthetic[f"end_balance_{name}"] = value
        self._save_settlement_no_notify(cycle_id, synthetic)

    def add_extra_member(self, cycle_id: str, member_name: str) -> None:
        cleaned_cycle = self._text(cycle_id)
        cleaned_name = self._text(member_name)
        if not cleaned_name:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["cycle_member_name_required"])
        if cleaned_name in RESERVED_MEMBER_NAMES:
            from ..ui_text import MESSAGES

            raise ValueError(MESSAGES["member_name_reserved"].format(name=cleaned_name))
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

        wear_incomplete = any(self._has_incomplete_balance(row) for row in range_rows)
        wear_values = [
            value for value in (self._resolve_wear(row) for row in range_rows)
            if value is not None
        ]
        wear_total = None if wear_incomplete else normalize_wear(sum(wear_values))
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
        profit_flow = None
        profit = None
        profit_delta = None
        if wear_total is not None:
            profit_flow = normalize_income(income_total - wear_total - redpacket_total)
        if is_settled and has_balances:
            profit = normalize_income(end_balance - start_balance - redpacket_total)
            if profit_flow is not None:
                profit_delta = normalize_income(profit - profit_flow)

        return {
            "member_name": member_name,
            "start_balance": start_balance_text,
            "end_balance": end_balance_text,
            "wear_total": wear_total,
            "wear_incomplete": wear_incomplete,
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
                "is_latest_cycle": False,
                "start_date": "",
                "settle_date": "",
                "settle_date_default": "",
                "next_start_date_default": self._today_iso(),
                "redpacket_dates": [],
                "rows": [],
                "totals": {
                    "profit": 0.0, "profit_flow": 0.0, "profit_delta": 0.0,
                    "wear": 0.0, "wear_incomplete": False, "income": 0.0, "redpacket": 0.0,
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
        any_wear_incomplete = False
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
            if member_row["wear_incomplete"]:
                any_wear_incomplete = True
            if member_row["wear_total"] is not None:
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

        # Whether the viewed cycle is the most recent one (drives whether the
        # "open next cycle" controls appear on this page view). Tiebreak by
        # created_at then id so duplicate-start-date cycles don't cause the
        # ordering to flip-flop between renders.
        ordered_cycle_ids = [
            c["id"]
            for c in sorted(
                cycles,
                key=lambda c: (
                    self._text(c.get("start_date")) or self._text(c.get("created_at")),
                    self._text(c.get("created_at")),
                    self._text(c.get("id")),
                ),
            )
        ]
        is_latest_cycle = bool(ordered_cycle_ids) and (
            selected_cycle["id"] == ordered_cycle_ids[-1]
        )

        # Default new-cycle start_date when the viewed cycle is settled and
        # latest: settle_date + 1. Used to prefill the inline create form.
        next_start_date_default = ""
        if is_latest_cycle and is_settled and settle_date:
            try:
                next_start_date_default = self._iso_next_day(settle_date)
            except ValueError:
                next_start_date_default = ""

        # Default settle_date for the unsettled-latest dialog: today.
        settle_date_default = self._today_iso() if is_latest_cycle and not is_settled else ""

        return {
            "cycles": cycles,
            "selected_cycle": selected_cycle,
            "has_cycle": True,
            "is_settled": is_settled,
            "is_latest_cycle": is_latest_cycle,
            "start_date": start_date,
            "settle_date": settle_date,
            "settle_date_default": settle_date_default,
            "next_start_date_default": next_start_date_default,
            "redpacket_dates": redpacket_dates,
            "rows": rows,
            "totals": {
                "profit": normalize_income(total_profit) if is_settled else 0.0,
                "profit_flow": None if any_wear_incomplete else normalize_income(total_profit_flow),
                "profit_delta": None if (is_settled and any_wear_incomplete) else (normalize_income(total_profit_delta) if is_settled else 0.0),
                "wear": None if any_wear_incomplete else normalize_wear(total_wear),
                "wear_incomplete": any_wear_incomplete,
                "income": normalize_income(total_income),
                "redpacket": normalize_expense(total_redpacket),
            },
        }
