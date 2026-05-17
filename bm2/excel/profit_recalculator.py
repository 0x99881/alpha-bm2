from __future__ import annotations

from datetime import date as _date_cls

from .header_locator import find_column
from .value_normalizer import normalize_expense, normalize_income, normalize_wear
from ..constants import NAME_HEADER


def _current_cycle_window(local_db) -> tuple[str, str]:
    """Return (start_iso, end_iso) for the current cycle, summed straight from
    SQLite's settlement_cycles table.

    Picks the latest unsettled cycle if any, otherwise falls back to the most
    recent settled cycle. When the DB has no cycles at all we return empty
    strings — the caller treats that as "include every entry" so a brand-new
    workbook still shows a meaningful profit column.
    """
    cycles = local_db.get_settlement_cycles()
    if not cycles:
        return "", ""

    def _text(value):
        return str(value or "").strip()

    ordered = sorted(
        cycles,
        key=lambda c: _text(c.get("start_date")) or _text(c.get("created_at")),
    )
    today_iso = _date_cls.today().strftime("%Y-%m-%d")
    for cycle in reversed(ordered):
        if not int(cycle.get("settled", 0) or 0):
            return _text(cycle.get("start_date")), today_iso
    latest = ordered[-1]
    settle = _text(latest.get("settle_date"))
    return _text(latest.get("start_date")), settle or today_iso


def recalculate_score_profits_from_db(
    *, local_db, workbook, score_sheet, profit_col: int,
) -> None:
    """Per-cycle flow-method profit (``income − wear − redpacket``).

    Each row in the score sheet shows the **current cycle's** running
    profit-loss for that member, summed straight from SQLite. Entries
    outside the cycle window are skipped so a settled cycle's totals don't
    drift forward into the next cycle's column.

    Header text comes from ``PROFIT_HEADER`` (``周期盈亏``) — the
    cycle-scope is what makes "累计" misleading and we dropped it.
    """
    from ..value_utils import to_float_or_none
    from ..domain.rules.daily_entry import resolve_wear_value, IncompleteBalanceInput
    from .value_normalizer import parse_decimal

    score_name_col = find_column(score_sheet, NAME_HEADER)
    if score_name_col is None:
        return

    start_iso, end_iso = _current_cycle_window(local_db)
    rows = local_db.get_score_rows()
    by_member: dict[str, dict[str, float]] = {}

    def _in_window(date_text: str) -> bool:
        if not start_iso:  # no cycle exists — include everything
            return True
        return start_iso <= date_text <= end_iso

    def _resolve_wear(row):
        try:
            value = resolve_wear_value(
                before_text=str(row.get("before_balance", "") or "").strip(),
                after_text=str(row.get("after_balance", "") or "").strip(),
                manual_wear_text=str(row.get("manual_wear", "") or "").strip(),
                parse_decimal=parse_decimal,
                manual_field="manual_wear",
                before_field="before",
                after_field="after",
            )
        except (IncompleteBalanceInput, ValueError):
            return None
        return None if value is None else normalize_wear(value)

    for row in rows:
        name = str(row.get("member_name", "") or "").strip()
        if not name:
            continue
        date_text = str(row.get("score_date", "") or "").strip()
        if not _in_window(date_text):
            continue
        bucket = by_member.setdefault(name, {"income": 0.0, "wear": 0.0, "expense": 0.0})
        income = to_float_or_none(str(row.get("income", "") or "").strip())
        expense = to_float_or_none(str(row.get("other_expense", "") or "").strip())
        wear = _resolve_wear(row)
        if income is not None:
            bucket["income"] += normalize_income(income)
        if expense is not None:
            bucket["expense"] += normalize_expense(expense)
        if wear is not None:
            bucket["wear"] += wear

    for row in range(2, score_sheet.max_row + 1):
        member_name = str(score_sheet.cell(row, score_name_col).value or "").strip()
        if not member_name:
            continue
        sums = by_member.get(member_name, {"income": 0.0, "wear": 0.0, "expense": 0.0})
        profit = normalize_income(
            normalize_income(sums["income"]) - normalize_wear(sums["wear"]) - normalize_expense(sums["expense"])
        )
        score_sheet.cell(row, profit_col, profit)


# Legacy alias for backward compatibility with call sites that pass the
# sheet handlers. The handlers are no longer needed for this aggregation —
# we read from DB.
def recalculate_score_profits(
    workbook, score_sheet, profit_col: int,
    income_sheet_handler, wear_sheet_handler, expense_sheet_handler,
    *, local_db=None,
) -> None:
    if local_db is None:
        # Best-effort: dig the local_db out of one of the handlers' stores.
        local_db = getattr(
            getattr(wear_sheet_handler, "store", None), "local_db", None,
        )
    if local_db is None:
        return
    recalculate_score_profits_from_db(
        local_db=local_db,
        workbook=workbook,
        score_sheet=score_sheet,
        profit_col=profit_col,
    )
