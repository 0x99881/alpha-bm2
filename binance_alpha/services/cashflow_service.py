from __future__ import annotations

import logging
from datetime import datetime
from typing import Any

from ..excel.cash_flow_sheet import regenerate_cash_flow_sheet
from ..ui_text import MESSAGES
from ..value_utils import to_float_or_none

LOGGER = logging.getLogger(__name__)


# 出金/入金 两个方向 + 分类下拉。分类是纯展示用的标签，改这里就能加/删项，
# 不影响已存数据（历史行照旧显示自己存下来的分类文本）。
DIRECTION_OUT = "out"
DIRECTION_IN = "in"
CASH_FLOW_CATEGORIES = ("红包支出", "日常消费", "转账借还", "入金", "其他")
DEFAULT_CATEGORY = "其他"


def _fmt_amount(value: float) -> str:
    """300.0 -> '300', 300.5 -> '300.5'. Keeps the ledger visually tidy."""
    rounded = round(float(value), 2)
    if rounded == int(rounded):
        return str(int(rounded))
    text = f"{rounded:.2f}".rstrip("0").rstrip(".")
    return text


class CashFlowService:
    """出入金流水账：录入、软删、按月/累计汇总，并把整张表重刷进 Excel。

    单位是 U。写库成功后才尝试写 Excel；Excel 被占用时写库照样成功，只把
    异常带回给路由提示，绝不因为 Excel 没写成而回滚已落库的这一笔——这是
    本项目对真实数据一贯的稳妥做法。
    """

    def __init__(self, local_db, workbook_repository, *, today_provider) -> None:
        self._local_db = local_db
        self._workbook_repository = workbook_repository
        self._today = today_provider

    # ---- helpers ----------------------------------------------------------
    @staticmethod
    def _text(value: Any) -> str:
        return str(value or "").strip()

    @staticmethod
    def _is_iso_date(value: str) -> bool:
        try:
            datetime.strptime(value, "%Y-%m-%d")
            return True
        except (TypeError, ValueError):
            return False

    # ---- queries ----------------------------------------------------------
    def get_view(self) -> dict[str, Any]:
        raw_rows = self._local_db.list_cash_flows()
        this_month = self._today()[:7]
        total_out = 0.0
        total_in = 0.0
        month_out = 0.0
        month_in = 0.0
        display_rows: list[dict[str, Any]] = []
        for row in raw_rows:
            amount = to_float_or_none(row.get("amount")) or 0.0
            direction = self._text(row.get("direction")) or DIRECTION_OUT
            is_out = direction != DIRECTION_IN
            entry_date = self._text(row.get("entry_date"))
            if is_out:
                total_out += amount
                if entry_date[:7] == this_month:
                    month_out += amount
            else:
                total_in += amount
                if entry_date[:7] == this_month:
                    month_in += amount
            display_rows.append({
                "id": self._text(row.get("id")),
                "entry_date": entry_date,
                "direction": DIRECTION_IN if not is_out else DIRECTION_OUT,
                "direction_label": MESSAGES["cashflow_in"] if not is_out else MESSAGES["cashflow_out"],
                "is_out": is_out,
                "amount": _fmt_amount(amount),
                "category": self._text(row.get("category")),
                "note": self._text(row.get("note")),
            })
        return {
            "rows": display_rows,
            "summary": {
                "month_out": _fmt_amount(month_out),
                "month_in": _fmt_amount(month_in),
                "total_out": _fmt_amount(total_out),
                "total_in": _fmt_amount(total_in),
                "net_out": _fmt_amount(total_out - total_in),
                "month_label": this_month,
            },
            "categories": list(CASH_FLOW_CATEGORIES),
            "today": self._today(),
        }

    # ---- commands ---------------------------------------------------------
    def add(self, form_data: Any) -> dict[str, Any]:
        entry_date = self._text(form_data.get("entry_date"))
        direction = self._text(form_data.get("direction")) or DIRECTION_OUT
        amount_text = self._text(form_data.get("amount"))
        category = self._text(form_data.get("category")) or DEFAULT_CATEGORY
        note = self._text(form_data.get("note"))

        if not self._is_iso_date(entry_date):
            raise ValueError(MESSAGES["cashflow_date_invalid"])
        if direction not in (DIRECTION_OUT, DIRECTION_IN):
            direction = DIRECTION_OUT
        amount_value = to_float_or_none(amount_text)
        if amount_value is None or amount_value <= 0:
            raise ValueError(MESSAGES["cashflow_amount_invalid"])

        self._local_db.add_cash_flow(
            entry_date=entry_date,
            direction=direction,
            amount=_fmt_amount(amount_value),
            category=category,
            note=note,
        )
        return {
            "export_error": self._regen_excel(),
            "direction_label": MESSAGES["cashflow_in"] if direction == DIRECTION_IN else MESSAGES["cashflow_out"],
            "amount": _fmt_amount(amount_value),
        }

    def delete(self, flow_id: str) -> dict[str, Any]:
        hit = self._local_db.soft_delete_cash_flow(self._text(flow_id))
        if not hit:
            raise ValueError(MESSAGES["cashflow_not_found"])
        return {"export_error": self._regen_excel()}

    # ---- excel ------------------------------------------------------------
    def _regen_excel(self) -> Exception | None:
        """Rewrite the 资金流水 sheet from the full ledger. ANY Excel failure —
        a locked workbook (ValueError), or the Excel layer choking on a
        corrupt/odd file — is caught and returned (not raised), so a bad
        workbook can never lose the DB write or 500 the page. The route turns a
        non-None result into a warning flash."""
        try:
            regenerate_cash_flow_sheet(
                self._workbook_repository, self._local_db.list_cash_flows()
            )
            return None
        except Exception as exc:  # noqa: BLE001 — Excel must never break the ledger
            LOGGER.exception("cash-flow Excel regen failed; DB write kept")
            return exc
