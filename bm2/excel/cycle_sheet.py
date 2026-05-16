"""Generates the per-cycle settlement Excel sheet.

Layout — one sheet per cycle, named like 周期MM-DD:

  Left:  姓名 | {start}余额 | {date1} | {date2} | ... | {settle}余额 | 红包合计 | 利润(余额) | 利润(流水) | 差额
  Gap
  Right: 姓名 | 磨损 | 收入

* Middle date columns are dynamic — only dates with red-packet activity get a column.
* 利润(余额) = 期末余额 - 期初余额 - 红包合计 (only after settle)
* 利润(流水) = 收入 - 磨损 - 红包合计 (only after settle)
* 差额 = 利润(余额) - 利润(流水), surfacing daily-entry/balance discrepancies.
* All profit cells show "未结算" until the cycle is settled.
"""
from __future__ import annotations

from datetime import datetime
from typing import Any

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from ..constants import NAME_HEADER


HEADER_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
PROFIT_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
PROFIT_FLOW_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
DELTA_FILL = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
WEAR_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
INCOME_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
UNSETTLED_FONT = Font(color="888888", italic=True)


def cycle_sheet_name(start_date: str) -> str:
    """周期MM-DD short tab name. Falls back to the raw text if parse fails."""
    try:
        d = datetime.strptime(start_date, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return f"周期{start_date}"
    return f"周期{d.month:02d}-{d.day:02d}"


def _zh_date(date_text: str) -> str:
    try:
        d = datetime.strptime(date_text, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return date_text
    return f"{d.month}月{d.day}日"


def _balance_header(date_text: str, fallback: str) -> str:
    zh = _zh_date(date_text)
    return f"{zh}余额" if zh and date_text else fallback


def remove_cycle_sheet(workbook, sheet_name: str) -> None:
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]


def delete_cycle_sheet(workbook_repository, start_date: str) -> str | None:
    sheet_name = cycle_sheet_name(start_date)
    book = workbook_repository.open()
    try:
        if sheet_name not in book.sheetnames:
            return None
        remove_cycle_sheet(book, sheet_name)
        workbook_repository.save(book)
        return sheet_name
    finally:
        book.close()


def regenerate_cycle_sheet(workbook_repository, cycle_data: dict[str, Any]) -> str | None:
    if not cycle_data.get("has_cycle"):
        return None
    book = workbook_repository.open()
    try:
        name = write_cycle_sheet(book, cycle_data)
        workbook_repository.save(book)
        return name
    finally:
        book.close()


def write_cycle_sheet(workbook, cycle_data: dict[str, Any]) -> str:
    cycle = cycle_data["selected_cycle"]
    start_date = str(cycle.get("start_date") or "").strip()
    settle_date = str(cycle.get("settle_date") or "").strip()
    is_settled = bool(int(cycle.get("settled", 0) or 0))
    redpacket_dates: list[str] = list(cycle_data.get("redpacket_dates") or [])
    rows: list[dict[str, Any]] = list(cycle_data.get("rows") or [])

    sheet_name = cycle_sheet_name(start_date)
    remove_cycle_sheet(workbook, sheet_name)
    sheet = workbook.create_sheet(title=sheet_name)

    start_header = _balance_header(start_date, "起始余额")
    settle_header = _balance_header(settle_date, "结算余额")

    # Left table headers
    left_headers = [NAME_HEADER, start_header]
    left_headers.extend(_zh_date(d) for d in redpacket_dates)
    left_headers.extend([settle_header, "红包合计", "利润(余额)", "利润(流水)", "差额"])

    for col_index, value in enumerate(left_headers, start=1):
        cell = sheet.cell(1, col_index, value)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    delta_col = len(left_headers)
    profit_flow_col = delta_col - 1
    profit_col = profit_flow_col - 1
    redpacket_total_col = profit_col - 1
    settle_balance_col = redpacket_total_col - 1
    date_columns = {date: 3 + index for index, date in enumerate(redpacket_dates)}

    # Right side small table: 姓名 | 磨损 | 收入
    gap_col = delta_col + 1
    right_name_col = gap_col + 1
    right_wear_col = right_name_col + 1
    right_income_col = right_name_col + 2
    for col_index, label in ((right_name_col, NAME_HEADER), (right_wear_col, "磨损"), (right_income_col, "收入")):
        c = sheet.cell(1, col_index, label)
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

    for row_index, row in enumerate(rows, start=2):
        name = row["member_name"]
        sheet.cell(row_index, 1, name).alignment = Alignment(horizontal="left")
        start_balance = _coerce_number(row.get("start_balance"))
        if start_balance is not None:
            sheet.cell(row_index, 2, start_balance)
        for date_text, amount in (row.get("redpacket_by_date") or {}).items():
            col = date_columns.get(date_text)
            if col is None or not amount:
                continue
            sheet.cell(row_index, col, _coerce_number(amount))
        end_balance = _coerce_number(row.get("end_balance"))
        if end_balance is not None:
            sheet.cell(row_index, settle_balance_col, end_balance)
        rp_total = row.get("redpacket_total") or 0
        rp_cell = sheet.cell(row_index, redpacket_total_col, rp_total if rp_total else None)
        if rp_total:
            rp_cell.fill = TOTAL_FILL

        if is_settled and row.get("profit") is not None:
            cell = sheet.cell(row_index, profit_col, row["profit"])
            cell.fill = PROFIT_FILL
            cell.font = Font(bold=True)
        else:
            cell = sheet.cell(row_index, profit_col, "未结算")
            cell.font = UNSETTLED_FONT
            cell.alignment = Alignment(horizontal="center")

        if is_settled and row.get("profit_flow") is not None:
            cell = sheet.cell(row_index, profit_flow_col, row["profit_flow"])
            cell.fill = PROFIT_FLOW_FILL
            cell.font = Font(bold=True)
        else:
            cell = sheet.cell(row_index, profit_flow_col, "未结算")
            cell.font = UNSETTLED_FONT
            cell.alignment = Alignment(horizontal="center")

        if is_settled and row.get("profit_delta") is not None:
            cell = sheet.cell(row_index, delta_col, row["profit_delta"])
            cell.fill = DELTA_FILL
        else:
            cell = sheet.cell(row_index, delta_col, "未结算")
            cell.font = UNSETTLED_FONT
            cell.alignment = Alignment(horizontal="center")

        # right side: name + wear + income
        sheet.cell(row_index, right_name_col, name).alignment = Alignment(horizontal="left")
        wear_total = row.get("wear_total") or 0
        wear_cell = sheet.cell(row_index, right_wear_col, wear_total if wear_total else None)
        if wear_total:
            wear_cell.fill = WEAR_FILL
        income_total = row.get("income_total") or 0
        income_cell = sheet.cell(row_index, right_income_col, income_total if income_total else None)
        if income_total:
            income_cell.fill = INCOME_FILL

    _autosize_columns(sheet, left_headers, right_name_col, right_wear_col, right_income_col)
    sheet.freeze_panes = "B2"
    return sheet_name


def _coerce_number(value: Any):
    if value in (None, ""):
        return None
    try:
        f = float(value)
    except (TypeError, ValueError):
        return value
    return int(f) if f.is_integer() else f


def _autosize_columns(sheet, left_headers, right_name_col, right_wear_col, right_income_col) -> None:
    for col_index, header in enumerate(left_headers, start=1):
        width = max(8, len(str(header)) * 2 + 2)
        sheet.column_dimensions[get_column_letter(col_index)].width = width
    sheet.column_dimensions[get_column_letter(right_name_col)].width = 10
    sheet.column_dimensions[get_column_letter(right_wear_col)].width = 10
    sheet.column_dimensions[get_column_letter(right_income_col)].width = 10
    gap_col = right_name_col - 1
    if gap_col > 0:
        sheet.column_dimensions[get_column_letter(gap_col)].width = 3
