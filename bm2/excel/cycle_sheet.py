"""Stacks every settlement cycle as one block inside a single Excel sheet.

Layout (one sheet, name = `周期盈亏记录`):

  [Block 1 — 周期 yyyy-mm-dd ~ yyyy-mm-dd  (已结算 | 未结算)]
    姓名 | {start}余额 | {date1} | … | {settle}余额 | 红包合计 | 利润 |    | 姓名 | 磨损 | 目前盈亏=收入减去磨损减去红包 | 差异=利润减去目前盈亏
    member rows
  [blank row]
  [Block 2 — 周期 …]
    …

* Left block ends with 利润 (= 结算余额 − 起始余额 − 红包合计, balance method).
* Right block carries 目前盈亏 (流水法: 收入 − 磨损 − 红包) and 差异
  (利润 − 目前盈亏) — the difference surfaces daily-entry inaccuracies.
* Each block keeps the per-cycle column layout (red-packet date columns vary).
* Blocks are written in chronological order (oldest first; new cycles appended).
* On any change we rewrite the whole sheet and drop the obsolete per-cycle
  tabs (周期MM-DD) created by the previous implementation.
"""
from __future__ import annotations

from datetime import datetime
from typing import Any

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from ..constants import NAME_HEADER


OVERVIEW_SHEET_NAME = "周期盈亏记录"

HEADER_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
PROFIT_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
PROFIT_FLOW_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
DELTA_FILL = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
WEAR_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
INCOME_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
TITLE_SETTLED_FILL = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
TITLE_UNSETTLED_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
UNSETTLED_FONT = Font(color="888888", italic=True)


def _zh_date(date_text: str) -> str:
    try:
        d = datetime.strptime(date_text, "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return date_text
    return f"{d.month}月{d.day}日"


def _balance_header(date_text: str, fallback: str) -> str:
    zh = _zh_date(date_text)
    return f"{zh}余额" if zh and date_text else fallback


def _drop_legacy_per_cycle_sheets(workbook) -> None:
    """Remove old `周期MM-DD` tabs from the previous one-sheet-per-cycle design."""
    legacy = [
        name for name in workbook.sheetnames
        if name.startswith("周期") and name != OVERVIEW_SHEET_NAME
    ]
    for name in legacy:
        del workbook[name]


def regenerate_overview_sheet(workbook_repository, cycles_data: list[dict[str, Any]]) -> str:
    """Rebuild the entire 周期盈亏记录 sheet from the given cycles."""
    book = workbook_repository.open()
    try:
        write_overview_sheet(book, cycles_data)
        workbook_repository.save(book)
    finally:
        book.close()
    return OVERVIEW_SHEET_NAME


# Backward-compatible aliases (callers still using the old names).
def regenerate_cycle_sheet(workbook_repository, cycles_data: list[dict[str, Any]]) -> str:
    return regenerate_overview_sheet(workbook_repository, cycles_data)


def delete_cycle_sheet(workbook_repository, _start_date: str) -> str:
    """Compatibility shim — actual cleanup happens via the regen path now."""
    return OVERVIEW_SHEET_NAME


def write_overview_sheet(workbook, cycles_data: list[dict[str, Any]]) -> str:
    _drop_legacy_per_cycle_sheets(workbook)
    if OVERVIEW_SHEET_NAME in workbook.sheetnames:
        del workbook[OVERVIEW_SHEET_NAME]
    sheet = workbook.create_sheet(title=OVERVIEW_SHEET_NAME)

    if not cycles_data:
        sheet.cell(1, 1, "暂无周期").font = UNSETTLED_FONT
        sheet.freeze_panes = "A2"
        return OVERVIEW_SHEET_NAME

    current_row = 1
    max_width_map: dict[int, int] = {}
    for cycle_data in cycles_data:
        rows_used, col_widths = _write_one_block(sheet, current_row, cycle_data)
        for col_index, width in col_widths.items():
            max_width_map[col_index] = max(max_width_map.get(col_index, 0), width)
        current_row += rows_used + 1  # 1 blank separator between blocks

    for col_index, width in max_width_map.items():
        sheet.column_dimensions[get_column_letter(col_index)].width = max(width, 8)

    sheet.freeze_panes = "A2"
    return OVERVIEW_SHEET_NAME


def _write_one_block(sheet, start_row: int, cycle_data: dict[str, Any]) -> tuple[int, dict[int, int]]:
    cycle = cycle_data["selected_cycle"]
    start_date = str(cycle.get("start_date") or "").strip()
    settle_date = str(cycle.get("settle_date") or "").strip()
    is_settled = bool(int(cycle.get("settled", 0) or 0))
    redpacket_dates: list[str] = list(cycle_data.get("redpacket_dates") or [])
    rows: list[dict[str, Any]] = list(cycle_data.get("rows") or [])

    start_header = _balance_header(start_date, "起始余额")
    settle_header = _balance_header(settle_date, "结算余额")

    left_headers = [NAME_HEADER, start_header]
    left_headers.extend(_zh_date(d) for d in redpacket_dates)
    left_headers.extend([settle_header, "红包合计", "利润"])

    profit_col = len(left_headers)
    redpacket_total_col = profit_col - 1
    settle_balance_col = redpacket_total_col - 1
    date_columns = {date: 3 + index for index, date in enumerate(redpacket_dates)}

    gap_col = profit_col + 1
    right_name_col = gap_col + 1
    right_wear_col = right_name_col + 1
    right_flow_col = right_name_col + 2
    right_delta_col = right_name_col + 3
    total_cols = right_delta_col

    # ---- title row (merged across the block) -----------------------------
    title_text = _title_for(start_date, settle_date, is_settled)
    title_cell = sheet.cell(start_row, 1, title_text)
    title_cell.font = Font(bold=True, size=12, color="1F2937")
    title_cell.fill = TITLE_SETTLED_FILL if is_settled else TITLE_UNSETTLED_FILL
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    sheet.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=total_cols)

    # ---- header row ------------------------------------------------------
    header_row = start_row + 1
    for col_index, value in enumerate(left_headers, start=1):
        cell = sheet.cell(header_row, col_index, value)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    right_columns = (
        (right_name_col, NAME_HEADER),
        (right_wear_col, "磨损"),
        (right_flow_col, "目前盈亏"),
        (right_delta_col, "差异"),
    )
    for col_index, label in right_columns:
        c = sheet.cell(header_row, col_index, label)
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")

    # ---- member rows -----------------------------------------------------
    body_start = header_row + 1
    for row_index, row in enumerate(rows, start=body_start):
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
        if is_settled and end_balance is not None:
            sheet.cell(row_index, settle_balance_col, end_balance)
        elif not is_settled:
            cell = sheet.cell(row_index, settle_balance_col, "未结算")
            cell.font = UNSETTLED_FONT
            cell.alignment = Alignment(horizontal="center")
        rp_total = row.get("redpacket_total") or 0
        rp_cell = sheet.cell(row_index, redpacket_total_col, rp_total if rp_total else None)
        if rp_total:
            rp_cell.fill = TOTAL_FILL

        _write_profit_cell(sheet, row_index, profit_col, is_settled, row.get("profit"), PROFIT_FILL)

        sheet.cell(row_index, right_name_col, name).alignment = Alignment(horizontal="left")
        wear_total = row.get("wear_total") or 0
        if is_settled:
            wear_cell = sheet.cell(row_index, right_wear_col, wear_total if wear_total else None)
            if wear_total:
                wear_cell.fill = WEAR_FILL
        else:
            cell = sheet.cell(row_index, right_wear_col, "未结算")
            cell.font = UNSETTLED_FONT
            cell.alignment = Alignment(horizontal="center")
        _write_profit_cell(sheet, row_index, right_flow_col, is_settled, row.get("profit_flow"), PROFIT_FLOW_FILL)
        _write_profit_cell(sheet, row_index, right_delta_col, is_settled, row.get("profit_delta"), DELTA_FILL, bold=False)

    rows_used = 2 + len(rows)  # title row + header row + member rows

    # Per-column width estimate for this block (we collect; outer fn aggregates)
    col_widths: dict[int, int] = {}
    for col_index, header in enumerate(left_headers, start=1):
        col_widths[col_index] = max(8, len(str(header)) * 2 + 2)
    col_widths[right_name_col] = 10
    col_widths[right_wear_col] = 10
    col_widths[right_flow_col] = 12
    col_widths[right_delta_col] = 10
    col_widths[gap_col] = 3

    return rows_used, col_widths


def _title_for(start_date: str, settle_date: str, is_settled: bool) -> str:
    start_label = start_date or "起始日期未填"
    if is_settled and settle_date:
        return f"周期 {start_label} ~ {settle_date}    已结算"
    return f"周期 {start_label} ~ 未结算"


def _write_profit_cell(sheet, row_index: int, col: int, is_settled: bool, value, fill, bold: bool = True) -> None:
    if is_settled and value is not None:
        cell = sheet.cell(row_index, col, value)
        cell.fill = fill
        if bold:
            cell.font = Font(bold=True)
    else:
        cell = sheet.cell(row_index, col, "未结算")
        cell.font = UNSETTLED_FONT
        cell.alignment = Alignment(horizontal="center")


def _coerce_number(value: Any):
    if value in (None, ""):
        return None
    try:
        f = float(value)
    except (TypeError, ValueError):
        return value
    return int(f) if f.is_integer() else f
