"""Rewrites the 资金流水 (personal cash-flow) sheet from the SQLite ledger.

One flat table, newest first — the same order the web page shows:

    日期 | 方向 | 金额(U) | 分类 | 备注
    …rows…
    [blank]
    合计 | 出金 X | 入金 Y | 净流出 Z

On every add/delete we drop the sheet and rewrite it whole, so the workbook
always mirrors the DB. A locked workbook raises ValueError from
``workbook_repository.save`` (excel_busy) — the caller keeps the DB write and
just reports it.
"""
from __future__ import annotations

from typing import Any

from openpyxl.styles import Alignment, Font, PatternFill

from ..value_utils import to_float_or_none


CASH_FLOW_SHEET_NAME = "资金流水"

HEADER_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
OUT_FILL = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
IN_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

HEADERS = ("日期", "方向", "金额(U)", "分类", "备注")


def regenerate_cash_flow_sheet(workbook_repository, rows: list[dict[str, Any]]) -> str:
    book = workbook_repository.open()
    try:
        write_cash_flow_sheet(book, rows)
        workbook_repository.save(book)
    finally:
        book.close()
    return CASH_FLOW_SHEET_NAME


def write_cash_flow_sheet(workbook, rows: list[dict[str, Any]]) -> str:
    if CASH_FLOW_SHEET_NAME in workbook.sheetnames:
        del workbook[CASH_FLOW_SHEET_NAME]
    sheet = workbook.create_sheet(title=CASH_FLOW_SHEET_NAME)

    for col_index, title in enumerate(HEADERS, start=1):
        cell = sheet.cell(1, col_index, title)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    total_out = 0.0
    total_in = 0.0
    row_index = 2
    for row in rows:
        amount = to_float_or_none(row.get("amount")) or 0.0
        direction = str(row.get("direction") or "out").strip()
        is_out = direction != "in"
        if is_out:
            total_out += amount
        else:
            total_in += amount

        sheet.cell(row_index, 1, str(row.get("entry_date") or "")).alignment = Alignment(horizontal="center")
        dir_cell = sheet.cell(row_index, 2, "出金" if is_out else "入金")
        dir_cell.alignment = Alignment(horizontal="center")
        dir_cell.fill = OUT_FILL if is_out else IN_FILL
        sheet.cell(row_index, 3, _coerce_number(amount)).alignment = Alignment(horizontal="right")
        sheet.cell(row_index, 4, str(row.get("category") or "")).alignment = Alignment(horizontal="center")
        sheet.cell(row_index, 5, str(row.get("note") or "")).alignment = Alignment(horizontal="left")
        row_index += 1

    # ---- totals row (blank separator, then 合计) --------------------------
    total_row = row_index + 1
    label = sheet.cell(total_row, 1, "合计")
    label.font = Font(bold=True)
    label.fill = TOTAL_FILL
    summary = sheet.cell(
        total_row, 2,
        f"出金 {_coerce_number(total_out)} / 入金 {_coerce_number(total_in)} / 净流出 {_coerce_number(total_out - total_in)}",
    )
    summary.font = Font(bold=True)
    summary.fill = TOTAL_FILL
    sheet.merge_cells(start_row=total_row, start_column=2, end_row=total_row, end_column=5)

    sheet.column_dimensions["A"].width = 12
    sheet.column_dimensions["B"].width = 8
    sheet.column_dimensions["C"].width = 12
    sheet.column_dimensions["D"].width = 12
    sheet.column_dimensions["E"].width = 30
    sheet.freeze_panes = "A2"
    return CASH_FLOW_SHEET_NAME


def _coerce_number(value: Any):
    if value in (None, ""):
        return None
    try:
        f = float(value)
    except (TypeError, ValueError):
        return value
    return int(f) if f.is_integer() else round(f, 2)
