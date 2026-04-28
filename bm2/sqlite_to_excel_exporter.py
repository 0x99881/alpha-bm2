"""
SQLiteToExcelExporter - the only sanctioned path from SQLite to Excel.

This exporter writes score data only. Wear / income / expense are not stored
in SQLite yet, so exporting scores must not create or shift unrelated value
sheet columns.
"""
from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

from .constants import META_SHEET, NAME_HEADER, PROFIT_HEADER, TOTAL_HEADER
from .excel.header_locator import find_column
from .excel.member_rows import ensure_member_rows_with_map
from .excel.profit_recalculator import recalculate_score_profits
from .excel.sheet_metadata import append_sheet_meta, read_sheet_meta

if TYPE_CHECKING:
    from .local_database import LocalDatabase

LOGGER = logging.getLogger(__name__)


class SQLiteToExcelExporter:
    """Reads SQLite and mirrors score data into the Excel workbook."""

    def __init__(self, local_db: LocalDatabase, excel_writer: Any) -> None:
        self._local_db = local_db
        self._excel_writer = excel_writer

    def _score_meta_number(self, workbook, date_text: str) -> int | None:
        target = str(date_text).strip()
        for saved_date, column_name in read_sheet_meta(workbook, META_SHEET):
            if str(saved_date).strip() != target:
                continue
            column_name = str(column_name).strip()
            if column_name.startswith("D") and column_name[1:].isdigit():
                return int(column_name[1:])
        return None

    def _target_column_for_date(self, store: Any, workbook, date_text: str) -> tuple[int, int, bool]:
        score_sheet = store._score_sheet(workbook)
        total_col = find_column(score_sheet, TOTAL_HEADER)
        if total_col is None:
            raise ValueError("Score sheet total column is missing")

        existing_number = self._score_meta_number(workbook, date_text)
        date_columns = store.score_sheet.date_columns(score_sheet)
        if existing_number is not None:
            for number, col in date_columns:
                if number == existing_number:
                    return col, number, True

        next_number = date_columns[-1][0] + 1 if date_columns else 1
        score_sheet.insert_cols(total_col, 1)
        target_col = total_col
        score_sheet.cell(1, target_col, str(date_text)[5:])
        append_sheet_meta(workbook, META_SHEET, date_text, f"D{next_number}")
        return target_col, next_number, False

    def _write_date(self, store: Any, workbook, *, date_text: str, score_map: dict[tuple[str, str], int], member_names: list[str]) -> bool:
        score_sheet = store._score_sheet(workbook)
        target_col, _, existed = self._target_column_for_date(store, workbook, date_text)
        total_col = find_column(score_sheet, TOTAL_HEADER)
        profit_col = find_column(score_sheet, PROFIT_HEADER)
        name_col = find_column(score_sheet, NAME_HEADER)
        if total_col is None or profit_col is None or name_col is None:
            raise ValueError("Score sheet summary columns are missing")

        row_map = ensure_member_rows_with_map(
            score_sheet,
            members=store.get_members(),
            name_col=name_col,
            total_col=total_col,
            value_columns=[col for _, col in store.score_sheet.date_columns(score_sheet)],
            extra_columns=[profit_col],
        )[0]

        changed = False
        for name in member_names:
            row = row_map[name]
            target_value = int(score_map.get((name, date_text), 0))
            current_value = score_sheet.cell(row, target_col).value
            current_numeric = int(current_value or 0)
            if current_numeric != target_value:
                score_sheet.cell(row, target_col, target_value)
                changed = True
        return changed or not existed

    def export_missing_dates(self, store: Any) -> int:
        rows = self._local_db.get_filtered_score_rows()
        if not rows:
            return 0

        dates = sorted({
            str(row.get("score_date", "")).strip()
            for row in rows
            if str(row.get("score_date", "")).strip()
        })
        score_map: dict[tuple[str, str], int] = {
            (str(row["member_name"]).strip(), str(row["score_date"]).strip()): int(row["score"] or 0)
            for row in rows
        }
        member_names = [str(member["name"]).strip() for member in self._local_db.get_active_member_rows()]

        workbook = store.workbook_repository.open()
        try:
            store._ensure_score_sheet_structure(workbook)
            exported = 0
            for date_text in dates:
                try:
                    if self._write_date(store, workbook, date_text=date_text, score_map=score_map, member_names=member_names):
                        exported += 1
                except Exception:
                    LOGGER.exception("Failed to export date %s to Excel", date_text)
            store.score_sheet.ensure_structure(workbook)
            score_sheet = store._score_sheet(workbook)
            total_col = find_column(score_sheet, TOTAL_HEADER)
            profit_col = find_column(score_sheet, PROFIT_HEADER)
            if total_col is not None and profit_col is not None:
                recalculate_score_profits(workbook, score_sheet, profit_col, store.income_sheet, store.wear_sheet)
            store._sync_member_visibility_in_workbook(workbook)
            store.workbook_repository.save(workbook)
            return exported
        finally:
            workbook.close()

