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
from .excel.sheet_metadata import replace_sheet_meta

if TYPE_CHECKING:
    from .local_database import LocalDatabase

LOGGER = logging.getLogger(__name__)


class SQLiteToExcelExporter:
    """Reads SQLite and mirrors score data into the Excel workbook."""

    def __init__(self, local_db: LocalDatabase, excel_writer: Any) -> None:
        self._local_db = local_db
        self._excel_writer = excel_writer

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
            score_sheet = store._score_sheet(workbook)
            if score_sheet.max_column:
                score_sheet.delete_cols(1, score_sheet.max_column)
            for column_index, date_text in enumerate(dates, start=1):
                score_sheet.cell(1, column_index, str(date_text)[5:])
            total_col = len(dates) + 1
            profit_col = total_col + 1
            name_col = profit_col + 1
            score_sheet.cell(1, total_col, TOTAL_HEADER)
            score_sheet.cell(1, profit_col, PROFIT_HEADER)
            score_sheet.cell(1, name_col, NAME_HEADER)
            replace_sheet_meta(
                workbook,
                META_SHEET,
                [(date_text, f"D{index}") for index, date_text in enumerate(dates, start=1)],
            )
            row_map = ensure_member_rows_with_map(
                score_sheet,
                members=store.get_members(),
                name_col=name_col,
                total_col=total_col,
                value_columns=[col for _, col in store.score_sheet.date_columns(score_sheet)],
                extra_columns=[profit_col],
            )[0]
            for name in member_names:
                row = row_map[name]
                for column_index, date_text in enumerate(dates, start=1):
                    score_sheet.cell(row, column_index, int(score_map.get((name, date_text), 0)))
            store.score_sheet.ensure_structure(workbook)
            score_sheet = store._score_sheet(workbook)
            total_col = find_column(score_sheet, TOTAL_HEADER)
            profit_col = find_column(score_sheet, PROFIT_HEADER)
            if total_col is not None and profit_col is not None:
                recalculate_score_profits(workbook, score_sheet, profit_col, store.income_sheet, store.wear_sheet)
            store._sync_member_visibility_in_workbook(workbook)
            store.workbook_repository.save(workbook)
            return len(dates)
        finally:
            workbook.close()

