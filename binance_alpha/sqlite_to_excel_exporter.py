"""
SQLiteToExcelExporter - the only sanctioned path from SQLite to Excel.

SQLite is the local source of truth. Export must mirror the stored score,
wear, income, and expense entry fields back into the workbook without treating
Excel as the primary write path.
"""
from __future__ import annotations

import logging
from decimal import Decimal, InvalidOperation
from typing import TYPE_CHECKING, Any

from .domain.rules.daily_entry import IncompleteBalanceInput, resolve_wear_value
from .excel.cycle_block_sheets import build_cycle_blocks, rewrite_value_sheet_with_blocks
from .excel.header_locator import find_column
from .excel.member_rows import ensure_member_rows_with_map
from .excel.profit_recalculator import recalculate_score_profits
from .excel.sheet_metadata import append_sheet_meta, read_sheet_meta, replace_sheet_meta
from .excel.value_normalizer import normalize_expense, normalize_income, normalize_wear, parse_decimal
from .excel.value_sheet_spec import get_value_sheet_spec
from .constants import META_SHEET, NAME_HEADER, PROFIT_HEADER, TOTAL_HEADER
from .ui_text import MESSAGES

if TYPE_CHECKING:
    from .local_database import LocalDatabase

LOGGER = logging.getLogger(__name__)


class SQLiteToExcelExporter:
    """Reads SQLite and mirrors saved entry data into the Excel workbook."""

    def __init__(self, local_db: LocalDatabase, excel_writer: Any) -> None:
        self._local_db = local_db
        self._excel_writer = excel_writer

    @staticmethod
    def _entry_text(row: dict[str, Any], key: str) -> str:
        return str(row.get(key, "") or "").strip()

    def _resolve_wear(self, row: dict[str, Any]) -> float | None:
        member_name = self._entry_text(row, "member_name")
        try:
            value = resolve_wear_value(
                before_text=self._entry_text(row, "before_balance"),
                after_text=self._entry_text(row, "after_balance"),
                manual_wear_text=self._entry_text(row, "manual_wear"),
                parse_decimal=parse_decimal,
                manual_field=MESSAGES["manual_wear_field"].format(name=member_name),
                before_field=MESSAGES["before_balance_field"].format(name=member_name),
                after_field=MESSAGES["after_balance_field"].format(name=member_name),
            )
        except (IncompleteBalanceInput, ValueError):
            return None
        if value is None:
            return None
        return normalize_wear(value)

    def _resolve_value(self, row: dict[str, Any], key: str) -> float | None:
        raw_value = self._entry_text(row, key)
        if not raw_value:
            return None
        try:
            value = Decimal(raw_value)
        except (InvalidOperation, ValueError):
            return None
        if key == "income":
            return normalize_income(value)
        return normalize_expense(value)

    def _detail_dates(self, rows: list[dict[str, Any]]) -> list[str]:
        dates = set()
        for row in rows:
            if (
                self._resolve_wear(row) is not None
                or self._resolve_value(row, "income") is not None
                or self._resolve_value(row, "other_expense") is not None
            ):
                date_text = self._entry_text(row, "score_date")
                if date_text:
                    dates.add(date_text)
        return sorted(dates)

    @staticmethod
    def _remove_meta_rows_for_date(workbook, sheet_name: str, date_text: str) -> None:
        if sheet_name not in workbook.sheetnames:
            return
        sheet = workbook[sheet_name]
        for row_index in range(sheet.max_row, 1, -1):
            if str(sheet.cell(row_index, 1).value or "").strip() == date_text:
                sheet.delete_rows(row_index, 1)

    @staticmethod
    def _last_named_row(sheet, name_col: int) -> int:
        last_row = 1
        for row_index in range(2, sheet.max_row + 1):
            if str(sheet.cell(row_index, name_col).value or "").strip():
                last_row = row_index
        return last_row

    def _write_score_date_notes(
        self,
        sheet,
        *,
        date_columns: list[tuple[str, int]],
        name_col: int,
        notes_map: dict[str, list[str]],
    ) -> None:
        if not date_columns:
            return
        note_start_row = self._last_named_row(sheet, name_col) + 1
        for _, column_index in date_columns:
            sheet.cell(note_start_row, column_index, "")
            sheet.cell(note_start_row + 1, column_index, "")
            sheet.cell(note_start_row + 2, column_index, "")
        for date_text, column_index in date_columns:
            notes = notes_map.get(date_text, [])
            for note_index, note_text in enumerate(notes[:3]):
                sheet.cell(note_start_row + note_index, column_index, note_text)

    @staticmethod
    def _meta_headers_for_date(workbook, meta_sheet: str, date_text: str) -> set[str]:
        return {
            str(column_name).strip()
            for saved_date, column_name in read_sheet_meta(workbook, meta_sheet)
            if str(saved_date).strip() == date_text
        }

    def _remove_existing_value_column(self, workbook, store: Any, sheet_type: str, date_text: str) -> None:
        helpers = store.value_sheet_helpers_for(sheet_type)
        sheet = helpers.sheet_getter(workbook)
        meta_sheet = helpers.spec["meta_sheet"]
        day_code = date_text[5:].replace("-", "")
        matching_headers = self._meta_headers_for_date(workbook, meta_sheet, date_text)
        columns_to_delete = []
        for _, column_index in helpers.columns_getter(sheet):
            header_text = str(sheet.cell(1, column_index).value or "").strip()
            normalized_header = header_text.zfill(4) if header_text.isdigit() else header_text
            if header_text in matching_headers or normalized_header == day_code:
                columns_to_delete.append(column_index)
        for column_index in sorted(columns_to_delete, reverse=True):
            sheet.delete_cols(column_index, 1)
        self._remove_meta_rows_for_date(workbook, meta_sheet, date_text)

    def _insert_value_column(
        self,
        sheet,
        *,
        total_header: str | None,
        name_header: str,
        header_value: str,
    ) -> tuple[int, int, int | None]:
        total_col = find_column(sheet, total_header) if total_header else None
        name_col = find_column(sheet, name_header)
        if name_col is None or (total_header and total_col is None):
            raise ValueError(f"Value sheet structure is invalid: {sheet.title}")
        insert_col = total_col if total_col is not None else name_col
        sheet.insert_cols(insert_col, 1)
        sheet.cell(1, insert_col, header_value)
        updated_total_col = insert_col + 1 if total_col is not None else None
        updated_name_col = insert_col + 2 if total_col is not None else insert_col + 1
        return insert_col, updated_name_col, updated_total_col

    def _write_value_sheet(
        self,
        workbook,
        store: Any,
        *,
        sheet_type: str,
        date_text: str,
        member_names: list[str],
        values: dict[str, float],
    ) -> None:
        helpers = store.value_sheet_helpers_for(sheet_type)
        sheet = helpers.sheet_getter(workbook)
        spec = get_value_sheet_spec(sheet_type)
        day_code = date_text[5:].replace("-", "")
        self._remove_existing_value_column(workbook, store, sheet_type, date_text)
        target_col, name_col, total_col = self._insert_value_column(
            sheet,
            total_header=spec["total_header"],
            name_header=spec["name_header"],
            header_value=day_code,
        )
        row_map = ensure_member_rows_with_map(
            sheet,
            members=store.get_members(),
            name_col=name_col,
            total_col=total_col,
            value_columns=[col for _, col in helpers.columns_getter(sheet)],
        )[0]
        for name in member_names:
            sheet.cell(row_map[name], target_col, values.get(name, 0))
        self._remove_meta_rows_for_date(workbook, spec["meta_sheet"], date_text)
        append_sheet_meta(workbook, spec["meta_sheet"], date_text, day_code)

    def _write_detail_sheets(
        self,
        workbook,
        store: Any,
        rows: list[dict[str, Any]],
        *,
        member_names: list[str],
        replace_dates: set[str] | None = None,
    ) -> None:
        """Rebuild wear / income / expense sheets with cycle-block layout.

        ``rows`` and ``replace_dates`` are ignored here: the rewrite reads
        directly from SQLite so the Excel file is always a fresh projection of
        the database, with the active cycle's table at the top and historical
        cycles laid out as blocks below.
        """
        for sheet_type in ("wear", "income", "expense"):
            helpers = store.value_sheet_helpers_for(sheet_type)
            sheet = helpers.sheet_getter(workbook)
            blocks = build_cycle_blocks(
                local_db=self._local_db,
                sheet_type=sheet_type,
                member_order=member_names,
            )
            rewrite_value_sheet_with_blocks(
                workbook,
                sheet_type=sheet_type,
                sheet=sheet,
                blocks=blocks,
                member_order=member_names,
            )

    def export_missing_dates(self, store: Any, *, detail_date_text: str | None = None) -> int:
        rows = self._local_db.get_filtered_score_rows()
        notes_map = self._local_db.get_filtered_score_date_notes()

        dates = sorted({
            str(row.get("score_date", "")).strip()
            for row in rows
            if str(row.get("score_date", "")).strip()
        })
        dates = sorted({*dates, *notes_map.keys()})
        score_map: dict[tuple[str, str], int] = {
            (str(row["member_name"]).strip(), str(row["score_date"]).strip()): int(row["score"] or 0)
            for row in rows
        }
        member_names = [str(member["name"]).strip() for member in self._local_db.get_active_member_rows()]

        workbook = store.workbook_repository.open()
        try:
            # We no longer run ensure_*_sheet_structure for the value sheets
            # here: the cycle-block rewrite below handles their layout end-to-
            # end. Running ensure on them would re-introduce the legacy 2D
            # member-row dedupe and clobber historical blocks.
            score_sheet = store.score_sheet_for(workbook)
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
            replace_dates = {detail_date_text.strip()} if detail_date_text and detail_date_text.strip() else None
            self._write_detail_sheets(workbook, store, rows, member_names=member_names, replace_dates=replace_dates)
            # Value sheets (wear/income/expense) are now fully rendered by
            # ``_write_detail_sheets`` (cycle-block layout); skip ensure_*.
            store.score_sheet.ensure_structure(workbook)
            score_sheet = store.score_sheet_for(workbook)
            total_col = find_column(score_sheet, TOTAL_HEADER)
            profit_col = find_column(score_sheet, PROFIT_HEADER)
            if total_col is not None and profit_col is not None:
                recalculate_score_profits(
                    workbook, score_sheet, profit_col,
                    store.income_sheet, store.wear_sheet, store.expense_sheet,
                    local_db=self._local_db,
                )
            name_col = find_column(score_sheet, NAME_HEADER)
            if name_col is not None:
                self._write_score_date_notes(
                    score_sheet,
                    date_columns=list(zip(dates, [col for _, col in store.score_sheet.date_columns(score_sheet)])),
                    name_col=name_col,
                    notes_map=notes_map,
                )
            store.sync_member_visibility_in_workbook(workbook)
            store.workbook_repository.save(workbook)
            return len(dates)
        finally:
            workbook.close()

    def sync_member_visibility(self, store: Any) -> None:
        workbook = store.workbook_repository.open()
        try:
            store.sync_member_visibility_in_workbook(workbook)
            store.workbook_repository.save(workbook)
        finally:
            workbook.close()

