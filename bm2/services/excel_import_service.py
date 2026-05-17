from __future__ import annotations

from datetime import date, datetime
from typing import Callable

from ..constants import DATA_START_ROW, META_SHEET, NAME_HEADER, TOTAL_HEADER, WORKBOOK_FILENAME_PREFIX
from ..excel.header_locator import find_column
from ..excel.sheet_metadata import read_sheet_meta
from ..excel.value_normalizer import normalize_expense, normalize_income, normalize_wear
from ..excel.value_sheet_spec import get_value_sheet_spec
from ..ui_text import MESSAGES


class ExcelImportService:
    def __init__(
        self,
        local_database,
        excel_store,
        config_members_provider: Callable[[], list[dict[str, str]]],
        excel_exporter,
    ) -> None:
        self._local_database = local_database
        self._excel_store = excel_store
        self._config_members_provider = config_members_provider
        self._excel_exporter = excel_exporter

    def bootstrap_from_excel_if_empty(self) -> bool:
        if self._local_database.get_member_rows(include_deleted=True):
            return False
        if self._local_database.get_score_rows(include_deleted=True):
            return False
        self.refresh_from_excel()
        return True

    def refresh_from_excel(self) -> None:
        score_rows, date_notes = self._score_rows_and_notes_from_workbook()
        self._seed_members_if_empty()
        score_rows = self._known_member_score_rows(score_rows)
        self._local_database.replace_score_entries_from_snapshot(score_rows)
        self._sync_score_date_notes(date_notes)
        # ``score_cycle_start_date`` sync_state was the per-cycle-file marker
        # from the old workflow; the rolling-window filter now derives the
        # boundary directly from score entries so we deliberately stop
        # writing it back to keep the two sources in sync.

    def _seed_members_if_empty(self) -> None:
        if self._local_database.get_member_rows(include_deleted=True):
            return
        self._local_database.sync_members(self._config_members_provider())

    def _known_member_score_rows(self, rows: list[dict[str, object]]) -> list[dict[str, object]]:
        member_names = {
            str(member.get("name", "")).strip()
            for member in self._local_database.get_member_rows()
        }
        if not member_names:
            return []
        return [
            row
            for row in rows
            if str(row.get("member_name", "")).strip() in member_names
        ]

    def after_supabase_pull(self) -> None:
        self._excel_exporter.export_missing_dates(self._excel_store)

    @staticmethod
    def _parse_score_header_date(header_text: str, year: int) -> date | None:
        try:
            return datetime.strptime(f"{year:04d}-{header_text}", "%Y-%m-%d").date()
        except ValueError:
            return None

    def _score_header_date_map(self, workbook, used_columns: list[tuple[int, int]]) -> dict[int, str]:
        score_sheet = self._excel_store.score_sheet_for(workbook)
        cycle_start = self._workbook_cycle_start()
        if cycle_start:
            year_hint = datetime.strptime(cycle_start, "%Y-%m-%d").year
        else:
            year_hint = datetime.now().year

        result: dict[int, str] = {}
        previous: date | None = None
        for _, column_index in used_columns:
            header_text = str(score_sheet.cell(1, column_index).value or "").strip()
            if not self._excel_store.score_sheet.is_date_header(header_text):
                continue
            current = self._parse_score_header_date(header_text, year_hint)
            if current is None:
                continue
            while previous is not None and current <= previous:
                current = date(current.year + 1, current.month, current.day)
            result[column_index] = current.strftime("%Y-%m-%d")
            previous = current
        return result

    def _score_date_map_from_workbook(self, workbook) -> dict[int, str]:
        score_sheet = self._excel_store.score_sheet_for(workbook)
        meta_by_number: dict[int, list[str]] = {}
        for saved_date, column_name in read_sheet_meta(workbook, META_SHEET):
            column_name = str(column_name).strip()
            if not column_name.startswith("D"):
                continue
            suffix = column_name[1:]
            if suffix.isdigit():
                meta_by_number.setdefault(int(suffix), []).append(str(saved_date).strip())

        used_columns = [
            (number, col)
            for number, col in self._excel_store.score_sheet.date_columns(score_sheet)
            if number in meta_by_number or self._excel_store.score_sheet.column_has_data(score_sheet, col)
        ]
        if not used_columns:
            return {}

        score_date_map: dict[int, str] = {}
        unresolved: list[tuple[int, int]] = []
        header_date_map = self._score_header_date_map(workbook, used_columns)
        for number, column_index in used_columns:
            saved_dates = {item.strip() for item in meta_by_number.get(number, []) if item.strip()}
            saved_date = next(iter(saved_dates)) if len(saved_dates) == 1 else ""
            header_date = header_date_map.get(column_index, "")
            header_text = str(score_sheet.cell(1, column_index).value or "").strip()
            if saved_date and (not header_text or saved_date[5:] == header_text):
                score_date_map[column_index] = saved_date
            elif header_date:
                score_date_map[column_index] = header_date
            else:
                unresolved.append((number, column_index))

        if unresolved:
            columns = ", ".join(f"D{number}" for number, _ in unresolved)
            raise ValueError(MESSAGES["excel_score_date_unresolved"].format(columns=columns))
        return score_date_map

    @staticmethod
    def _note_text(value) -> str:
        return str(value or "").strip()

    def _last_member_row(self, sheet, name_col: int) -> int:
        last_row = DATA_START_ROW - 1
        for row_index in range(DATA_START_ROW, sheet.max_row + 1):
            if self._note_text(sheet.cell(row_index, name_col).value):
                last_row = row_index
        return last_row

    def _score_date_notes_from_sheet(
        self,
        sheet,
        *,
        score_date_map: dict[int, str],
        name_col: int,
    ) -> dict[str, dict[str, str]]:
        note_start_row = self._last_member_row(sheet, name_col) + 1
        notes: dict[str, dict[str, str]] = {}
        for column_index, score_date in score_date_map.items():
            values = [
                self._note_text(sheet.cell(note_start_row, column_index).value),
                self._note_text(sheet.cell(note_start_row + 1, column_index).value),
                self._note_text(sheet.cell(note_start_row + 2, column_index).value),
            ]
            compacted = [value for value in values if value]
            notes[score_date] = {
                "note1": compacted[0] if len(compacted) >= 1 else "",
                "note2": compacted[1] if len(compacted) >= 2 else "",
                "note3": compacted[2] if len(compacted) >= 3 else "",
            }
        return notes

    def _sync_score_date_notes(self, date_notes: dict[str, dict[str, str]]) -> None:
        for score_date, notes in date_notes.items():
            self._local_database.record_score_date_notes(score_date, notes, source="local")

    def _score_rows_and_notes_from_workbook(self) -> tuple[list[dict[str, object]], dict[str, dict[str, str]]]:
        workbook = self._excel_store.workbook_repository.open()
        try:
            self._excel_store.ensure_score_sheet_structure(workbook)
            sheet = self._excel_store.score_sheet_for(workbook)
            name_col = find_column(sheet, NAME_HEADER)
            total_col = find_column(sheet, TOTAL_HEADER)
            if name_col is None or total_col is None:
                return [], {}

            score_date_map = self._score_date_map_from_workbook(workbook)
            date_notes = self._score_date_notes_from_sheet(
                sheet,
                score_date_map=score_date_map,
                name_col=name_col,
            )
            rows: list[dict[str, object]] = []
            for row in range(2, sheet.max_row + 1):
                member_name = str(sheet.cell(row, name_col).value or "").strip()
                if not member_name:
                    continue
                for col, score_date in score_date_map.items():
                    value = sheet.cell(row, col).value
                    try:
                        score_value = int(value or 0)
                    except (TypeError, ValueError):
                        score_value = 0
                    rows.append({"member_name": member_name, "score_date": score_date, "score": score_value})
            self._attach_value_sheet_fields(workbook, rows)
            return rows, date_notes
        finally:
            workbook.close()

    def _read_value_sheet_fields(self, workbook, sheet_type: str) -> dict[tuple[str, str], str]:
        """Read user-edited values across every cycle block on the sheet.

        Each block (active at top, historical below) has its own local
        header row with date columns, so edits in any block — even a
        long-settled one — get picked up by refresh-from-Excel.
        """
        from ..excel.cycle_block_sheets import iter_blocks

        helpers = self._excel_store.value_sheet_helpers_for(sheet_type)
        sheet = helpers.sheet_getter(workbook)
        spec = get_value_sheet_spec(sheet_type)
        normalizer = {
            "wear": normalize_wear,
            "income": normalize_income,
            "expense": normalize_expense,
        }[sheet_type]
        values: dict[tuple[str, str], str] = {}
        for block in iter_blocks(sheet, spec["name_header"], spec["total_header"]):
            for row_index in block["data_rows"]:
                member_name = str(sheet.cell(row_index, block["name_col"]).value or "").strip()
                if not member_name:
                    continue
                for col, score_date in block["date_by_col"].items():
                    raw_value = sheet.cell(row_index, col).value
                    if raw_value in (None, ""):
                        # Empty cell is treated as "user didn't touch this
                        # row x date" so DB keeps its existing value. To
                        # clear an entry to zero, type ``0`` explicitly —
                        # an empty cell isn't distinguishable from a padded
                        # day where the member just didn't participate.
                        continue
                    try:
                        numeric = normalizer(raw_value)
                    except (TypeError, ValueError):
                        continue
                    # Note: ``0`` IS surfaced. Previously the reader skipped
                    # zeros, which made ``5 → 0`` edits look like no-ops
                    # because the (member, date) pair never made it into
                    # the snapshot. Now zero rewrites land in DB just like
                    # any other edit.
                    values[(member_name, score_date)] = str(numeric)
        return values

    def _attach_value_sheet_fields(self, workbook, rows: list[dict[str, object]]) -> None:
        rows_by_key = {
            (str(row.get("member_name", "")).strip(), str(row.get("score_date", "")).strip()): row
            for row in rows
        }
        field_map = {
            "wear": "manual_wear",
            "income": "income",
            "expense": "other_expense",
        }
        for sheet_type, field_name in field_map.items():
            for key, value in self._read_value_sheet_fields(workbook, sheet_type).items():
                row = rows_by_key.get(key)
                if row is None:
                    member_name, score_date = key
                    row = {"member_name": member_name, "score_date": score_date, "score": 0}
                    rows.append(row)
                    rows_by_key[key] = row
                row[field_name] = value

    def _workbook_cycle_start(self) -> str:
        stem = self._excel_store.workbook_path.stem
        if stem.startswith(WORKBOOK_FILENAME_PREFIX):
            suffix = stem[len(WORKBOOK_FILENAME_PREFIX):]
            suffix = suffix.split("_", 1)[0].strip()
            try:
                return datetime.strptime(suffix, "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:
                return ""
        return ""

    @staticmethod
    def _score_rows_cycle_start(rows: list[dict[str, object]]) -> str:
        dates = []
        for row in rows:
            raw_date = str(row.get("score_date", "")).strip()
            if not raw_date:
                continue
            try:
                dates.append(datetime.strptime(raw_date, "%Y-%m-%d").date())
            except ValueError:
                continue
        return min(dates).strftime("%Y-%m-%d") if dates else ""
