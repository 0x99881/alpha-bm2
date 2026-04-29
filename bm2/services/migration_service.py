from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Callable

from ..constants import META_SHEET, NAME_HEADER, TOTAL_HEADER, WORKBOOK_FILENAME_PREFIX
from ..excel.sheet_metadata import read_sheet_meta


class LocalDataMigrationService:
    def __init__(
        self,
        local_database,
        legacy_store,
        legacy_members_provider: Callable[[], list[dict[str, str]]],
        excel_exporter,
    ) -> None:
        self._local_database = local_database
        self._legacy_store = legacy_store
        self._legacy_members_provider = legacy_members_provider
        self._excel_exporter = excel_exporter

    def bootstrap_from_legacy_sources(self) -> None:
        score_rows = self._score_rows_from_workbook()
        self._local_database.sync_members(self._legacy_members_provider())
        self._local_database.replace_score_entries_from_snapshot(score_rows)
        cycle_start = self._score_rows_cycle_start(score_rows) or self._workbook_cycle_start()
        if cycle_start:
            self._local_database.set_sync_state("score_cycle_start_date", cycle_start)

    def refresh_from_legacy_store(self) -> None:
        self.bootstrap_from_legacy_sources()

    def after_supabase_pull(self) -> None:
        self._excel_exporter.export_missing_dates(self._legacy_store)
        self.refresh_from_legacy_store()

    @staticmethod
    def _parse_score_header_date(header_text: str, year: int) -> date | None:
        try:
            return datetime.strptime(f"{year:04d}-{header_text}", "%Y-%m-%d").date()
        except ValueError:
            return None

    def _score_header_date_map(self, workbook, used_columns: list[tuple[int, int]]) -> dict[int, str]:
        score_sheet = self._legacy_store._score_sheet(workbook)
        cycle_start = self._workbook_cycle_start()
        if cycle_start:
            year_hint = datetime.strptime(cycle_start, "%Y-%m-%d").year
        else:
            year_hint = datetime.now().year

        result: dict[int, str] = {}
        previous: date | None = None
        for _, column_index in used_columns:
            header_text = str(score_sheet.cell(1, column_index).value or "").strip()
            if not self._legacy_store.score_sheet.is_date_header(header_text):
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
        score_sheet = self._legacy_store._score_sheet(workbook)
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
            for number, col in self._legacy_store.score_sheet.date_columns(score_sheet)
            if number in meta_by_number or self._legacy_store.score_sheet.column_has_data(score_sheet, col)
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

        if not unresolved:
            return score_date_map

        latest_date = self._legacy_store.score_sheet.latest_used_date(workbook)
        if latest_date is None:
            return score_date_map

        total_used = len(used_columns)
        fallback_by_number = {
            number: (latest_date - timedelta(days=(total_used - index - 1))).strftime("%Y-%m-%d")
            for index, (number, _) in enumerate(used_columns)
        }
        for number, column_index in unresolved:
            score_date_map[column_index] = fallback_by_number[number]
        return score_date_map

    def _score_rows_from_workbook(self) -> list[dict[str, object]]:
        workbook = self._legacy_store.workbook_repository.open()
        try:
            self._legacy_store._ensure_score_sheet_structure(workbook)
            sheet = self._legacy_store._score_sheet(workbook)
            name_col = self._legacy_store._find_column(sheet, NAME_HEADER)
            total_col = self._legacy_store._find_column(sheet, TOTAL_HEADER)
            if name_col is None or total_col is None:
                return []

            score_date_map = self._score_date_map_from_workbook(workbook)
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
            return rows
        finally:
            workbook.close()

    def _workbook_cycle_start(self) -> str:
        stem = self._legacy_store.workbook_path.stem
        if stem.startswith(WORKBOOK_FILENAME_PREFIX):
            suffix = stem[len(WORKBOOK_FILENAME_PREFIX):]
            suffix = suffix.split("_", 1)[0].strip()
            try:
                return datetime.strptime(suffix, "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:
                pass
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
