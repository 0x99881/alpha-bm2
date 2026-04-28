from __future__ import annotations

from datetime import datetime, timedelta
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
        self._local_database.sync_members(self._legacy_members_provider())
        self._local_database.replace_score_entries_from_snapshot(self._score_rows_from_workbook())
        cycle_start = self._workbook_cycle_start()
        if cycle_start:
            self._local_database.set_sync_state("score_cycle_start_date", cycle_start)

    def refresh_from_legacy_store(self) -> None:
        self.bootstrap_from_legacy_sources()

    def after_supabase_pull(self) -> None:
        self._excel_exporter.export_missing_dates(self._legacy_store)
        self.refresh_from_legacy_store()

    def _score_date_map_from_workbook(self, workbook) -> dict[int, str]:
        score_sheet = self._legacy_store._score_sheet(workbook)
        meta_by_number: dict[int, str] = {}
        for saved_date, column_name in read_sheet_meta(workbook, META_SHEET):
            column_name = str(column_name).strip()
            if not column_name.startswith("D"):
                continue
            suffix = column_name[1:]
            if suffix.isdigit():
                meta_by_number[int(suffix)] = str(saved_date).strip()

        used_columns = [
            (number, col)
            for number, col in self._legacy_store.score_sheet.date_columns(score_sheet)
            if number in meta_by_number or self._legacy_store.score_sheet.column_has_data(score_sheet, col)
        ]
        if not used_columns:
            return {}

        score_date_map: dict[int, str] = {}
        unresolved: list[tuple[int, int]] = []
        for number, column_index in used_columns:
            saved_date = meta_by_number.get(number, "").strip()
            if saved_date:
                score_date_map[column_index] = saved_date
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
