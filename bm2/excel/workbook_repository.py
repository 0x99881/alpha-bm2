from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

from openpyxl import load_workbook

from ..constants import (
    DATA_FILE_PATTERNS,
    WORKBOOK_FILENAME_PREFIX,
)
from ..ui_text import MESSAGES


class WorkbookRepository:
    def __init__(self, base_dir: Path, config_repository, *, read_only: bool = False) -> None:
        self._base_dir = Path(base_dir)
        self._config_repository = config_repository
        self._read_only = read_only
        self.workbook_path = self._resolve_workbook_path()

    def _resolve_workbook_path(self) -> Path:
        config = self._config_repository.load()
        configured = config.get("excel_filename")
        if configured:
            candidate = self._base_dir / configured
            if candidate.exists():
                return candidate

        existing_files: list[Path] = []
        for pattern in DATA_FILE_PATTERNS:
            existing_files.extend(self._base_dir.glob(pattern))
        existing_files = sorted(set(existing_files))
        if existing_files:
            target = self._best_existing_workbook(existing_files, str(configured or ""))
        else:
            target = self._base_dir / f"{WORKBOOK_FILENAME_PREFIX}{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        config["excel_filename"] = target.name
        self._config_repository.save(config)
        return target

    @staticmethod
    def _date_from_filename(filename: str) -> datetime | None:
        match = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
        if not match:
            return None
        try:
            return datetime.strptime(match.group(1), "%Y-%m-%d")
        except ValueError:
            return None

    def _best_existing_workbook(self, existing_files: list[Path], configured_name: str) -> Path:
        configured_date = self._date_from_filename(configured_name)
        if configured_date is not None:
            same_date_files = [
                path
                for path in existing_files
                if self._date_from_filename(path.name) == configured_date
            ]
            if same_date_files:
                return sorted(same_date_files)[-1]

        dated_files = [
            (self._date_from_filename(path.name), path)
            for path in existing_files
        ]
        dated_files = [(date_value, path) for date_value, path in dated_files if date_value is not None]
        if dated_files:
            return max(dated_files, key=lambda item: (item[0], item[1].name))[1]
        return existing_files[-1]

    def open(self):
        return load_workbook(self.workbook_path)

    def save(self, workbook) -> None:
        try:
            workbook.save(self.workbook_path)
        except PermissionError as exc:
            raise ValueError(MESSAGES["excel_busy"].format(filename=self.workbook_path.name)) from exc
        except OSError as exc:
            raise ValueError(MESSAGES["excel_save_failed"].format(filename=self.workbook_path.name)) from exc

    def update_workbook(self, change_func) -> None:
        workbook = self.open()
        try:
            change_func(workbook)
            self.save(workbook)
        finally:
            workbook.close()
