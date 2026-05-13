from __future__ import annotations

from datetime import datetime, timedelta
import logging
from pathlib import Path
import re

from openpyxl import Workbook, load_workbook

from ..constants import (
    DATA_FILE_PATTERNS,
    NAME_HEADER,
    SCORE_SHEET,
    TOTAL_HEADER,
    WINDOW_SIZE,
    WORKBOOK_FILENAME_PREFIX,
)
from ..ui_text import MESSAGES

LOGGER = logging.getLogger(__name__)


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

    def create_new_cycle(self, structure_owner, start_date_text: str) -> str:
        raw_text = (start_date_text or "").strip()
        if not raw_text:
            raise ValueError(MESSAGES["new_cycle_date_required"])
        try:
            start_date = datetime.strptime(raw_text, "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError(MESSAGES["new_cycle_date_invalid"]) from exc

        date_text = start_date.strftime("%Y-%m-%d")
        new_path = self._base_dir / f"{WORKBOOK_FILENAME_PREFIX}{date_text}.xlsx"
        counter = 2
        while new_path.exists():
            new_path = self._base_dir / f"{WORKBOOK_FILENAME_PREFIX}{date_text}_{counter}.xlsx"
            counter += 1

        workbook = Workbook()
        score_sheet = workbook.active
        score_sheet.title = SCORE_SHEET
        for column_index in range(1, WINDOW_SIZE + 1):
            header_date = start_date - timedelta(days=(WINDOW_SIZE - column_index))
            score_sheet.cell(1, column_index, header_date.strftime("%m-%d"))
        score_sheet.cell(1, WINDOW_SIZE + 1, TOTAL_HEADER)
        score_sheet.cell(1, WINDOW_SIZE + 2, NAME_HEADER)

        try:
            workbook.save(new_path)
        except PermissionError as exc:
            raise ValueError(MESSAGES["excel_busy"].format(filename=new_path.name)) from exc
        except OSError as exc:
            raise ValueError(MESSAGES["excel_save_failed"].format(filename=new_path.name)) from exc
        finally:
            workbook.close()

        previous_path = self.workbook_path
        self.workbook_path = new_path
        try:
            structure_owner.workbook_path = new_path
            structure_owner.ensure_workbook()
        except (OSError, ValueError, KeyError, RuntimeError):
            self.workbook_path = previous_path
            structure_owner.workbook_path = previous_path
            try:
                new_path.unlink()
            except OSError as cleanup_error:
                LOGGER.warning("Failed to remove incomplete workbook %s: %s", new_path, cleanup_error)
            raise

        config = self._config_repository.load()
        config["excel_filename"] = new_path.name
        config["default_score_date"] = date_text
        self._config_repository.save(config)
        return new_path.name
