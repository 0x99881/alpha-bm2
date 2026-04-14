from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .constants import (
    DATA_FILE_PATTERNS,
    DEFAULT_MEMBERS,
    DEFAULT_QUICK_SCORES,
    DISABLED,
    DISABLED_ALIASES,
    ENABLED,
    EXPENSE_NAME_HEADER,
    INCOME_NAME_HEADER,
    NAME_HEADER,
    WEAR_ABNORMAL_THRESHOLD,
    WEAR_NAME_HEADER,
    WORKBOOK_FILENAME_PREFIX,
)
from .ui_text import MESSAGES


class StoreBaseMixin:
    def __init__(self, base_dir: Path) -> None:
        self.base_dir = base_dir
        self.config_path = base_dir / 'system_config.json'
        self._ensure_member_config()
        self.workbook_path = self._resolve_workbook_path()
        self._ensure_workbook()
        workbook = self._open_workbook()
        self._sync_member_visibility_in_workbook(workbook)
        self._save_workbook(workbook)
        workbook.close()

    def _timestamp(self) -> str:
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    def _normalize_status(self, raw: str) -> str:
        if raw in DISABLED_ALIASES:
            return DISABLED
        return ENABLED

    def _load_config(self) -> dict[str, Any]:
        if self.config_path.exists():
            return json.loads(self.config_path.read_text(encoding='utf-8'))
        return {}

    def _save_config(self, config: dict[str, Any]) -> None:
        self.config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding='utf-8')

    def _default_members(self) -> list[dict[str, str]]:
        now_text = self._timestamp()
        return [
            {'name': name, 'status': ENABLED, 'note': '', 'created_at': now_text, 'disabled_at': '', 'sort_order': index}
            for index, name in enumerate(DEFAULT_MEMBERS, start=1)
        ]

    def _ensure_member_config(self) -> None:
        config = self._load_config()
        changed = False
        members = config.get('members')
        if not isinstance(members, list):
            config['members'] = self._default_members()
            changed = True
        else:
            next_sort_order = 1
            known_names = {str(item.get('name', '')).strip() for item in members if isinstance(item, dict)}
            for item in members:
                if not isinstance(item, dict):
                    continue
                original_sort_order = item.get('sort_order')
                try:
                    item['sort_order'] = int(original_sort_order or next_sort_order)
                except (TypeError, ValueError):
                    item['sort_order'] = next_sort_order
                if item.get('sort_order') != original_sort_order:
                    changed = True
                next_sort_order = max(next_sort_order, int(item['sort_order']) + 1)
            for name in DEFAULT_MEMBERS:
                if name not in known_names:
                    members.append(
                        {
                            'name': name,
                            'status': ENABLED,
                            'note': '',
                            'created_at': self._timestamp(),
                            'disabled_at': '',
                            'sort_order': next_sort_order,
                        }
                    )
                    next_sort_order += 1
                    changed = True
            config['members'] = members
        if 'quick_scores' not in config:
            config['quick_scores'] = DEFAULT_QUICK_SCORES
            changed = True
        if 'wear_abnormal_threshold' not in config:
            config['wear_abnormal_threshold'] = WEAR_ABNORMAL_THRESHOLD
            changed = True
        if changed:
            self._save_config(config)

    def get_members(self) -> list[dict[str, str]]:
        self._ensure_member_config()
        config = self._load_config()
        result: list[dict[str, str]] = []
        for item in config.get('members', []):
            if not isinstance(item, dict):
                continue
            name = str(item.get('name', '')).strip()
            if not name:
                continue
            result.append(
                {
                    'name': name,
                    'status': self._normalize_status(str(item.get('status') or ENABLED)),
                    'note': str(item.get('note') or ''),
                    'created_at': str(item.get('created_at') or ''),
                    'disabled_at': str(item.get('disabled_at') or ''),
                    'sort_order': int(item.get('sort_order') or 0),
                }
            )
        result.sort(key=lambda item: (item['status'] != ENABLED, item['sort_order'], item['name']))
        return result

    def get_member(self, name: str) -> dict[str, str] | None:
        for member in self.get_members():
            if member['name'] == name:
                return member
        return None

    def get_active_members(self) -> list[dict[str, str]]:
        return [item for item in self.get_members() if item['status'] == ENABLED]

    def _sync_member_visibility_in_workbook(self, workbook) -> None:
        self._ensure_score_sheet_structure(workbook)
        self._ensure_wear_sheet_structure(workbook)
        self._ensure_income_sheet_structure(workbook)
        self._ensure_expense_sheet_structure(workbook)

        score_sheet = self._score_sheet(workbook)
        wear_sheet = self._wear_sheet(workbook)
        income_sheet = self._income_sheet(workbook)
        expense_sheet = self._expense_sheet(workbook)
        score_name_col = self._find_column(score_sheet, NAME_HEADER)
        wear_name_col = self._find_column(wear_sheet, WEAR_NAME_HEADER)
        income_name_col = self._find_column(income_sheet, INCOME_NAME_HEADER)
        expense_name_col = self._find_column(expense_sheet, EXPENSE_NAME_HEADER)

        for member in self.get_members():
            should_hide = member['status'] == DISABLED
            if score_name_col is not None:
                self._set_member_row_hidden(score_sheet, score_name_col, member['name'], should_hide)
            if wear_name_col is not None:
                self._set_member_row_hidden(wear_sheet, wear_name_col, member['name'], should_hide)
            if income_name_col is not None:
                self._set_member_row_hidden(income_sheet, income_name_col, member['name'], should_hide)
            if expense_name_col is not None:
                self._set_member_row_hidden(expense_sheet, expense_name_col, member['name'], should_hide)


    def _delete_member_from_workbook(self, workbook, member_name: str) -> None:
        score_sheet = self._score_sheet(workbook)
        wear_sheet = self._wear_sheet(workbook)
        income_sheet = self._income_sheet(workbook)
        expense_sheet = self._expense_sheet(workbook)
        score_name_col = self._find_column(score_sheet, NAME_HEADER)
        wear_name_col = self._find_column(wear_sheet, WEAR_NAME_HEADER)
        income_name_col = self._find_column(income_sheet, INCOME_NAME_HEADER)
        expense_name_col = self._find_column(expense_sheet, EXPENSE_NAME_HEADER)

        if expense_name_col is not None:
            self._delete_member_row(expense_sheet, expense_name_col, member_name)
        if income_name_col is not None:
            self._delete_member_row(income_sheet, income_name_col, member_name)
        if wear_name_col is not None:
            self._delete_member_row(wear_sheet, wear_name_col, member_name)

        if score_name_col is not None:
            self._delete_member_row(score_sheet, score_name_col, member_name)

    def add_member(self, name: str, note: str = '') -> None:
        cleaned_name = name.strip()
        if not cleaned_name:
            raise ValueError(MESSAGES['member_name_required'])
        members = self.get_members()
        if any(item['name'] == cleaned_name for item in members):
            raise ValueError(MESSAGES['member_exists'])

        config = self._load_config()
        config.setdefault('members', members)
        config['members'].append(
            {
                'name': cleaned_name,
                'status': ENABLED,
                'note': note.strip(),
                'created_at': self._timestamp(),
                'disabled_at': '',
                'sort_order': max((item['sort_order'] for item in members), default=0) + 1,
            }
        )
        self._save_config(config)

        workbook = self._open_workbook()
        self._sync_member_visibility_in_workbook(workbook)
        self._save_workbook(workbook)
        workbook.close()

    def update_member(self, name: str, note: str, status: str | None = None) -> None:
        config = self._load_config()
        members = config.get('members', [])
        found = False
        for item in members:
            if not isinstance(item, dict):
                continue
            if str(item.get('name', '')).strip() != name:
                continue
            found = True
            item['note'] = note.strip()
            if status in {ENABLED, DISABLED}:
                previous_status = self._normalize_status(str(item.get('status') or ENABLED))
                item['status'] = status
                if status == DISABLED and previous_status != DISABLED:
                    item['disabled_at'] = self._timestamp()
                if status == ENABLED:
                    item['disabled_at'] = ''
            break
        if not found:
            raise ValueError(MESSAGES['member_missing'])

        config['members'] = members
        self._save_config(config)

        workbook = self._open_workbook()
        self._sync_member_visibility_in_workbook(workbook)
        self._save_workbook(workbook)
        workbook.close()

    def reorder_active_members(self, ordered_names: list[str]) -> None:
        current_members = self.get_members()
        active_members = [item for item in current_members if item['status'] == ENABLED]
        disabled_members = [item for item in current_members if item['status'] == DISABLED]
        active_names = [item['name'] for item in active_members]
        unique_requested = []
        for name in ordered_names:
            if name in active_names and name not in unique_requested:
                unique_requested.append(name)
        final_active_names = unique_requested + [name for name in active_names if name not in unique_requested]
        sort_map = {}
        for index, name in enumerate(final_active_names + [item['name'] for item in disabled_members], start=1):
            sort_map[name] = index

        config = self._load_config()
        members = config.get('members', [])
        for item in members:
            if not isinstance(item, dict):
                continue
            name = str(item.get('name', '')).strip()
            if name in sort_map:
                item['sort_order'] = sort_map[name]
        config['members'] = members
        self._save_config(config)

    def delete_member(self, name: str) -> None:
        config = self._load_config()
        members = config.get('members', [])
        remaining_members = []
        found = False
        for item in members:
            if not isinstance(item, dict):
                continue
            if str(item.get('name', '')).strip() == name:
                found = True
                continue
            remaining_members.append(item)
        if not found:
            raise ValueError(MESSAGES['member_missing'])

        config['members'] = remaining_members
        self._save_config(config)

        workbook = self._open_workbook()
        self._ensure_score_sheet_structure(workbook)
        self._ensure_wear_sheet_structure(workbook)
        self._ensure_income_sheet_structure(workbook)
        self._ensure_expense_sheet_structure(workbook)
        self._delete_member_from_workbook(workbook, name)
        self._save_workbook(workbook)
        workbook.close()

    def get_quick_scores(self) -> list[int]:
        config = self._load_config()
        return [int(value) for value in config.get('quick_scores', DEFAULT_QUICK_SCORES)]

    def get_default_score_date(self) -> str:
        config = self._load_config()
        default_date = str(config.get('default_score_date') or '').strip()
        if default_date:
            try:
                return datetime.strptime(default_date, '%Y-%m-%d').strftime('%Y-%m-%d')
            except ValueError:
                pass
        return datetime.now().strftime('%Y-%m-%d')

    def get_wear_abnormal_threshold(self) -> float:
        override_value = getattr(self, '_wear_abnormal_threshold_override', None)
        if override_value is not None:
            return float(override_value)
        config = self._load_config()
        raw_value = config.get('wear_abnormal_threshold', WEAR_ABNORMAL_THRESHOLD)
        try:
            return float(raw_value)
        except (TypeError, ValueError):
            return float(WEAR_ABNORMAL_THRESHOLD)

    def set_wear_abnormal_threshold(self, value_text: str) -> float:
        threshold = float(value_text.strip())
        self._wear_abnormal_threshold_override = threshold
        workbook = self._open_workbook()
        try:
            self._ensure_wear_sheet_structure(workbook)
            self._save_workbook(workbook)
        finally:
            workbook.close()
            if hasattr(self, '_wear_abnormal_threshold_override'):
                delattr(self, '_wear_abnormal_threshold_override')

        config = self._load_config()
        config['wear_abnormal_threshold'] = threshold
        self._save_config(config)
        return threshold

    def set_default_score_date(self, date_text: str) -> str:
        saved_date = datetime.strptime(date_text.strip(), '%Y-%m-%d').strftime('%Y-%m-%d')
        config = self._load_config()
        config['default_score_date'] = saved_date
        self._save_config(config)
        return saved_date

    def _resolve_workbook_path(self) -> Path:
        config = self._load_config()
        configured = config.get('excel_filename')
        if configured:
            candidate = self.base_dir / configured
            if candidate.exists():
                return candidate

        existing_files: list[Path] = []
        for pattern in DATA_FILE_PATTERNS:
            existing_files.extend(self.base_dir.glob(pattern))
        existing_files = sorted(set(existing_files))
        target = existing_files[0] if existing_files else self.base_dir / f"{WORKBOOK_FILENAME_PREFIX}{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        config['excel_filename'] = target.name
        self._save_config(config)
        return target

    def _open_workbook(self):
        return load_workbook(self.workbook_path)

    def _save_workbook(self, workbook) -> None:
        try:
            workbook.save(self.workbook_path)
        except PermissionError as exc:
            raise ValueError(MESSAGES['excel_busy'].format(filename=self.workbook_path.name)) from exc
