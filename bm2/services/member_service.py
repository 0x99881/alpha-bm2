from __future__ import annotations

from datetime import datetime

from ..constants import DISABLED, ENABLED, normalize_status
from ..ui_text import MESSAGES


class MemberService:
    def __init__(self, members_provider, config_repository, sync_members_to_workbook, delete_member_from_workbook) -> None:
        self._members_provider = members_provider
        self._config_repository = config_repository
        self._sync_members_to_workbook = sync_members_to_workbook
        self._delete_member_from_workbook = delete_member_from_workbook

    def _timestamp(self) -> str:
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    def get_active_members(self) -> list[dict[str, str]]:
        return [item for item in self._members_provider() if item['status'] == ENABLED]

    def add_member(self, name: str, note: str = '') -> None:
        cleaned_name = name.strip()
        if not cleaned_name:
            raise ValueError(MESSAGES['member_name_required'])
        members = self._members_provider()
        if any(item['name'] == cleaned_name for item in members):
            raise ValueError(MESSAGES['member_exists'])

        config = self._config_repository.load()
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
        self._config_repository.save(config)
        self._sync_members_to_workbook()

    def update_member(self, name: str, note: str, status: str | None = None) -> None:
        config = self._config_repository.load()
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
                previous_status = normalize_status(str(item.get('status') or ENABLED))
                item['status'] = status
                if status == DISABLED and previous_status != DISABLED:
                    item['disabled_at'] = self._timestamp()
                if status == ENABLED:
                    item['disabled_at'] = ''
            break
        if not found:
            raise ValueError(MESSAGES['member_missing'])

        config['members'] = members
        self._config_repository.save(config)
        self._sync_members_to_workbook()

    def reorder_active_members(self, ordered_names: list[str]) -> None:
        current_members = self._members_provider()
        active_members = [item for item in current_members if item['status'] == ENABLED]
        disabled_members = [item for item in current_members if item['status'] == DISABLED]
        active_names = [item['name'] for item in active_members]
        unique_requested = []
        for name in ordered_names:
            if name in active_names and name not in unique_requested:
                unique_requested.append(name)
        final_active_names = unique_requested + [name for name in active_names if name not in unique_requested]
        sort_map = {}
        for index, member_name in enumerate(final_active_names + [item['name'] for item in disabled_members], start=1):
            sort_map[member_name] = index

        config = self._config_repository.load()
        members = config.get('members', [])
        for item in members:
            if not isinstance(item, dict):
                continue
            member_name = str(item.get('name', '')).strip()
            if member_name in sort_map:
                item['sort_order'] = sort_map[member_name]
        config['members'] = members
        self._config_repository.save(config)

    def delete_member(self, name: str) -> None:
        config = self._config_repository.load()
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
        self._config_repository.save(config)
        self._delete_member_from_workbook(name)
