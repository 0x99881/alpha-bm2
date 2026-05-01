from __future__ import annotations

from ..constants import DISABLED, ENABLED, normalize_status
from ..ui_text import MESSAGES


class MemberService:
    def __init__(
        self,
        members_provider,
        member_repository,
        after_change=None,
    ) -> None:
        self._members_provider = members_provider
        self._member_repository = member_repository
        self._after_change = after_change

    def get_active_members(self) -> list[dict[str, str]]:
        return [item for item in self._members_provider() if item['status'] == ENABLED]

    def get_members(self) -> list[dict[str, str]]:
        return self._members_provider()

    def _run_after_change(self) -> None:
        if callable(self._after_change):
            self._after_change()

    def add_member(self, name: str, note: str = '') -> None:
        cleaned_name = name.strip()
        if not cleaned_name:
            raise ValueError(MESSAGES['member_name_required'])
        members = self._members_provider()
        if any(item['name'] == cleaned_name for item in members):
            raise ValueError(MESSAGES['member_exists'])

        self._member_repository.add_member(cleaned_name, note)
        self._run_after_change()

    def update_member(self, name: str, note: str, status: str | None = None) -> None:
        normalized_status = normalize_status(str(status or ENABLED)) if status in {ENABLED, DISABLED} else status
        updated = self._member_repository.update_member(name, note, status=normalized_status)
        if not updated:
            raise ValueError(MESSAGES['member_missing'])
        self._run_after_change()

    def reorder_active_members(self, ordered_names: list[str]) -> None:
        self._member_repository.reorder_active_members(ordered_names)
        self._run_after_change()

    def delete_member(self, name: str) -> None:
        deleted = self._member_repository.delete_member(name)
        if not deleted:
            raise ValueError(MESSAGES['member_missing'])
        self._run_after_change()
