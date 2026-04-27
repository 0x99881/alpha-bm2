from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Callable

from ..constants import DEFAULT_MEMBERS, DEFAULT_QUICK_SCORES, ENABLED, WEAR_ABNORMAL_THRESHOLD


TimestampFactory = Callable[[], str]
StatusNormalizer = Callable[[str], str]


class ConfigRepository:
    def __init__(self, config_path: Path) -> None:
        self.config_path = config_path

    def load(self) -> dict[str, Any]:
        if self.config_path.exists():
            return json.loads(self.config_path.read_text(encoding='utf-8'))
        return {}

    def save(self, config: dict[str, Any]) -> None:
        self.config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding='utf-8')

    def default_members(self, timestamp_factory: TimestampFactory) -> list[dict[str, str | int]]:
        now_text = timestamp_factory()
        return [
            {'name': name, 'status': ENABLED, 'note': '', 'created_at': now_text, 'disabled_at': '', 'sort_order': index}
            for index, name in enumerate(DEFAULT_MEMBERS, start=1)
        ]

    def ensure_member_config(self, timestamp_factory: TimestampFactory) -> None:
        config = self.load()
        changed = False
        members = config.get('members')
        if not isinstance(members, list):
            config['members'] = self.default_members(timestamp_factory)
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
                            'created_at': timestamp_factory(),
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
            self.save(config)

    def load_members(self, normalize_status: StatusNormalizer) -> list[dict[str, str | int]]:
        config = self.load()
        result: list[dict[str, str | int]] = []
        for item in config.get('members', []):
            if not isinstance(item, dict):
                continue
            name = str(item.get('name', '')).strip()
            if not name:
                continue
            result.append(
                {
                    'name': name,
                    'status': normalize_status(str(item.get('status') or ENABLED)),
                    'note': str(item.get('note') or ''),
                    'created_at': str(item.get('created_at') or ''),
                    'disabled_at': str(item.get('disabled_at') or ''),
                    'sort_order': int(item.get('sort_order') or 0),
                }
            )
        result.sort(key=lambda member: (member['status'] != ENABLED, member['sort_order'], member['name']))
        return result
