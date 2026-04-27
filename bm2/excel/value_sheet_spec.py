from __future__ import annotations

from typing import Any

from ..constants import VALUE_SHEET_SPECS


def get_value_sheet_spec(sheet_type: str) -> dict[str, Any]:
    return VALUE_SHEET_SPECS[sheet_type]
