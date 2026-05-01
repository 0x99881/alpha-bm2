from __future__ import annotations


def member_id(name: str) -> str:
    return f"member:{name.strip()}"


def score_entry_id(member_name: str, score_date: str) -> str:
    return f"score:{score_date}:{member_name.strip()}"


def count_wear_entries(entries: list[dict[str, str]]) -> int:
    count = 0
    for entry in entries:
        has_manual = str(entry.get("manual_wear", "")).strip()
        has_before = str(entry.get("before_balance", "")).strip()
        has_after = str(entry.get("after_balance", "")).strip()
        if has_manual or (has_before and has_after):
            count += 1
    return count
