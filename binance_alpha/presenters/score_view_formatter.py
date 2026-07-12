from __future__ import annotations

from typing import Any

from ..constants import NAME_HEADER, PROFIT_HEADER, TOTAL_HEADER, WINDOW_SIZE

LOW_SCORE_THRESHOLD = 10


def score_header_kind(value: Any) -> str:
    if value == TOTAL_HEADER:
        return "total"
    if value == PROFIT_HEADER:
        return "profit"
    if value == NAME_HEADER:
        return "name"
    return "date"


def _numeric_value(value: Any) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _sorted_score_rows(headers: list[Any], raw_rows: list[list[Any]]) -> list[list[Any]]:
    try:
        total_index = headers.index(TOTAL_HEADER)
        name_index = headers.index(NAME_HEADER)
    except ValueError:
        return raw_rows

    member_rows: list[list[Any]] = []
    other_rows: list[list[Any]] = []
    for row in raw_rows:
        name = str((row[name_index] if name_index < len(row) else "") or "").strip()
        if name:
            member_rows.append(row)
        else:
            other_rows.append(row)
    member_rows.sort(
        key=lambda row: (
            -_numeric_value(row[total_index] if total_index < len(row) else 0),
            str((row[name_index] if name_index < len(row) else "") or ""),
        )
    )
    return member_rows + other_rows


def _is_low_score(value: Any, kind: str) -> bool:
    if kind != "date":
        return False
    try:
        return float(value) < LOW_SCORE_THRESHOLD
    except (TypeError, ValueError):
        return False


def build_score_summary_view(summary: dict[str, Any]) -> dict[str, Any]:
    return {
        "latest_column": summary["latest_column"],
        "window_size": min(WINDOW_SIZE, summary["column_count"]),
        "rankings": summary["rankings"][:999],
    }


def build_score_sheet_view(headers: list[Any], raw_rows: list[list[Any]]) -> dict[str, Any]:
    sorted_rows = _sorted_score_rows(headers, raw_rows)
    view_headers = [
        {
            "value": value,
            "kind": score_header_kind(value),
        }
        for value in headers
    ]
    return {
        "headers": view_headers,
        "rows": [
            [
                {
                    "value": value,
                    "kind": view_headers[index]["kind"],
                    "is_low_score": _is_low_score(value, view_headers[index]["kind"]),
                }
                for index, value in enumerate(row)
            ]
            for row in sorted_rows
        ],
        "row_count": len(sorted_rows),
        "column_count": len(view_headers),
    }
