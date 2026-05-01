from __future__ import annotations

from typing import Any

from ..constants import NAME_HEADER, PROFIT_HEADER, TOTAL_HEADER, WINDOW_SIZE


def score_header_kind(value: Any) -> str:
    if value == TOTAL_HEADER:
        return "total"
    if value == PROFIT_HEADER:
        return "profit"
    if value == NAME_HEADER:
        return "name"
    return "date"


def build_score_summary_view(summary: dict[str, Any]) -> dict[str, Any]:
    return {
        "latest_column": summary["latest_column"],
        "window_size": min(WINDOW_SIZE, summary["column_count"]),
        "rankings": summary["rankings"][:999],
    }


def build_score_sheet_view(headers: list[Any], raw_rows: list[list[Any]]) -> dict[str, Any]:
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
                }
                for index, value in enumerate(row)
            ]
            for row in raw_rows
        ],
        "row_count": len(raw_rows),
        "column_count": len(view_headers),
    }
