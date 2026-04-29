from __future__ import annotations

import logging
from datetime import datetime, timedelta
from typing import Callable

from ..constants import ENABLED
from ..source_metadata import parse_source_profit


LOGGER = logging.getLogger(__name__)


class ApplicationService:
    def __init__(self, local_database, supabase_client, local_next_score_date: Callable[[], str]) -> None:
        self._local_database = local_database
        self._supabase_client = supabase_client
        self._local_next_score_date = local_next_score_date

    def _online_member_rows(self) -> list[dict]:
        if not self._supabase_client.is_configured():
            raise RuntimeError("线上数据库未配置，不能读取线上成员数据。")
        rows = self._supabase_client.pull_members()
        rows.sort(
            key=lambda item: (
                int(item.get("deleted", 0) or 0),
                int(item.get("sort_order", 0) or 0),
                str(item.get("name", "")),
            )
        )
        return rows

    def _online_score_rows(self) -> list[dict]:
        if not self._supabase_client.is_configured():
            raise RuntimeError("线上数据库未配置，不能读取线上积分数据。")
        return self._supabase_client.pull_score_entries()

    def _active_members_from_rows(self, member_rows: list[dict]) -> list[dict[str, str]]:
        rows = [
            row
            for row in member_rows
            if str(row.get("status", "")).strip() == ENABLED
            and int(row.get("deleted", 0) or 0) == 0
        ]
        return [
            {
                "name": str(item.get("name", "")),
                "status": str(item.get("status", ENABLED)),
                "note": str(item.get("note", "")),
                "created_at": str(item.get("created_at", "")),
                "disabled_at": str(item.get("disabled_at", "")),
                "sort_order": int(item.get("sort_order", 0) or 0),
            }
            for item in rows
        ]

    def _live_score_rows(self, score_rows: list[dict]) -> list[dict]:
        return [row for row in score_rows if int(row.get("deleted", 0) or 0) == 0]

    def _score_map(self, score_rows: list[dict]) -> dict[tuple[str, str], int]:
        return {
            (str(row.get("member_name", "")).strip(), str(row.get("score_date", "")).strip()): int(row.get("score", 0) or 0)
            for row in self._live_score_rows(score_rows)
        }

    def _score_profit_map(self, score_rows: list[dict]) -> dict[str, float]:
        profit_map: dict[str, float] = {}
        for row in self._live_score_rows(score_rows):
            member_name = str(row.get("member_name", "")).strip()
            if not member_name:
                continue
            profit = parse_source_profit(row.get("source", ""))
            if profit is not None:
                profit_map[member_name] = profit
        return profit_map

    def online_score_rows_for_date(self, score_date: str) -> list[dict]:
        return [
            row
            for row in self._online_score_rows()
            if str(row.get("score_date", "")).strip() == score_date
            and int(row.get("deleted", 0) or 0) == 0
        ]

    def online_active_members(self) -> list[dict[str, str]]:
        return self._active_members_from_rows(self._online_member_rows())

    def _online_window_dates(self, score_rows: list[dict]) -> list[str]:
        live_rows = [row for row in score_rows if int(row.get("deleted", 0) or 0) == 0]
        if live_rows:
            latest = max(
                datetime.strptime(str(row["score_date"]).strip(), "%Y-%m-%d").date()
                for row in live_rows
            )
        else:
            latest = datetime.now().date()
        return [(latest - timedelta(days=offset)).strftime("%Y-%m-%d") for offset in range(14, -1, -1)]

    def _score_summary_data(self, members: list[dict[str, str]], score_rows: list[dict]) -> dict:
        window_dates = self._online_window_dates(score_rows)
        score_map = self._score_map(score_rows)
        rankings = []
        for member in members:
            name = str(member["name"])
            total = sum(score_map.get((name, score_date), 0) for score_date in window_dates)
            rankings.append({"name": name, "total": total})
        rankings.sort(key=lambda item: (-int(item["total"]), str(item["name"])))
        return {
            "latest_column": window_dates[-1][5:],
            "column_count": len(window_dates),
            "rankings": rankings,
        }

    def _score_summary(self, summary: dict):
        return {
            "latest_column": summary["latest_column"],
            "window_size": min(15, summary["column_count"]),
            "rankings": summary["rankings"][:999],
        }

    def _score_sheet_view(self, members: list[dict[str, str]], score_rows: list[dict]):
        window_dates = self._online_window_dates(score_rows)
        profit_map = self._score_profit_map(score_rows)
        headers = [{"value": score_date[5:], "kind": "date"} for score_date in window_dates]
        headers.extend(
            [
                {"value": "总积分", "kind": "total"},
                {"value": "累计盈亏", "kind": "profit"},
                {"value": "姓名", "kind": "name"},
            ]
        )
        score_map = self._score_map(score_rows)
        raw_rows = []
        for member in members:
            name = str(member["name"])
            date_scores = [score_map.get((name, score_date), 0) for score_date in window_dates]
            raw_rows.append([*date_scores, sum(date_scores), profit_map.get(name, 0.0), name])
        raw_rows.sort(key=lambda row: (-int(row[-3]), str(row[-1])))
        return {
            "headers": headers,
            "rows": [
                [
                    {"value": value, "kind": headers[index]["kind"]}
                    for index, value in enumerate(row)
                ]
                for row in raw_rows
            ],
            "row_count": len(raw_rows),
            "column_count": len(headers),
        }

    def _next_score_date(self, score_rows: list[dict]) -> str:
        score_rows = self._live_score_rows(score_rows)
        if score_rows:
            latest = max(
                datetime.strptime(str(row["score_date"]).strip(), "%Y-%m-%d").date()
                for row in score_rows
            )
            return (latest + timedelta(days=1)).strftime("%Y-%m-%d")
        return self._local_next_score_date()

    def online_score_summary_data(self) -> dict:
        return self._score_summary_data(self.online_active_members(), self._online_score_rows())

    def online_score_summary(self):
        return self._score_summary(self.online_score_summary_data())

    def online_score_sheet_view(self):
        return self._score_sheet_view(self.online_active_members(), self._online_score_rows())

    def online_next_score_date(self) -> str:
        return self._next_score_date(self._online_score_rows())

    def mobile_overview(self) -> dict:
        members = self._active_members_from_rows(self._online_member_rows())
        score_rows = self._online_score_rows()
        summary_data = self._score_summary_data(members, score_rows)
        return {
            "score_sheet_view": self._score_sheet_view(members, score_rows),
            "score_summary": self._score_summary(summary_data),
            "active_members": members,
            "next_score_date": self._next_score_date(score_rows),
        }
