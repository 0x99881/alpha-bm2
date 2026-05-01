from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Callable

from ..constants import ENABLED, NAME_HEADER, PROFIT_HEADER, TOTAL_HEADER
from ..presenters.score_view_formatter import build_score_sheet_view, build_score_summary_view
from ..source_metadata import parse_source_profit


class ApplicationService:
    def __init__(self, local_database, supabase_client, local_next_score_date: Callable[[], str]) -> None:
        self._local_database = local_database
        self._supabase_client = supabase_client
        self._local_next_score_date = local_next_score_date

    def _online_member_rows(self) -> list[dict]:
        if not self._supabase_client.is_configured():
            raise RuntimeError("\u7ebf\u4e0a\u6570\u636e\u5e93\u672a\u914d\u7f6e\uff0c\u4e0d\u80fd\u8bfb\u53d6\u7ebf\u4e0a\u6210\u5458\u6570\u636e\u3002")
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
            raise RuntimeError("\u7ebf\u4e0a\u6570\u636e\u5e93\u672a\u914d\u7f6e\uff0c\u4e0d\u80fd\u8bfb\u53d6\u7ebf\u4e0a\u79ef\u5206\u6570\u636e\u3002")
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

    def _latest_score_date(self, score_rows: list[dict]) -> date | None:
        live_rows = self._live_score_rows(score_rows)
        if not live_rows:
            return None
        return max(
            datetime.strptime(str(row["score_date"]).strip(), "%Y-%m-%d").date()
            for row in live_rows
        )

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
        latest = self._latest_score_date(score_rows) or datetime.now().date()
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
        return build_score_summary_view(summary)

    def _score_sheet_view(self, members: list[dict[str, str]], score_rows: list[dict]):
        window_dates = self._online_window_dates(score_rows)
        profit_map = self._score_profit_map(score_rows)
        headers = [score_date[5:] for score_date in window_dates]
        headers.extend([TOTAL_HEADER, PROFIT_HEADER, NAME_HEADER])
        score_map = self._score_map(score_rows)
        raw_rows = []
        for member in members:
            name = str(member["name"])
            date_scores = [score_map.get((name, score_date), 0) for score_date in window_dates]
            raw_rows.append([*date_scores, sum(date_scores), profit_map.get(name, 0.0), name])
        raw_rows.sort(key=lambda row: (-int(row[-3]), str(row[-1])))
        return build_score_sheet_view(headers, raw_rows)

    def _next_score_date(self, score_rows: list[dict]) -> str:
        latest = self._latest_score_date(score_rows)
        if latest is not None:
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
