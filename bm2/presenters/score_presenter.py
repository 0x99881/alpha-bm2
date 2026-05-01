from __future__ import annotations

from .score_view_formatter import build_score_sheet_view, build_score_summary_view


class ScorePresenter:
    def __init__(self, store) -> None:
        self.store = store

    def build_score_summary(self):
        return build_score_summary_view(self.store.get_score_summary_data())

    def build_score_sheet_view(self):
        snapshot = self.store.get_score_sheet_snapshot()
        return build_score_sheet_view(snapshot["headers"], snapshot["raw_rows"])
