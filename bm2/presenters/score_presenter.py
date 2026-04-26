from __future__ import annotations

from ..constants import WINDOW_SIZE


class ScorePresenter:
    def __init__(self, store) -> None:
        self.store = store

    def build_score_summary(self):
        summary = self.store.get_score_summary_data()
        return {
            'latest_column': summary['latest_column'],
            'window_size': min(WINDOW_SIZE, summary['column_count']),
            'rankings': summary['rankings'][:999],
        }
