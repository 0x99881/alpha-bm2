from __future__ import annotations

from ..constants import WINDOW_SIZE


class ScorePresenter:
    def __init__(self, repository) -> None:
        self.repository = repository

    def build_score_summary(self):
        column_count = self.repository.get_score_column_count()
        return {
            'latest_column': self.repository.get_score_latest_column(),
            'window_size': min(WINDOW_SIZE, column_count),
            'rankings': self.repository.get_score_rankings(limit=999),
        }
