from __future__ import annotations

from ..constants import NAME_HEADER, PROFIT_HEADER, TOTAL_HEADER
from ..constants import WINDOW_SIZE


class ScorePresenter:
    def __init__(self, store) -> None:
        self.store = store

    def _header_kind(self, value):
        if value == TOTAL_HEADER:
            return 'total'
        if value == PROFIT_HEADER:
            return 'profit'
        if value == NAME_HEADER:
            return 'name'
        return 'date'

    def build_score_summary(self):
        summary = self.store.get_score_summary_data()
        return {
            'latest_column': summary['latest_column'],
            'window_size': min(WINDOW_SIZE, summary['column_count']),
            'rankings': summary['rankings'][:999],
        }

    def build_score_sheet_view(self):
        snapshot = self.store.get_score_sheet_snapshot()
        headers = [
            {
                'value': value,
                'kind': self._header_kind(value),
            }
            for value in snapshot['headers']
        ]
        return {
            'headers': headers,
            'rows': [
                [
                    {
                        'value': value,
                        'kind': headers[index]['kind'],
                    }
                    for index, value in enumerate(row)
                ]
                for row in snapshot['raw_rows']
            ],
            'row_count': len(snapshot['raw_rows']),
            'column_count': len(headers),
        }
