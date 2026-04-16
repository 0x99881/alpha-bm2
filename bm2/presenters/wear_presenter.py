from __future__ import annotations

from ..ui_text import UI_TEXT
from ..value_utils import build_threshold_abnormal_flags, round_wear, to_float_or_none


class WearPresenter:
    def __init__(self, repository) -> None:
        self.repository = repository

    def build_wear_sheet_view(self):
        snapshot = self.repository.get_wear_sheet_snapshot()
        headers = list(snapshot['headers'])
        headers.append(UI_TEXT['wear_member_avg'])
        raw_rows = snapshot['raw_rows']
        wear_col_indices = snapshot['wear_col_indices']
        threshold = snapshot['threshold']
        abnormal_flags = build_threshold_abnormal_flags(raw_rows, wear_col_indices, threshold)
        wear_values = [
            numeric_value
            for row in raw_rows
            for col_index in wear_col_indices
            if (numeric_value := to_float_or_none(row[col_index])) is not None
        ]
        abnormal_count = sum(1 for value in wear_values if value > threshold)
        rows = []
        for row_index, row in enumerate(raw_rows):
            row_cells = [
                {
                    'value': value,
                    'is_abnormal': abnormal_flags.get((row_index, col_index), False),
                }
                for col_index, value in enumerate(row)
            ]
            row_wear_values = [
                numeric_value
                for col_index in wear_col_indices
                if (numeric_value := to_float_or_none(row[col_index])) is not None
            ]
            row_cells.append(
                {
                    'value': round_wear(sum(row_wear_values) / len(row_wear_values)) if row_wear_values else 0.0,
                    'is_abnormal': False,
                }
            )
            rows.append(row_cells)
        return {
            'headers': headers,
            'rows': rows,
            'row_count': len(rows),
            'column_count': len(headers),
            'avg_daily_wear': round_wear(sum(wear_values) / len(wear_values)) if wear_values else 0.0,
            'abnormal_threshold': threshold,
            'abnormal_count': abnormal_count,
        }
