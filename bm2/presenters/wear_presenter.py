from __future__ import annotations

from ..constants import WEAR_NAME_HEADER, WEAR_TOTAL_HEADER
from ..excel.value_normalizer import normalize_wear
from ..ui_text import UI_TEXT
from ..value_utils import build_threshold_abnormal_flags, to_float_or_none


class WearPresenter:
    def __init__(self, repository) -> None:
        self.repository = repository

    def _column_kind(self, header, col_index: int, wear_col_indices: list[int]) -> str:
        if col_index in wear_col_indices:
            return 'date'
        if header == WEAR_NAME_HEADER:
            return 'name'
        if header in (WEAR_TOTAL_HEADER, UI_TEXT['wear_member_avg']):
            return 'total'
        return 'date'

    def _cell(self, *, value, is_abnormal: bool, kind: str) -> dict:
        return {'value': value, 'is_abnormal': is_abnormal, 'kind': kind}

    def _daily_average_row(self, headers: list, raw_rows: list, wear_col_indices: list[int], threshold: float) -> list[dict]:
        label = UI_TEXT['wear_daily_member_avg']
        name_index = headers.index(WEAR_NAME_HEADER) if WEAR_NAME_HEADER in headers else None
        data_rows = [
            row
            for row in raw_rows
            if name_index is None or str(row[name_index] or '').strip() != label
        ]
        daily_averages: dict[int, float] = {}
        for col_index in wear_col_indices:
            values = [
                numeric_value
                for row in data_rows
                if (numeric_value := to_float_or_none(row[col_index])) is not None
            ]
            daily_averages[col_index] = normalize_wear(sum(values) / len(values)) if values else 0.0
        entered_daily_values = [
            value
            for col_index, value in daily_averages.items()
            if any(to_float_or_none(row[col_index]) is not None for row in data_rows)
        ]
        total_average = normalize_wear(sum(daily_averages.values()))
        overall_average = normalize_wear(sum(entered_daily_values) / len(entered_daily_values)) if entered_daily_values else 0.0
        row = []
        for col_index, header in enumerate(headers):
            kind = self._column_kind(header, col_index, wear_col_indices)
            if col_index in daily_averages:
                value = daily_averages[col_index]
                row.append(self._cell(value=value, is_abnormal=value > threshold, kind=kind))
            elif header == WEAR_TOTAL_HEADER:
                row.append(self._cell(value=total_average, is_abnormal=False, kind=kind))
            elif header == WEAR_NAME_HEADER:
                row.append(self._cell(value=label, is_abnormal=False, kind=kind))
            elif header == UI_TEXT['wear_member_avg']:
                row.append(self._cell(value=overall_average, is_abnormal=False, kind=kind))
            else:
                row.append(self._cell(value='', is_abnormal=False, kind=kind))
        return row

    def build_wear_sheet_view(self):
        snapshot = self.repository.get_wear_sheet_snapshot()
        headers = list(snapshot['headers'])
        headers.append(UI_TEXT['wear_member_avg'])
        raw_rows = snapshot['raw_rows']
        wear_col_indices = snapshot['wear_col_indices']
        column_kinds = [
            self._column_kind(header, col_index, wear_col_indices)
            for col_index, header in enumerate(headers)
        ]
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
                self._cell(
                    value=value,
                    is_abnormal=abnormal_flags.get((row_index, col_index), False),
                    kind=column_kinds[col_index],
                )
                for col_index, value in enumerate(row)
            ]
            row_wear_values = [
                numeric_value
                for col_index in wear_col_indices
                if (numeric_value := to_float_or_none(row[col_index])) is not None
            ]
            row_cells.append(
                self._cell(
                    value=normalize_wear(sum(row_wear_values) / len(row_wear_values)) if row_wear_values else 0.0,
                    is_abnormal=False,
                    kind='total',
                )
            )
            rows.append(row_cells)
        return {
            'headers': headers,
            'column_kinds': column_kinds,
            'rows': rows,
            'daily_average_row': self._daily_average_row(headers, raw_rows, wear_col_indices, threshold),
            'row_count': len(rows),
            'column_count': len(headers),
            'avg_daily_wear': normalize_wear(sum(wear_values) / len(wear_values)) if wear_values else 0.0,
            'abnormal_threshold': threshold,
            'abnormal_count': abnormal_count,
        }
