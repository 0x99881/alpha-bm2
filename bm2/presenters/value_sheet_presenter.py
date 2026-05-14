from __future__ import annotations

from ..constants import VALUE_SHEET_SPECS
from ..value_utils import to_float_or_none


class ValueSheetPresenter:
    def __init__(self, store) -> None:
        self.store = store

    def _column_kind(self, header, col_index: int, value_col_indices: list[int], name_header: str) -> str:
        if col_index in value_col_indices:
            return "date"
        if header == name_header:
            return "name"
        return "date"

    def _cell(self, *, value, is_marked: bool, kind: str) -> dict:
        return {"value": value, "is_marked": is_marked, "kind": kind}

    def _total_label(self, sheet_type: str) -> str:
        if sheet_type == "income":
            return "总收入"
        return "总支出"

    def _normalizer(self, sheet_type: str):
        if sheet_type == "wear":
            return self.store.format_wear_value
        if sheet_type == "income":
            return self.store.format_income_value
        return self.store.format_expense_value

    def build_sheet_view(self, sheet_type: str) -> dict:
        spec = VALUE_SHEET_SPECS[sheet_type]
        snapshot = self.store.get_value_sheet_snapshot(sheet_type)
        headers = list(snapshot["headers"])
        raw_rows = snapshot["raw_rows"]
        value_col_indices = snapshot["value_col_indices"]
        normalizer = self._normalizer(sheet_type)

        column_kinds = [
            self._column_kind(header, col_index, value_col_indices, spec["name_header"])
            for col_index, header in enumerate(headers)
        ]

        date_totals = {col_index: 0.0 for col_index in value_col_indices}
        marked_count = 0
        rows = []
        for raw_row in raw_rows:
            row_total = 0.0
            row_cells = []
            for col_index, value in enumerate(raw_row):
                numeric_value = to_float_or_none(value)
                has_value = col_index in value_col_indices and numeric_value is not None and numeric_value != 0
                if has_value:
                    row_total += float(numeric_value)
                    date_totals[col_index] += float(numeric_value)
                    marked_count += 1
                row_cells.append(
                    self._cell(
                        value=value,
                        is_marked=has_value,
                        kind=column_kinds[col_index],
                    )
                )
            row_cells.append(
                self._cell(
                    value=normalizer(row_total),
                    is_marked=row_total != 0,
                    kind="total",
                )
            )
            rows.append(row_cells)

        chart_items = [
            {
                "label": str(headers[col_index] or ""),
                "value": normalizer(total),
                "is_marked": total != 0,
            }
            for col_index, total in date_totals.items()
        ][-15:]
        max_chart_value = max((abs(float(item["value"] or 0)) for item in chart_items), default=0.0)
        for item in chart_items:
            item["percent"] = 0 if max_chart_value == 0 else max(6, abs(float(item["value"] or 0)) / max_chart_value * 100)

        total_value = normalizer(sum(date_totals.values()))
        headers.append(self._total_label(sheet_type))
        column_kinds.append("total")
        return {
            "headers": headers,
            "column_kinds": column_kinds,
            "rows": rows,
            "chart_items": chart_items,
            "row_count": len(rows),
            "column_count": len(headers),
            "marked_count": marked_count,
            "total_value": total_value,
            "sheet_type": sheet_type,
        }
