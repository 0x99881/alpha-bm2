from __future__ import annotations

from datetime import date, datetime

from ..constants import VALUE_SHEET_SPECS
from ..value_utils import to_float_or_none


def _label_to_iso(label, today_iso: str) -> str | None:
    """Parse a value-sheet column header into an ISO date string, or None.

    Accepts: ``YYYY-MM-DD``, ``MM-DD`` (with current/previous year as needed),
    ``MMDD`` (4 digits, current year). Returns None for non-date labels like
    the 名称/合计 columns.
    """
    if isinstance(label, (date, datetime)):
        return label.strftime("%Y-%m-%d") if not isinstance(label, datetime) else label.date().strftime("%Y-%m-%d")
    text = str(label or "").strip()
    if not text:
        return None
    try:
        return datetime.strptime(text, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        pass
    try:
        parsed = datetime.strptime(text, "%m-%d")
    except ValueError:
        parsed = None
    if parsed is None and len(text) == 4 and text.isdigit():
        try:
            parsed = datetime.strptime(text, "%m%d")
        except ValueError:
            parsed = None
    if parsed is None:
        return None
    year = int(today_iso[:4])
    candidate = parsed.replace(year=year).strftime("%Y-%m-%d")
    if candidate > today_iso:
        candidate = parsed.replace(year=year - 1).strftime("%Y-%m-%d")
    return candidate


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
            return "总空投收入"
        return "总支出"

    def _normalizer(self, sheet_type: str):
        if sheet_type == "wear":
            return self.store.format_wear_value
        if sheet_type == "income":
            return self.store.format_income_value
        return self.store.format_expense_value

    def build_sheet_view_for_cycle(
        self,
        sheet_type: str,
        *,
        score_rows: list,
        member_order: list,
        start_iso: str,
        end_iso: str,
    ) -> dict:
        """Rebuild the chart-page view from SQLite for a cycle window.

        Used by the income/expense chart pages so the Excel-style preview
        table actually changes when the user picks a different cycle.
        Empty bounds ⇒ full history (= 全部历史 view).
        """
        spec = VALUE_SHEET_SPECS[sheet_type]
        name_header = spec["name_header"]
        normalizer = self._normalizer(sheet_type)
        field_key = {"income": "income", "expense": "other_expense"}[sheet_type]

        date_set: set[str] = set()
        per_cell: dict[tuple[str, str], float] = {}
        for row in score_rows:
            d = str(row.get("score_date") or "").strip()
            if not d:
                continue
            if start_iso and end_iso and not (start_iso <= d <= end_iso):
                continue
            name = str(row.get("member_name") or "").strip()
            if not name:
                continue
            raw = str(row.get(field_key) or "").strip()
            if not raw:
                continue
            v = to_float_or_none(raw)
            if v is None or v == 0:
                continue
            per_cell[(name, d)] = normalizer(v)
            date_set.add(d)

        dates = sorted(date_set)
        date_headers = [d[5:].replace("-", "") for d in dates]
        headers = list(date_headers) + [name_header, self._total_label(sheet_type)]
        value_col_indices = list(range(len(date_headers)))
        column_kinds = [self._column_kind(h, i, value_col_indices, name_header) for i, h in enumerate(headers)]

        date_totals = {i: 0.0 for i in value_col_indices}
        marked_count = 0
        built_rows: list[tuple[float, int, list]] = []
        for member_index, name in enumerate(member_order):
            row_cells = []
            row_total = 0.0
            for i, d in enumerate(dates):
                v = per_cell.get((name, d))
                cell_val = v if v is not None else None
                has = v is not None and v != 0
                if has:
                    row_total += float(v)
                    date_totals[i] += float(v)
                    marked_count += 1
                row_cells.append(self._cell(value=cell_val, is_marked=has, kind=column_kinds[i]))
            row_cells.append(self._cell(value=name, is_marked=False, kind="name"))
            row_cells.append(self._cell(value=normalizer(row_total), is_marked=row_total != 0, kind="total"))
            built_rows.append((row_total, member_index, row_cells))
        # 预览表按各成员合计金额从大到小排，一眼看出谁多谁少；合计相同（含全为 0
        # 的成员）时保持原成员顺序，保证排序稳定、可预期。
        built_rows.sort(key=lambda item: (-item[0], item[1]))
        rows = [cells for _, _, cells in built_rows]

        chart_items = [
            {"label": str(headers[i] or ""), "value": normalizer(t), "is_marked": t != 0}
            for i, t in date_totals.items()
        ]
        max_chart = max((abs(float(it["value"] or 0)) for it in chart_items), default=0.0)
        for it in chart_items:
            it["percent"] = 0 if max_chart == 0 else max(6, abs(float(it["value"] or 0)) / max_chart * 100)

        total_value = normalizer(sum(float(it["value"] or 0) for it in chart_items))

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

    def build_sheet_view(
        self,
        sheet_type: str,
        *,
        cycle_window: dict | None = None,
    ) -> dict:
        """Build the value-sheet view for the chart pages.

        When ``cycle_window`` is provided (dict with ``start_date`` and
        ``end_date`` ISO strings), chart bars and the headline total are
        restricted to columns whose label date falls inside the window. The
        Excel preview table below the chart is unaffected — it always shows the
        raw sheet so the user can still spot historical anomalies.
        """
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

        # Cycle-window filter: keep only columns whose date falls in
        # [start_date, end_date]. When labels can't be parsed (non-date
        # columns) they are kept so the chart degrades gracefully.
        if cycle_window and cycle_window.get("has_cycle"):
            start = cycle_window.get("start_date") or ""
            end = cycle_window.get("end_date") or ""
            today_iso = date.today().strftime("%Y-%m-%d")
            kept = []
            for item in chart_items:
                iso = _label_to_iso(item["label"], today_iso)
                if iso is None or (start <= iso <= end):
                    kept.append(item)
            chart_items = kept

        max_chart_value = max((abs(float(item["value"] or 0)) for item in chart_items), default=0.0)
        for item in chart_items:
            item["percent"] = 0 if max_chart_value == 0 else max(6, abs(float(item["value"] or 0)) / max_chart_value * 100)

        if cycle_window and cycle_window.get("has_cycle"):
            total_value = normalizer(sum(float(item["value"] or 0) for item in chart_items))
        else:
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
