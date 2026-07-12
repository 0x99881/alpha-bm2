from __future__ import annotations

from typing import Any

from ..constants import DATA_START_ROW


def build_name_row_map(sheet, name_col: int) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for row in range(DATA_START_ROW, sheet.max_row + 1):
        name = sheet.cell(row, name_col).value
        if name:
            mapping[str(name).strip()] = row
    return mapping


def ensure_member_rows(
    sheet,
    *,
    members: list[dict[str, Any]],
    name_col: int,
    total_col: int | None,
    value_columns: list[int],
    extra_columns: list[int] | None = None,
) -> bool:
    _, changed = ensure_member_rows_with_map(
        sheet,
        members=members,
        name_col=name_col,
        total_col=total_col,
        value_columns=value_columns,
        extra_columns=extra_columns,
    )
    return changed


def ensure_member_rows_with_map(
    sheet,
    *,
    members: list[dict[str, Any]],
    name_col: int,
    total_col: int | None,
    value_columns: list[int],
    extra_columns: list[int] | None = None,
) -> tuple[dict[str, int], bool]:
    changed = False
    extra_columns = extra_columns or []
    protected_cols = sorted({*value_columns, *extra_columns, *([total_col] if total_col is not None else []), name_col})

    if name_col < sheet.max_column:
        sheet.delete_cols(name_col + 1, sheet.max_column - name_col)
        changed = True

    first_rows: dict[str, int] = {}
    duplicate_rows: list[int] = []
    for row in range(DATA_START_ROW, sheet.max_row + 1):
        raw_name = sheet.cell(row, name_col).value
        if not raw_name:
            continue
        name = str(raw_name).strip()
        if name not in first_rows:
            first_rows[name] = row
            continue
        keeper_row = first_rows[name]
        for col in protected_cols:
            current = sheet.cell(keeper_row, col).value
            candidate = sheet.cell(row, col).value
            if _should_replace_cell(current, candidate):
                sheet.cell(keeper_row, col, candidate)
                changed = True
        keeper_hidden = bool(sheet.row_dimensions[keeper_row].hidden)
        row_hidden = bool(sheet.row_dimensions[row].hidden)
        if keeper_hidden and not row_hidden:
            sheet.row_dimensions[keeper_row].hidden = False
            changed = True
        duplicate_rows.append(row)
    for row in reversed(duplicate_rows):
        sheet.delete_rows(row, 1)
        changed = True

    row_map = build_name_row_map(sheet, name_col)
    for member in members:
        name = member["name"]
        if name in row_map:
            continue
        row = sheet.max_row + 1
        for col in value_columns:
            sheet.cell(row, col, 0)
        if total_col is not None:
            sheet.cell(row, total_col, 0)
        for col in extra_columns:
            sheet.cell(row, col, 0)
        sheet.cell(row, name_col, name)
        row_map[name] = row
        changed = True
    return row_map, changed


def sort_named_rows(sheet, total_col: int, name_col: int) -> None:
    rows = _read_named_rows(sheet, name_col)
    rows.sort(
        key=lambda row: (
            -(float(row[total_col - 1]) if isinstance(row[total_col - 1], (int, float)) else 0.0),
            str(row[name_col - 1]),
        )
    )
    _write_rows(sheet, rows)


def sort_rows_by_name(sheet, name_col: int) -> None:
    rows = _read_named_rows(sheet, name_col)
    rows.sort(key=lambda row: str(row[name_col - 1]))
    _write_rows(sheet, rows)


def _read_named_rows(sheet, name_col: int) -> list[list[Any]]:
    rows = []
    for row in range(DATA_START_ROW, sheet.max_row + 1):
        values = [sheet.cell(row, col).value for col in range(1, sheet.max_column + 1)]
        if values[name_col - 1]:
            rows.append(values)
    return rows


def _write_rows(sheet, rows: list[list[Any]]) -> None:
    for row_index, values in enumerate(rows, start=DATA_START_ROW):
        for col_index, value in enumerate(values, start=1):
            sheet.cell(row_index, col_index, value)


def _is_meaningful_value(value: Any) -> bool:
    return value not in (None, '')


def _should_replace_cell(current: Any, candidate: Any) -> bool:
    if not _is_meaningful_value(candidate):
        return False
    if not _is_meaningful_value(current):
        return True
    if isinstance(current, (int, float)) and current == 0 and isinstance(candidate, (int, float)) and candidate != 0:
        return True
    return False
