from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest
from werkzeug.datastructures import MultiDict

from .conftest import add_members, score_form


def _workbook_mtime(store) -> float | None:
    path = Path(store.workbook_path)
    return path.stat().st_mtime if path.exists() else None


def test_normal_and_multi_member_score_save_writes_only_sqlite(store):
    members = add_members(store, ["Score_A", "Score_B"])
    before_excel = _workbook_mtime(store)

    result = store.daily_entry_service.process_submission(
        members,
        score_form(["Score_A", "Score_B"], 7),
        "2026-07-01",
    )

    assert result["ok"] is True
    rows = store._context.local_db.get_score_rows_for_date("2026-07-01")
    saved = {row["member_name"]: row["score"] for row in rows if row["member_name"].startswith("Score_")}
    assert saved == {"Score_A": 7, "Score_B": 7}
    assert _workbook_mtime(store) == before_excel


def test_empty_zero_negative_and_duplicate_scores(store):
    members = add_members(store, ["Score_C", "Score_D", "Score_E"])

    result = store.daily_entry_service.process_submission(
        members,
        MultiDict(
            {
                "score_Score_C": "",
                "score_Score_D": "0",
                "score_Score_E": "-2",
            }
        ),
        "2026-07-02",
    )
    assert result["ok"] is True
    rows = store._context.local_db.get_score_rows_for_date("2026-07-02")
    saved = {row["member_name"]: row["score"] for row in rows if row["member_name"].startswith("Score_")}
    assert saved == {"Score_C": 0, "Score_D": 0, "Score_E": -2}

    duplicate = store.daily_entry_service.process_submission(
        members,
        score_form(["Score_C", "Score_D", "Score_E"], 9),
        "2026-07-02",
    )
    assert duplicate["ok"] is True
    updated_rows = [
        row for row in store._context.local_db.get_score_rows_for_date("2026-07-02")
        if row["member_name"].startswith("Score_")
    ]
    assert len(updated_rows) == 3
    assert {row["score"] for row in updated_rows} == {9}


@pytest.mark.parametrize(
    ("selected_date", "form_data"),
    [
        ("", MultiDict({})),
        ("2026-99-01", score_form(["Score_F"], 1)),
        ("2026-07-03", score_form(["Unknown_Member"], 1)),
        ("2026-07-03", score_form(["Score_F"], "abc")),
    ],
)
def test_invalid_score_submissions_do_not_write_rows(store, selected_date, form_data):
    active_members = add_members(store, ["Score_F"])
    before_rows = list(store._context.local_db.get_score_rows(include_deleted=True))
    before_excel = _workbook_mtime(store)

    result = store.daily_entry_service.process_submission(active_members, form_data, selected_date)

    assert result["ok"] is False
    assert store._context.local_db.get_score_rows(include_deleted=True) == before_rows
    assert _workbook_mtime(store) == before_excel


def test_sqlite_write_failure_is_not_hidden(store):
    active_members = add_members(store, ["Score_G"])
    database_path = store._context.local_db.db_path
    connection = sqlite3.connect(database_path)
    connection.execute("DROP TABLE score_entries")
    connection.commit()
    connection.close()

    with pytest.raises(sqlite3.OperationalError):
        store.daily_entry_service.process_submission(
            active_members,
            score_form(["Score_G"], 5),
            "2026-07-04",
        )

    connection = sqlite3.connect(database_path)
    try:
        tables = {
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
        }
    finally:
        connection.close()
    assert "score_entries" not in tables
