from __future__ import annotations

import sqlite3

import pytest

from bm2.constants import ENABLED
from bm2.services.store_application import StoreApplication
from .conftest import add_members, score_form


def test_bad_json_reorder_returns_400(client):
    response = client.post("/members/reorder", data="[1, 2]", content_type="application/json")

    assert response.status_code == 400
    assert response.get_json() == {"ok": False}


def test_missing_score_form_fields_do_not_crash(client, route_store):
    response = client.post("/scores/save", data={"date": "2026-07-20"})

    assert response.status_code == 302
    assert route_store._context.local_db.get_score_rows_for_date("2026-07-20")


def test_service_none_form_data_fails_clearly(store):
    active = add_members(store, ["Failure_A"])

    with pytest.raises(AttributeError):
        store.daily_entry_service.process_submission(active, None, "2026-07-21")


def test_sqlite_missing_table_fails_loudly(store):
    active = add_members(store, ["Failure_B"])
    connection = sqlite3.connect(store._context.local_db.db_path)
    connection.execute("DROP TABLE score_entries")
    connection.commit()
    connection.close()

    with pytest.raises(sqlite3.OperationalError):
        store.daily_entry_service.process_submission(active, score_form(["Failure_B"], 1), "2026-07-22")


def test_sqlite_lock_fails_without_excel_fallback(store, monkeypatch):
    active = add_members(store, ["Failure_C"])
    lock_connection = sqlite3.connect(store._context.local_db.db_path)
    lock_connection.execute("BEGIN EXCLUSIVE")

    original_connect = store._context.local_db._connect

    def connect_without_wait():
        connection = sqlite3.connect(store._context.local_db.db_path, timeout=0)
        connection.row_factory = sqlite3.Row
        return connection

    monkeypatch.setattr(store._context.local_db, "_connect", connect_without_wait)
    try:
        with pytest.raises(sqlite3.OperationalError):
            store.daily_entry_service.process_submission(active, score_form(["Failure_C"], 1), "2026-07-23")
    finally:
        monkeypatch.setattr(store._context.local_db, "_connect", original_connect)
        lock_connection.rollback()
        lock_connection.close()

    assert store._context.local_db.get_score_rows_for_date("2026-07-23") == []


def test_startup_creates_missing_base_dir(test_workspace):
    missing_base = test_workspace / "missing" / "nested"

    app = StoreApplication(missing_base)

    assert missing_base.exists()
    assert app._context.local_db.db_path.exists()


def test_corrupt_excel_import_fails_without_sqlite_pollution(store):
    before_members = store._context.local_db.get_member_rows(include_deleted=True)
    before_scores = store._context.local_db.get_score_rows(include_deleted=True)
    store.workbook_path.write_bytes(b"not an xlsx")

    with pytest.raises(Exception):
        store.refresh_local_database()

    assert store._context.local_db.get_member_rows(include_deleted=True) == before_members
    assert store._context.local_db.get_score_rows(include_deleted=True) == before_scores


def test_supabase_failure_does_not_change_local_rows(client, route_store, monkeypatch):
    route_store.member_service.add_member("Cloud_Failure")
    local_save = route_store.daily_entry_service.process_submission(
        [{"name": "Cloud_Failure", "status": ENABLED}],
        score_form(["Cloud_Failure"], 2),
        "2026-07-24",
    )
    assert local_save["ok"] is True
    before_rows = route_store._context.local_db.get_score_rows(include_deleted=True)

    monkeypatch.setattr(route_store, "is_supabase_configured", lambda: True)

    def fail_push(*, force_full=True):
        raise RuntimeError("cloud unavailable")

    monkeypatch.setattr(route_store, "supabase_push", fail_push)
    response = client.post("/supabase-push")

    assert response.status_code == 302
    assert route_store._context.local_db.get_score_rows(include_deleted=True) == before_rows


def test_empty_and_invalid_date_pages_do_not_crash(client):
    assert client.get("/score-overview").status_code == 200
    assert client.get("/members").status_code == 200
    assert client.get("/scores?date=not-a-date").status_code == 200
    assert client.get("/api/score-overview").status_code == 200
    assert client.get("/not-found").status_code == 404
