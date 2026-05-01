from __future__ import annotations

from pathlib import Path

import pytest

from bm2.constants import DISABLED, ENABLED
from .conftest import add_members


def _workbook_mtime(store) -> float | None:
    path = Path(store.workbook_path)
    return path.stat().st_mtime if path.exists() else None


def test_add_member_trims_name_and_uses_sqlite_only(store):
    before_excel = _workbook_mtime(store)

    store.member_service.add_member("  Alice  ", "note")

    members = store.member_service.get_members()
    alice = next(item for item in members if item["name"] == "Alice")
    assert alice["note"] == "note"
    assert alice["status"] == ENABLED
    assert _workbook_mtime(store) == before_excel


@pytest.mark.parametrize("name", ["", "   "])
def test_blank_member_names_are_rejected(store, name):
    before_members = store.member_service.get_members()

    with pytest.raises(ValueError):
        store.member_service.add_member(name)

    assert store.member_service.get_members() == before_members


def test_duplicate_chinese_and_special_member_names(store):
    store.member_service.add_member("张三")
    store.member_service.add_member("Name-特殊_#1")

    with pytest.raises(ValueError):
        store.member_service.add_member("张三")

    names = {item["name"] for item in store.member_service.get_members()}
    assert "张三" in names
    assert "Name-特殊_#1" in names


def test_update_missing_and_delete_missing_members(store):
    store.member_service.add_member("Member_To_Update", "old")

    store.member_service.update_member("Member_To_Update", "new", status=DISABLED)
    updated = next(item for item in store.member_service.get_members() if item["name"] == "Member_To_Update")
    assert updated["note"] == "new"
    assert updated["status"] == DISABLED

    with pytest.raises(ValueError):
        store.member_service.update_member("Missing_Member", "note", status=ENABLED)

    with pytest.raises(ValueError):
        store.member_service.delete_member("Missing_Member")


def test_delete_member_keeps_history_in_sqlite(store):
    active = add_members(store, ["History_Member"])
    store.daily_entry_service.process_submission(
        active,
        {"score_History_Member": "6"},
        "2026-07-05",
    )

    store.member_service.delete_member("History_Member")

    assert "History_Member" not in [item["name"] for item in store.member_service.get_active_members()]
    history = store._context.local_db.get_score_rows(include_deleted=True)
    assert any(row["member_name"] == "History_Member" for row in history)


def test_active_member_reorder_is_stable(store):
    add_members(store, ["Order_A", "Order_B", "Order_C"])

    store.member_service.reorder_active_members(["Order_C", "Order_A", "Order_B"])

    active_names = [
        item["name"]
        for item in store.member_service.get_active_members()
        if item["name"].startswith("Order_")
    ]
    assert active_names == ["Order_C", "Order_A", "Order_B"]


def test_member_routes_do_not_write_excel(client, route_store):
    before_excel = _workbook_mtime(route_store)

    response = client.post("/members/add", data={"name": "Route_Member", "note": "via route"})

    assert response.status_code == 302
    assert client.get("/members").status_code == 200
    assert any(item["name"] == "Route_Member" for item in route_store.member_service.get_members())
    assert _workbook_mtime(route_store) == before_excel
