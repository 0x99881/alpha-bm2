from __future__ import annotations

import random

from bm2.constants import DISABLED, ENABLED
from .conftest import score_form


def test_lightweight_deterministic_stress(store):
    rng = random.Random(20260501)
    names = [f"Stress_{index}" for index in range(10)]
    for name in names:
        store.member_service.add_member(name)

    for operation in range(100):
        active_names = [
            item["name"]
            for item in store.member_service.get_active_members()
            if item["name"].startswith("Stress_")
        ]
        all_names = [
            item["name"]
            for item in store.member_service.get_members()
            if item["name"].startswith("Stress_")
        ]
        action = rng.choice(["save", "disable", "enable", "query", "export"])

        if action == "save" and active_names:
            selected = rng.sample(active_names, k=min(len(active_names), rng.randint(1, 3)))
            score = rng.choice([0, 1, 5, -2, 9999])
            day = rng.randint(1, 28)
            result = store.daily_entry_service.process_submission(
                [{"name": name, "status": ENABLED} for name in selected],
                score_form(selected, score),
                f"2026-09-{day:02d}",
            )
            assert result["ok"] is True
        elif action == "disable" and active_names:
            store.member_service.update_member(rng.choice(active_names), "stress", status=DISABLED)
        elif action == "enable" and all_names:
            store.member_service.update_member(rng.choice(all_names), "stress", status=ENABLED)
        elif action == "query":
            store.get_score_summary()
            store.get_score_sheet_view()
        else:
            store._context.export_service.export_to_excel()

    store.get_score_summary()
    store.get_member_profit_calendar("all", 2026, 9)
    rows = store._context.local_db.get_score_rows(include_deleted=True)
    assert len(rows) >= 0
    assert store._context.local_db.db_path.parent == store._context.base_dir
    assert store.workbook_path.parent == store._context.base_dir
