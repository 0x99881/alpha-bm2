from __future__ import annotations

import importlib
import json
import re
import shutil
import sys
from pathlib import Path

sys.dont_write_bytecode = True


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
TEST_TEMP_ROOT = PROJECT_ROOT / ".tmp_test_workspaces"


MODULES_TO_IMPORT = [
    "bm2.services.store_application",
    "bm2.services.daily_entry_service",
    "bm2.presenters.score_presenter",
    "bm2.presenters.wear_presenter",
    "bm2.presenters.profit_calendar_presenter",
]


class SmokeCheckRunner:
    def __init__(self) -> None:
        self.results: list[tuple[bool, str, str]] = []
        self._store = None
        self._temp_root: Path | None = None

    def check(self, name: str, func) -> None:
        try:
            detail = func()
            self.results.append((True, name, "" if detail is None else str(detail)))
        except Exception as exc:  # pragma: no cover - smoke output path
            self.results.append((False, name, f"{type(exc).__name__}: {exc}"))

    def build_store(self):
        if self._store is not None:
            return self._store
        from bm2.services.store_application import StoreApplication

        temp_dir = TEST_TEMP_ROOT / "smoke"
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        temp_dir.mkdir(parents=True)
        self._temp_root = temp_dir
        self._store = StoreApplication(temp_dir)
        return self._store

    def check_module_imports(self) -> None:
        for module_name in MODULES_TO_IMPORT:
            self.check(f"import {module_name}", lambda module_name=module_name: importlib.import_module(module_name).__name__)

    def check_application_init(self) -> None:
        def _run():
            store = self.build_store()
            return f"workbook={store.workbook_path.name}"

        self.check("StoreApplication 初始化", _run)

    def check_store_read_entry(self) -> None:
        def _run():
            store = self.build_store()
            next_date = store.get_next_score_date()
            wear_records = store.get_member_wear_records("bb")
            return f"next_date={next_date}, wear_records={len(wear_records)}"

        self.check("repository 入口", _run)

    def check_presenter_entry(self) -> None:
        def _run():
            store = self.build_store()
            summary = store.get_score_summary()
            required_keys = {"latest_column", "window_size", "rankings"}
            missing = required_keys - set(summary.keys())
            if missing:
                raise ValueError(f"missing keys: {sorted(missing)}")
            return f"summary_keys={sorted(summary.keys())}"

        self.check("presenter 入口", _run)

    def check_score_overview_sorting_and_low_score(self) -> None:
        def _run():
            from bm2.services.store_application import StoreApplication

            temp_dir = TEST_TEMP_ROOT / "score_overview_sorting"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir)
                members = store.get_active_members()
                high_member = members[0]["name"]
                low_member = members[1]["name"]
                target_date = store.get_next_score_date()
                form_data = {f"score_{member['name']}": "0" for member in members}
                form_data[f"score_{high_member}"] = "17"
                form_data[f"score_{low_member}"] = "9"
                result = store.daily_entry_service.process_submission(members, form_data, target_date)
                if not result.get("ok"):
                    raise ValueError(result.get("error"))

                score_sheet = store.get_score_sheet_view()
                name_index = next(
                    index
                    for index, header in enumerate(score_sheet["headers"])
                    if header["kind"] == "name"
                )
                latest_index = next(
                    index
                    for index, header in enumerate(score_sheet["headers"])
                    if header["value"] == target_date[5:]
                )
                rows = score_sheet["rows"]
                if rows[0][name_index]["value"] != high_member:
                    raise ValueError("score overview is not sorted by total score")
                low_row = next(row for row in rows if row[name_index]["value"] == low_member)
                high_row = next(row for row in rows if row[name_index]["value"] == high_member)
                if not low_row[latest_index].get("is_low_score"):
                    raise ValueError("low daily score is not marked")
                if high_row[latest_index].get("is_low_score"):
                    raise ValueError("normal daily score was marked low")
                return f"{high_member}>{low_member}"
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("score overview sorting and low score", _run)

    def check_online_score_overview_sorting_and_low_score(self) -> None:
        def _run():
            from bm2.constants import ENABLED
            from bm2.services.application_service import ApplicationService

            class FakeSupabase:
                def is_configured(self):
                    return True

                def pull_members(self):
                    return [
                        {"name": "low", "status": ENABLED, "deleted": 0, "sort_order": 1},
                        {"name": "high", "status": ENABLED, "deleted": 0, "sort_order": 2},
                    ]

                def pull_score_entries(self):
                    return [
                        {"member_name": "low", "score_date": "2026-05-12", "score": 9, "deleted": 0},
                        {"member_name": "high", "score_date": "2026-05-12", "score": 17, "deleted": 0},
                    ]

            service = ApplicationService(None, FakeSupabase(), lambda: "2026-05-12")
            view = service.mobile_overview()["score_sheet_view"]
            headers = view["headers"]
            name_index = next(index for index, header in enumerate(headers) if header["kind"] == "name")
            latest_index = next(index for index, header in enumerate(headers) if header["value"] == "05-12")
            rows = view["rows"]
            if rows[0][name_index]["value"] != "high":
                raise ValueError("online score overview is not sorted by total score")
            low_row = next(row for row in rows if row[name_index]["value"] == "low")
            high_row = next(row for row in rows if row[name_index]["value"] == "high")
            if not low_row[latest_index].get("is_low_score"):
                raise ValueError("online low daily score is not marked")
            if high_row[latest_index].get("is_low_score"):
                raise ValueError("online normal daily score was marked low")
            return "high>low"

        self.check("online score overview sorting and low score", _run)

    def check_online_wear_view_uses_synced_wear(self) -> None:
        def _run():
            from bm2.constants import ENABLED
            from bm2.services.application_service import ApplicationService

            class FakeSupabase:
                def is_configured(self):
                    return True

                def pull_members(self):
                    return [
                        {"name": "manual", "status": ENABLED, "deleted": 0, "sort_order": 1},
                        {"name": "balance", "status": ENABLED, "deleted": 0, "sort_order": 2},
                    ]

                def pull_score_entries(self):
                    return [
                        {
                            "member_name": "manual",
                            "score_date": "2026-05-12",
                            "score": 8,
                            "manual_wear": "3",
                            "before_balance": "",
                            "after_balance": "",
                            "deleted": 0,
                        },
                        {
                            "member_name": "balance",
                            "score_date": "2026-05-12",
                            "score": 8,
                            "manual_wear": "",
                            "before_balance": "10",
                            "after_balance": "8",
                            "deleted": 0,
                        },
                    ]

            service = ApplicationService(None, FakeSupabase(), lambda: "2026-05-12")
            view = service.online_wear_sheet_view()
            date_index = view["headers"].index("05-12")
            name_index = view["headers"].index("姓名")
            manual_row = next(row for row in view["rows"] if row[name_index]["value"] == "manual")
            balance_row = next(row for row in view["rows"] if row[name_index]["value"] == "balance")
            if manual_row[date_index]["value"] != 3.0:
                raise ValueError("manual wear is not shown online")
            if not manual_row[date_index]["is_abnormal"]:
                raise ValueError("online abnormal wear is not marked")
            if balance_row[date_index]["value"] != 2.0:
                raise ValueError("balance-derived wear is not shown online")
            return "manual=3 balance=2"

        self.check("online wear view uses synced wear", _run)

    def check_service_entry(self) -> None:
        def _run():
            store = self.build_store()
            active_members = store.get_active_members()
            selected_date = store.get_next_score_date()
            result = store.daily_entry_service.process_submission(active_members, {}, selected_date)
            if "ok" not in result:
                raise ValueError("missing ok flag")
            if result["ok"]:
                raise ValueError("empty submission should not save")
            if store.get_score_rows_for_date(selected_date):
                raise ValueError("empty submission created rows")
            return f"ok={result['ok']}"

        self.check("service 入口", _run)

    def check_application_entry(self) -> None:
        def _run():
            store = self.build_store()
            summary = store.get_score_summary()
            wear_view = store.get_wear_sheet_view()
            if "rankings" not in summary:
                raise ValueError("score summary missing rankings")
            if "rows" not in wear_view:
                raise ValueError("wear view missing rows")
            return f"rankings={len(summary['rankings'])}, wear_rows={wear_view['row_count']}"

        self.check("application 入口", _run)

    def check_static_assets(self) -> None:
        def _run():
            base_template = (PROJECT_ROOT / "templates" / "base.html").read_text(encoding="utf-8")
            referenced_assets = re.findall(r"filename='([^']+)'", base_template)
            missing = [asset for asset in referenced_assets if not (PROJECT_ROOT / "static" / asset).exists()]
            if missing:
                raise FileNotFoundError(", ".join(missing))
            if not (PROJECT_ROOT / "static" / "app.js").exists():
                raise FileNotFoundError("static/app.js")
            return f"assets={len(referenced_assets)}"

        self.check("模板静态资源", _run)

    def check_page_routes(self) -> None:
        def _run():
            from flask import Flask

            from bm2.services.store_application import StoreApplication
            from bm2.web import register_routes

            temp_dir = TEST_TEMP_ROOT / "routes"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                app = Flask(
                    __name__,
                    template_folder=str(PROJECT_ROOT / "templates"),
                    static_folder=str(PROJECT_ROOT / "static"),
                )
                app.secret_key = "smoke"
                register_routes(app, StoreApplication(temp_dir))
                client = app.test_client()
                routes = ["/scores", "/score-overview", "/wear", "/profit-calendar", "/members"]
                failures = []
                for route in routes:
                    response = client.get(route)
                    if response.status_code != 200:
                        failures.append(f"{route}={response.status_code}")
                    if route == "/scores":
                        if response.data.count(b"data-risk-confirm") != 3:
                            failures.append("/scores risk confirmation buttons missing")
                if failures:
                    raise ValueError(", ".join(failures))
                return f"routes={len(routes)}"
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("页面路由", _run)

    def check_read_only_routes_do_not_open_excel(self) -> None:
        def _run():
            from flask import Flask

            from bm2.services.store_application import StoreApplication
            from bm2.web import register_routes

            temp_dir = TEST_TEMP_ROOT / "read_only_routes"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir, read_only=True)
                if not store.read_only:
                    raise ValueError("read_only was not exposed")
                store.get_online_next_score_date = lambda: "2026-05-14"
                store.get_score_rows_for_date = lambda _date: []
                store.get_online_active_members = lambda: []
                store.get_online_score_summary = lambda: {"rankings": [], "latest_column": "", "window_size": 15}

                app = Flask(
                    __name__,
                    template_folder=str(PROJECT_ROOT / "templates"),
                    static_folder=str(PROJECT_ROOT / "static"),
                )
                app.secret_key = "smoke"
                register_routes(app, store)
                response = app.test_client().get("/scores")
                if response.status_code != 200:
                    raise ValueError(f"/scores={response.status_code}")
                return "read_only=/scores"
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("read only routes avoid excel", _run)

    def check_excel_date_notes_import(self) -> None:
        def _run():
            from openpyxl import load_workbook

            from bm2.constants import NAME_HEADER, SCORE_SHEET
            from bm2.services.store_application import StoreApplication

            temp_dir = TEST_TEMP_ROOT / "excel_date_notes_import"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir)
                members = store.get_active_members()
                target_date = store.get_next_score_date()
                first_name = members[0]["name"]
                form_data = {f"score_{member['name']}": "1" for member in members}
                form_data["date_note_1"] = "old note"
                result = store.daily_entry_service.process_submission(
                    members,
                    form_data,
                    target_date,
                    existing_notes=store.get_score_date_notes(target_date),
                )
                if not result.get("ok"):
                    raise ValueError(result.get("error"))
                store.export_to_excel(target_date)

                def _note_cells():
                    workbook = load_workbook(store.workbook_path)
                    try:
                        sheet = workbook[SCORE_SHEET]
                        date_col = next(
                            col
                            for col in range(1, sheet.max_column + 1)
                            if str(sheet.cell(1, col).value or "") == target_date[5:]
                        )
                        name_col = next(
                            col
                            for col in range(1, sheet.max_column + 1)
                            if str(sheet.cell(1, col).value or "") == NAME_HEADER
                        )
                        named_rows = [
                            row
                            for row in range(2, sheet.max_row + 1)
                            if str(sheet.cell(row, name_col).value or "").strip()
                        ]
                        if not named_rows:
                            raise ValueError("no member rows")
                        if not any(str(sheet.cell(row, name_col).value or "") == first_name for row in named_rows):
                            raise ValueError("member row missing")
                        return workbook, sheet, date_col, max(named_rows) + 1
                    except (KeyError, StopIteration, ValueError):
                        workbook.close()
                        raise

                workbook, sheet, date_col, note_row = _note_cells()
                try:
                    sheet.cell(note_row, date_col, "manual note 1")
                    sheet.cell(note_row + 1, date_col, "manual note 2")
                    sheet.cell(note_row + 2, date_col, "manual note 3")
                    workbook.save(store.workbook_path)
                finally:
                    workbook.close()

                refresh_result = store.refresh_local_database()
                if not refresh_result.get("changed"):
                    raise ValueError("manual Excel note change was not reported")
                if store.get_score_date_notes(target_date) != {"note1": "manual note 1", "note2": "manual note 2", "note3": "manual note 3"}:
                    raise ValueError("manual Excel notes were not imported")

                workbook, sheet, date_col, note_row = _note_cells()
                try:
                    sheet.cell(note_row, date_col, "")
                    sheet.cell(note_row + 1, date_col, "")
                    sheet.cell(note_row + 2, date_col, "")
                    workbook.save(store.workbook_path)
                finally:
                    workbook.close()

                refresh_result = store.refresh_local_database()
                if not refresh_result.get("changed"):
                    raise ValueError("blank Excel note change was not reported")
                if store.get_score_date_notes(target_date) != {"note1": "", "note2": "", "note3": ""}:
                    raise ValueError("blank Excel notes did not clear local notes")
                return target_date
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("excel date notes import", _run)

    def check_excel_refresh_side_effect_guards(self) -> None:
        def _run():
            from openpyxl import load_workbook

            from bm2.constants import NAME_HEADER, SCORE_SHEET
            from bm2.services.store_application import StoreApplication

            temp_dir = TEST_TEMP_ROOT / "excel_refresh_side_effects"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir)
                members = store.get_active_members()
                target_date = store.get_next_score_date()
                first_name = members[0]["name"]
                form_data = {f"score_{member['name']}": "12" for member in members}
                form_data[f"before_{first_name}"] = "100"
                form_data[f"after_{first_name}"] = "90"
                form_data[f"income_{first_name}"] = "33"
                form_data[f"other_expense_{first_name}"] = "2"
                result = store.daily_entry_service.process_submission(members, form_data, target_date)
                if not result.get("ok"):
                    raise ValueError(result.get("error"))
                store.export_to_excel(target_date)

                first_refresh = store.refresh_local_database()
                second_refresh = store.refresh_local_database()
                if first_refresh.get("changed") or second_refresh.get("changed"):
                    raise ValueError("unchanged Excel refresh was reported as changed")
                preserved = next(row for row in store.get_score_rows_for_date(target_date) if row["member_name"] == first_name)
                expected = {
                    "before_balance": "100",
                    "after_balance": "90",
                    "manual_wear": "",
                    "income": "33",
                    "other_expense": "2",
                }
                actual = {key: str(preserved.get(key, "")) for key in expected}
                if actual != expected:
                    raise ValueError(f"visible fields changed during refresh: {actual}")

                def _score_cells():
                    workbook = load_workbook(store.workbook_path)
                    try:
                        sheet = workbook[SCORE_SHEET]
                        date_col = next(
                            col
                            for col in range(1, sheet.max_column + 1)
                            if str(sheet.cell(1, col).value or "") == target_date[5:]
                        )
                        name_col = next(
                            col
                            for col in range(1, sheet.max_column + 1)
                            if str(sheet.cell(1, col).value or "") == NAME_HEADER
                        )
                        first_row = next(
                            row
                            for row in range(2, sheet.max_row + 1)
                            if str(sheet.cell(row, name_col).value or "").strip() == first_name
                        )
                        return workbook, sheet, date_col, name_col, first_row
                    except (KeyError, StopIteration, ValueError):
                        workbook.close()
                        raise

                workbook, sheet, date_col, name_col, first_row = _score_cells()
                try:
                    unknown_row = sheet.max_row + 1
                    sheet.cell(unknown_row, name_col, "not-a-member")
                    sheet.cell(unknown_row, date_col, 99)
                    workbook.save(store.workbook_path)
                finally:
                    workbook.close()
                store.refresh_local_database()
                names = {row["member_name"] for row in store.get_score_rows_for_date(target_date)}
                if "not-a-member" in names:
                    raise ValueError("unknown Excel member was imported")

                workbook, sheet, date_col, name_col, first_row = _score_cells()
                try:
                    sheet.cell(first_row, date_col, 18)
                    workbook.save(store.workbook_path)
                finally:
                    workbook.close()
                changed_refresh = store.refresh_local_database()
                changed_row = next(row for row in store.get_score_rows_for_date(target_date) if row["member_name"] == first_name)
                if not changed_refresh.get("changed") or int(changed_row["score"]) != 18:
                    raise ValueError("real Excel score change was not imported")
                return target_date
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("excel refresh side effect guards", _run)

    def check_member_status_survives_excel_refresh(self) -> None:
        def _run():
            from openpyxl import load_workbook

            from bm2.constants import DISABLED, NAME_HEADER, SCORE_SHEET
            from bm2.services.store_application import StoreApplication

            temp_dir = TEST_TEMP_ROOT / "member_status_refresh"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir)
                member_name = store.get_active_members()[0]["name"]
                store.member_service.update_member(member_name, "", status=DISABLED)
                if any(row["name"] == member_name for row in store.get_active_members()):
                    raise ValueError("member is still active after disable")
                workbook = load_workbook(store.workbook_path)
                try:
                    sheet = workbook[SCORE_SHEET]
                    name_col = next(
                        col
                        for col in range(1, sheet.max_column + 1)
                        if str(sheet.cell(1, col).value or "") == NAME_HEADER
                    )
                    member_row = next(
                        row
                        for row in range(2, sheet.max_row + 1)
                        if str(sheet.cell(row, name_col).value or "").strip() == member_name
                    )
                    if not bool(sheet.row_dimensions[member_row].hidden):
                        raise ValueError("disabled member row is still visible in Excel")
                    sheet.row_dimensions[member_row].hidden = False
                    workbook.save(store.workbook_path)
                finally:
                    workbook.close()

                store = StoreApplication(temp_dir)
                workbook = load_workbook(store.workbook_path)
                try:
                    sheet = workbook[SCORE_SHEET]
                    name_col = next(
                        col
                        for col in range(1, sheet.max_column + 1)
                        if str(sheet.cell(1, col).value or "") == NAME_HEADER
                    )
                    member_row = next(
                        row
                        for row in range(2, sheet.max_row + 1)
                        if str(sheet.cell(row, name_col).value or "").strip() == member_name
                    )
                    if not bool(sheet.row_dimensions[member_row].hidden):
                        raise ValueError("startup did not restore disabled member row hiding")
                finally:
                    workbook.close()

                if any(
                    any(str(cell.get("value", "")).strip() == member_name for cell in row)
                    for row in store.get_wear_sheet_view()["rows"]
                ):
                    raise ValueError("disabled member is still visible in wear preview")

                store.refresh_local_database()
                after = next(row for row in store.get_members() if row["name"] == member_name)
                if after["status"] != DISABLED:
                    raise ValueError("Excel refresh restored disabled member")
                if any(row["name"] == member_name for row in store.get_active_members()):
                    raise ValueError("disabled member returned to active members")
                return member_name
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("member disable survives excel refresh", _run)

    def check_existing_config_members_are_not_reseeded(self) -> None:
        def _run():
            from bm2.services.store_application import StoreApplication

            temp_dir = TEST_TEMP_ROOT / "config_members_no_reseed"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                config_path = temp_dir / "system_config.json"
                config_path.write_text(
                    json.dumps(
                        {
                            "members": [
                                {
                                    "name": "only-one",
                                    "status": "启用",
                                    "note": "",
                                    "created_at": "2026-05-12 00:00:00",
                                    "disabled_at": "",
                                    "sort_order": 1,
                                }
                            ],
                            "quick_scores": [17],
                            "wear_abnormal_threshold": 2.5,
                        },
                        ensure_ascii=False,
                    ),
                    encoding="utf-8",
                )
                store = StoreApplication(temp_dir)
                names = [member["name"] for member in store.get_members()]
                if names != ["only-one"]:
                    raise ValueError(f"default members were reseeded: {names[:5]}")
                return names[0]
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("existing config members are not reseeded", _run)

    def check_supabase_profit_uses_formal_column(self) -> None:
        def _run():
            from bm2.constants import ENABLED
            from bm2.local_database import LocalDatabase

            class FakeSupabase:
                def __init__(self) -> None:
                    self.score_rows: list[dict] = []

                def ensure_score_entries_schema(self) -> None:
                    return None

                def push_members(self, rows):
                    return len(rows)

                def push_score_entries(self, rows):
                    self.score_rows = rows
                    return len(rows)

            temp_dir = TEST_TEMP_ROOT / "supabase_profit_column"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                db = LocalDatabase(temp_dir)
                db.sync_members(
                    [
                        {
                            "name": "profit-user",
                            "status": ENABLED,
                            "note": "",
                            "created_at": "2026-05-12 00:00:00",
                            "disabled_at": "",
                            "sort_order": 1,
                        }
                    ]
                )
                db.record_score_entries("2026-05-12", [{"name": "profit-user", "score": "12"}])
                fake = FakeSupabase()
                db.push_to_supabase(fake, force_full=False, score_profit_map={"profit-user": 12.34})
                if not fake.score_rows:
                    raise ValueError("no score rows were pushed")
                row = fake.score_rows[0]
                if row.get("profit") != "12.3":
                    raise ValueError(f"profit column not set: {row}")
                if "profit=" in str(row.get("source", "")):
                    raise ValueError("profit leaked into source")
                return row["profit"]
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("supabase profit formal column", _run)

    def check_online_entry_reports_supabase_request_failure(self) -> None:
        def _run():
            import httpx

            from bm2.services.supabase_entry_writer import SupabaseEntryWriter

            class FailingSupabase:
                def is_configured(self) -> bool:
                    return True

                def fetch_score_entries_by_ids(self, ids):
                    raise httpx.TimeoutException("forced timeout")

            writer = SupabaseEntryWriter(FailingSupabase())
            try:
                writer.save_scores_and_wear(
                    "2026-05-12",
                    [{"name": "online-user", "score": "12"}],
                )
            except ValueError as exc:
                if "线上数据库暂时连接失败" not in str(exc):
                    raise ValueError(f"unexpected error text: {exc}") from exc
                return "request failure reported"
            raise ValueError("supabase request failure was not reported")

        self.check("online entry request failure", _run)

    def check_wear_threshold_uses_saved_config(self) -> None:
        def _run():
            from bm2.services.store_application import StoreApplication

            temp_dir = TEST_TEMP_ROOT / "wear_threshold_config"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir)
                threshold = store.set_wear_abnormal_threshold("1.5")
                if threshold != 1.5:
                    raise ValueError("threshold return value is wrong")
                view = store.get_wear_sheet_view()
                if float(view["abnormal_threshold"]) != 1.5:
                    raise ValueError("wear sheet did not read saved threshold")
                return threshold
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("wear threshold saved config", _run)

    def check_excel_refresh_rejects_unknown_score_date(self) -> None:
        def _run():
            from openpyxl import load_workbook

            from bm2.constants import NAME_HEADER, SCORE_SHEET
            from bm2.services.store_application import StoreApplication

            temp_dir = TEST_TEMP_ROOT / "excel_unknown_score_date"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir)
                members = store.get_active_members()
                target_date = store.get_next_score_date()
                form_data = {f"score_{member['name']}": "11" for member in members}
                result = store.daily_entry_service.process_submission(members, form_data, target_date)
                if not result.get("ok"):
                    raise ValueError(result.get("error"))
                store.export_to_excel(target_date)

                workbook = load_workbook(store.workbook_path)
                try:
                    sheet = workbook[SCORE_SHEET]
                    name_col = next(
                        col
                        for col in range(1, sheet.max_column + 1)
                        if str(sheet.cell(1, col).value or "") == NAME_HEADER
                    )
                    first_row = next(
                        row
                        for row in range(2, sheet.max_row + 1)
                        if str(sheet.cell(row, name_col).value or "").strip() == members[0]["name"]
                    )
                    sheet.cell(1, 1, "not-a-date")
                    sheet.cell(first_row, 1, 11)
                    workbook.save(store.workbook_path)
                finally:
                    workbook.close()

                try:
                    store.refresh_local_database()
                except ValueError as exc:
                    if "D1" not in str(exc):
                        raise ValueError(f"unexpected error text: {exc}") from exc
                    return "D1 rejected"
                raise ValueError("invalid score date column was accepted")
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("excel unknown score date rejected", _run)

    def check_delete_score_date(self) -> None:
        def _run():
            from openpyxl import load_workbook

            from bm2.constants import EXPENSE_SHEET, INCOME_SHEET, SCORE_SHEET, WEAR_SHEET

            store = self.build_store()
            target_date = "2026-05-02"
            active_members = store.get_active_members()
            first_name = active_members[0]["name"]
            form_data = {f"score_{member['name']}": "1" for member in active_members}
            form_data[f"manual_wear_{first_name}"] = "9"
            form_data[f"income_{first_name}"] = "30"
            form_data[f"other_expense_{first_name}"] = "4"
            result = store.daily_entry_service.process_submission(active_members, form_data, target_date)
            if not result.get("ok"):
                raise ValueError(result.get("error"))
            store.export_to_excel(target_date)
            deleted = store.delete_score_date(target_date)
            if not deleted.get("deleted_rows"):
                raise ValueError("no rows deleted")
            if store.get_score_rows_for_date(target_date):
                raise ValueError("database rows still visible")

            workbook = load_workbook(store.workbook_path)
            try:
                score_headers = [
                    str(workbook[SCORE_SHEET].cell(1, col).value or "")
                    for col in range(1, workbook[SCORE_SHEET].max_column + 1)
                ]
                if "05-02" in score_headers:
                    raise ValueError("score date still in workbook")
                for sheet_name in (WEAR_SHEET, INCOME_SHEET, EXPENSE_SHEET):
                    headers = [
                        str(workbook[sheet_name].cell(1, col).value or "").zfill(4)
                        for col in range(1, workbook[sheet_name].max_column + 1)
                    ]
                    if "0502" in headers:
                        raise ValueError(f"{sheet_name} date still in workbook")
            finally:
                workbook.close()
            return f"deleted_rows={deleted['deleted_rows']}"

        self.check("delete score date", _run)

    def check_delete_score_date_rolls_back_on_export_error(self) -> None:
        def _run():
            from bm2.services.store_application import StoreApplication

            class FailingExportService:
                def export_to_excel(self, date_text: str | None = None) -> int:
                    raise ValueError("forced export failure")

            temp_dir = TEST_TEMP_ROOT / "delete_date_rollback"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir)
                target_date = store.get_next_score_date()
                active_members = store.get_active_members()
                form_data = {f"score_{member['name']}": "7" for member in active_members}
                form_data["date_note_1"] = "rollback note"
                result = store.daily_entry_service.process_submission(active_members, form_data, target_date)
                if not result.get("ok"):
                    raise ValueError(result.get("error"))
                store.export_to_excel(target_date)

                original_export_service = store._context.export_service
                store._context.export_service = FailingExportService()
                try:
                    failed_as_expected = False
                    try:
                        store.delete_score_date(target_date)
                    except ValueError:
                        failed_as_expected = True
                    if not failed_as_expected:
                        raise ValueError("delete did not fail")
                finally:
                    store._context.export_service = original_export_service

                if not store.get_score_rows_for_date(target_date):
                    raise ValueError("score rows were not restored")
                if store.get_score_date_notes(target_date).get("note1") != "rollback note":
                    raise ValueError("date note was not restored")
                return target_date
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("delete score date rollback on export error", _run)

    def check_supabase_push_saves_posted_entry_once(self) -> None:
        def _run():
            from flask import Flask

            from bm2.services.store_application import StoreApplication
            from bm2.web import register_routes

            class FakeSyncService:
                def __init__(self) -> None:
                    self.push_calls = 0

                def is_configured(self) -> bool:
                    return True

                def push(self, *, force_full: bool = True, score_profit_map=None):
                    self.push_calls += 1
                    return {"members": 0, "score_entries": 1, "score_entries_deleted": 0}

                def pull(self):
                    return {"members": 0, "score_entries": 0}

            temp_dir = TEST_TEMP_ROOT / "supabase_push_saves_posted_entry"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir)
                fake_sync = FakeSyncService()
                store._context.sync_service = fake_sync
                app = Flask(
                    __name__,
                    template_folder=str(PROJECT_ROOT / "templates"),
                    static_folder=str(PROJECT_ROOT / "static"),
                )
                app.secret_key = "smoke"
                register_routes(app, store)
                client = app.test_client()

                target_date = store.get_next_score_date()
                first_member = store.get_active_members()[0]["name"]
                response = client.post(
                    "/supabase-push",
                    data={
                        "date": target_date,
                        f"score_{first_member}": "13",
                        "date_note_1": "push note",
                    },
                )
                if response.status_code != 302:
                    raise ValueError(f"unexpected status: {response.status_code}")
                if fake_sync.push_calls != 1:
                    raise ValueError("push was not called once")
                saved_rows = store.get_score_rows_for_date(target_date)
                saved_row = next(row for row in saved_rows if row["member_name"] == first_member)
                if int(saved_row["score"]) != 13:
                    raise ValueError("posted score was not saved before push")
                if store.get_score_date_notes(target_date).get("note1") != "push note":
                    raise ValueError("posted note was not saved before push")
                return target_date
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("supabase push saves posted entry once", _run)

    def check_online_manual_wear_syncs_to_local(self) -> None:
        def _run():
            from bm2.entry_helpers import member_id, score_entry_id
            from bm2.local_database import LocalDatabase
            from bm2.services.supabase_entry_writer import SupabaseEntryWriter

            class FakeSupabase:
                def __init__(self) -> None:
                    self.rows = []

                def is_configured(self) -> bool:
                    return True

                def fetch_score_entries_by_ids(self, ids):
                    return []

                def push_score_entries(self, rows):
                    self.rows.extend(rows)
                    return len(rows)

            class FakePullSupabase:
                def __init__(self, row) -> None:
                    self.row = row

                def pull_members(self, since_updated_at=None):
                    return [
                        {
                            "id": member_id("online-user"),
                            "name": "online-user",
                            "status": "启用",
                            "note": "",
                            "sort_order": 1,
                            "created_at": "2026-05-13 00:00:00",
                            "disabled_at": "",
                            "updated_at": self.row["updated_at"],
                            "version": 1,
                            "deleted": 0,
                        }
                    ]

                def pull_score_entries(self, since_updated_at=None):
                    return [self.row]

            fake = FakeSupabase()
            writer = SupabaseEntryWriter(fake)
            writer.save_scores_and_wear(
                "2026-05-13",
                [{"name": "online-user", "score": "8", "manual_wear": "3", "income": "", "other_expense": ""}],
            )
            if not fake.rows:
                raise ValueError("online save did not push a row")
            pushed_row = fake.rows[0]
            if pushed_row["manual_wear"] != "3":
                raise ValueError("online manual wear was not saved")
            if pushed_row["id"] != score_entry_id("online-user", "2026-05-13"):
                raise ValueError("unexpected online row id")

            temp_dir = TEST_TEMP_ROOT / "online_manual_wear_sync"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                db = LocalDatabase(temp_dir)
                db.pull_from_supabase(FakePullSupabase(pushed_row))
                saved = db.get_score_rows_for_date("2026-05-13")
                if not saved:
                    raise ValueError("online row was not pulled into local database")
                if str(saved[0].get("manual_wear", "")) != "3":
                    raise ValueError("manual wear did not sync to local database")
                return saved[0]["manual_wear"]
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("online manual wear syncs to local", _run)

    def check_recent_low_score_format(self) -> None:
        def _run():
            from datetime import datetime, timedelta
            from openpyxl import load_workbook

            from bm2.constants import NAME_HEADER, SCORE_LOW_DAILY_FONT_COLOR, SCORE_SHEET
            from bm2.services.store_application import StoreApplication

            temp_dir = TEST_TEMP_ROOT / "low_score_format"
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
            temp_dir.mkdir(parents=True)
            try:
                store = StoreApplication(temp_dir)
                members = store.get_active_members()
                low_member = members[0]["name"]
                start = datetime.strptime(store.get_next_score_date(), "%Y-%m-%d").date()
                for offset in range(16):
                    current = (start + timedelta(days=offset)).strftime("%Y-%m-%d")
                    form_data = {f"score_{member['name']}": "12" for member in members}
                    form_data[f"score_{low_member}"] = "9"
                    result = store.daily_entry_service.process_submission(members, form_data, current)
                    if not result.get("ok"):
                        raise ValueError(result.get("error"))
                store.export_to_excel("2026-05-16")
                workbook = load_workbook(store.workbook_path)
                try:
                    sheet = workbook[SCORE_SHEET]
                    name_col = next(
                        col
                        for col in range(1, sheet.max_column + 1)
                        if str(sheet.cell(1, col).value or "") == NAME_HEADER
                    )
                    target_row = next(
                        row
                        for row in range(2, sheet.max_row + 1)
                        if str(sheet.cell(row, name_col).value or "") == low_member
                    )
                    by_header = {
                        str(sheet.cell(1, col).value or ""): col
                        for col in range(1, sheet.max_column + 1)
                    }

                    def color_text(header: str) -> str:
                        color = sheet.cell(target_row, by_header[header]).font.color
                        return str(getattr(color, "rgb", "") or "")

                    recent_header = (start + timedelta(days=1)).strftime("%m-%d")
                    old_header = start.strftime("%m-%d")
                    if color_text(recent_header)[-6:] != SCORE_LOW_DAILY_FONT_COLOR:
                        raise ValueError("recent low score is not red")
                    if color_text(old_header)[-6:] == SCORE_LOW_DAILY_FONT_COLOR:
                        raise ValueError("older low score should not be red")
                    high_member = members[1]["name"]
                    high_row = next(
                        row
                        for row in range(2, sheet.max_row + 1)
                        if str(sheet.cell(row, name_col).value or "") == high_member
                    )
                    high_cell = sheet.cell(high_row, by_header[recent_header])
                    high_fill = str(getattr(high_cell.fill.fgColor, "rgb", "") or "")[-6:]
                    if int(high_cell.value or 0) != 12:
                        raise ValueError("score value changed during formatting")
                    if high_fill not in ("", "000000"):
                        raise ValueError("normal score should not have special fill")
                finally:
                    workbook.close()
                return "recent low score red"
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("recent low score format", _run)

    def check_architecture_rules(self) -> None:
        def _run():
            from scripts.architecture_check import run_checks

            failures = run_checks()
            if failures:
                raise ValueError("; ".join(failures))
            return "rules=3"

        self.check("架构依赖规则", _run)

    def run(self) -> int:
        try:
            self.check_module_imports()
            self.check_application_init()
            self.check_store_read_entry()
            self.check_presenter_entry()
            self.check_score_overview_sorting_and_low_score()
            self.check_online_score_overview_sorting_and_low_score()
            self.check_online_wear_view_uses_synced_wear()
            self.check_service_entry()
            self.check_application_entry()
            self.check_static_assets()
            self.check_page_routes()
            self.check_read_only_routes_do_not_open_excel()
            self.check_excel_date_notes_import()
            self.check_excel_refresh_side_effect_guards()
            self.check_member_status_survives_excel_refresh()
            self.check_existing_config_members_are_not_reseeded()
            self.check_supabase_profit_uses_formal_column()
            self.check_online_entry_reports_supabase_request_failure()
            self.check_wear_threshold_uses_saved_config()
            self.check_excel_refresh_rejects_unknown_score_date()
            self.check_delete_score_date()
            self.check_delete_score_date_rolls_back_on_export_error()
            self.check_supabase_push_saves_posted_entry_once()
            self.check_online_manual_wear_syncs_to_local()
            self.check_recent_low_score_format()
            self.check_architecture_rules()
        finally:
            if self._temp_root and self._temp_root.exists():
                shutil.rmtree(self._temp_root, ignore_errors=True)
            if TEST_TEMP_ROOT.exists():
                shutil.rmtree(TEST_TEMP_ROOT, ignore_errors=True)

        passed = 0
        failed = 0
        for ok, name, detail in self.results:
            if ok:
                passed += 1
                print(f"[PASS] {name}")
                if detail:
                    print(f"       {detail}")
            else:
                failed += 1
                print(f"[FAIL] {name}")
                print(f"       {detail}")

        print()
        print(f"总计: {passed} 通过, {failed} 失败")
        return 1 if failed else 0


def main() -> int:
    runner = SmokeCheckRunner()
    return runner.run()


if __name__ == "__main__":
    raise SystemExit(main())
