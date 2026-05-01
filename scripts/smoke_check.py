from __future__ import annotations

import importlib
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
            summary = store.score_presenter.build_score_summary()
            required_keys = {"latest_column", "window_size", "rankings"}
            missing = required_keys - set(summary.keys())
            if missing:
                raise ValueError(f"missing keys: {sorted(missing)}")
            return f"summary_keys={sorted(summary.keys())}"

        self.check("presenter 入口", _run)

    def check_service_entry(self) -> None:
        def _run():
            store = self.build_store()
            active_members = store.get_active_members()
            selected_date = store.get_next_score_date()
            result = store.daily_entry_service.process_submission(active_members, {}, selected_date)
            if "ok" not in result:
                raise ValueError("missing ok flag")
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
                if failures:
                    raise ValueError(", ".join(failures))
                return f"routes={len(routes)}"
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

        self.check("页面路由", _run)

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
            self.check_service_entry()
            self.check_application_entry()
            self.check_static_assets()
            self.check_page_routes()
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
