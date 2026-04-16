from __future__ import annotations

import importlib
import re
import shutil
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


MODULES_TO_IMPORT = [
    "bm2.store",
    "bm2.workbook_manager",
    "bm2.record_repository",
    "bm2.services.score_service",
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
        from bm2.store import ExcelStore

        temp_dir = PROJECT_ROOT / ".smoke_tmp"
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        temp_dir.mkdir()
        self._temp_root = temp_dir
        self._store = ExcelStore(temp_dir)
        return self._store

    def check_module_imports(self) -> None:
        for module_name in MODULES_TO_IMPORT:
            self.check(f"import {module_name}", lambda module_name=module_name: importlib.import_module(module_name).__name__)

    def check_excel_store_init(self) -> None:
        def _run():
            store = self.build_store()
            return f"workbook={store.workbook_path.name}"

        self.check("ExcelStore 初始化", _run)

    def check_repository_entry(self) -> None:
        def _run():
            store = self.build_store()
            next_date = store.record_repository.get_next_score_date()
            wear_records = store.record_repository.get_member_wear_records("bb")
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
            from bm2.services.score_service import process_score_submission

            store = self.build_store()
            active_members = store.get_active_members()
            selected_date = store.get_next_score_date()
            result = process_score_submission(store, active_members, {}, selected_date)
            if "ok" not in result:
                raise ValueError("missing ok flag")
            return f"ok={result['ok']}"

        self.check("service 入口", _run)

    def check_store_entry(self) -> None:
        def _run():
            store = self.build_store()
            summary = store.get_score_summary()
            wear_view = store.get_wear_sheet_view()
            if "rankings" not in summary:
                raise ValueError("score summary missing rankings")
            if "rows" not in wear_view:
                raise ValueError("wear view missing rows")
            return f"rankings={len(summary['rankings'])}, wear_rows={wear_view['row_count']}"

        self.check("store 门面入口", _run)

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

    def run(self) -> int:
        try:
            self.check_module_imports()
            self.check_excel_store_init()
            self.check_repository_entry()
            self.check_presenter_entry()
            self.check_service_entry()
            self.check_store_entry()
            self.check_static_assets()
        finally:
            if self._temp_root and self._temp_root.exists():
                shutil.rmtree(self._temp_root, ignore_errors=True)

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
