from __future__ import annotations

from pathlib import Path

from scripts.architecture_check import run_checks


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_architecture_check_has_no_failures():
    assert run_checks() == []


def test_retired_entry_files_do_not_exist():
    retired = [
        PROJECT_ROOT / "bm2" / "store.py",
        PROJECT_ROOT / "bm2" / "services" / "score_service.py",
        PROJECT_ROOT / "bm2" / "store_reader_facade.py",
        PROJECT_ROOT / "bm2" / "store_writer_facade.py",
    ]
    assert [path for path in retired if path.exists()] == []


def test_retired_score_wrappers_and_magic_forwarding_do_not_return():
    blocked = (
        "__getattr__",
        "record_local_score_submission",
        "write_entries_to_excel",
        "process_score_submission",
    )
    offenders: list[str] = []
    for path in (PROJECT_ROOT / "bm2").rglob("*.py"):
        if "__pycache__" in path.parts:
            continue
        text = path.read_text(encoding="utf-8-sig")
        for marker in blocked:
            if marker in text:
                offenders.append(f"{path.relative_to(PROJECT_ROOT)}:{marker}")
    assert offenders == []
