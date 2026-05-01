from __future__ import annotations

import compileall
import os
import shutil
import subprocess
import sys
import uuid
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEST_TEMP_ROOT = PROJECT_ROOT / ".tmp_test_workspaces"


def _env() -> dict[str, str]:
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env.setdefault("PYTEST_DISABLE_PLUGIN_AUTOLOAD", "1")
    env.setdefault("UV_CACHE_DIR", str(TEST_TEMP_ROOT / "uv-cache"))
    for key in ("SUPABASE_URL", "SUPABASE_SERVICE_ROLE_KEY", "SUPABASE_KEY", "BM2_READ_ONLY", "VERCEL"):
        env.pop(key, None)
    return env


def _print_step(name: str) -> None:
    print()
    print(f"== {name} ==", flush=True)


def _run_command(name: str, command: list[str]) -> None:
    _print_step(name)
    print(" ".join(command), flush=True)
    result = subprocess.run(command, cwd=PROJECT_ROOT, env=_env(), text=True)
    if result.returncode != 0:
        raise SystemExit(result.returncode)


def _python_command() -> str:
    candidates = [shutil.which("python"), getattr(sys, "_base_executable", None), sys.executable]
    temp_root_text = str(TEST_TEMP_ROOT).lower()
    for candidate in candidates:
        if candidate and temp_root_text not in str(candidate).lower():
            return str(candidate)
    return sys.executable


def _compile_temp_copy() -> None:
    _print_step("compileall")
    TEST_TEMP_ROOT.mkdir(exist_ok=True)
    target = TEST_TEMP_ROOT / f"compile_{uuid.uuid4().hex}"
    target.mkdir()
    try:
        for name in ("app.py", "bm2", "scripts", "static", "templates", "docs", "README.md", "AGENTS.md"):
            source = PROJECT_ROOT / name
            if not source.exists():
                continue
            destination = target / name
            if source.is_dir():
                shutil.copytree(
                    source,
                    destination,
                    ignore=shutil.ignore_patterns("__pycache__", ".pytest_cache", ".tmp_test_workspaces"),
                )
            else:
                shutil.copy2(source, destination)
        if not compileall.compile_dir(target, quiet=1):
            raise SystemExit(1)
        print("compileall temp copy: PASS")
    finally:
        shutil.rmtree(target, ignore_errors=True)


def _pytest_command() -> list[str]:
    python = _python_command()
    probe = subprocess.run(
        [python, "-m", "pytest", "--version"],
        cwd=PROJECT_ROOT,
        env=_env(),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    if probe.returncode == 0:
        return [python, "-m", "pytest", "-q"]
    uv = shutil.which("uv")
    if uv:
        return [uv, "run", "--quiet", "--with", "pytest", "pytest", "-q"]
    return [python, "-m", "pytest", "-q"]


def _keyword_scan() -> None:
    _print_step("retired keyword scan")
    blocked = (
        "__getattr__",
        "record_local_score_submission",
        "write_entries_to_excel",
        "process_score_submission",
        "bm2_cloud",
        "cloud_entry_writer",
        "legacy_json_sync",
        "json_sync",
        "blob_sync",
    )
    offenders: list[str] = []
    for path in (PROJECT_ROOT / "bm2").rglob("*.py"):
        if "__pycache__" in path.parts:
            continue
        text = path.read_text(encoding="utf-8-sig")
        for marker in blocked:
            if marker in text:
                offenders.append(f"{path.relative_to(PROJECT_ROOT)} contains {marker}")
    retired_files = (
        PROJECT_ROOT / "bm2" / "store.py",
        PROJECT_ROOT / "bm2" / "services" / "score_service.py",
        PROJECT_ROOT / "bm2" / "store_reader_facade.py",
        PROJECT_ROOT / "bm2" / "store_writer_facade.py",
    )
    offenders.extend(f"{path.relative_to(PROJECT_ROOT)} exists" for path in retired_files if path.exists())
    if offenders:
        for offender in offenders:
            print(f"[FAIL] {offender}")
        raise SystemExit(1)
    print("retired keyword scan: PASS")


def _cleanup_python_caches() -> None:
    _print_step("cleanup python caches")
    removed = 0
    for path in PROJECT_ROOT.rglob("__pycache__"):
        if ".tmp_test_workspaces" in path.parts:
            continue
        shutil.rmtree(path, ignore_errors=True)
        removed += 1
    for path in PROJECT_ROOT.rglob("*.pyc"):
        if ".tmp_test_workspaces" in path.parts:
            continue
        path.unlink(missing_ok=True)
        removed += 1
    print(f"removed={removed}")


def main() -> int:
    python = _python_command()
    _cleanup_python_caches()
    _compile_temp_copy()
    _run_command("smoke_check", [python, "scripts/smoke_check.py"])
    _run_command("architecture_check", [python, "scripts/architecture_check.py"])
    _run_command("pytest", _pytest_command())
    _cleanup_python_caches()
    _run_command("architecture_check after pytest", [python, "scripts/architecture_check.py"])
    _keyword_scan()
    print()
    print("[PASS] all checks")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
