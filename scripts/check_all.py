import os
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEST_TEMP_ROOT = PROJECT_ROOT / ".tmp_test_workspaces"


def _env() -> dict[str, str]:
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
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
    _run_command("smoke_check", [python, "scripts/smoke_check.py"])
    _run_command("architecture_check", [python, "scripts/architecture_check.py"])
    _cleanup_python_caches()
    print()
    print("[PASS] all checks")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
