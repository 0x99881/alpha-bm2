from __future__ import annotations

from pathlib import Path
import shutil


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
PUBLIC_DIR = ROOT / "public"


def main() -> None:
    if PUBLIC_DIR.exists():
        shutil.rmtree(PUBLIC_DIR)
    shutil.copytree(STATIC_DIR, PUBLIC_DIR)


if __name__ == "__main__":
    main()
