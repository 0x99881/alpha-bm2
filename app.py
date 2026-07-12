from pathlib import Path
import locale
import os
import sys

from flask import Flask, abort, send_from_directory

from bm2.services.store_application import StoreApplication
from bm2.web import register_routes

BASE_DIR = Path(__file__).resolve().parent


if sys.platform.startswith("win"):
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    try:
        locale.setlocale(locale.LC_ALL, "")
    except locale.Error:
        locale.setlocale(locale.LC_CTYPE, "C")
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def create_app() -> Flask:
    static_folder = "static"
    app = Flask(__name__, static_folder=static_folder)
    app.config["TEMPLATES_AUTO_RELOAD"] = True
    app.secret_key = "bm2-local-secret"
    asset_dir = BASE_DIR / static_folder

    @app.get("/assets/<path:filename>")
    def asset_file(filename: str):
        if not (asset_dir / filename).is_file():
            abort(404)
        return send_from_directory(asset_dir, filename)

    read_only = os.environ.get("BM2_READ_ONLY") == "1" or bool(os.environ.get("VERCEL"))
    store = StoreApplication(BASE_DIR, read_only=read_only)
    register_routes(app, store)
    return app


app = create_app()

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
