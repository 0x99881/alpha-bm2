from pathlib import Path
import locale
import os
import sys

from flask import Flask

from bm2.store import ExcelStore
from bm2.web import register_routes

BASE_DIR = Path(__file__).resolve().parent


if sys.platform.startswith("win"):
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    try:
        locale.setlocale(locale.LC_ALL, "")
    except locale.Error:
        pass
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


app = Flask(__name__)
app.secret_key = "bm2-local-secret"
store = ExcelStore(BASE_DIR)
register_routes(app, store)

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=False)
