from __future__ import annotations

import sys
import shutil
import uuid
from pathlib import Path

import pytest
from flask import Flask
from werkzeug.datastructures import MultiDict

sys.dont_write_bytecode = True

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture(autouse=True)
def isolated_environment(monkeypatch):
    for key in (
        "SUPABASE_URL",
        "SUPABASE_SERVICE_ROLE_KEY",
        "SUPABASE_KEY",
        "BM2_READ_ONLY",
        "VERCEL",
        "BM2_PASSWORD",
    ):
        monkeypatch.delenv(key, raising=False)
    yield


@pytest.fixture
def test_workspace():
    root = PROJECT_ROOT / ".tmp_test_workspaces" / "pytest"
    root.mkdir(parents=True, exist_ok=True)
    workspace = root / uuid.uuid4().hex
    workspace.mkdir()
    try:
        yield workspace
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


@pytest.fixture
def store_base(test_workspace):
    return test_workspace / "bm2_app"


@pytest.fixture
def store(store_base):
    from bm2.services.store_application import StoreApplication

    return StoreApplication(store_base)


@pytest.fixture
def local_db(store):
    return store._context.local_db


@pytest.fixture
def route_store(test_workspace):
    from bm2.services.store_application import StoreApplication

    return StoreApplication(test_workspace / "routes_app")


@pytest.fixture
def flask_app(route_store):
    from bm2.web import register_routes

    app = Flask(
        __name__,
        template_folder=str(PROJECT_ROOT / "templates"),
        static_folder=str(PROJECT_ROOT / "static"),
    )
    app.secret_key = "pytest"
    register_routes(app, route_store)
    return app


@pytest.fixture
def client(flask_app):
    return flask_app.test_client()


def score_form(names: list[str], value: object = 5, *, date: str | None = None) -> MultiDict:
    data: dict[str, str] = {}
    if date is not None:
        data["date"] = date
    for name in names:
        data[f"score_{name}"] = str(value)
        data[f"before_{name}"] = ""
        data[f"after_{name}"] = ""
        data[f"manual_wear_{name}"] = ""
        data[f"income_{name}"] = ""
        data[f"other_expense_{name}"] = ""
    return MultiDict(data)


def add_members(store, names: list[str]) -> list[dict[str, str]]:
    for name in names:
        store.member_service.add_member(name)
    return [{"name": name, "status": "启用"} for name in names]
