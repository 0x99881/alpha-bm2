from __future__ import annotations

from datetime import datetime
from typing import Callable

from flask import abort, flash, jsonify, redirect, render_template, request, url_for

from .constants import DISABLED, ENABLED
from .ui_text import MESSAGES


def register_member_routes(
    app,
    store,
    *,
    read_only_mode: bool,
    flash_remote_sync_needed: Callable[[], None],
) -> None:
    @app.post("/members/reorder")
    def reorder_members():
        if read_only_mode:
            abort(403)
        payload = request.get_json(silent=True)
        if not isinstance(payload, dict):
            return jsonify({"ok": False}), 400
        ordered_names = payload.get("ordered_names") or []
        if not isinstance(ordered_names, list):
            return jsonify({"ok": False}), 400
        store.member_service.reorder_active_members([str(name).strip() for name in ordered_names])
        return jsonify({"ok": True})

    @app.route("/members")
    def members():
        if read_only_mode:
            return redirect(url_for("score_overview"))
        all_members = store.member_service.get_members()
        enabled_count = sum(1 for item in all_members if item["status"] == ENABLED)
        return render_template("members.html", members=all_members, enabled_count=enabled_count)

    @app.post("/members/add")
    def add_member():
        if read_only_mode:
            abort(403)
        name = request.form.get("name", "").strip()
        note = request.form.get("note", "").strip()
        try:
            store.member_service.add_member(name, note)
            flash(MESSAGES["member_added"].format(name=name), "success")
            flash_remote_sync_needed()
        except ValueError as exc:
            flash(str(exc), "error")
        return redirect(url_for("members"))

    @app.post("/members/update")
    def update_member():
        if read_only_mode:
            abort(403)
        name = request.form.get("name", "").strip()
        note = request.form.get("note", "").strip()
        status = request.form.get("status", "").strip()
        try:
            store.member_service.update_member(name, note, status=status)
            if status == ENABLED:
                flash(MESSAGES["member_restored"].format(name=name), "success")
            elif status == DISABLED:
                flash(MESSAGES["member_disabled"].format(name=name), "success")
            else:
                flash(MESSAGES["member_note_updated"].format(name=name), "success")
            flash_remote_sync_needed()
        except ValueError as exc:
            flash(str(exc), "error")
        return redirect(url_for("members"))

    @app.post("/members/delete")
    def delete_member():
        if read_only_mode:
            abort(403)
        name = request.form.get("name", "").strip()
        try:
            store.member_service.delete_member(name)
            flash(MESSAGES["member_deleted"].format(name=name), "success")
            flash_remote_sync_needed()
        except ValueError as exc:
            flash(str(exc), "error")
        return redirect(url_for("members"))

    @app.route("/members/<name>")
    def member_detail(name: str):
        if read_only_mode:
            return redirect(url_for("score_overview"))
        year = request.args.get("year", type=int) or datetime.now().year
        month = request.args.get("month", type=int) or datetime.now().month
        return redirect(url_for("profit_calendar", name=name, year=year, month=month))
