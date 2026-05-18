from __future__ import annotations

from datetime import datetime
import os
import subprocess
import sys
from typing import Any, Callable

from flask import abort, flash, jsonify, redirect, render_template, request, url_for

from .ui_text import MESSAGES


def register_score_routes(
    app,
    store,
    *,
    read_only_mode: bool,
    flash_remote_sync_needed: Callable[[], None],
) -> None:
    def _build_score_page_members(entries: list[dict[str, str]] | None = None) -> list[dict[str, str]]:
        base_members = store.get_online_active_members() if read_only_mode else store.get_active_members()
        active_members = [dict(member) for member in base_members]
        entry_map = {item["name"]: item for item in (entries or [])}
        for member in active_members:
            saved_entry = entry_map.get(member["name"], {})
            member["score"] = saved_entry.get("score", "")
            member["before_balance"] = saved_entry.get("before_balance", "")
            member["after_balance"] = saved_entry.get("after_balance", "")
            member["manual_wear"] = saved_entry.get("manual_wear", "")
            member["income"] = saved_entry.get("income", "")
            member["other_expense"] = saved_entry.get("other_expense", "")
        return active_members

    def _score_rows_to_entries(rows: list[dict[str, Any]]) -> list[dict[str, str]]:
        entries = []
        for row in rows:
            member_name = str(row.get("member_name", "")).strip()
            if not member_name:
                continue
            entries.append(
                {
                    "name": member_name,
                    "score": str(row.get("score", "") or ""),
                    "before_balance": str(row.get("before_balance", "") or ""),
                    "after_balance": str(row.get("after_balance", "") or ""),
                    "manual_wear": str(row.get("manual_wear", "") or ""),
                    "income": str(row.get("income", "") or ""),
                    "other_expense": str(row.get("other_expense", "") or ""),
                }
            )
        return entries

    def _render_score_entry(
        *,
        selected_date: str,
        entries: list[dict[str, str]] | None = None,
        date_notes: dict[str, str] | None = None,
        overwrite_prompt: dict[str, Any] | None = None,
    ):
        template_name = "mobile_scores.html" if read_only_mode else "scores.html"
        page_entries = entries
        if page_entries is None:
            page_entries = _score_rows_to_entries(store.get_score_rows_for_date(selected_date))
        page_notes = date_notes
        if page_notes is None:
            page_notes = {"note1": "", "note2": "", "note3": ""} if read_only_mode else store.get_score_date_notes(selected_date)
        return render_template(
            template_name,
            active_members=_build_score_page_members(page_entries),
            score_summary=store.get_online_score_summary() if read_only_mode else store.get_score_summary(),
            selected_date=selected_date,
            date_notes=page_notes,
            overwrite_prompt=overwrite_prompt,
            quick_wear_values=[1, 2, 3, 4, 5],
        )

    def _get_requested_score_date() -> str:
        raw_date = request.args.get("date", "").strip()
        if raw_date:
            return raw_date
        return store.get_online_next_score_date() if read_only_mode else store.get_next_score_date()

    def _get_selected_date_from_form() -> str:
        return request.form.get("date", "").strip() or datetime.now().strftime("%Y-%m-%d")

    def _flash_save_score_result(selected_date: str, result: dict[str, Any]) -> None:
        if result["saved_date"] != selected_date:
            flash(MESSAGES["score_saved_shifted"].format(selected_date=selected_date, **result), "success")
        else:
            flash(MESSAGES["score_saved"].format(**result), "success")

    @app.route("/")
    def index():
        return redirect(url_for("score_entry"))

    @app.route("/scores")
    def score_entry():
        return _render_score_entry(selected_date=_get_requested_score_date())

    @app.route("/score-overview")
    def score_overview():
        if read_only_mode:
            data = store.get_mobile_overview()
            return render_template(
                "mobile_score_overview.html",
                score_sheet=data["score_sheet_view"],
                score_summary=data["score_summary"],
            )
        return render_template(
            "score_overview.html",
            score_sheet=store.get_score_sheet_view(),
            score_summary=store.get_score_summary(),
        )

    @app.route("/api/score-overview")
    def score_overview_api():
        if not read_only_mode:
            data = {
                "score_sheet_view": store.get_score_sheet_view(),
                "score_summary": store.get_score_summary(),
                "next_score_date": store.get_next_score_date(),
                "active_members": store.get_active_members(),
            }
        else:
            data = store.get_mobile_overview()
        return jsonify(data)

    @app.post("/scores/open-excel")
    def open_excel_file():
        if read_only_mode:
            abort(403)
        workbook_path = store.workbook_path
        try:
            if os.name == "nt" and hasattr(os, "startfile"):
                os.startfile(workbook_path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(workbook_path)])
            else:
                subprocess.Popen(["xdg-open", str(workbook_path)])
            flash(MESSAGES["excel_opened"].format(filename=workbook_path.name), "success")
        except OSError:
            app.logger.exception("Failed to open workbook: %s", workbook_path)
            flash(MESSAGES["excel_open_failed"].format(filename=workbook_path.name), "error")
        return redirect(url_for("score_entry"))

    @app.post("/scores/save")
    def save_scores():
        selected_date = _get_selected_date_from_form()
        existing_notes = {"note1": "", "note2": "", "note3": ""} if read_only_mode else store.get_score_date_notes(selected_date)
        submission = store.save_daily_entry(
            request.form,
            selected_date,
            existing_notes=existing_notes,
        )
        if not submission["ok"]:
            flash(submission["error"], "error")
            return _render_score_entry(
                selected_date=selected_date,
                entries=submission["entries"],
                date_notes=submission.get("notes"),
            )
        _flash_save_score_result(selected_date, submission["result"])
        if not read_only_mode:
            if submission.get("export_error") is not None:
                app.logger.error("Failed to update workbook after score save: %s", submission["export_error"])
                flash(MESSAGES["excel_save_failed"].format(filename=store.workbook_path.name), "error")
            flash_remote_sync_needed()
        return redirect(url_for("score_entry", date=submission["result"]["saved_date"], draft_saved="1"))

    @app.post("/scores/delete-date")
    def delete_score_date():
        if read_only_mode:
            abort(403)
        selected_date = _get_selected_date_from_form()
        date_error = store.daily_entry_service.validate_selected_date(selected_date)
        if date_error is not None:
            flash(date_error, "error")
            return redirect(url_for("score_entry", date=selected_date))
        try:
            result = store.delete_score_date(selected_date)
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("score_entry", date=selected_date))
        if result.get("deleted_rows", 0) or result.get("deleted_notes", 0):
            flash(MESSAGES["score_date_deleted"].format(date=selected_date), "success")
            flash_remote_sync_needed()
        else:
            flash(MESSAGES["score_date_delete_empty"].format(date=selected_date), "info")
        return redirect(url_for("score_entry", date=selected_date))

    @app.post("/scores/refresh-from-excel")
    def refresh_from_excel():
        if read_only_mode:
            abort(403)
        try:
            result = store.refresh_local_database()
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("score_entry"))
        if result.get("changed"):
            flash(MESSAGES["excel_refresh_changed"], "warning")
        else:
            flash(MESSAGES["excel_refresh_unchanged"], "success")
        return redirect(url_for("score_entry"))

