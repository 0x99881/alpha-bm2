from __future__ import annotations

from flask import abort, flash, redirect, request, url_for

from .repositories.supabase_client import SUPABASE_REQUEST_ERRORS
from .ui_text import MESSAGES, display_workbook_filename


def _is_transient_network_error(exc: BaseException) -> bool:
    return exc.__class__.__module__.startswith("httpx")


def register_sync_routes(app, store, *, read_only_mode: bool) -> None:
    def _redirect_back_to_score_entry():
        return redirect(request.referrer or url_for("score_entry"))

    def _posted_score_form_date() -> str:
        return request.form.get("date", "").strip()

    def _looks_like_score_form_post() -> bool:
        if not _posted_score_form_date():
            return False
        prefixes = (
            "score_",
            "before_",
            "after_",
            "manual_wear_",
            "income_",
            "other_expense_",
            "date_note_",
        )
        return any(
            str(key).startswith(prefixes) and str(request.form.get(key, "") or "").strip() != ""
            for key in request.form.keys()
        )

    def _save_posted_score_form_if_present() -> tuple[bool, str]:
        if not _looks_like_score_form_post():
            return True, ""
        selected_date = _posted_score_form_date()
        submission = store.save_daily_entry(
            request.form,
            selected_date,
        )
        if not submission["ok"]:
            flash(submission["error"], "error")
            return False, selected_date
        saved_date = submission["result"]["saved_date"]
        if submission.get("export_error") is not None:
            app.logger.error("Failed to update workbook before Supabase push: %s", submission["export_error"])
            flash(MESSAGES["excel_save_failed"].format(filename=display_workbook_filename(store.workbook_path.name)), "error")
        return True, saved_date

    def _ensure_supabase_configured() -> bool:
        if store.is_supabase_configured():
            return True
        flash(MESSAGES["supabase_not_configured"], "error")
        return False

    @app.post("/supabase-push")
    def supabase_push_now():
        if read_only_mode:
            abort(403)
        saved_ok, saved_date = _save_posted_score_form_if_present()
        if not saved_ok:
            return redirect(url_for("score_entry", date=saved_date) if saved_date else url_for("score_entry"))
        if not _ensure_supabase_configured():
            return redirect(url_for("score_entry", date=saved_date) if saved_date else url_for("score_entry"))
        try:
            push = store.supabase_push()
        except ValueError as exc:
            flash(str(exc), "error")
            return _redirect_back_to_score_entry()
        except SUPABASE_REQUEST_ERRORS as exc:
            app.logger.exception("Supabase push failed")
            if _is_transient_network_error(exc):
                flash(MESSAGES["supabase_upload_failed_timeout"], "error")
            else:
                flash(MESSAGES["supabase_upload_failed"], "error")
            return _redirect_back_to_score_entry()
        flash(
            MESSAGES["supabase_upload_done"].format(
                push_members=push["members"],
                push_score_entries=push["score_entries"],
            ),
            "success",
        )
        return _redirect_back_to_score_entry()

    @app.post("/supabase-pull")
    def supabase_pull_now():
        if read_only_mode:
            abort(403)
        if not _ensure_supabase_configured():
            return _redirect_back_to_score_entry()
        force_full = request.form.get("force_full", "").strip() == "1"
        try:
            pull = store.supabase_pull(force_full=force_full)
        except SUPABASE_REQUEST_ERRORS:
            app.logger.exception("Supabase pull failed")
            flash(MESSAGES["supabase_download_failed"], "error")
            return _redirect_back_to_score_entry()
        flash(
            MESSAGES["supabase_download_done"].format(
                pull_members=pull["members"],
                pull_score_entries=pull["score_entries"],
            ),
            "success",
        )
        return _redirect_back_to_score_entry()
