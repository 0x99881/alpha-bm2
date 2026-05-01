from __future__ import annotations

from flask import abort, flash, redirect, request, url_for

from .ui_text import MESSAGES


def register_sync_routes(app, store, *, read_only_mode: bool) -> None:
    def _redirect_back_to_score_entry():
        return redirect(request.referrer or url_for("score_entry"))

    def _ensure_supabase_configured() -> bool:
        if store.is_supabase_configured():
            return True
        flash(MESSAGES["supabase_not_configured"], "error")
        return False

    @app.post("/supabase-push")
    def supabase_push_now():
        if read_only_mode:
            abort(403)
        if not _ensure_supabase_configured():
            return _redirect_back_to_score_entry()
        try:
            push = store.supabase_push()
        except Exception:
            app.logger.exception("Supabase push failed")
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
        try:
            pull = store.supabase_pull()
        except Exception:
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
