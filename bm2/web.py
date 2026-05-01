from __future__ import annotations

from datetime import datetime
import hmac
import os
from typing import Any

from flask import abort, flash, redirect, render_template, request, Response, send_from_directory, url_for

from .constants import DISABLED, ENABLED
from .ui_text import JS_UI_TEXT, MESSAGES, UI_TEXT
from .web_member_routes import register_member_routes
from .web_score_routes import register_score_routes
from .web_sync_routes import register_sync_routes


def register_routes(app, store) -> None:
    read_only_mode = bool(getattr(store, 'read_only', False))

    _app_password = os.environ.get("BM2_PASSWORD", "")
    if _app_password:
        @app.before_request
        def _require_auth():
            auth = request.authorization
            if not auth or not hmac.compare_digest(str(auth.password), _app_password):
                return Response(
                    "BM2 requires authentication.",
                    401,
                    {"WWW-Authenticate": 'Basic realm="BM2"'},
                )

    if 'asset_file' not in app.view_functions and app.static_folder:
        @app.get('/assets/<path:filename>')
        def asset_file(filename: str):
            asset_root = app.static_folder
            if not asset_root or not os.path.isfile(os.path.join(asset_root, filename)):
                abort(404)
            return send_from_directory(asset_root, filename)

    def _flash_remote_sync_needed() -> None:
        if store.is_supabase_configured():
            flash(MESSAGES['remote_sync_needed'], 'warning')
        else:
            flash(MESSAGES['remote_sync_needed_no_config'], 'warning')

    @app.context_processor
    def inject_shared_data() -> dict[str, Any]:
        return {
            'excel_filename': store.workbook_path.name,
            'quick_scores': store.get_quick_scores(),
            'asset_version': '20260428-01',
            'ui': UI_TEXT,
            'js_ui_text': JS_UI_TEXT,
            'enabled_status': ENABLED,
            'disabled_status': DISABLED,
            'read_only_mode': read_only_mode,
        }

    @app.after_request
    def disable_cache(response):
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
        if response.mimetype == 'text/html':
            response.headers['Content-Type'] = 'text/html; charset=utf-8'
        return response

    @app.route('/wear')
    def wear_entry():
        if read_only_mode:
            return redirect(url_for('score_overview'))
        return render_template('wear.html', wear_sheet=store.get_wear_sheet_view())

    @app.post('/wear/threshold')
    def save_wear_threshold():
        if read_only_mode:
            abort(403)
        threshold_text = request.form.get('wear_abnormal_threshold', '').strip()
        try:
            threshold = store.set_wear_abnormal_threshold(threshold_text)
            flash(MESSAGES['wear_threshold_saved'].format(threshold=threshold), 'success')
        except ValueError as exc:
            error_text = str(exc).strip()
            if error_text:
                flash(error_text, 'error')
            else:
                flash(MESSAGES['wear_threshold_invalid'], 'error')
        return redirect(url_for('wear_entry'))

    @app.route('/profit-calendar')
    def profit_calendar():
        if read_only_mode:
            return redirect(url_for('score_overview'))
        members = store.member_service.get_members()
        if not members:
            abort(404)
        member_names = [item['name'] for item in members]
        selected_name = request.args.get('name', '').strip() or 'all'
        if selected_name not in ['all', *member_names]:
            selected_name = 'all'
        year = request.args.get('year', type=int) or datetime.now().year
        month = request.args.get('month', type=int) or datetime.now().month
        return render_template(
            'profit_calendar.html',
            members=members,
            selected_name=selected_name,
            calendar_data=store.get_member_profit_calendar(selected_name, year, month),
        )

    register_score_routes(
        app,
        store,
        read_only_mode=read_only_mode,
        flash_remote_sync_needed=_flash_remote_sync_needed,
    )
    register_member_routes(
        app,
        store,
        read_only_mode=read_only_mode,
        flash_remote_sync_needed=_flash_remote_sync_needed,
    )
    register_sync_routes(app, store, read_only_mode=read_only_mode)
