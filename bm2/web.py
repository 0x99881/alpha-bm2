from __future__ import annotations

from datetime import datetime
import hmac
import os
import subprocess
import sys
from typing import Any

from flask import abort, flash, jsonify, redirect, render_template, request, Response, send_from_directory, url_for

from .services import process_score_submission
from .store import DISABLED, ENABLED, ExcelStore
from .ui_text import JS_UI_TEXT, MESSAGES, UI_TEXT


def register_routes(app, store: ExcelStore) -> None:
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

    def _build_score_page_members(entries: list[dict[str, str]] | None = None) -> list[dict[str, str]]:
        base_members = store.get_online_active_members() if read_only_mode else store.get_active_members()
        active_members = [dict(member) for member in base_members]
        entry_map = {item['name']: item for item in (entries or [])}
        for member in active_members:
            saved_entry = entry_map.get(member['name'], {})
            member['score'] = saved_entry.get('score', '')
            member['before_balance'] = saved_entry.get('before_balance', '')
            member['after_balance'] = saved_entry.get('after_balance', '')
            member['manual_wear'] = saved_entry.get('manual_wear', '')
            member['income'] = saved_entry.get('income', '')
            member['other_expense'] = saved_entry.get('other_expense', '')
        return active_members

    def _render_score_entry(
        *,
        selected_date: str,
        entries: list[dict[str, str]] | None = None,
        overwrite_prompt: dict[str, Any] | None = None,
    ):
        template_name = 'mobile_scores.html' if read_only_mode else 'scores.html'
        return render_template(
            template_name,
            active_members=_build_score_page_members(entries),
            score_summary=store.get_online_score_summary() if read_only_mode else store.get_score_summary(),
            selected_date=selected_date,
            overwrite_prompt=overwrite_prompt,
        )

    def _get_requested_score_date() -> str:
        raw_date = request.args.get('date', '').strip()
        if raw_date:
            return raw_date
        return store.get_online_next_score_date() if read_only_mode else store.get_next_score_date()

    def _get_selected_date_from_form() -> str:
        return request.form.get('date', '').strip() or datetime.now().strftime('%Y-%m-%d')

    def _get_overwrite_step_from_form() -> int:
        raw_value = str(request.form.get('overwrite_step', '0') or '0').strip()
        try:
            return max(0, int(raw_value))
        except ValueError:
            return 0

    def _flash_save_score_result(selected_date: str, result: dict[str, Any]) -> None:
        if result['saved_date'] != selected_date:
            flash(MESSAGES['score_saved_shifted'].format(selected_date=selected_date, **result), 'success')
        else:
            flash(MESSAGES['score_saved'].format(**result), 'success')

    def _redirect_back_to_score_entry():
        return redirect(request.referrer or url_for('score_entry'))

    def _ensure_supabase_configured() -> bool:
        if store.is_supabase_configured():
            return True
        flash(MESSAGES['supabase_not_configured'], 'error')
        return False

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

    @app.route('/')
    def index():
        return redirect(url_for('score_entry'))

    @app.route('/scores')
    def score_entry():
        return _render_score_entry(selected_date=_get_requested_score_date())

    @app.route('/score-overview')
    def score_overview():
        return render_template(
            'mobile_score_overview.html' if read_only_mode else 'score_overview.html',
            score_sheet=store.get_online_score_sheet_view() if read_only_mode else store.get_score_sheet_view(),
            score_summary=store.get_online_score_summary() if read_only_mode else store.get_score_summary(),
        )

    @app.route('/api/score-overview')
    def score_overview_api():
        if not read_only_mode:
            data = {
                'score_sheet_view': store.get_score_sheet_view(),
                'score_summary': store.get_score_summary(),
                'next_score_date': store.get_next_score_date(),
                'active_members': store.get_active_members(),
            }
        else:
            data = store.get_mobile_overview()
        return jsonify(data)

    @app.post('/scores/open-excel')
    def open_excel_file():
        if read_only_mode:
            abort(403)
        workbook_path = store.workbook_path
        try:
            if os.name == 'nt' and hasattr(os, 'startfile'):
                os.startfile(workbook_path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(workbook_path)])
            else:
                subprocess.Popen(['xdg-open', str(workbook_path)])
            flash(MESSAGES['excel_opened'].format(filename=workbook_path.name), 'success')
        except OSError:
            app.logger.exception("Failed to open workbook: %s", workbook_path)
            flash(MESSAGES['excel_open_failed'].format(filename=workbook_path.name), 'error')
        return redirect(url_for('score_entry'))

    @app.post('/scores/save')
    def save_scores():
        selected_date = _get_selected_date_from_form()
        active_members = store.get_online_active_members() if read_only_mode else store.get_active_members()
        entries = store.daily_entry_service.build_entries(active_members, request.form)
        validation_error = store.daily_entry_service.validate_submission(active_members, request.form, selected_date, entries)
        if validation_error is not None:
            flash(validation_error, 'error')
            return _render_score_entry(selected_date=selected_date, entries=entries)
        overwrite_step = _get_overwrite_step_from_form()
        if store.has_existing_score_date(selected_date) and overwrite_step < 2:
            prompt_message = (
                MESSAGES['score_overwrite_stage_one']
                if overwrite_step == 0
                else MESSAGES['score_overwrite_stage_two']
            )
            return _render_score_entry(
                selected_date=selected_date,
                entries=entries,
                overwrite_prompt={
                    'message': prompt_message,
                    'next_step': overwrite_step + 1,
                    'confirm_label': MESSAGES['score_overwrite_confirm'],
                },
            )
        submission = process_score_submission(store, active_members, request.form, selected_date)
        if not submission['ok']:
            flash(submission['error'], 'error')
            return _render_score_entry(selected_date=selected_date, entries=submission['entries'])
        sync_result = None
        if not read_only_mode:
            sync_result = store.after_local_score_submission(submission['result'])
        _flash_save_score_result(selected_date, submission['result'])
        if not read_only_mode:
            if sync_result and sync_result.get("ok"):
                push = sync_result.get("push", {})
                flash(
                    MESSAGES['remote_auto_sync_done'].format(
                        members=push.get("members", 0),
                        score_entries=push.get("score_entries", 0),
                    ),
                    'success',
                )
            elif sync_result and not sync_result.get("configured"):
                flash(MESSAGES['remote_auto_sync_not_configured'], 'warning')
            else:
                flash(MESSAGES['remote_auto_sync_failed'], 'warning')
        return redirect(url_for('score_overview' if read_only_mode else 'score_entry'))

    @app.post('/scores/refresh-from-excel')
    def refresh_from_excel():
        if read_only_mode:
            abort(403)
        result = store.refresh_local_database()
        if result.get("changed"):
            flash(MESSAGES['excel_refresh_changed'], 'warning')
        else:
            flash(MESSAGES['excel_refresh_unchanged'], 'success')
        return redirect(url_for('score_entry'))

    @app.post('/cycles/new')
    def create_new_cycle():
        if read_only_mode:
            abort(403)
        date_text = request.form.get('start_date', '').strip()
        try:
            new_filename = store.create_new_cycle(date_text)
            flash(MESSAGES['new_cycle_created'].format(filename=new_filename), 'success')
            _flash_remote_sync_needed()
            return redirect(url_for('score_entry', date=date_text))
        except ValueError as exc:
            flash(str(exc), 'error')
        return _render_score_entry(selected_date=date_text or store.get_next_score_date())

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
        members = store.get_members()
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

    @app.post('/members/reorder')
    def reorder_members():
        if read_only_mode:
            abort(403)
        payload = request.get_json(silent=True) or {}
        ordered_names = payload.get('ordered_names') or []
        if not isinstance(ordered_names, list):
            return jsonify({'ok': False}), 400
        store.member_service.reorder_active_members([str(name).strip() for name in ordered_names])
        return jsonify({'ok': True})

    @app.route('/members')
    def members():
        if read_only_mode:
            return redirect(url_for('score_overview'))
        all_members = store.get_members()
        enabled_count = sum(1 for item in all_members if item['status'] == ENABLED)
        return render_template('members.html', members=all_members, enabled_count=enabled_count)

    @app.post('/members/add')
    def add_member():
        if read_only_mode:
            abort(403)
        name = request.form.get('name', '').strip()
        note = request.form.get('note', '').strip()
        try:
            store.member_service.add_member(name, note)
            flash(MESSAGES['member_added'].format(name=name), 'success')
            _flash_remote_sync_needed()
        except ValueError as exc:
            flash(str(exc), 'error')
        return redirect(url_for('members'))

    @app.post('/members/update')
    def update_member():
        if read_only_mode:
            abort(403)
        name = request.form.get('name', '').strip()
        note = request.form.get('note', '').strip()
        status = request.form.get('status', '').strip()
        try:
            store.member_service.update_member(name, note, status=status)
            if status == ENABLED:
                flash(MESSAGES['member_restored'].format(name=name), 'success')
            elif status == DISABLED:
                flash(MESSAGES['member_disabled'].format(name=name), 'success')
            else:
                flash(MESSAGES['member_note_updated'].format(name=name), 'success')
            _flash_remote_sync_needed()
        except ValueError as exc:
            flash(str(exc), 'error')
        return redirect(url_for('members'))

    @app.post('/members/delete')
    def delete_member():
        if read_only_mode:
            abort(403)
        name = request.form.get('name', '').strip()
        try:
            store.member_service.delete_member(name)
            flash(MESSAGES['member_deleted'].format(name=name), 'success')
            _flash_remote_sync_needed()
        except ValueError as exc:
            flash(str(exc), 'error')
        return redirect(url_for('members'))

    @app.post('/cloud-sync')
    def sync_now():
        if read_only_mode:
            abort(403)
        if not _ensure_supabase_configured():
            return _redirect_back_to_score_entry()
        result = store.supabase_sync()
        pull, push = result['pull'], result['push']
        flash(
            MESSAGES['supabase_sync_scores_done'].format(
                pull_score_entries=pull["score_entries"],
                push_score_entries=push["score_entries"],
            ),
            'success',
        )
        return _redirect_back_to_score_entry()

    @app.post('/supabase-sync')
    def supabase_sync_now():
        if read_only_mode:
            abort(403)
        if not _ensure_supabase_configured():
            return _redirect_back_to_score_entry()
        result = store.supabase_sync()
        pull, push = result['pull'], result['push']
        flash(
            MESSAGES['supabase_sync_done'].format(
                pull_members=pull["members"],
                pull_score_entries=pull["score_entries"],
                push_members=push["members"],
                push_score_entries=push["score_entries"],
            ),
            'success',
        )
        return _redirect_back_to_score_entry()

    @app.route('/members/<name>')
    def member_detail(name: str):
        if read_only_mode:
            return redirect(url_for('score_overview'))
        year = request.args.get('year', type=int) or datetime.now().year
        month = request.args.get('month', type=int) or datetime.now().month
        return redirect(url_for('profit_calendar', name=name, year=year, month=month))
