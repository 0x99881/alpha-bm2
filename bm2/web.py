from __future__ import annotations

from datetime import datetime
import os
import subprocess
import sys
from typing import Any

from flask import abort, flash, jsonify, redirect, render_template, request, url_for

from .store import DISABLED, ENABLED, ExcelStore
from .ui_text import JS_UI_TEXT, MESSAGES, UI_TEXT


def register_routes(app, store: ExcelStore) -> None:
    def _build_score_page_members(entries: list[dict[str, str]] | None = None) -> list[dict[str, str]]:
        active_members = [dict(member) for member in store.get_active_members()]
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

    def _render_score_entry(*, selected_date: str, entries: list[dict[str, str]] | None = None):
        return render_template(
            'scores.html',
            active_members=_build_score_page_members(entries),
            score_summary=store.get_score_summary(),
            selected_date=selected_date,
        )

    @app.context_processor
    def inject_shared_data() -> dict[str, Any]:
        return {
            'excel_filename': store.workbook_path.name,
            'quick_scores': store.get_quick_scores(),
            'asset_version': '20260411-10',
            'ui': UI_TEXT,
            'js_ui_text': JS_UI_TEXT,
            'enabled_status': ENABLED,
            'disabled_status': DISABLED,
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
        return _render_score_entry(selected_date=datetime.now().strftime('%Y-%m-%d'))

    @app.post('/scores/open-excel')
    def open_excel_file():
        workbook_path = store.workbook_path
        try:
            if os.name == 'nt' and hasattr(os, 'startfile'):
                os.startfile(workbook_path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(workbook_path)])
            else:
                subprocess.Popen(['xdg-open', str(workbook_path)])
            flash(MESSAGES['excel_opened'].format(filename=workbook_path.name), 'success')
        except Exception:
            flash(MESSAGES['excel_open_failed'].format(filename=workbook_path.name), 'error')
        return redirect(url_for('score_entry'))

    @app.post('/scores/save')
    def save_scores():
        selected_date = request.form.get('date', '').strip() or datetime.now().strftime('%Y-%m-%d')
        entries = []
        for member in store.get_active_members():
            name = member['name']
            score_text = request.form.get(f'score_{name}', '').strip()
            entry = {
                'name': name,
                'score': score_text,
                'before_balance': request.form.get(f'before_{name}', '').strip(),
                'after_balance': request.form.get(f'after_{name}', '').strip(),
                'manual_wear': request.form.get(f'manual_wear_{name}', '').strip(),
                'income': request.form.get(f'income_{name}', '').strip(),
                'other_expense': request.form.get(f'other_expense_{name}', '').strip(),
            }
            entries.append(entry)
            if score_text:
                try:
                    int(score_text)
                except ValueError:
                    flash(MESSAGES['score_must_integer'].format(name=name), 'error')
                    return _render_score_entry(selected_date=selected_date, entries=entries)
        try:
            result = store.save_scores_and_wear(selected_date, entries)
        except ValueError as exc:
            flash(str(exc), 'error')
            return _render_score_entry(selected_date=selected_date, entries=entries)
        if result['saved_date'] != selected_date:
            flash(MESSAGES['score_saved_shifted'].format(selected_date=selected_date, **result), 'success')
        else:
            flash(MESSAGES['score_saved'].format(**result), 'success')
        return redirect(url_for('score_entry'))

    @app.route('/wear')
    def wear_entry():
        return render_template('wear.html', wear_sheet=store.get_wear_sheet_view())

    @app.post('/wear/threshold')
    def save_wear_threshold():
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
        payload = request.get_json(silent=True) or {}
        ordered_names = payload.get('ordered_names') or []
        if not isinstance(ordered_names, list):
            return jsonify({'ok': False}), 400
        store.reorder_active_members([str(name).strip() for name in ordered_names])
        return jsonify({'ok': True})

    @app.route('/members')
    def members():
        all_members = store.get_members()
        enabled_count = sum(1 for item in all_members if item['status'] == ENABLED)
        return render_template('members.html', members=all_members, enabled_count=enabled_count)

    @app.post('/members/add')
    def add_member():
        name = request.form.get('name', '').strip()
        note = request.form.get('note', '').strip()
        try:
            store.add_member(name, note)
            flash(MESSAGES['member_added'].format(name=name), 'success')
        except ValueError as exc:
            flash(str(exc), 'error')
        return redirect(url_for('members'))

    @app.post('/members/update')
    def update_member():
        name = request.form.get('name', '').strip()
        note = request.form.get('note', '').strip()
        status = request.form.get('status', '').strip()
        try:
            store.update_member(name, note, status=status)
            if status == ENABLED:
                flash(MESSAGES['member_restored'].format(name=name), 'success')
            elif status == DISABLED:
                flash(MESSAGES['member_disabled'].format(name=name), 'success')
            else:
                flash(MESSAGES['member_note_updated'].format(name=name), 'success')
        except ValueError as exc:
            flash(str(exc), 'error')
        return redirect(url_for('members'))

    @app.post('/members/delete')
    def delete_member():
        name = request.form.get('name', '').strip()
        try:
            store.delete_member(name)
            flash(MESSAGES['member_deleted'].format(name=name), 'success')
        except ValueError as exc:
            flash(str(exc), 'error')
        return redirect(url_for('members'))

    @app.route('/members/<name>')
    def member_detail(name: str):
        year = request.args.get('year', type=int) or datetime.now().year
        month = request.args.get('month', type=int) or datetime.now().month
        return redirect(url_for('profit_calendar', name=name, year=year, month=month))


