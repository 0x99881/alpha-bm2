from __future__ import annotations

from flask import abort, flash, jsonify, redirect, render_template, request, url_for

from .ui_text import MESSAGES


def register_cycle_routes(app, store, *, read_only_mode: bool) -> None:
    def _flash_action_error(exc: Exception) -> None:
        if isinstance(exc, ValueError):
            flash(str(exc), "error")
        else:
            flash(MESSAGES["cycle_excel_locked"], "error")

    @app.route("/cycle-profit")
    def cycle_profit():
        if read_only_mode:
            return redirect(url_for("score_overview"))
        cycle_id = request.args.get("cycle", "").strip() or None
        return render_template(
            "cycle_profit.html",
            cycle_data=store.get_cycle_profit_view(cycle_id),
        )

    @app.post("/cycle-profit/create")
    def create_cycle_profit():
        if read_only_mode:
            abort(403)
        start_date = request.form.get("start_date", "").strip()
        try:
            cycle_id = store.create_settlement_cycle(start_date)
            flash(MESSAGES["cycle_created"].format(name=start_date), "success")
            return redirect(url_for("cycle_profit", cycle=cycle_id))
        except (ValueError, OSError) as exc:
            _flash_action_error(exc)
        return redirect(url_for("cycle_profit"))

    @app.post("/cycle-profit/save")
    def save_cycle_profit():
        if read_only_mode:
            abort(403)
        cycle_id = request.form.get("cycle_id", "").strip()
        try:
            store.save_cycle_settlement(cycle_id, request.form)
            flash(MESSAGES["cycle_saved"], "success")
        except (ValueError, OSError) as exc:
            _flash_action_error(exc)
        if cycle_id:
            return redirect(url_for("cycle_profit", cycle=cycle_id))
        return redirect(url_for("cycle_profit"))

    @app.post("/cycle-profit/settle")
    def settle_cycle_profit():
        if read_only_mode:
            abort(403)
        cycle_id = request.form.get("cycle_id", "").strip()
        settle_date = request.form.get("settle_date", "").strip()
        try:
            store.save_cycle_settlement(cycle_id, request.form)
            store.settle_cycle(cycle_id, settle_date)
            flash(MESSAGES["cycle_settled"], "success")
        except (ValueError, OSError) as exc:
            _flash_action_error(exc)
        if cycle_id:
            return redirect(url_for("cycle_profit", cycle=cycle_id))
        return redirect(url_for("cycle_profit"))

    @app.post("/cycle-profit/settle-and-create-next")
    def settle_and_create_next_cycle():
        """Settle current cycle and open next one atomically.

        end_balance_<name> form fields are picked up and persisted before
        settlement. On Excel failure all DB changes are rolled back.
        """
        if read_only_mode:
            abort(403)
        cycle_id = request.form.get("cycle_id", "").strip()
        settle_date = request.form.get("settle_date", "").strip()
        end_balances: dict[str, str] = {}
        for key, value in request.form.items():
            if key.startswith("end_balance_"):
                end_balances[key[len("end_balance_"):]] = value
        try:
            new_cycle_id = store.settle_and_create_next_cycle(
                cycle_id, settle_date, end_balances
            )
            window = store.get_cycle_window(new_cycle_id)
            flash(
                MESSAGES["cycle_settle_and_create_next_done"].format(
                    start_date=window.get("start_date", "")
                ),
                "success",
            )
            if new_cycle_id:
                return redirect(url_for("cycle_profit", cycle=new_cycle_id))
        except (ValueError, OSError) as exc:
            _flash_action_error(exc)
        if cycle_id:
            return redirect(url_for("cycle_profit", cycle=cycle_id))
        return redirect(url_for("cycle_profit"))

    @app.post("/cycle-profit/delete")
    def delete_cycle_profit():
        if read_only_mode:
            abort(403)
        cycle_id = request.form.get("cycle_id", "").strip()
        try:
            store.delete_cycle(cycle_id)
            flash(MESSAGES["cycle_deleted"], "success")
        except (ValueError, OSError) as exc:
            _flash_action_error(exc)
        return redirect(url_for("cycle_profit"))

    @app.post("/cycle-profit/add-member")
    def add_cycle_member_route():
        if read_only_mode:
            abort(403)
        cycle_id = request.form.get("cycle_id", "").strip()
        member_name = request.form.get("member_name", "").strip()
        try:
            store.add_cycle_member(cycle_id, member_name)
            flash(MESSAGES["cycle_member_added"].format(name=member_name), "success")
        except (ValueError, OSError) as exc:
            _flash_action_error(exc)
        if cycle_id:
            return redirect(url_for("cycle_profit", cycle=cycle_id))
        return redirect(url_for("cycle_profit"))

    @app.post("/cycle-profit/remove-member")
    def remove_cycle_member_route():
        if read_only_mode:
            abort(403)
        cycle_id = request.form.get("cycle_id", "").strip()
        member_name = request.form.get("member_name", "").strip()
        try:
            store.remove_cycle_member(cycle_id, member_name)
            flash(MESSAGES["cycle_member_removed"].format(name=member_name), "success")
        except (ValueError, OSError) as exc:
            _flash_action_error(exc)
        if cycle_id:
            return redirect(url_for("cycle_profit", cycle=cycle_id))
        return redirect(url_for("cycle_profit"))

    @app.post("/cycle-profit/reorder")
    def reorder_cycle_members_route():
        if read_only_mode:
            abort(403)
        payload = request.get_json(silent=True)
        if not isinstance(payload, dict):
            return jsonify({"ok": False, "error": "bad payload"}), 400
        cycle_id = str(payload.get("cycle_id", "")).strip()
        ordered_names = payload.get("ordered_names") or []
        if not isinstance(ordered_names, list):
            return jsonify({"ok": False, "error": "bad payload"}), 400
        try:
            store.reorder_cycle_members(cycle_id, [str(n) for n in ordered_names])
            return jsonify({"ok": True})
        except (ValueError, OSError) as exc:
            return jsonify({"ok": False, "error": str(exc)}), 400
