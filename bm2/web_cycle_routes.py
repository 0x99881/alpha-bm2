from __future__ import annotations

from flask import abort, flash, redirect, render_template, request, url_for

from .ui_text import MESSAGES


def register_cycle_routes(app, store, *, read_only_mode: bool) -> None:
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
        name = request.form.get("name", "").strip()
        try:
            cycle_id = store.create_settlement_cycle(name)
            flash(MESSAGES["cycle_created"].format(name=name), "success")
            return redirect(url_for("cycle_profit", cycle=cycle_id))
        except ValueError as exc:
            flash(str(exc), "error")
        return redirect(url_for("cycle_profit"))

    @app.post("/cycle-profit/save")
    def save_cycle_profit():
        if read_only_mode:
            abort(403)
        cycle_id = request.form.get("cycle_id", "").strip()
        try:
            store.save_cycle_settlement(cycle_id, request.form)
            flash(MESSAGES["cycle_saved"], "success")
        except ValueError as exc:
            flash(str(exc), "error")
        if cycle_id:
            return redirect(url_for("cycle_profit", cycle=cycle_id))
        return redirect(url_for("cycle_profit"))
