from __future__ import annotations

from flask import abort, flash, redirect, render_template, request, url_for

from .ui_text import MESSAGES


def register_cashflow_routes(app, store, *, read_only_mode: bool) -> None:
    @app.route("/cash-flow")
    def cash_flow():
        if read_only_mode:
            return redirect(url_for("score_overview"))
        return render_template("cash_flow.html", cash_flow=store.get_cash_flow_view())

    @app.post("/cash-flow/add")
    def add_cash_flow_route():
        if read_only_mode:
            abort(403)
        try:
            result = store.add_cash_flow(request.form)
            flash(
                MESSAGES["cashflow_added"].format(
                    direction=result["direction_label"], amount=result["amount"]
                ),
                "success",
            )
            # DB write already succeeded; a locked Excel only means the sheet
            # lags behind. Tell the user, don't fail the record.
            if result.get("export_error") is not None:
                flash(MESSAGES["cycle_excel_locked"], "warning")
        except (ValueError, OSError) as exc:
            _flash_action_error(exc)
        return redirect(url_for("cash_flow"))

    @app.post("/cash-flow/delete")
    def delete_cash_flow_route():
        if read_only_mode:
            abort(403)
        flow_id = request.form.get("flow_id", "").strip()
        try:
            result = store.delete_cash_flow(flow_id)
            flash(MESSAGES["cashflow_deleted"], "success")
            if result.get("export_error") is not None:
                flash(MESSAGES["cycle_excel_locked"], "warning")
        except (ValueError, OSError) as exc:
            _flash_action_error(exc)
        return redirect(url_for("cash_flow"))

    def _flash_action_error(exc: Exception) -> None:
        if isinstance(exc, ValueError):
            flash(str(exc), "error")
        else:
            flash(MESSAGES["cycle_excel_locked"], "error")
