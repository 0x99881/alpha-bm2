from __future__ import annotations


class ExcelExportService:
    def __init__(self, context) -> None:
        self._context = context

    def export_to_excel(self, date_text: str | None = None) -> int:
        if self._context.read_only:
            return 0
        return self._context.excel_exporter.export_missing_dates(self._context, detail_date_text=date_text)

    def sync_member_visibility(self) -> None:
        if self._context.read_only:
            return
        self._context.excel_exporter.sync_member_visibility(self._context)
