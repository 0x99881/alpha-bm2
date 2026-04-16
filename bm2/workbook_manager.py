from __future__ import annotations


class WorkbookManager:
    def __init__(self, store) -> None:
        self.store = store

    def initialize_workbook(self) -> None:
        self.store._ensure_workbook()
        workbook = self.open_workbook()
        try:
            self.sync_member_visibility(workbook)
            self.save_workbook(workbook)
        finally:
            workbook.close()

    def open_workbook(self):
        return self.store._open_workbook()

    def save_workbook(self, workbook) -> None:
        self.store._save_workbook(workbook)

    def sync_member_visibility(self, workbook) -> None:
        self.store._sync_member_visibility_in_workbook(workbook)

    def delete_member_data(self, workbook, member_name: str) -> None:
        self.store._ensure_score_sheet_structure(workbook)
        self.store._ensure_wear_sheet_structure(workbook)
        self.store._ensure_income_sheet_structure(workbook)
        self.store._ensure_expense_sheet_structure(workbook)
        self.store._delete_member_from_workbook(workbook, member_name)
