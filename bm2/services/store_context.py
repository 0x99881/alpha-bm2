from __future__ import annotations

from ..excel.workbook_structure import StoreStructureMixin
from ..store_sheet_utils import StoreSheetUtilsMixin


class StoreContext(StoreSheetUtilsMixin, StoreStructureMixin):
    def __getattr__(self, name):
        for service_name in ("query_service", "command_service", "export_service"):
            service = self.__dict__.get(service_name)
            if service is not None and hasattr(service, name):
                return getattr(service, name)
        raise AttributeError(name)

    def get_members(self):
        return self.query_service.get_members()

    def get_active_members(self):
        return self.query_service.get_active_members()

    def get_wear_abnormal_threshold(self) -> float:
        return self.query_service.get_wear_abnormal_threshold()

    def _round_wear(self, value):
        return self._writer._round_wear(value)

    def _round_income(self, value):
        return self._writer._round_income(value)

    def _round_expense(self, value):
        return self._writer._round_expense(value)
