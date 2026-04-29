from __future__ import annotations

# Thin facade only. Real wiring and behavior live in services/store_application.py.

from .constants import DISABLED, ENABLED
from .services.store_application import StoreApplication


class ExcelStore:
    def __init__(self, base_dir, read_only: bool = False):
        object.__setattr__(self, "_app", StoreApplication(base_dir, read_only=read_only))

    def __getattr__(self, name):
        return getattr(self._app, name)

    def __setattr__(self, name, value):
        if name == "_app" or "_app" not in self.__dict__:
            object.__setattr__(self, name, value)
            return
        setattr(self._app, name, value)


__all__ = ["ExcelStore", "ENABLED", "DISABLED"]
