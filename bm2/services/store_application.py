from __future__ import annotations

from .store_bootstrap_service import StoreBootstrapService


class StoreApplication:
    def __init__(self, base_dir, read_only: bool = False):
        self._context = StoreBootstrapService().create_context(base_dir, read_only=read_only)

    def __getattr__(self, name):
        for target in (
            self._context.command_service,
            self._context.query_service,
            self._context.export_service,
            self._context,
        ):
            if hasattr(target, name):
                return getattr(target, name)
        raise AttributeError(name)

    def __setattr__(self, name, value):
        if name == "_context" or "_context" not in self.__dict__:
            object.__setattr__(self, name, value)
            return
        setattr(self._context, name, value)


__all__ = ["StoreApplication"]
