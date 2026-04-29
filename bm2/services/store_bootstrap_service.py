from __future__ import annotations

from datetime import datetime
from pathlib import Path

from ..excel import ExpenseSheet, IncomeSheet, ScoreSheet, WearSheet
from ..excel.workbook_repository import WorkbookRepository
from ..local_database import LocalDatabase
from ..presenters import ProfitCalendarPresenter, ScorePresenter, WearPresenter
from ..repositories import ConfigRepository, SupabaseClient
from ..sqlite_to_excel_exporter import SQLiteToExcelExporter
from ..store_reader_facade import StoreReaderFacade
from ..store_writer_facade import StoreWriterFacade
from .application_service import ApplicationService
from .daily_entry_service import DailyEntryService
from .excel_export_service import ExcelExportService
from .member_service import MemberService
from .migration_service import LocalDataMigrationService
from .sqlite_entry_writer import SQLiteEntryWriter
from .store_command_service import StoreCommandService
from .store_context import StoreContext
from .store_query_service import StoreQueryService
from .supabase_entry_writer import SupabaseEntryWriter
from .sync_service import SyncService


class StoreBootstrapService:
    def timestamp(self) -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def create_context(self, base_dir, *, read_only: bool = False) -> StoreContext:
        context = StoreContext()
        context.base_dir = Path(base_dir)
        context.read_only = read_only
        context.bootstrap_service = self
        context.config_repository = ConfigRepository(context.base_dir / "system_config.json", read_only=read_only)
        context.config_repository.ensure_member_config(self.timestamp)
        context.workbook_repository = WorkbookRepository(context.base_dir, context.config_repository, read_only=read_only)
        context.workbook_path = context.workbook_repository.workbook_path
        context.score_sheet = ScoreSheet(context)
        context.wear_sheet = WearSheet(context)
        context.income_sheet = IncomeSheet(context)
        context.expense_sheet = ExpenseSheet(context)
        context._reader = StoreReaderFacade(context)
        context._writer = StoreWriterFacade(context)
        context.query_service = StoreQueryService(context)
        context.supabase = SupabaseClient(context.base_dir)
        if read_only:
            context.local_db = None
            context.excel_exporter = None
            context.export_service = ExcelExportService(context)
            context.migration_service = None
            context.sync_service = SyncService(context.local_db, context.supabase)
            context.application_service = ApplicationService(
                context.local_db,
                context.supabase,
                lambda: datetime.now().strftime("%Y-%m-%d"),
            )
            entry_writer = SupabaseEntryWriter(context.supabase)
            context._sqlite_writer = None
        else:
            context._ensure_workbook()
            context.local_db = LocalDatabase(context.base_dir, read_only=read_only)
            context.excel_exporter = SQLiteToExcelExporter(context.local_db, context._writer)
            context.export_service = ExcelExportService(context)
            context.migration_service = LocalDataMigrationService(
                context.local_db,
                context,
                context.query_service.legacy_config_members,
                context.excel_exporter,
            )
            context.migration_service.bootstrap_from_legacy_sources()
            context.sync_service = SyncService(context.local_db, context.supabase, after_pull=context.migration_service.after_supabase_pull)
            context.application_service = self.create_online_application_service(context)
            context._sqlite_writer = SQLiteEntryWriter(context.local_db)
            entry_writer = context._sqlite_writer
        context.daily_entry_service = DailyEntryService(entry_writer)
        context.member_service = MemberService(context.get_members, context.local_db)
        context.command_service = StoreCommandService(context)
        context.score_presenter = ScorePresenter(context)
        context.wear_presenter = WearPresenter(context)
        context.profit_calendar_presenter = ProfitCalendarPresenter(context)
        return context

    def create_online_application_service(self, context) -> ApplicationService:
        return ApplicationService(context.local_db, context.supabase, context.query_service.get_next_score_date)
