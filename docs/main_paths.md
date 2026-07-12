# Binance Alpha Main Paths

This document records the only supported runtime paths. Code that creates a
second active path for the same action should be removed rather than wrapped.

## Score Submission

Current entry: `POST /scores/save`

Target path:

```text
web_score_routes -> DailyEntryService.process_submission
  -> SQLiteEntryWriter -> LocalDatabase/repositories -> SQLite
```

Rules:

- Web extracts request data and renders the response.
- DailyEntryService validates the request and builds entries.
- SQLiteEntryWriter writes the local score rows.
- Excel is not part of score submission.

Known removal target:

- None. Score submission now writes only through the entry writer.

## Score Reads And Page Data

Current entry: score pages, score overview API, presenters.

Target path:

```text
web routes -> StoreQueryService -> Presenter -> LocalDatabase/repositories
```

Read-only/mobile online views may read from `ApplicationService` and Supabase.
Local page reads should use SQLite as the primary source.

## Member Management

Current entry: member routes.

Target path:

```text
web_member_routes -> MemberService -> LocalDatabase/member repository -> SQLite
```

Rules:

- Member add, update, delete, reorder, and query use SQLite locally.
- Excel does not own member state.
- Store command proxies are not a member-management entry point.

## Excel Import

Current entry: refresh and bootstrap import.

Target path:

```text
StoreCommandService -> ExcelImportService -> excel helpers -> LocalDatabase
```

Excel import is a historical migration path into SQLite.

## Excel Export

Current entry: explicit export service or sync after pull.

Target path:

```text
ExcelExportService -> SQLiteToExcelExporter -> ExcelWorkbookWriter
```

Excel export is an output path only. Export failures must not change whether a
SQLite score submission succeeded.

## Supabase Sync

Current entry: sync routes.

Target path:

```text
web_sync_routes -> StoreCommandService -> SyncService
  -> LocalDatabase/repositories <-> SupabaseClient
```

Supabase failures should be reported clearly. They must not create an alternate
local write path.
