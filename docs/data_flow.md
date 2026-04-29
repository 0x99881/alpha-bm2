# Data Flow And Architecture Boundary

Last updated: 2026-04-28

This document describes the current architecture after cleanup.

## Main Rule

There is only one active data direction:

```text
Local use:   Flask -> service -> SQLite
Online use:  Flask -> service -> Supabase
Sync:        SQLite <-> SyncService <-> Supabase
Excel:       import/export only
```

Excel, JSON files, cloud blobs, `bm2_cloud`, and old cloud sync wrappers are not primary data paths.

## Active Write Paths

| Action | Active path | Target |
|---|---|---|
| Local score submit | `web.py -> DailyEntryService -> SQLiteEntryWriter -> LocalDatabase` | SQLite |
| Online/mobile score submit | `web.py -> DailyEntryService -> SupabaseEntryWriter -> SupabaseClient` | Supabase |
| Member add/update/delete | `web.py -> MemberService -> SQLite member repository` | SQLite |
| Member reorder | `web.py -> MemberService -> SQLite member repository` | SQLite |
| Supabase push/pull/sync | `Store facade -> StoreCommandService -> SyncService -> SQLite/Supabase` | SQLite/Supabase |
| Excel export | `ExcelExportService -> SQLiteToExcelExporter` | Excel file |
| Legacy Excel import | `StoreBootstrapService -> LocalDataMigrationService -> SQLite` | SQLite |

## Active Read Paths

| Action | Active path | Source |
|---|---|---|
| Local rankings | `web.py -> StoreQueryService -> SQLite report repository` | SQLite |
| Online rankings | `web.py -> StoreQueryService -> ApplicationService -> SupabaseClient` | Supabase, fallback SQLite |
| Score overview | `web.py -> StoreQueryService -> presenter` | SQLite/Supabase |
| Wear/profit calendar | `web.py -> StoreQueryService -> reader/presenter` | Excel compatibility data |

## Layer Boundaries

`web.py`:

- Parses request input.
- Calls store facade or service methods.
- Returns HTML, JSON, redirect, or error.
- Must not directly import or operate SQLite, Supabase SDK, Excel, or JSON sync files.

`bm2/store.py`:

- Thin facade only.
- Keeps the historical `ExcelStore` entry point.
- Must stay under 120 lines.
- Must not import repositories, Excel modules, Supabase SDK, or SQLite internals.

`bm2/services/store_application.py`:

- Thin coordinator only.
- Must stay under 100 lines.
- Delegates to focused services.
- Must not directly operate repositories, Excel, Supabase SDK, or SQLite.

Application services:

- `StoreBootstrapService`: initialization and startup migration wiring.
- `StoreCommandService`: write commands and sync commands.
- `StoreQueryService`: read/query methods.
- `ExcelExportService`: SQLite-to-Excel export.
- `SyncService`: SQLite and Supabase synchronization flow.
- `MemberService`: member business actions.
- `DailyEntryService`: score submission validation and processing.

`repositories/`:

- Own concrete data access.
- SQLite access belongs in SQLite repository modules.
- Supabase SDK access belongs in `repositories/supabase_client.py`.

`excel/`:

- Own workbook structure, import/export helpers, and Excel compatibility code.
- Excel is not a primary business write path.

`static/` and `public/`:

- `static/` is the hand-maintained frontend source.
- `public/` is generated output only.
- Do not hand-edit or commit `public/` as source.

## Retired Paths

These are retired and must not be imported by active source code:

- `bm2_cloud/`
- `score-records.json`
- `bm2/cloud_sync.py`
- `bm2/services/cloud_entry_writer.py`
- `legacy_json_sync`
- `json_sync`
- `blob_sync`
- JSON/blob/file based score sync
- `bm2/store_base.py`
- active `public/` source directory

If any retired file is needed for investigation, keep it outside the active runtime or restore it only as read-only legacy material.

## Guard Rails

Run this before packaging or deployment:

```bash
python scripts/architecture_check.py
python scripts/smoke_check.py
```

`architecture_check.py` blocks the old routes from coming back.

## Remaining Known Debt

- `local_database.py` still owns the SQLite connection entry point.
- Some Excel compatibility helpers are still needed for historical import/export and wear/profit calendar views.

These are known boundaries, not active old cloud-sync routes.
