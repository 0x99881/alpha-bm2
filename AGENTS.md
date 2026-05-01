# BM2 Project Instructions

This file is for AI coding agents and future maintainers.

## Current Architecture

The project has one main architecture:

```text
Local data:       SQLite
Online/mobile:    Supabase
Sync bridge:      SQLite <-> SyncService <-> Supabase
Excel:            import/export/legacy source only
Frontend source:  static/
Generated output: public/ if build output exists
```

Do not reintroduce old routes.

## Hard Rules

- SQLite is the local source of truth.
- Supabase is the online/mobile sync source.
- Excel is not a primary business write path.
- JSON file sync is retired.
- `bm2_cloud`, `score-records.json`, `cloud_sync`, `cloud_entry_writer`, `legacy_json_sync`, `json_sync`, and `blob_sync` must not return to active source.
- `public/` must not be hand-maintained as source.
- Do not modify real local data files unless explicitly asked.

Real local data files include:

```text
bm2_local.db
BM2记录_*.xlsx
system_config.json
.env.local
```

## Layer Responsibilities

`bm2/web*.py`

- HTTP request parsing.
- Calls explicit application or service methods.
- Returns response.
- Must not import or directly operate SQLite, Supabase SDK, Excel, or JSON sync files.

`bm2/services/store_application.py`

- Explicit application entry and thin coordinator.
- Must stay under 100 lines.
- Delegates to focused services.
- Must not become another business center.

`bm2/services/`

- Business use cases and flow coordination.
- No Flask imports.
- No direct `sqlite3`.
- No direct Supabase SDK imports.
- No `openpyxl` except migration/import/export boundary code when explicitly allowed by architecture checks.

`bm2/repositories/`

- Concrete data access.
- SQLite SQL belongs in SQLite repository modules.
- Supabase SDK belongs in `supabase_client.py`.

`bm2/excel/`

- Workbook structure, import/export helpers, Excel compatibility logic.
- Excel must not become a primary write path again.

`bm2/presenters/`

- Page/view data assembly.
- No writes.

`static/`

- Frontend source.

`templates/`

- HTML templates.

## Current Application Services

- `StoreBootstrapService`: app startup, object wiring, startup migration.
- `StoreCommandService`: configuration, import/export, and sync commands.
- `StoreQueryService`: read/query methods.
- `ExcelExportService`: SQLite-to-Excel export.
- `SyncService`: SQLite/Supabase push, pull, sync.
- `MemberService`: member business actions.
- `DailyEntryService`: daily score validation and submission.
- `ApplicationService`: online/mobile score and ranking reads.

## Before Editing

Read the relevant files first. Prefer small changes.

Do not:

- Add a new wrapper just to hide old logic.
- Reconnect Excel as the main database.
- Reconnect JSON/blob/cloud sync.
- Restore `bm2/store.py` or the old `ExcelStore` entry point.
- Put repository logic into services.
- Put service logic into repositories.
- Restore quarantined or retired files into active runtime.

## Required Checks

Run these after changes:

```bash
python scripts/architecture_check.py
python scripts/smoke_check.py
```


## Packaging Rules

Code package may include:

```text
bm2/
scripts/
static/
templates/
docs/
app.py
requirements.txt
start.bat
vercel.json
README.md
```

Do not package:

```text
.env.local
bm2_local.db
BM2记录_*.xlsx
__pycache__/
.tmp_*/
public/
bm2_cloud/
_legacy_quarantine_*/
```

For a lightweight runtime package, `scripts/` and `docs/` may be excluded. For a development package, keep them.
