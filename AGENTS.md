# Binance Alpha Multi-Account Manager Instructions

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
币安Alpha记录_*.xlsx
system_config.json
.env.local
```

## Daily Entry Data Safety

The daily entry form is user-entered business data. If a field is visible on
the entry page, it must not be treated as temporary UI-only state.

Required fields for daily entry persistence:

```text
score
before_balance
after_balance
manual_wear
income
other_expense
```

Hard rule: do not save only `score` while leaving the other visible entry
fields out of the durable save/read path. This already caused real user data
loss after the SQLite refactor: score values survived, but wear/income inputs
were not persisted and could not be recovered after refresh/upload.

Any change touching daily entry, score submission, SQLite schema, Supabase sync,
Excel import/export, or the score entry page must verify:

- Save all visible daily entry fields.
- Reopen the same date and confirm all fields are populated.
- Do not clear entered values when redirecting after save.
- Do not claim upload/sync protects fields that are not actually persisted.
- When touching this flow, manually verify save and reload for score,
  balances, manual wear, income, and other expense.

## Layer Responsibilities

`binance_alpha/web*.py`

- HTTP request parsing.
- Calls explicit application or service methods.
- Returns response.
- Must not import or directly operate SQLite, Supabase SDK, Excel, or JSON sync files.

`binance_alpha/services/store_application.py`

- Explicit application entry and thin coordinator.
- Must stay under 100 lines.
- Delegates to focused services.
- Must not become another business center.

`binance_alpha/services/`

- Business use cases and flow coordination.
- No Flask imports.
- No direct `sqlite3`.
- No direct Supabase SDK imports.
- No `openpyxl` except migration/import/export boundary code when explicitly allowed by architecture checks.

`binance_alpha/repositories/`

- Concrete data access.
- SQLite SQL belongs in SQLite repository modules.
- Supabase SDK belongs in `supabase_client.py`.

`binance_alpha/excel/`

- Workbook structure, import/export helpers, Excel compatibility logic.
- Excel must not become a primary write path again.

`binance_alpha/presenters/`

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
- Restore `binance_alpha/store.py` or the old `ExcelStore` entry point.
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
binance_alpha/
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
币安Alpha记录_*.xlsx
__pycache__/
.tmp_*/
public/
bm2_cloud/
_legacy_quarantine_*/
```

For a lightweight runtime package, `scripts/` and `docs/` may be excluded. For a development package, keep them.
