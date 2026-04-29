# Development Rules

Last updated: 2026-04-28

This file explains how to extend BM2 without drifting back to the old mixed architecture.

## Golden Path

Use these paths for new work:

```text
Local write:
web.py -> service -> SQLite repository/writer -> SQLite

Online/mobile write:
web.py -> service -> Supabase client -> Supabase

Sync:
SQLite <-> SyncService <-> Supabase

Excel:
SQLite -> Excel export
Excel -> migration/import -> SQLite
```

## Where To Put New Code

New route:

- Add request parsing in `bm2/web.py`.
- Call an existing service or add a focused service method.
- Do not access database, Excel, Supabase SDK, or JSON files in `web.py`.

New business action:

- Add to `bm2/services/`.
- Keep the service focused on one use case.
- Use repositories/writers for data access.

New SQLite query or write:

- Add to `bm2/repositories/sqlite_*.py`.
- Keep SQL inside repository files.

New Supabase operation:

- Add to `bm2/repositories/supabase_client.py`.
- Do not import Supabase SDK elsewhere.

New Excel export/import behavior:

- Add to `bm2/excel/`, `SQLiteToExcelExporter`, or migration/import service.
- Do not make Excel a primary data store.

New page display transformation:

- Add to `bm2/presenters/`.
- Presenters should not write data.

New frontend source:

- Add to `static/` and `templates/`.
- Do not edit generated `public/` as source.

## Store Boundary

`bm2/store.py` is intentionally tiny.

It should only expose the old `ExcelStore` entry point and forward to the application layer.

Do not add:

- Business rules.
- Sync flows.
- Excel operations.
- SQLite calls.
- Supabase calls.
- Repository imports.

`bm2/services/store_application.py` is also intentionally tiny.

It coordinates:

- `StoreBootstrapService`
- `StoreCommandService`
- `StoreQueryService`
- `ExcelExportService`

Do not grow it into another large service.

## Retired Code

These must stay retired:

```text
bm2_cloud/
score-records.json
cloud_sync
cloud_entry_writer
legacy_json_sync
json_sync
blob_sync
store_base.py
public/ as source
```

If historical data is needed, treat it as read-only migration input.

## Data Files

Do not casually edit or delete:

```text
bm2_local.db
BM2记录_*.xlsx
system_config.json
.env.local
```

These may contain real local data or secrets.

## Checks

Before handing work back, run:

```bash
python scripts/architecture_check.py
python scripts/smoke_check.py
```

Use `architecture_check.py` as the source of truth for enforced boundaries.

## If Unsure

Choose the path that keeps:

- SQLite local-first.
- Supabase online-sync only.
- Excel import/export only.
- Store facade thin.
- Services focused.
- Repositories owning data access.
