# Development Rules

Last updated: 2026-05-01

This file explains how to extend the Binance Alpha multi-account manager without drifting back to the old mixed architecture.

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

- Add request parsing in `binance_alpha/web.py`.
- Call an existing service or add a focused service method.
- Do not access database, Excel, Supabase SDK, or JSON files in `web.py`.

New business action:

- Add to `binance_alpha/services/`.
- Keep the service focused on one use case.
- Use repositories/writers for data access.

New SQLite query or write:

- Add to `binance_alpha/repositories/sqlite_*.py`.
- Keep SQL inside repository files.

New Supabase operation:

- Add to `binance_alpha/repositories/supabase_client.py`.
- Do not import Supabase SDK elsewhere.

New Excel export/import behavior:

- Add to `binance_alpha/excel/`, `SQLiteToExcelExporter`, or migration/import service.
- Do not make Excel a primary data store.

New page display transformation:

- Add to `binance_alpha/presenters/`.
- Presenters should not write data.

New frontend source:

- Add to `static/` and `templates/`.
- Do not edit generated `public/` as source.

## Application Boundary

`binance_alpha/services/store_application.py` is also intentionally tiny.

It exposes explicit application methods and coordinates:

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
- Application entry thin and explicit.
- Services focused.
- Repositories owning data access.
