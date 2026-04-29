from __future__ import annotations

import ast
import re
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
IGNORED_DIR_NAMES = {'.git', '.vercel'}


def _is_ignored(path: Path) -> bool:
    parts = path.relative_to(PROJECT_ROOT).parts
    return any(part in IGNORED_DIR_NAMES or part.startswith('_legacy_quarantine_') for part in parts)


def _iter_files(*, suffixes: tuple[str, ...] | None = None) -> list[Path]:
    files: list[Path] = []
    for path in PROJECT_ROOT.rglob('*'):
        if _is_ignored(path):
            continue
        if not path.is_file():
            continue
        if suffixes and path.suffix not in suffixes:
            continue
        files.append(path)
    return files


def _module_name(path: Path) -> str:
    return '.'.join(path.relative_to(PROJECT_ROOT).with_suffix('').parts)


def _resolve_import(current_module: str, node: ast.ImportFrom) -> str:
    if node.level == 0:
        return node.module or ''
    parts = current_module.split('.')[:-node.level]
    if node.module:
        parts.extend(node.module.split('.'))
    return '.'.join(parts)


def _python_imports(path: Path) -> list[str]:
    module_name = _module_name(path)
    tree = ast.parse(_read_text(path))
    imports: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imports.extend(alias.name for alias in node.names)
        if isinstance(node, ast.ImportFrom):
            imports.append(_resolve_import(module_name, node))
    return imports


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding='utf-8')
    except UnicodeDecodeError:
        return path.read_text(encoding='utf-8-sig')


def check_workspace_artifacts() -> list[str]:
    failures: list[str] = []
    for path in PROJECT_ROOT.rglob('*'):
        if _is_ignored(path):
            continue
        rel = path.relative_to(PROJECT_ROOT)
        if path.is_dir() and path.name == '__pycache__':
            failures.append(f'{rel} is a Python cache directory')
        if path.is_file() and path.suffix == '.pyc':
            failures.append(f'{rel} is a Python bytecode cache')
        if len(rel.parts) == 1 and path.name.startswith('.tmp_'):
            failures.append(f'{rel} is a temporary root artifact')
        if len(rel.parts) == 1 and path.suffix == '.zip' and not path.name.startswith('alpha_baseline_'):
            failures.append(f'{rel} is an old root package artifact')
    return failures


def check_domain_dependencies() -> list[str]:
    failures: list[str] = []
    blocked_prefixes = ('bm2.web', 'bm2.presenters', 'bm2.features')
    blocked_exact = {'flask'}
    for path in (PROJECT_ROOT / 'bm2' / 'domain').rglob('*.py'):
        if _is_ignored(path):
            continue
        for imported in _python_imports(path):
            if imported in blocked_exact or imported.startswith(blocked_prefixes):
                failures.append(f'{path.relative_to(PROJECT_ROOT)} imports {imported}')
    return failures


def check_feature_dependencies() -> list[str]:
    failures: list[str] = []
    feature_root = PROJECT_ROOT / 'bm2' / 'features'
    for feature_dir in feature_root.iterdir() if feature_root.exists() else []:
        if not feature_dir.is_dir() or feature_dir.name == '__pycache__':
            continue
        current_feature = feature_dir.name
        for path in feature_dir.rglob('*.py'):
            if _is_ignored(path):
                continue
            for imported in _python_imports(path):
                if not imported.startswith('bm2.features.'):
                    continue
                parts = imported.split('.')
                if len(parts) >= 3 and parts[2] != current_feature:
                    failures.append(f'{path.relative_to(PROJECT_ROOT)} imports {imported}')
    return failures


def check_frontend_requests() -> list[str]:
    failures: list[str] = []
    for path in (PROJECT_ROOT / 'static').rglob('*.js'):
        if _is_ignored(path):
            continue
        text = _read_text(path)
        if path.name == 'request.js':
            continue
        if re.search(r'\bfetch\s*\(', text):
            failures.append(f'{path.relative_to(PROJECT_ROOT)} uses fetch directly')
    return failures


def check_retired_sync_keywords_in_source() -> list[str]:
    failures: list[str] = []
    retired = (
        'bm2_cloud',
        'cloud_sync',
        'cloud_entry_writer',
        'legacy_json_sync',
        'json_sync',
        'blob_sync',
    )
    for path in (PROJECT_ROOT / 'bm2').rglob('*.py'):
        if _is_ignored(path):
            continue
        rel = path.relative_to(PROJECT_ROOT)
        if rel.parts[:2] == ('bm2', 'legacy'):
            continue
        text = _read_text(path)
        for keyword in retired:
            if keyword in text:
                failures.append(f'{rel} contains retired sync keyword: {keyword}')
    return failures


def check_services_boundaries() -> list[str]:
    failures: list[str] = []
    blocked = ('openpyxl', 'sqlite3', 'supabase', 'flask')
    for path in (PROJECT_ROOT / 'bm2' / 'services').rglob('*.py'):
        if _is_ignored(path):
            continue
        for imported in _python_imports(path):
            if imported in blocked or any(imported.startswith(f'{item}.') for item in blocked):
                failures.append(f'{path.relative_to(PROJECT_ROOT)} imports blocked dependency {imported}')
    return failures


def check_web_boundaries() -> list[str]:
    failures: list[str] = []
    path = PROJECT_ROOT / 'bm2' / 'web.py'
    if not path.exists():
        return failures
    imports = _python_imports(path)
    blocked_imports = ('sqlite3', 'openpyxl', 'supabase', 'json')
    for imported in imports:
        if imported in blocked_imports or any(imported.startswith(f'{item}.') for item in blocked_imports):
            failures.append(f'bm2/web.py imports blocked dependency {imported}')
    text = _read_text(path)
    blocked_calls = (
        r'\blocal_db\.',
        r'\bsupabase\.',
        r'\bstore\.supabase\b',
        r'\bSQLiteEntryWriter\b',
        r'\bStoreWriterFacade\b',
        r'\bjson\.dump\b',
        r'\bjson\.load\b',
    )
    for pattern in blocked_calls:
        if re.search(pattern, text):
            failures.append(f'bm2/web.py directly uses blocked data/write pattern: {pattern}')
    return failures


def check_public_matches_static() -> list[str]:
    failures: list[str] = []
    public_root = PROJECT_ROOT / 'public'
    static_root = PROJECT_ROOT / 'static'
    if not public_root.exists():
        return failures
    failures.append('public/ exists in the source tree; static/ is the source and public/ must be generated')
    for path in public_root.rglob('*'):
        if not path.is_file() or path.suffix not in {'.js', '.css'}:
            continue
        rel = path.relative_to(public_root)
        source = static_root / rel
        if not source.exists():
            failures.append(f'public/{rel} has no matching static source')
            continue
        if path.read_bytes() != source.read_bytes():
            failures.append(f'public/{rel} differs from static/{rel}; public must be generated')
    return failures


def check_retired_write_wrappers() -> list[str]:
    failures: list[str] = []
    web_path = PROJECT_ROOT / 'bm2' / 'web.py'
    if web_path.exists() and 'record_local_score_submission' in _read_text(web_path):
        failures.append('bm2/web.py calls retired local score wrapper record_local_score_submission')
    store_path = PROJECT_ROOT / 'bm2' / 'store.py'
    if store_path.exists() and 'def record_local_score_submission' in _read_text(store_path):
        failures.append('bm2/store.py defines retired local score wrapper record_local_score_submission')
    member_service = PROJECT_ROOT / 'bm2' / 'services' / 'member_service.py'
    if member_service.exists():
        text = _read_text(member_service)
        for marker in ('sync_members_to_workbook', 'delete_member_from_workbook'):
            if marker in text:
                failures.append(f'bm2/services/member_service.py still references Excel member write callback {marker}')
    return failures


def check_store_base_retired() -> list[str]:
    failures: list[str] = []
    if (PROJECT_ROOT / 'bm2' / 'store_base.py').exists():
        failures.append('bm2/store_base.py is retired and must stay in quarantine')
    for path in (PROJECT_ROOT / 'bm2').rglob('*.py'):
        if _is_ignored(path):
            continue
        rel = path.relative_to(PROJECT_ROOT)
        text = _read_text(path)
        if 'store_base' in text or 'StoreBaseMixin' in text:
            failures.append(f'{rel} imports or references retired store_base')
    return failures


def check_local_database_is_composition_entry() -> list[str]:
    failures: list[str] = []
    path = PROJECT_ROOT / 'bm2' / 'local_database.py'
    if not path.exists():
        return failures
    tree = ast.parse(_read_text(path))
    allowed_methods = {'__init__', '_connect'}
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef) and node.name == 'LocalDatabase':
            for item in node.body:
                if isinstance(item, ast.FunctionDef) and item.name not in allowed_methods:
                    failures.append(f'bm2/local_database.py defines logic method {item.name}; keep only connection/composition')
    text = _read_text(path)
    forbidden_markers = (
        'sync_scores_from_workbook',
        'push_to_supabase',
        'pull_from_supabase',
        'workbook_repository',
        'workbook.',
        'META_SHEET',
        'openpyxl',
    )
    for marker in forbidden_markers:
        if marker in text:
            failures.append(f'bm2/local_database.py contains forbidden non-composition marker: {marker}')
    return failures


def check_store_is_thin_facade() -> list[str]:
    failures: list[str] = []
    path = PROJECT_ROOT / 'bm2' / 'store.py'
    if not path.exists():
        return failures
    text = _read_text(path)
    lines = text.splitlines()
    if len(lines) >= 120:
        failures.append(f'bm2/store.py has {len(lines)} lines; store.py must stay under 120 lines')
    forbidden_imports = (
        'bm2.repositories',
        'bm2.excel',
        'bm2.local_database',
        'bm2.sqlite_to_excel_exporter',
        'supabase',
        'openpyxl',
    )
    for imported in _python_imports(path):
        if imported in forbidden_imports or any(imported.startswith(f'{item}.') for item in forbidden_imports):
            failures.append(f'bm2/store.py imports forbidden active dependency {imported}')
    forbidden_text = (
        'ConfigRepository',
        'SupabaseClient',
        'LocalDatabase',
        'WorkbookRepository',
        'SQLiteToExcelExporter',
        'StoreWriterFacade',
        'StoreReaderFacade',
        'local_db.',
        'supabase.',
    )
    for marker in forbidden_text:
        if marker in text:
            failures.append(f'bm2/store.py contains forbidden facade marker: {marker}')
    return failures


def check_store_application_is_thin_coordinator() -> list[str]:
    failures: list[str] = []
    path = PROJECT_ROOT / 'bm2' / 'services' / 'store_application.py'
    if not path.exists():
        return failures
    text = _read_text(path)
    lines = text.splitlines()
    if len(lines) >= 100:
        failures.append(f'bm2/services/store_application.py has {len(lines)} lines; keep it under 100 lines')
    forbidden_imports = (
        'bm2.repositories',
        'bm2.excel',
        'bm2.local_database',
        'bm2.sqlite_to_excel_exporter',
        'supabase',
        'openpyxl',
    )
    for imported in _python_imports(path):
        if imported in forbidden_imports or any(imported.startswith(f'{item}.') for item in forbidden_imports):
            failures.append(f'bm2/services/store_application.py imports forbidden dependency {imported}')
    forbidden_text = (
        'ConfigRepository',
        'SupabaseClient',
        'LocalDatabase',
        'WorkbookRepository',
        'SQLiteToExcelExporter',
        'StoreWriterFacade',
        'StoreReaderFacade',
        'local_db.',
        'supabase.',
    )
    for marker in forbidden_text:
        if marker in text:
            failures.append(f'bm2/services/store_application.py contains forbidden coordinator marker: {marker}')
    return failures


def check_supabase_sdk_boundary() -> list[str]:
    failures: list[str] = []
    allowed = {
        Path('bm2/repositories/supabase_client.py'),
    }
    for path in (PROJECT_ROOT / 'bm2').rglob('*.py'):
        if _is_ignored(path):
            continue
        rel = path.relative_to(PROJECT_ROOT)
        text = _read_text(path)
        if ('from supabase import' in text or 'import supabase' in text) and rel not in allowed:
            failures.append(f'{rel} imports Supabase SDK outside repository/client boundary')
    if (PROJECT_ROOT / 'bm2' / 'supabase_sync.py').exists():
        failures.append('bm2/supabase_sync.py is retired; use bm2/repositories/supabase_client.py')
    return failures


def check_excel_boundary() -> list[str]:
    failures: list[str] = []
    blocked_files = [
        PROJECT_ROOT / 'bm2' / 'web.py',
        PROJECT_ROOT / 'bm2' / 'store.py',
    ]
    blocked_files.extend(
        path
        for path in (PROJECT_ROOT / 'bm2' / 'services').rglob('*.py')
        if path.name != 'migration_service.py'
    )
    patterns = (
        r'\bopenpyxl\b',
        r'\bWorkbook\b',
        r'\bload_workbook\b',
        r'\b_open_workbook\s*\(',
        r'\b_save_workbook\s*\(',
        r'\bworkbook\.',
        r'\bsheet\.',
        r'\bworksheet\b',
    )
    for path in blocked_files:
        if not path.exists() or _is_ignored(path):
            continue
        rel = path.relative_to(PROJECT_ROOT)
        text = _read_text(path)
        for pattern in patterns:
            if re.search(pattern, text):
                failures.append(f'{rel} directly operates Excel/workbook outside Excel boundary: {pattern}')
    return failures


def check_docs_deprecate_json_sync() -> list[str]:
    failures: list[str] = []
    retired = ('JSON blob', 'bm2_cloud', 'cloud_sync', 'cloud_entry_writer', 'score-records.json')
    docs_root = PROJECT_ROOT / 'docs'
    if not docs_root.exists():
        return failures
    for path in docs_root.rglob('*.md'):
        text = _read_text(path)
        if not any(item in text for item in retired):
            continue
        markers = ('deprecated', 'retired', '废弃', '不得扩展')
        if not any(marker in text.lower() for marker in markers[:2]) and not any(marker in text for marker in markers[2:]):
            failures.append(f'{path.relative_to(PROJECT_ROOT)} mentions retired JSON sync without deprecated marker')
    return failures


def collect_debt_warnings() -> list[str]:
    warnings: list[str] = []
    store_path = PROJECT_ROOT / 'bm2' / 'store.py'
    if store_path.exists():
        lines = _read_text(store_path).splitlines()
        if len(lines) > 250:
            warnings.append(f'bm2/store.py is still thick ({len(lines)} lines); keep thinning in later rounds')
        text = '\n'.join(lines)
        for marker in ('local_db.', 'supabase.', 'export_to_excel(', 'get_online_score_sheet_view'):
            if marker in text:
                warnings.append(f'bm2/store.py still coordinates {marker}; migrate to services/repositories')

    transitional_files = {
        'bm2/local_database.py': ('sqlite3',),
        'bm2/store_base.py': ('openpyxl', 'system_config.json', 'workbook'),
    }
    for rel, markers in transitional_files.items():
        path = PROJECT_ROOT / rel
        if not path.exists():
            continue
        text = _read_text(path)
        for marker in markers:
            if marker in text:
                warnings.append(f'{rel} still contains transitional responsibility: {marker}')
    return warnings


def run_checks() -> list[str]:
    failures: list[str] = []
    failures.extend(check_workspace_artifacts())
    failures.extend(check_domain_dependencies())
    failures.extend(check_feature_dependencies())
    failures.extend(check_frontend_requests())
    failures.extend(check_retired_sync_keywords_in_source())
    failures.extend(check_services_boundaries())
    failures.extend(check_web_boundaries())
    failures.extend(check_public_matches_static())
    failures.extend(check_docs_deprecate_json_sync())
    failures.extend(check_retired_write_wrappers())
    failures.extend(check_store_base_retired())
    failures.extend(check_local_database_is_composition_entry())
    failures.extend(check_store_is_thin_facade())
    failures.extend(check_store_application_is_thin_coordinator())
    failures.extend(check_supabase_sdk_boundary())
    failures.extend(check_excel_boundary())
    return failures


def run_checks_with_warnings() -> tuple[list[str], list[str]]:
    return run_checks(), collect_debt_warnings()


def main() -> int:
    failures, warnings = run_checks_with_warnings()
    if failures:
        print('[FAIL] architecture checks')
        for failure in failures:
            print(f'       {failure}')
    else:
        print('[PASS] architecture checks')
    if warnings:
        print('[WARN] architecture debt')
        for warning in warnings:
            print(f'       {warning}')
    return 1 if failures else 0


if __name__ == '__main__':
    raise SystemExit(main())
