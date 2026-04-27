from __future__ import annotations

import ast
import re
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


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
    tree = ast.parse(path.read_text(encoding='utf-8'))
    imports: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imports.extend(alias.name for alias in node.names)
        if isinstance(node, ast.ImportFrom):
            imports.append(_resolve_import(module_name, node))
    return imports


def check_domain_dependencies() -> list[str]:
    failures: list[str] = []
    blocked_prefixes = ('bm2.web', 'bm2.presenters', 'bm2.features')
    blocked_exact = {'flask'}
    for path in (PROJECT_ROOT / 'bm2' / 'domain').rglob('*.py'):
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
        text = path.read_text(encoding='utf-8')
        if path.name == 'request.js':
            continue
        if re.search(r'\bfetch\s*\(', text):
            failures.append(f'{path.relative_to(PROJECT_ROOT)} uses fetch directly')
    return failures


def run_checks() -> list[str]:
    failures: list[str] = []
    failures.extend(check_domain_dependencies())
    failures.extend(check_feature_dependencies())
    failures.extend(check_frontend_requests())
    return failures


def main() -> int:
    failures = run_checks()
    if failures:
        print('[FAIL] architecture checks')
        for failure in failures:
            print(f'       {failure}')
        return 1
    print('[PASS] architecture checks')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
