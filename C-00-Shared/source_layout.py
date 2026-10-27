"""Resolve the physical inputs of a product without copying shared implementations."""
from pathlib import Path
import json

SHARED_ROOT = Path(__file__).resolve().parent


def manifest():
    return json.loads((SHARED_ROOT / 'source-manifest.json').read_text(encoding='utf-8'))


def resolve_source(project_root, reference, *, addin=False):
    root = Path(project_root).resolve()
    value = str(reference).replace('\\', '/')
    if value.startswith('@shared/'):
        relative, base = value[8:], SHARED_ROOT
    else:
        relative, base = value, root
        if addin and '/' not in relative:
            relative = 'src/addin/' + relative
        # Compatibility for maintained entry scripts and source-location probes.
        if relative in manifest()['linked']:
            base = SHARED_ROOT
    raw = Path(relative)
    if raw.is_absolute() or '..' in raw.parts or ':' in relative:
        raise ValueError('Source path escapes allowed roots: ' + value)
    target = (base / raw).resolve()
    if base not in target.parents or not target.is_file():
        raise ValueError('Missing or invalid source: ' + value)
    return target


def addin_sources(project_root, project=None):
    root = Path(project_root).resolve()
    project = project or json.loads((root / 'project.json').read_text(encoding='utf-8-sig'))
    return [resolve_source(root, item, addin=True) for item in project['sources']]


def run_entry(namespace, source, project_root):
    root = Path(project_root).resolve()
    namespace['PROJECT_ROOT'] = root
    namespace['source_path'] = lambda reference, addin=False: resolve_source(root, reference, addin=addin)
    namespace['source_files'] = lambda: addin_sources(root)
    target = SHARED_ROOT / source
    exec(compile(target.read_bytes(), str(target), 'exec'), namespace)
