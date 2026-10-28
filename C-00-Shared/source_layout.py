"""Resolve the physical inputs of a product without copying shared implementations."""
from pathlib import Path
import json

SHARED_ROOT = Path(__file__).resolve().parent


def manifest():
    return json.loads((SHARED_ROOT / 'source-manifest.json').read_text(encoding='utf-8'))


def project_data(project_root):
    return json.loads((Path(project_root) / 'project.json').read_text(encoding='utf-8-sig'))


def source_references(project_root, reference):
    """An explicit legacy read alias may span several real compiled source files."""
    project = project_data(project_root)
    value = str(reference).replace('\\', '/')
    if '/' not in value or value.startswith('src/addin/'):
        name = value.rsplit('/', 1)[-1]
        aliases = project.get('source_aliases', {})
        if name in aliases:
            references = aliases[name]
            if not references or any(ref not in project['sources'] for ref in references):
                raise ValueError('Source read alias must refer to compiled sources: ' + value)
            return references
        candidates = [ref for ref in project['sources'] if ref.rsplit('/', 1)[-1] == name]
        if len(candidates) > 1:
            raise ValueError('Ambiguous source name: ' + value)
        if candidates:
            return candidates
    return [value]


def source_text(project_root, reference):
    return '\n'.join(resolve_source(project_root, ref).read_text(encoding='utf-8-sig')
                     for ref in source_references(project_root, reference))


def resolve_source(project_root, reference, *, addin=False):
    root = Path(project_root).resolve()
    value = str(reference).replace('\\', '/')
    if value.startswith('@shared/'):
        relative, base = value[8:], SHARED_ROOT
    else:
        relative, base = value, root
        if addin and '/' not in relative:
            references = source_references(root, relative)
            if len(references) != 1:
                raise ValueError('Split source requires source_text: ' + value)
            if references[0] != relative:
                return resolve_source(root, references[0])
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
    namespace['source_text'] = lambda reference: source_text(root, reference)
    target = SHARED_ROOT / source
    exec(compile(target.read_bytes(), str(target), 'exec'), namespace)
