"""Exercise shared source confinement and both products' cache dependencies."""
from pathlib import Path
import argparse
import sys
from unittest.mock import patch

sys.dont_write_bytecode = True
COMMON = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(COMMON))
from source_layout import resolve_source, addin_sources


def verify(project_root):
    root = Path(project_root).resolve()
    repository = root.parent
    checks = 0

    def check(condition, label):
        nonlocal checks
        if not condition:
            raise AssertionError(label)
        checks += 1

    sources = addin_sources(root)
    check(len(sources) == len(set(sources)), 'a product cannot compile the same implementation twice')
    check(all(path.is_file() for path in sources), 'all compiled source references exist')
    for reference in ('@shared/../verify_shared.py', '../AGENTS.md', '@shared/missing-source.cs'):
        try:
            resolve_source(root, reference)
        except ValueError:
            checks += 1
        else:
            raise AssertionError('invalid reference accepted: ' + reference)
    site = repository / 'C-03-Website·Xycabin'
    sys.path.insert(0, str(site / 'src'))
    import build_cache
    snapshot = site / 'AA-Program/XyCabin.html'
    products = [repository / name for name in ('C-01-ExcelIterate', 'C-02-WordUpdater')]
    existing = [product for product in products if product.is_dir()]
    before = [build_cache.inputs(product, snapshot) for product in existing]
    digest = build_cache.digest
    shared_source = resolve_source(root, '@shared/src/addin/UiKit.cs')

    def changed_digest(path):
        return 'changed-shared-source' if Path(path).resolve() == shared_source else digest(path)

    # Inject a hash change instead of editing source during a build or test.
    with patch.object(build_cache, 'digest', side_effect=changed_digest):
        after = [build_cache.inputs(product, snapshot) for product in existing]
    for old, new in zip(before, after):
        check(old != new, 'shared source change invalidates each product cache')
    print('SourceLayoutProbe: %d PASS, 0 FAIL' % checks)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--project-root', type=Path, required=True)
    verify(parser.parse_args().project_root)
