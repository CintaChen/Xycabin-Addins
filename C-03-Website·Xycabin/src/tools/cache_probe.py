"""Exercise cache receipts/transaction boundaries only in disposable fixtures."""
from pathlib import Path
import json
import os
import sys
import tempfile
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import build_cache as cache


def main():
    checks = 0
    def check(value, label):
        nonlocal checks
        if not value:
            raise AssertionError(label)
        checks += 1
    with tempfile.TemporaryDirectory(prefix='xy-cache-中文·') as temporary:
        base = Path(temporary)
        original_site = cache.SITE
        original_pending = os.environ.get('XY_CACHE_PENDING')
        cache.SITE = base / 'C-03-Website·Xycabin'
        target = cache.SITE / 'build/incremental-cache/check.json'
        output = base / 'artifact.bin'
        output.write_bytes(b'original')
        try:
            os.environ.pop('XY_CACHE_PENDING', None)
            cache.save(target, 'input-A', [output], 'PASS original')
            hit = cache.receipt(target, 'input-A', [output])
            check(hit and hit['output'] == 'PASS original', 'matching bytes and evidence reuse')
            check(cache.receipt(target, 'input-B', [output]) is None, 'input change invalidates')
            check(cache.receipt(target, 'input-A', []) is None, 'expected output set must match')
            old_time = output.stat().st_mtime_ns
            output.write_bytes(b'modified')
            os.utime(output, ns=(old_time, old_time))
            check(cache.receipt(target, 'input-A', [output]) is None, 'same size/time changed bytes invalidate')
            output.unlink()
            check(cache.receipt(target, 'input-A', [output]) is None, 'missing output invalidates')
            output.write_bytes(b'original')
            data = json.loads(target.read_text(encoding='utf-8'))
            data['output'] = 'altered evidence'
            cache.atomic_json(target, data)
            check(cache.receipt(target, 'input-A', [output]) is None, 'altered evidence invalidates')
            target.write_text('{broken', encoding='utf-8')
            check(cache.receipt(target, 'input-A', [output]) is None, 'broken JSON invalidates')
            target.unlink()
            pending = base / 'pending'
            os.environ['XY_CACHE_PENDING'] = str(pending)
            cache.save(target, 'input-A', [output], 'PASS queued')
            check(not target.exists(), 'queued receipt invisible before commit')
            cache.commit(pending)
            check(cache.receipt(target, 'input-A', [output]) is not None, 'successful commit publishes receipt')
            cache.atomic_json(pending / 'escape.json', {'target': str(base / 'outside.json'), 'receipt': {}})
            try:
                cache.commit(pending)
                raise AssertionError('escaped target accepted')
            except ValueError:
                check(not (base / 'outside.json').exists(), 'commit cannot escape cache roots')
        finally:
            cache.SITE = original_site
            if original_pending is None:
                os.environ.pop('XY_CACHE_PENDING', None)
            else:
                os.environ['XY_CACHE_PENDING'] = original_pending
    print('CacheProbe: %d PASS, 0 FAIL' % checks, flush=True)


if __name__ == '__main__':
    main()
