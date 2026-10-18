"""Verify reviewed single-file cleanup protections using isolated fake evidence."""
from pathlib import Path
import json
import sys
import tempfile
sys.dont_write_bytecode = True
from file_inventory import NAMES, collect, digest, source_fingerprint, scan, apply


def main():
    checks = 0
    with tempfile.TemporaryDirectory(prefix='xy-reviewed-cleanup-') as temporary:
        parent = Path(temporary)
        base, _, site = [parent / name for name in NAMES]
        for name in NAMES:
            (parent / name / 'build').mkdir(parents=True)
            (parent / name / 'src').mkdir()
        source = base / 'src/engine.cs'
        source.write_text('source', encoding='utf-8')
        checker = parent / 'verify_shared.py'
        checker.write_text('shared checker', encoding='utf-8')
        candidate = base / 'build/obsolete.dll'
        review_path = site / 'build/cleanup-review.json'
        log = site / 'build/fixture-build.log'
        proof = site / 'build/fixture-proof.json'
        def setup(target=candidate):
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(b'disposable')
            files, _ = collect(parent)
            fingerprint = source_fingerprint(files, [target])
            log.write_text('simulated successful build for protection unit tests only', encoding='utf-8')
            proof.write_text(json.dumps({'status': 'passed', 'source_fingerprint': fingerprint,
                'excluded_targets': [str(target)], 'build_log': str(log), 'build_log_sha256': digest(log)}), encoding='utf-8')
            review_path.write_text(json.dumps({'schema': 1, 'source_fingerprint': fingerprint,
                'proof': str(proof), 'proof_sha256': digest(proof), 'experience_merged': True,
                'reason': 'isolated test fixture', 'targets': [{'path': str(target), 'base': str(base),
                'identities': {target.name: digest(target)}}]}), encoding='utf-8')
        def rejected(work, label):
            nonlocal checks
            try:
                work()
            except (ValueError, OSError):
                checks += 1
                return
            raise AssertionError(label)
        setup()
        _, entries, _, fingerprint = scan(parent)
        assert len(entries) == 1
        checks += 1
        candidate.write_bytes(b'changed')
        rejected(lambda: apply(entries, parent, fingerprint), 'changed target accepted')
        setup()
        _, entries, _, fingerprint = scan(parent)
        source.write_text('source changed', encoding='utf-8')
        rejected(lambda: apply(entries, parent, fingerprint), 'changed source accepted')
        setup()
        _, entries, _, fingerprint = scan(parent)
        checker.write_text('changed shared checker', encoding='utf-8')
        rejected(lambda: apply(entries, parent, fingerprint), 'changed parent build gate accepted')
        setup()
        log.write_text('tampered', encoding='utf-8')
        rejected(lambda: scan(parent), 'changed isolation log accepted')
        setup(base / 'build/Excel_Iterate.dll')
        rejected(lambda: scan(parent), 'current payload accepted')
        (base / 'build/Excel_Iterate.dll').unlink()
        setup(base / 'build/private.snk')
        rejected(lambda: scan(parent), 'key accepted')
        (base / 'build/private.snk').unlink()
        setup(base / 'build/unknown-recovery.dll')
        rejected(lambda: scan(parent), 'recovery data accepted')
        (base / 'build/unknown-recovery.dll').unlink()
        source.write_text('path = "build/obsolete.dll"', encoding='utf-8')
        setup()
        rejected(lambda: scan(parent), 'current reference accepted')
        source.write_text('source', encoding='utf-8')
        setup()
        _, entries, _, fingerprint = scan(parent)
        apply(entries, parent, fingerprint)
        assert not candidate.exists() and entries[0]['status'] == 'deleted'
        checks += 1
        assert not scan(parent)[1]
        checks += 1
    print('CleanupReviewProbe: %d PASS, 0 FAIL' % checks, flush=True)


if __name__ == '__main__':
    main()
