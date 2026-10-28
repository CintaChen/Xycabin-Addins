"""Validate the canonical page and both current published add-in packages."""
from pathlib import Path
import argparse
import sys
sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'src'))
from delivery import projects, run, verify_payload, RUN_STAGES
from site_core import verify_installers
from build_cache import atomic_json

MODE = 'full'
CLASSIFICATION = {'rule': 'R-30', 'status': 'not_run', 'cached': False}


def write_summary(status):
    classification = dict(CLASSIFICATION)
    if classification['status'] == 'running':
        classification['status'] = 'fail'
    atomic_json(ROOT / 'build/validate-run-summary.json',
                {'mode': MODE, 'status': status, 'stages': RUN_STAGES,
                 'classification_gate': classification})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--mode', choices=('incremental', 'full'), default='full')
    parser.add_argument('--site-only', action='store_true')
    parser.add_argument('--skip-ui', action='store_true')
    parser.add_argument('--background', action='store_true', help='不运行会显示或激活窗口的探针，前台项目列为待验收')
    global MODE, CLASSIFICATION
    args = parser.parse_args()
    MODE = args.mode
    sys.path.insert(0, str(ROOT.parent / 'C-00-Shared'))
    from feature_index import check as check_features
    if not args.site_only:
        CLASSIFICATION['status'] = 'running'
        CLASSIFICATION = check_features(ROOT.parent)
    run([sys.executable, ROOT / 'src/tools/cleanup_probe.py'], ROOT, ROOT / 'build/cleanup-validation.log')
    run([sys.executable, ROOT / 'src/tools/cache_probe.py'], ROOT, ROOT / 'build/cache-validation.log')
    run([sys.executable, ROOT / 'src/tools/cleanup_review_probe.py'], ROOT, ROOT / 'build/cleanup-review-validation.log')
    run([sys.executable, ROOT / 'src/tools/devlog_probe.py'], ROOT, ROOT / 'build/devlog-validation.log')
    site = ROOT / 'AA-Program/XyCabin.html'
    run([sys.executable, ROOT / 'src/tools/site_probe.py', '--html', site], ROOT, ROOT / 'build/site-validation.log')
    if args.site_only:
        return
    from tools.layout_probe import verify as verify_layout
    verify_layout(ROOT.parent)
    projects(ROOT)
    from delivery import version
    bases=projects(ROOT)
    verify_installers(site.read_text(encoding='utf-8'),{
        'word' if assembly=='WordUpdater' else 'excel':(base/'AA-Program'/(assembly+'_Setup.exe'),version(base))
        for base,assembly in bases})
    run([sys.executable, ROOT / 'src/tools/workflow_probe.py'], ROOT, ROOT / 'build/workflow-validation.log')
    checker = ROOT.parent / 'verify_shared.py'
    if not checker.is_file():
        raise FileNotFoundError('缺少共享检查入口：' + str(checker))
    run([sys.executable, checker], ROOT.parent, ROOT / 'build/shared-source.log')
    for base, assembly in projects(ROOT):
        installer = base / 'AA-Program' / (assembly + '_Setup.exe')
        command = [sys.executable, base / 'validate.py', '--mode', args.mode, '--site-html', site, '--installer', installer]
        if args.skip_ui:
            command.append('--skip-ui')
        if args.background:
            command.append('--background')
        run(command, base, ROOT / ('build/' + assembly + '-validation.log'))
        verify_payload(base, assembly, installer, site, ROOT)
    print('PASS: 独立网页中的两包与正式 EXE 一致；移除安装包数据后与两 DLL 网页一致', flush=True)
    write_summary('success')


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, RuntimeError) as error:
        try:
            write_summary('failed')
        except OSError:
            pass
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
