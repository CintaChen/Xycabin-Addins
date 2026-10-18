"""Read-only gate for the three delivery roles and unique maintenance files."""
from pathlib import Path
import argparse
import re
import sys
import tempfile
from urllib.parse import unquote

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'src'))
from document_assets import collect as collect_document_assets
PROGRAM, DOCS, DESIGN = 'AP1-程序', 'AP3-说明', 'AP2-美工'
PROJECTS = {'C-01-ExcelIterate': 'Excel_Iterate', 'C-02-WordUpdater': 'WordUpdater', 'C-03-Website·Xycabin': None}


def verify(parent=ROOT.parent, source_only=False, folders=None, strict_evidence=False):
    count = 0
    # Only these exact outputs may be absent before compilation. Missing inputs
    # and other files in AP1 still fail; final delivery checks require all outputs.
    pending_outputs = {(Path(parent) / folder / PROGRAM /
                        (assembly + '_Setup.exe' if assembly else 'XyCabin.html')).resolve()
                       for folder, assembly in PROJECTS.items()}
    for folder, assembly in PROJECTS.items():
        if folders is not None and folder not in folders:
            continue
        base = Path(parent) / folder
        if not base.is_dir():
            continue  # An add-in + Website pair is a supported checkout.
        program = {'Configure_program.py', '一键重建并打包.bat', assembly + '_Setup.exe'} if assembly else {'Configure_Site.py', '一键重建并打包.bat', 'XyCabin.html'}
        expected = {PROGRAM: program, DOCS: {'开发日志.md', '使用说明.md', 'AGENTS.md'}}
        for directory, names in expected.items():
            path = base / directory
            actual = {p.name for p in path.iterdir()} if path.is_dir() else set()
            required = names - ({assembly + '_Setup.exe'} if assembly else {'XyCabin.html'}) if source_only and directory == PROGRAM else names
            if not path.is_dir() or not required.issubset(actual) or not actual.issubset(names) or any(not p.is_file() or p.is_symlink() for p in path.iterdir()):
                raise ValueError('输出目录必须只包含规定文件：' + str(path) + '；应为 ' + ', '.join(sorted(names)))
            count += len(names)
        art = base / DESIGN
        if not art.is_dir() or any(not p.is_file() or p.is_symlink() or p.suffix.lower() != '.png' for p in art.iterdir()):
            raise ValueError('美工目录只允许当前 PNG：' + str(art))
        for obsolete in ('输出文件—程序整合', '输出文件—美工设计', '输出文件—使用说明', '输出', 'README.md', '开发日志.md', 'AGENTS.md', 'Configure_program.py', 'Configure_Site.py', '一键重建并打包.bat'):
            if (base / obsolete).exists():
                raise ValueError('旧维护位置必须消失：' + str(base / obsolete))
        unique = {name: base / directory / name for directory, names in expected.items() for name in names}
        for file in base.rglob('*'):
            if not file.is_file():
                continue
            # A fixed compiler input exists only during central construction;
            # old products must not block a new source build. Final gates forbid it.
            if source_only and any(part.startswith('stage-') for part in file.relative_to(base).parts):
                continue
            if file.name in unique and file.resolve() != unique[file.name].resolve():
                raise ValueError('重复维护或交付文件：' + str(file))
            if file.name in ('README.md', 'readme_probe.py', 'Configure.py', 'AgentReadme.md', 'AgentUserGuide.md', 'AgentHome.html'):
                raise ValueError('遗留维护副本：' + str(file))
            if assembly and file.name in ('XyCabin.html', 'Configure_Site.py'):
                raise ValueError('加载项不得保留网站副本：' + str(file))
        for doc in (base / DOCS).glob('*.md'):
            text = doc.read_text(encoding='utf-8-sig')
            audits = [m.span() for m in re.finditer(r'<!-- delivery-audit:start -->.*?<!-- delivery-audit:end -->', text, re.S)]
            for link in re.finditer(r'(?<!!)\[[^\]]*\]\(([^)]+)\)', text):
                target = link.group(1)
                target = unquote(target.split('#', 1)[0].strip('<>'))
                if not target or re.match(r'^[a-zA-Z]+:', target):
                    continue
                destination = (doc.parent / target).resolve()
                # Documentation may refer to the other host in an independently
                # copied add-in + Website checkout; that project is optional.
                optional = next((Path(parent) / name for name in PROJECTS if name in destination.parts), None)
                if optional is not None and not optional.exists():
                    continue
                if not destination.exists():
                    if source_only and destination in pending_outputs:
                        continue
                    # Audit evidence is optional when carrying source to a new machine.
                    # Only marked audit references into project build directories qualify;
                    # source/config/manual links remain mandatory.
                    in_audit = any(start <= link.start() < end for start, end in audits)
                    generated = optional is not None and destination.is_relative_to(optional / 'build')
                    if in_audit and generated and not strict_evidence:
                        print('INFO: 可选审计证据未携带，不作为重建输入：' + str(destination), flush=True)
                        continue
                    raise ValueError('文档链接失效：' + str(doc) + ' -> ' + target)
        if assembly:
            collect_document_assets(base)  # Includes inline PNGs, SVGs, and code-aware path/content checks.
        launcher = (base / PROGRAM / '一键重建并打包.bat').read_bytes()
        if launcher.startswith(b'\xef\xbb\xbf') or b'\n' in launcher.replace(b'\r\n', b'') or b'chcp 65001' not in launcher or b'PYTHONUTF8=1' not in launcher or b'PYTHONIOENCODING=utf-8' not in launcher:
            raise ValueError('批处理必须 UTF-8 无 BOM、CRLF、65001 和 Python UTF-8：' + str(base))
        launcher.decode('utf-8', 'strict')
    print('PASS: 三类输出、唯一维护位置、旧路径、文档链接及 UTF-8 批处理；%d 项维护/交付文件' % count, flush=True)


def self_test():
    with tempfile.TemporaryDirectory(prefix='xycabin-layout-中文·') as temporary:
        parent = Path(temporary)
        base = parent / 'C-03-Website·Xycabin'
        for directory in (PROGRAM, DOCS, DESIGN):
            (base / directory).mkdir(parents=True)
        for name in ('Configure_Site.py','XyCabin.html'):
            (base / PROGRAM / name).write_text('', encoding='utf-8')
        (base / PROGRAM / '一键重建并打包.bat').write_bytes(b'@echo off\r\nchcp 65001\r\nset PYTHONUTF8=1\r\nset PYTHONIOENCODING=utf-8\r\n')
        for name in ('开发日志.md','使用说明.md','AGENTS.md'):
            (base / DOCS / name).write_text('[日志](开发日志.md)', encoding='utf-8')
        verify(parent)
        page = base / PROGRAM / 'XyCabin.html'
        page.unlink()
        (base / DOCS / '使用说明.md').write_text('[页面](../AP1-程序/XyCabin.html)', encoding='utf-8')
        verify(parent, source_only=True)
        try:
            verify(parent)
        except ValueError:
            pass
        else:
            raise AssertionError('Final gate accepted an absent artifact')
        for host, assembly in PROJECTS.items():
            if assembly is None:
                continue
            host_base = parent / host
            host_base.mkdir()
            guide = base / DOCS / '使用说明.md'
            guide.write_text('[安装包](../../' + host + '/AP1-程序/' + assembly + '_Setup.exe)', encoding='utf-8')
            verify(parent, source_only=True, folders={base.name})
            try:
                verify(parent, folders={base.name})
            except ValueError:
                pass
            else:
                raise AssertionError('Final gate accepted a missing host installer link')
        (base / DOCS / '使用说明.md').write_text('[配置](../AP1-程序/missing.py)', encoding='utf-8')
        try:
            verify(parent, source_only=True, folders={base.name})
        except ValueError:
            pass
        else:
            raise AssertionError('Source gate ignored a missing AP1 input')
        for host, assembly in PROJECTS.items():
            if assembly is not None:
                (parent / host).rmdir()
        (base / DOCS / '使用说明.md').write_text('[日志](开发日志.md)', encoding='utf-8')
        page.write_text('', encoding='utf-8')
        for invalid in (base/'README.md', base/PROGRAM/'SetupDoctor.exe', base/DESIGN/'bad.txt', base/'build/Configure_Site.py'):
            invalid.parent.mkdir(exist_ok=True)
            invalid.write_text('', encoding='utf-8')
            try:
                verify(parent)
            except ValueError:
                pass
            else:
                raise AssertionError('Accepted invalid output: ' + str(invalid))
            invalid.unlink()
        log = base/DOCS/'开发日志.md'
        log.write_text('<!-- delivery-audit:start -->\n[可选证据](../build/old-proof.log)\n<!-- delivery-audit:end -->', encoding='utf-8')
        verify(parent, source_only=True)
        verify(parent)
        try:
            verify(parent, strict_evidence=True)
        except ValueError:
            pass
        else:
            raise AssertionError('Strict evidence gate ignored an absent audit reference')
        log.write_text('[必需入口](../build/missing-tool.py)', encoding='utf-8')
        try:
            verify(parent, source_only=True)
        except ValueError:
            pass
        else:
            raise AssertionError('Source gate ignored a mandatory non-audit link')
        log.write_text('[日志](开发日志.md)', encoding='utf-8')
        (base/DOCS/'AGENTS.md').write_text('[失效](../README.md)', encoding='utf-8')
        try:
            verify(parent)
        except ValueError:
            pass
        else:
            raise AssertionError('Accepted missing document link')
    print('LayoutProbe self-test: 17 PASS, 0 FAIL', flush=True)


if __name__ == '__main__':
    try:
        parser = argparse.ArgumentParser(description=__doc__)
        parser.add_argument('--self-test', action='store_true')
        parser.add_argument('--strict-evidence', action='store_true', help='正式审计同时要求所有证据链接存在')
        args = parser.parse_args()
        if args.self_test:
            self_test()
        verify(strict_evidence=args.strict_evidence)
    except (OSError, ValueError, AssertionError) as error:
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
