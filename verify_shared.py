"""Check shared source and documentation; never compare old build artifacts."""
from pathlib import Path
import sys

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
SHARED_FILES = (
    'src/tools/AgentDocumentCoreProbe.cs', 'src/tools/GripProbe.cs', 'src/tools/LiveProbeSafety.cs', 'src/testharness/LiveProbeSafetyFixture.cs',
    'src/addin/CustomizationRestart.cs', 'src/shared/RestartProtocol.cs', 'src/tools/RestartHelper.cs',
    'src/tools/window_thread_probe.py', 'src/testharness/HostDialogProbe.cs', 'src/addin/HostDialogService.cs', 'src/addin/UiDiagnostics.cs', 'src/addin/CustomizationSaving.cs', 'src/addin/FormatSession.cs', 'src/addin/FormatEditorThread.cs',
    'src/testharness/Feature099Probe.cs', 'src/testharness/GridHeadersProbe.cs',
    'src/addin/AgentCustomization.cs', 'src/addin/DialogFormatting.cs', 'src/addin/Connect.Customization.cs',
    'build.py', 'build_installer.py', 'gen_manifest.py',
    'lib/Interop.Microsoft.Office.Core.dll',
    'lib/Microsoft.Office.Tools.Common.dll',
    'lib/Microsoft.Office.Tools.Common.v4.0.Utilities.dll',
    'lib/Microsoft.Office.Tools.dll',
    'lib/Microsoft.VisualStudio.Tools.Applications.Runtime.dll',
    'project_support.py', 'rebuild.py', 'resign.py',
    'src/addin/AgentDocumentForm.cs', 'src/addin/AgentLocalFiles.cs',
    'src/addin/Connect.Agent.cs', 'src/addin/DialogSizes.cs', 'src/addin/DialogSizeEditor.cs', 'src/addin/DialogPreviewFactory.cs', 'src/addin/InteropDefs.cs', 'src/addin/UiKit.cs',
    'src/configuration/configure_core.py', 'src/installer/Installer.cs',
    'src/installer/InstallFlow.cs', 'src/tools/AddInProbe.cs',
    'src/tools/AgentExtensionProbe.cs', 'src/tools/DialogSizesProbe.cs', 'src/tools/dialog_catalog_probe.py', 'src/tools/configure_probe.py',
    'src/tools/devlog_probe.py', 'src/tools/markdown_gate.py', 'src/tools/GetPkt.cs', 'src/tools/InstallFlowProbe.cs',
    'src/tools/MakeCertCapi.cs', 'src/tools/RibbonProbe.cs', 'src/tools/SetupDoctor.cs',
    'src/tools/SignCmi.cs', 'src/tools/TrustAdd.cs', 'src/tools/UiPreviewProbe.cs',
    'src/tools/ValidateManifest.cs', 'src/tools/ValidateXmlSig.cs',
    'validate.py', 'verify_package.py',
    'AA-Program/一键重建并打包.bat',
)

DOCS = {'B0-Entry.md', 'A-UserGuide.md', 'B1-Project-Files.md',
        'B2-DevExperience.md', 'B3-Gate Control.md'}
SHARED_TITLES = ('# 一. Word与Excel共享开发经验', '# 一. Word与Excel共享门禁')


# rule=R-01
def check_documents():
    """Documents live once in C-04-Devlog; projects keep no copies."""
    agents = ROOT / 'AGENTS.md'
    if not agents.is_file():
        raise ValueError('缺少仓库根 AGENTS.md 指针')
    pointer = agents.read_text(encoding='utf-8-sig')
    if (len(pointer.encode('utf-8')) > 300 or pointer.count('](') != 1
            or '](C-04-Devlog/B0-Entry.md)' not in pointer):
        raise ValueError('AGENTS.md 必须只放一个指向 C-04-Devlog/B0-Entry.md 的链接')
    if '只作指针' not in pointer or '<!--' in pointer:
        raise ValueError('AGENTS.md 不得维护规则或包含注释')
    docs = ROOT / 'C-04-Devlog'
    actual = {p.name for p in docs.iterdir()} if docs.is_dir() else set()
    if actual != DOCS:
        raise ValueError('C-04-Devlog 必须恰好包含五份文档：' + ', '.join(sorted(actual)))
    texts = {name: (docs / name).read_text(encoding='utf-8-sig') for name in DOCS}
    for title in SHARED_TITLES:
        count = sum(1 for text in texts.values() for line in text.splitlines() if line.strip() == title)
        if count != 1:
            raise ValueError('共享章节必须只保留一份：' + title)
    for name, text in texts.items():
        if 'shared-wps-experience' in text:
            raise ValueError(name + ': 旧的共享标记必须删除')
    for project in (ROOT / 'C-01-ExcelIterate', ROOT / 'C-02-WordUpdater', ROOT / 'C-03-Website·Xycabin'):
        for name in ('开发日志.md', '使用说明.md', 'AGENTS.md'):
            leftovers = [p for p in project.rglob(name) if 'build' not in p.parts]
            if leftovers:
                raise ValueError('项目内不得保留文档副本：' + str(leftovers[0]))


# rule=R-15
def main():
    left, right = ROOT / 'C-01-ExcelIterate', ROOT / 'C-02-WordUpdater'
    for project in (left, right):
        if not project.is_dir():
            raise FileNotFoundError('共享检查需要两个加载项目录：' + str(project))
    errors = []
    for relative in SHARED_FILES:
        paths = [project / relative for project in (left, right)]
        missing = [str(path) for path in paths if not path.is_file()]
        if missing:
            errors.append('缺少公共文件：' + '；'.join(missing))
        elif paths[0].read_bytes() != paths[1].read_bytes():
            errors.append('公共文件内容不同：' + relative)
    try:
        check_documents()
    except (OSError, ValueError) as error:
        errors.append(str(error))
    for relative in ('src/document_assets.py', 'src/site_core.py', 'src/site_template.html',
                     'AA-Program/C3-XycabinSite_Configure.py'):
        path = ROOT / 'C-03-Website·Xycabin' / relative
        if not path.is_file():
            errors.append('缺少集中网站接入文件：' + str(path))
    if errors:
        raise ValueError('\n'.join(errors))
    print('PASS: C-04 documents and %d shared files are consistent' % len(SHARED_FILES))


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError) as error:
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
