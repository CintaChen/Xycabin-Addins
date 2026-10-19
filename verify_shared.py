"""Check shared source and documentation; never compare old build artifacts."""
from pathlib import Path
import sys

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
SHARED_FILES = (
    'src/pdf/build_trial.py',
    'src/tools/PdfTrialLauncher.cs',
    'src/testharness/PdfToolsProbe.cs',
    'src/testharness/PdfCurrentDocumentProbe.cs',

    'src/addin/Connect.Pdf.cs',
    'src/addin/PdfToolsForm.cs',
    'src/addin/PdfCurrentDocument.cs',
    'src/pdf/PdfWorker.cs',
    'src/pdf/PdfWorker.exe.config',
    'src/pdf/build_worker.py',
    'src/pdf/packages.lock.json',
    'src/pdf/THIRD-PARTY-NOTICES.txt',

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
    'AP3-说明/AGENTS.md', 'AP1-程序/一键重建并打包.bat',
)


def shared_text(project):
    path = project / 'AP3-说明/开发日志.md'
    text = path.read_text(encoding='utf-8-sig')
    start, end = '<!-- shared-wps-experience:start -->', '<!-- shared-wps-experience:end -->'
    if text.count(start) != 1 or text.count(end) != 1 or text.index(start) >= text.index(end):
        raise ValueError(str(path) + ': 共享规范标记必须唯一且有序')
    block = text.split(start, 1)[1].split(end, 1)[0]
    if not block.strip():
        raise ValueError(str(path) + ': 共享规范不能为空')
    return block


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
    if shared_text(left) != shared_text(right):
        errors.append('两份开发日志的共享规范文字不同')
    for relative in ('src/document_assets.py', 'src/site_core.py', 'src/site_template.html',
                     'AP1-程序/Configure_Site.py'):
        path = ROOT / 'C-03-Website·Xycabin' / relative
        if not path.is_file():
            errors.append('缺少集中网站接入文件：' + str(path))
    if errors:
        raise ValueError('\n'.join(errors))
    print('PASS: shared devlog text and %d shared files are identical' % len(SHARED_FILES))


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError) as error:
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
