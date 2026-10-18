"""Read-only gate: every add-in-owned dialog construction has a registered size identity."""
from pathlib import Path
import json
import re
import sys
sys.dont_write_bytecode=True
from markdown_gate import authored_lines
ROOT=Path(__file__).resolve().parents[2]

def verify_developer_entries(text):
    """Require maintained implementation entries rather than one table layout."""
    text = re.sub(r'<!-- delivery-audit:start -->.*?<!-- delivery-audit:end -->', '', text, flags=re.S)
    start, end = '<!-- shared-wps-experience:start -->', '<!-- shared-wps-experience:end -->'
    if text.count(start) != 1 or text.count(end) != 1 or text.index(start) >= text.index(end):
        raise ValueError('开发日志.md 缺少唯一连续共享开发/维护区')
    shared = '\n'.join(authored_lines(text.split(start)[1].split(end)[0]))
    # Markdown editors may escape punctuation inside visible implementation names.
    # Remove ordinary text escapes only after fenced code has been excluded.
    shared = re.sub(r'\\([_.*])', r'\1', shared)
    topic = re.search(r'^## \d+\. Agent-Evo、内嵌文档与网站\s*\n(.*?)(?=^#{1,2}\s|\Z)', shared, re.M | re.S)
    if topic is None:
        raise ValueError('开发日志.md 缺少Agent-Evo/内嵌文档/网站开发入口')
    required = ('EXTRA_TABS', 'agent-overview', 'delivery-audit', 'WpsShared.Agent.Home.html', 'RequirementsForm')
    missing = [token for token in required if token not in topic[1]]
    lifecycle = ('DialogCatalog', 'DialogSizing.Prepare', 'DialogPreviewFactory', 'DialogTree', 'ThisAddIn', 'DialogSizes.xml')
    missing.extend(token for token in lifecycle if token not in shared)
    if missing:
        raise ValueError('开发日志.md 缺少弹窗寿命或内嵌资源维护入口：' + ', '.join(missing))


def self_test():
    agent = ('EXTRA_TABS', 'agent-overview', 'delivery-audit', 'WpsShared.Agent.Home.html', 'RequirementsForm')
    lifecycle = ('DialogCatalog', 'DialogSizing.Prepare', 'DialogPreviewFactory', 'DialogTree', 'ThisAddIn', 'DialogSizes.xml')
    fixture = ('<!-- shared-wps-experience:start -->\n## 2. 弹窗接线\n' + ' '.join(lifecycle)
               + '\n## 6. Agent-Evo、内嵌文档与网站\n' + ' '.join(agent) + '\n<!-- shared-wps-experience:end -->')
    verify_developer_entries(fixture)
    verify_developer_entries(fixture.replace('## 6.', '## 9.'))
    verify_developer_entries(fixture.replace('EXTRA_TABS', r'EXTRA\_TABS').replace('DialogSizing.Prepare', r'DialogSizing\.Prepare'))
    invalid = [fixture.replace(token, '', 1) for token in agent + lifecycle]
    invalid += [fixture.replace('## 6. Agent-Evo、内嵌文档与网站', '## 6. 其他经验'),
                fixture.replace('## 6.', '```md\n## 6.').replace('<!-- shared-wps-experience:end -->', '```\n<!-- shared-wps-experience:end -->'),
                '<!-- delivery-audit:start -->' + fixture + '<!-- delivery-audit:end -->',
                fixture + '\n<!-- shared-wps-experience:start -->']
    for example in invalid:
        try:
            verify_developer_entries(example)
        except ValueError:
            continue
        raise AssertionError('开发入口缺失/只在代码或历史审计中时应拒绝')

def main():
 project=json.loads((ROOT/'project.json').read_text(encoding='utf-8-sig'))
 expected={'AgentDocumentForm','DeveloperPasswordForm','UiMessageForm','DialogSizesForm','CommandCustomizationForm','RequirementsForm'}
 expected |= {'LinkViewForm'} if project['host']=='Word' else {'TrackerForm','ErrorSourceForm','ColumnFilterForm','FormulaOptionsForm','CustomFunctionForm','LinkAssistantForm','OrganizeFilesForm','RenameFilesForm','FileNamesForm','OcrPrefixesForm','OcrFieldsForm','MergeHeadersForm','EliminationForm'}
 found=set()
 for name in project['sources']:
  text=(ROOT/'src/addin'/name).read_text(encoding='utf-8-sig')
  found.update(re.findall(r'class\s+(\w+Form)\s*:\s*(?:Form|FileToolForm)\b',text))
 found.discard('FileToolForm')
 if found!=expected:raise ValueError('弹窗登记需要复核：'+str(found^expected))
 checks={'AgentDocumentForm.cs':['agent.user','agent.developer','agent.developerpassword'],'UiKit.cs':['message.ok','message.yesno','message.okcancel'], 'DialogSizeEditor.cs':['agent.dialogsizes','agent.developerpassword'], 'AgentCustomization.cs':['agent.notes','agent.'], 'DialogFormatting.cs':['agent.format']}
 if project['host']=='Word':checks['LinkViewForm.cs']=['word.links']
 else:checks.update({'TrackerForm.cs':['trace.precedents','trace.dependents'],'ErrorSourceForm.cs':['trace.errorsources'],'ColumnFilterForm.cs':['columns.filter'],'FormulaOptionsForm.cs':['"formula." + operation.ToString().ToLowerInvariant()'],'CustomFunctionForm.cs':['formula.custom'],'LinkAssistant.cs':['links.assistant'],'FileToolForms.cs':['file.organize','file.rename'],'Connect.FileNames.cs':['file.names'],'MergeHeadersForm.cs':['merge.headers'],'EliminationForm.cs':['elimination.receivable','elimination.revenuecost']})
 for name,tokens in checks.items():
  text=(ROOT/'src/addin'/name).read_text(encoding='utf-8-sig')
  if any(token not in text for token in tokens):raise ValueError('缺少稳定尺寸标识：'+name)
 import importlib.util
 spec=importlib.util.spec_from_file_location('program',ROOT/'AP1-程序/Configure_program.py');module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
 buttons=module.EXTRA_TABS[0]['buttons']
 assert [b['label'] for b in buttons]==['调用API','开发流程','自定配置','实验计划']
 assert [b['group'][1] for b in buttons]==['详情','详情','操作','操作']
 assert buttons[2]['children'][3]['callback']=='OnOpenAgentDialogSizes'
 assert buttons[3]['children'][0]['enabled'] is False
 developer=(ROOT/'AP3-说明/开发日志.md').read_text(encoding='utf-8-sig')
 self_test()
 verify_developer_entries(developer)
 guide=(ROOT/'AP3-说明/使用说明.md').read_text(encoding='utf-8-sig')
 guide_entries=['Agent-Evo','Agent扩展','交互弹窗','开发预览','用户预览','全文']
 guide_entries.extend(button['label'] for button in buttons)
 assert all(label in guide for label in guide_entries),'使用说明.md 缺少Agent-Evo操作入口'
 print('PASS: all dialog constructors, Agent-Evo groups, developer lifecycle/embedded-resource entries and user operation entries registered')
if __name__=='__main__':
 try:main()
 except (OSError,ValueError,AssertionError) as ex:print('FAIL: '+str(ex));sys.exit(1)
