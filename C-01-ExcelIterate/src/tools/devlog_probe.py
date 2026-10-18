"""Read-only development-log handoff gate. Summaries must be authored, never truncated."""
from pathlib import Path
import re
import sys
sys.dont_write_bytecode = True

from markdown_gate import hierarchy as numbering, verify_guide, verify_entry, verify_local_anchors, self_test as numbering_test

LIMIT = 200


def visible_text(text):
    escaped = []
    def protect(match):
        escaped.append(match.group(1))
        return '\ue000' + str(len(escaped) - 1) + '\ue001'
    text = re.sub(r'\\([\\`*_{}\[\]()#+.!|>~-])', protect, text)
    text = re.sub(r'!?\[([^\]]*)\]\([^)]*\)', r'\1', text)
    text = re.sub(r'<[^>]*>', '', text)
    text = re.sub(r'(`+)(.*?)\1', r'\2', text)
    for token in ('**', '__', '*', '_', '~~'):
        text = re.sub(r'(?<!\w)' + re.escape(token) + r'([^\n]+?)' + re.escape(token) + r'(?!\w)', r'\1', text)
    text = re.sub('\ue000(\\d+)\ue001', lambda m: escaped[int(m.group(1))], text)
    text = re.sub(r'(?<!\d)\d+(?:\.\d+){3}(?!\d)', '', text)
    return re.sub(r'\s', '', text)


def inspect(text, label):
    errors, counts = [], []
    # API versions and dated source URLs are technical identifiers, not updates.
    date_text = re.sub(r'(?ms)^```.*?^```[ \t]*$', '', text)
    date_text = re.sub(r'https?://[^\s)]+|`[^`\n]+`', '', date_text)
    dates = re.findall(r'\d{4}-\d{2}-\d{2}', date_text)
    top = re.findall(r'^最后更新：(\d{4}-\d{2}-\d{2})', text, re.M)
    if len(top) != 1 or len(dates) != 1 or dates != top:
        errors.append(label + ': 仅允许顶部一处“最后更新”日期，旧维护日期应归并或删除')
    section = re.search(r'^## \d+\. 版本摘要\s*\n(.*?)(?=^## |\Z)', text, re.M | re.S)
    if section is None:
        return errors + [label + ': 缺少版本摘要章节'], counts
    seen = set()
    previous = None
    for line in section.group(1).splitlines():
        row = re.match(r'^\s*\|\s*(\d+\.\d+\.\d+\.\d+)\s*\|\s*(.*?)\s*\|\s*$', line)
        if not row:
            continue
        version, summary = row.groups()
        if version in seen:
            errors.append(label + ': 重复版本 ' + version + '；同版本须人工归并为一条')
        seen.add(version)
        numeric = tuple(map(int, version.split('.')))
        if previous is not None and numeric < previous:
            errors.append(label + ': 版本摘要须按四段数字升序排列，新版本追加在下方：' + version)
        previous = numeric
        count = len(visible_text(summary))
        counts.append((version, count))
        if count > LIMIT:
            errors.append(label + ': ' + version + ' 摘要 ' + str(count) + ' 字，超过200字；请先归并重要变更再交付，不可直接截断')
    if not seen:
        errors.append(label + ': 版本摘要没有可识别的版本行')
    return errors, counts


def verify(path):
    text = Path(path).read_text(encoding='utf-8-sig')
    verify_local_anchors(Path(path))
    errors, counts = inspect(text, str(path))
    errors += hierarchy(text, str(path))
    if errors:
        raise ValueError('\n'.join(errors))
    verify_guide(Path(path).with_name('使用说明.md'))
    verify_entry(Path(path).with_name('AGENTS.md'))
    return counts


def hierarchy(text, label):
    errors, actual = numbering(text, label)
    expected = ['一. Agent及开发者接手概览', '二. 项目档案', '三. Word与Excel共享开发经验', '四. Word与Excel共享维护纪律']
    if actual[:4] != expected or len(actual) != 5 or actual[-1] not in ('五. Excel专属开发经验', '五. Word专属开发经验'):
        errors.append(label + ': 必须包含约定的五个一级章节')
    shared = re.search(r'<!-- shared-wps-experience:start -->(.*?)<!-- shared-wps-experience:end -->', text, re.S)
    if shared is None or numbering(shared.group(1), label)[1] != expected[2:4]:
        errors.append(label + ': 共享标记须连续包围第三、第四章')
    for start, end in (('agent-overview:start', 'agent-overview:end'), ('shared-wps-experience:start', 'shared-wps-experience:end'), ('delivery-audit:start', 'delivery-audit:end')):
        a, b = '<!-- ' + start + ' -->', '<!-- ' + end + ' -->'
        if text.count(a) != 1 or text.count(b) != 1 or text.find(a) >= text.find(b):
            errors.append(label + ': 标记缺失、重复或乱序：' + start)
    if re.search(r'^# 第[一二三]部分|第二部分[「“]共享开发规范|第三部分[ ·]+WPS (?:表格|文字)业务经验', text, re.M):
        errors.append(label + ': 仍有旧章节名称，须同步更新引用')
    return errors


def self_test():
    numbering_test()
    def doc(summary):
        return '最后更新：2026-10-04\n## 2. 版本摘要\n| 0.9.7.0 | ' + summary + ' |\n## 3. 源码地图\n'
    assert not inspect(doc('中' * 200), 'boundary')[0]
    assert '201' in inspect(doc('中' * 201), 'boundary')[0][0]
    assert len(visible_text('**中文** [ABC](https://example.com) `2.0` ！')) == 9
    assert visible_text('  A\n B\tC  ') == 'ABC'
    assert visible_text('WPS_WORD') == 'WPS_WORD'
    assert visible_text('WPS_WORD_EXCEL') == 'WPS_WORD_EXCEL'
    assert visible_text('维护 **0.9.7.0**：修复A1') == '维护：修复A1'
    assert visible_text(r'\*文本\*') == '*文本*'
    assert any('重复版本' in e for e in inspect(doc('a').replace('## 3.', '| 0.9.7.0 | b |\n## 3.'), 'duplicate')[0])
    assert inspect(doc('a') + '\n旧维护：2026-10-01', 'old-date')[0]
    assert not inspect(doc('a') + '\nAPI版本 `2021-07-07`，参考[接口](https://example.com/api-2021-07-07)', 'api-version')[0]
    assert inspect(doc('a').replace('最后更新：', '更新：'), 'missing')[0]
    assert inspect(doc('a') + '\n最后更新：2026-10-04', 'duplicate-date')[0]
    ordered = doc('a').replace('## 3.', '| 0.9.8.0 | b |\n| 0.9.10.0 | c |\n## 3.')
    assert not inspect(ordered, 'numeric-order')[0]
    reversed_rows = doc('a').replace('## 3.', '| 0.9.6.0 | b |\n## 3.')
    assert any('升序' in e for e in inspect(reversed_rows, 'reversed')[0])
    fixture = '\n'.join(['# 一. Agent及开发者接手概览', '<!-- agent-overview:start -->', '概览', '<!-- agent-overview:end -->', '# 二. 项目档案', '## 1. 项目定位', '### (1). 用途', '<!-- delivery-audit:start -->', '审计', '<!-- delivery-audit:end -->', '<!-- shared-wps-experience:start -->', '# 三. Word与Excel共享开发经验', '## 1. 经验', '# 四. Word与Excel共享维护纪律', '## 1. 维护', '<!-- shared-wps-experience:end -->', '# 五. Excel专属开发经验', '## 1. 单元格', '### (1). 公式'])
    assert not hierarchy(fixture, 'valid-hierarchy')
    assert hierarchy(fixture.replace('### (1). 公式', '### (2). 公式'), 'heading-gap')
    assert hierarchy(fixture.replace('# 四. Word与Excel共享维护纪律', '<!-- shared-wps-experience:end -->\n# 四. Word与Excel共享维护纪律').replace('维护\n<!-- shared-wps-experience:end -->', '维护'), 'shared-missing-discipline')
    assert hierarchy(fixture.replace('# 五. Excel专属开发经验', '# 六. Excel专属开发经验'), 'main-title')
    assert hierarchy(fixture + '\n<!-- agent-overview:start -->', 'marker-duplicate')
    assert not hierarchy(fixture.replace('Excel专属开发经验', 'Word专属开发经验'), 'word-hierarchy')



def main():
    self_test()
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parents[2] / 'AP3-说明/开发日志.md'
    counts = verify(path)
    print('PASS: 五部分、三份文档五级编号及本地标题锚点、连续共享经验/纪律、版本升序、200字边界、重复版本与日期；' + '，'.join(v + '=' + str(n) + '字' for v, n in counts))


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, AssertionError) as ex:
        print('FAIL:', ex)
        sys.exit(1)
