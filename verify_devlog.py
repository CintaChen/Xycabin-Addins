"""Read-only gate for the five documents in C-04-Devlog.

Replaces the per-project five-chapter devlog rule: documents now live once at
the repository root, so the gate checks each file's own structure instead of
comparing two copies.
"""
from pathlib import Path
import re
import sys

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
DOCS = ROOT / 'C-04-Devlog'
FILES = ('B0-Entry.md', 'A-UserGuide.md', 'B1-Project-Files.md',
         'B2-DevExperience.md', 'B3-Uphold-Discipline.md')
LIMIT = 200


def markdown_gate():
    for folder in ('C-01-ExcelIterate', 'C-02-WordUpdater', 'C-03-Website·Xycabin'):
        tools = ROOT / folder / 'src/tools'
        if (tools / 'markdown_gate.py').is_file():
            sys.path.insert(0, str(tools))
            import markdown_gate as module
            return module
    raise FileNotFoundError('找不到 markdown_gate.py')


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


def chapter_blocks(text):
    blocks, current = [], None
    fence = None
    for line in text.splitlines():
        if re.match(r'^\s{0,3}(`{3,}|~{3,})', line):
            fence = None if fence else line.strip()[:3]
        if fence is None and line.startswith('# '):
            if current:
                blocks.append(current)
            current = [line, []]
            continue
        if current:
            current[1].append(line)
    if current:
        blocks.append(current)
    return blocks


def check_dates(text, label, required):
    cleaned = re.sub(r'(?ms)^```.*?^```[ \t]*$', '', text)
    cleaned = re.sub(r'https?://[^\s)]+|`[^`\n]+`', '', cleaned)
    dates = re.findall(r'\d{4}-\d{2}-\d{2}', cleaned)
    top = re.findall(r'^最后更新：(\d{4}-\d{2}-\d{2})', text, re.M)
    if len(top) > 1:
        return [label + ': 只允许一处“最后更新”日期']
    if dates and (len(top) != 1 or dates != top):
        return [label + ': 正文不得出现“最后更新”以外的日期']
    if required and len(top) != 1:
        return [label + ': 缺少顶部“最后更新”日期']
    return []


def check_scopes(text, label):
    errors, stack = [], []
    for match in re.finditer(r'<!--\s*scope:([a-z]+):(start|end)\s*-->', text):
        name, kind = match.groups()
        if kind == 'start':
            stack.append(name)
        elif not stack or stack.pop() != name:
            errors.append(label + ': scope 标记未成对：' + name)
    if stack:
        errors.append(label + ': scope 标记未闭合：' + ', '.join(stack))
    return errors


def version_rows(block, label):
    errors, seen, previous = [], set(), None
    section = re.search(r'^## \d+\. 版本摘要\s*\n(.*?)(?=^## |\Z)', block, re.M | re.S)
    if section is None:
        return [label + ': 缺少版本摘要章节']
    for line in section.group(1).splitlines():
        row = re.match(r'^\s*\|\s*(\d+\.\d+\.\d+\.\d+)\s*\|\s*(.*?)\s*\|\s*$', line)
        if not row:
            continue
        version, summary = row.groups()
        if version in seen:
            errors.append(label + ': 重复版本 ' + version)
        seen.add(version)
        numeric = tuple(map(int, version.split('.')))
        if previous is not None and numeric < previous:
            errors.append(label + ': 版本摘要须升序：' + version)
        previous = numeric
        count = len(visible_text(summary))
        if count > LIMIT:
            errors.append(label + ': ' + version + ' 摘要 ' + str(count) + ' 字，超过 200 字')
    if not seen:
        errors.append(label + ': 版本摘要没有可识别的版本行')
    return errors


def verify_file(path):
    gate = markdown_gate()
    text = path.read_text(encoding='utf-8-sig')
    name = path.name
    errors, titles = gate.hierarchy(text, name)
    try:
        gate.verify_local_anchors(path)
    except ValueError as error:
        errors += str(error).splitlines()
    errors += check_scopes(text, name)
    errors += check_dates(text, name, required=(name == 'B1-Project-Files.md'))
    blocks = chapter_blocks(text)

    if name == 'B0-Entry.md':
        required = ('入口与阅读顺序', '三个输出的关系', '共享与同步', '交付边界')
        missing = [item for item in required if not re.search(r'(?m)^## \d+\. ' + re.escape(item) + r'$', text)]
        if missing:
            errors.append(name + ': 缺少章节 ' + ', '.join(missing))
    elif name == 'A-UserGuide.md':
        heads = [block[0][2:] for block in blocks]
        expected = ('Excel-Iterate', 'WordUpdater', 'XyCabin')
        if tuple(item.split('. ', 1)[-1] for item in heads) != expected:
            errors.append(name + ': 一级章节必须依次为三个产品')
        if text.count('<!-- agent-overview:start -->') != 1 or text.count('<!-- agent-overview:end -->') != 1:
            errors.append(name + ': 概览标记必须唯一')
    elif name == 'B1-Project-Files.md':
        if len(blocks) != 4:
            errors.append(name + ': 必须包含四个一级章节')
        if text.count('<!-- agent-overview:start -->') != 1 or text.count('<!-- agent-overview:end -->') != 1:
            errors.append(name + ': 概览标记必须唯一')
        starts = text.count('<!-- delivery-audit:start -->')
        ends = text.count('<!-- delivery-audit:end -->')
        if starts != ends or not 1 <= starts <= 3:
            errors.append(name + ': 交付审计标记须按项目成对，数量 1-3')
        for product in ('Excel-Iterate', 'WordUpdater'):
            block = next((item for item in blocks if product + ' 项目档案' in item[0]), None)
            if block is None:
                errors.append(name + ': 缺少 ' + product + ' 项目档案章节')
            else:
                errors += version_rows('\n'.join(block[1]), name + '/' + product)
    elif name == 'B2-DevExperience.md':
        if len(blocks) != 4:
            errors.append(name + ': 必须包含四个一级章节')
    elif name == 'B3-Uphold-Discipline.md':
        if len(blocks) != 2:
            errors.append(name + ': 必须包含两个一级章节')
    return errors


def verify(files=None):
    paths = [DOCS / name for name in FILES] if files is None else [Path(item) for item in files]
    errors = []
    for path in paths:
        if not path.is_file():
            errors.append('缺少文档：' + str(path))
            continue
        errors += verify_file(path)
    if errors:
        raise ValueError('\n'.join(errors))
    print('PASS: C-04 五份文档的编号、锚点、范围标记、版本表与日期；' + str(len(paths)) + ' 份')


def self_test():
    import tempfile
    gate = markdown_gate()
    assert gate.hierarchy('# 一. 甲\n## 1. 乙\n', 'ok')[0] == []
    assert len(visible_text('**中文** [ABC](https://example.com) `2.0` ！')) == 9
    assert version_rows('## 2. 版本摘要\n| 0.9.1.0 | 甲 |\n| 0.9.2.0 | 乙 |\n', 'ok') == []
    assert version_rows('## 2. 版本摘要\n| 0.9.2.0 | 甲 |\n| 0.9.1.0 | 乙 |\n', 'bad')
    assert version_rows('## 2. 版本摘要\n| 0.9.1.0 | ' + '中' * 201 + ' |\n', 'long')
    assert check_scopes('<!-- scope:excel:start -->\n甲\n<!-- scope:excel:end -->\n', 'ok') == []
    assert check_scopes('<!-- scope:excel:start -->\n甲\n', 'bad')
    with tempfile.TemporaryDirectory() as temporary:
        path = Path(temporary) / 'B2-DevExperience.md'
        path.write_text('# 一. 甲\n## 1. 乙\n# 二. 丙\n# 三. 丁\n# 四. 戊\n', encoding='utf-8')
        assert verify_file(path) == []
    print('VerifyDevlog self-test: 8 PASS, 0 FAIL', flush=True)


def main():
    arguments = sys.argv[1:]
    if arguments and arguments[0] == '--self-test':
        self_test()
        return
    verify(arguments or None)


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, AssertionError) as error:
        print('FAIL: ' + str(error), flush=True)
        sys.exit(1)
