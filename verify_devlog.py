"""Read-only gate for the five documents in C-04-Devlog.

Structure is derived from level-one chapter titles; no invisible HTML markers
are used. Scopes, overview and delivery-evidence sections are computed by the
same helpers that compose the embedded documents, so gate and build agree.
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
AUDIT_HEAD = re.compile(r'^###\s+\(\d+\)\.\s*本轮交付证据\s*$')
EXPECTED_SCOPES = {
    'A-UserGuide.md': ['excel', 'word', 'site'],
    'B1-Project-Files.md': ['shared', 'excel', 'word', 'site'],
    'B2-DevExperience.md': ['shared', 'excel', 'word', 'site'],
    'B3-Uphold-Discipline.md': ['shared', 'site'],
}


def markdown_gate():
    for folder in ('C-01-ExcelIterate', 'C-02-WordUpdater', 'C-03-Website·Xycabin'):
        tools = ROOT / folder / 'src/tools'
        if (tools / 'markdown_gate.py').is_file():
            sys.path.insert(0, str(tools))
            import markdown_gate as module
            return module
    raise FileNotFoundError('找不到 markdown_gate.py')


def document_rules():
    source = ROOT / 'C-03-Website·Xycabin/src'
    if (source / 'document_assets.py').is_file():
        if str(source) not in sys.path:
            sys.path.insert(0, str(source))
        import document_assets
        return document_assets
    raise FileNotFoundError('找不到 document_assets.py')


def check_no_comments(text, label):
    return [label + ': 文档不得包含 HTML 注释'] if '<!--' in text else []


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
        count = len(summary.strip())
        if count > LIMIT:
            errors.append(label + ': ' + version + ' 摘要 ' + str(count) + ' 字，超过 200 字')
    if not seen:
        errors.append(label + ': 版本摘要没有可识别的版本行')
    return errors


def chapter_overview(lines):
    index = 0
    while index < len(lines) and (not lines[index].strip() or lines[index].startswith('#')):
        index += 1
    return index < len(lines) and bool(lines[index].strip())


def verify_file(path):
    gate = markdown_gate()
    rules = document_rules()
    text = path.read_text(encoding='utf-8-sig')
    name = path.name
    errors, titles = gate.hierarchy(text, name)
    try:
        gate.verify_local_anchors(path)
    except ValueError as error:
        errors += str(error).splitlines()
    errors += check_no_comments(text, name)
    errors += check_dates(text, name, required=(name == 'B1-Project-Files.md'))
    try:
        rules.strip_delivery_audit(text)
    except ValueError as error:
        errors.append(name + ': ' + str(error))

    if name == 'B0-Entry.md':
        required = ('入口与阅读顺序', '三个输出的关系', '共享与同步', '交付边界')
        missing = [item for item in required if not re.search(r'(?m)^## \d+\. ' + re.escape(item) + r'$', text)]
        if missing:
            errors.append(name + ': 缺少章节 ' + ', '.join(missing))
        return errors

    try:
        _, chapters = rules.split_chapters(text)
        scopes = [rules.chapter_scope(title, name) for title, _ in chapters]
    except ValueError as error:
        return errors + [str(error)]
    if scopes != EXPECTED_SCOPES[name]:
        errors.append(name + ': 一级章节范围应为 ' + ', '.join(EXPECTED_SCOPES[name])
                      + '，实际为 ' + ', '.join(scopes))

    if name == 'A-UserGuide.md':
        products = [re.sub(r'^[一二三四五六七八九十]+\.\s*', '', title) for title, _ in chapters]
        if products != ['Excel-Iterate', 'WordUpdater', 'XyCabin']:
            errors.append(name + ': 一级章节必须依次为三个产品')
        for title, lines in chapters:
            if not chapter_overview(lines):
                errors.append(name + ': 章节缺少导言：' + title)
        try:
            rules.document_overview(text)
        except ValueError as error:
            errors.append(name + ': ' + str(error))
    elif name == 'B1-Project-Files.md':
        try:
            rules.document_overview(text)
        except ValueError as error:
            errors.append(name + ': ' + str(error))
        if len(chapters) != 4:
            errors.append(name + ': 必须包含四个一级章节')
        audits = sum(1 for line in text.split('\n') if AUDIT_HEAD.match(line))
        if audits != 3:
            errors.append(name + ': 「本轮交付证据」必须恰好三处，实际 ' + str(audits))
        for title, lines in chapters[1:]:
            body = '\n'.join(lines)
            found = sum(1 for line in body.split('\n') if AUDIT_HEAD.match(line))
            if found != 1:
                errors.append(name + ': 每个项目档案须有一处交付证据：' + title)
        for title, lines in chapters:
            if rules.chapter_scope(title, name) in ('excel', 'word'):
                errors += version_rows('\n'.join(lines), name + '/' + title)
    elif name == 'B2-DevExperience.md':
        if len(chapters) != 4:
            errors.append(name + ': 必须包含四个一级章节')
    elif name == 'B3-Uphold-Discipline.md':
        if len(chapters) != 2:
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
    print('PASS: C-04 五份文档的章节范围、概览、交付证据、版本表与日期；' + str(len(paths)) + ' 份')


def self_test():
    import tempfile
    gate = markdown_gate()
    rules = document_rules()
    assert gate.hierarchy('# 一. 甲\n## 1. 乙\n', 'ok')[0] == []
    assert version_rows('## 2. 版本摘要\n| 0.9.1.0 | 甲 |\n| 0.9.2.0 | 乙 |\n', 'ok') == []
    assert version_rows('## 2. 版本摘要\n| 0.9.2.0 | 甲 |\n| 0.9.1.0 | 乙 |\n', 'bad')
    assert version_rows('## 2. 版本摘要\n| 0.9.1.0 | ' + '中' * 201 + ' |\n', 'long')
    assert check_no_comments('<!-- x -->', 'bad')
    assert rules.chapter_scope('一. Agent及开发者接手概览', 'B1-Project-Files.md') == 'shared'
    assert rules.chapter_scope('二. Excel-Iterate 项目档案', 'B1-Project-Files.md') == 'excel'
    try:
        rules.chapter_scope('一. 未知章节', 'B1-Project-Files.md')
    except ValueError:
        pass
    else:
        raise AssertionError('unknown chapter accepted')
    try:
        rules.document_overview('# 一. 甲\n\n## 1. 乙\n')
    except ValueError:
        pass
    else:
        raise AssertionError('empty overview accepted')
    with tempfile.TemporaryDirectory() as temporary:
        path = Path(temporary) / 'B2-DevExperience.md'
        path.write_text('# 一. Word与Excel共享开发经验\n## 1. 乙\n# 二. Excel-Iterate 专属\n# 三. WordUpdater 专属\n# 四. XyCabin 专属\n', encoding='utf-8')
        assert verify_file(path) == []
    print('VerifyDevlog self-test: 10 PASS, 0 FAIL', flush=True)


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
