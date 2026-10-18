"""Shared read-only numbering gate for authored Markdown, excluding fenced code."""
import re
from urllib.parse import unquote, urlsplit
from pathlib import Path
from tempfile import TemporaryDirectory


def chinese(n):
    if n < 10:
        return '一二三四五六七八九'[n - 1]
    if n < 20:
        return '十' + (chinese(n - 10) if n > 10 else '')
    return chinese(n // 10) + '十' + (chinese(n % 10) if n % 10 else '')


def roman(n):
    text = ''
    for value, token in ((1000,'m'),(900,'cm'),(500,'d'),(400,'cd'),(100,'c'),(90,'xc'),(50,'l'),(40,'xl'),(10,'x'),(9,'ix'),(5,'v'),(4,'iv'),(1,'i')):
        while n >= value:
            text += token; n -= value
    return text


def hierarchy(text, label):
    errors, titles, counts, active = [], [], [0] * 6, [False] * 6
    fence = None
    for line in text.splitlines():
        code = re.match(r'^\s{0,3}(`{3,}|~{3,})(.*)$', line)
        if code:
            token, rest = code.groups()
            if fence is None:
                fence = token
            elif token[0] == fence[0] and len(token) >= len(fence) and not rest.strip():
                fence = None
            continue
        if fence is not None:
            continue
        match = re.match(r'^(#+)\s+(.+)$', line)
        if not match:
            continue
        level, title = len(match[1]), match[2]
        if level > 5:
            errors.append(label + ': 最多使用五级标题：' + title); continue
        if level > 1 and not active[level - 1]:
            errors.append(label + ': 标题不得跳级：' + title)
        counts[level] += 1; active[level] = True
        for child in range(level + 1, 6):
            counts[child] = 0; active[child] = False
        n = counts[level]
        prefixes = {1: chinese(n) + '. ', 2: str(n) + '. ', 3: '(' + str(n) + '). ', 4: chr(96 + n) + '. ', 5: roman(n) + '. '}
        remaining = title[len(prefixes[level]):]
        if not title.startswith(prefixes[level]) or not remaining.strip() or remaining[0].isspace():
            errors.append(label + ': 标题编号或格式错误：' + title)
        if level == 1:
            titles.append(title)
    if fence is not None:
        errors.append(label + ': 代码块未闭合')
    if not titles:
        errors.append(label + ': 缺少编号一级标题')
    return errors, titles


def authored_lines(text):
    """Omit fenced examples with the same fence rules as the numbering gate."""
    lines, fence = [], None
    for line in text.splitlines():
        code = re.match(r'^\s{0,3}(`{3,}|~{3,})(.*)$', line)
        if code:
            token, rest = code.groups()
            if fence is None:
                fence = token
            elif token[0] == fence[0] and len(token) >= len(fence) and not rest.strip():
                fence = None
            continue
        if fence is None:
            lines.append(line)
    return lines


def heading_anchors(text):
    """GitHub-style authored heading IDs, including duplicate heading suffixes."""
    anchors, seen = set(), {}
    for line in authored_lines(text):
        match = re.match(r'^#{1,5}\s+(.+)$', line)
        if not match:
            continue
        title = re.sub(r'!?\[([^\]]*)\]\([^)]*\)', r'\1', match[1])
        title = re.sub(r'<[^>]*>', '', title)
        slug = re.sub(r'[^\w\s-]', '', title.lower()).replace(' ', '-')
        count = seen.get(slug, 0)
        seen[slug] = count + 1
        anchors.add(slug if count == 0 else slug + '-' + str(count))
    return anchors


def verify_local_anchors(path):
    """Check only local .md fragments; file-only artifact links use layout_probe."""
    text = path.read_text(encoding='utf-8-sig')
    text = '\n'.join(authored_lines(text))
    errors = []
    for target in re.findall(r'!?\[[^\]]*\]\(([^)\s]+)\)', text):
        parts = urlsplit(target.replace('\\_', '_'))
        if parts.scheme or parts.netloc or not parts.fragment:
            continue
        linked = path if not parts.path else path.parent / unquote(parts.path)
        if linked.suffix.lower() != '.md':
            continue
        if not linked.is_file():
            errors.append(str(path) + ': 章节链接文件缺失：' + target)
        elif unquote(parts.fragment) not in heading_anchors(linked.read_text(encoding='utf-8-sig')):
            errors.append(str(path) + ': 标题锚点不存在：' + target)
    if errors:
        raise ValueError('\n'.join(errors))


def verify_guide(path):
    text = path.read_text(encoding='utf-8-sig')
    verify_local_anchors(path)
    errors, titles = hierarchy(text, str(path))
    for marker in ('agent-overview',):
        a,b = '<!-- '+marker+':start -->', '<!-- '+marker+':end -->'
        if text.count(a) != 1 or text.count(b) != 1 or text.index(a) >= text.index(b):
            errors.append(str(path) + ': 概览标记缺失、重复或乱序')
    if errors:
        raise ValueError('\n'.join(errors))
    return titles


def verify_entry(path):
    """The AI entry uses numbered headings, without embedded-document markers."""
    verify_local_anchors(path)
    errors, titles = hierarchy(path.read_text(encoding='utf-8-sig'), str(path))
    if errors:
        raise ValueError('\n'.join(errors))
    return titles


def self_test():
    anchor_self_test()
    good = '# 一. 概览\n## 1. 内容\n### (1). 步骤\n#### a. 条件\n##### i. 细节\n##### ii. 细节\n##### iii. 细节\n#### b. 条件\n##### i. 子级重新开始\n### (2). 步骤\n#### a. 子级重新开始\n## 2. 内容\n### (1). 子级重新开始\n# 二. 档案\n## 1. 内容\n~~~md\n# 无需编号\n~~~~\n```python\n# 注释保持原样\n```\n'
    assert not hierarchy(good, 'valid')[0]
    for old,new in [('##### ii.', '##### iv.'), ('#### b.', '#### d.'), ('# 二.', '# 三.'), ('### (2).', '### (3).')]:
        assert hierarchy(good.replace(old,new),'gap')[0]
    for old, new in [('# 一. ', '# 一、'), ('### (1). ', '### （1）'),
                     ('### (1). ', '### (1) '), ('# 一. ', '# 一.'),
                     ('## 1. ', '## 1.  '), ('#### a. ', '#### A. '),
                     ('##### i. ', '##### I. ')]:
        assert hierarchy(good.replace(old, new), 'old-or-invalid-format')[0]
    assert hierarchy('# 一. 概览\n### (1). 跳级','skip')[0]
    assert hierarchy(good + '\n###### 1. 超过五级', 'six-levels')[0]
    assert hierarchy(good + '\n####### 1. 超过五级', 'seven-levels')[0]
    assert hierarchy('# 一. 概览\n```\n','unclosed')[0]
    assert hierarchy('# 一. ', 'empty-title')[0]
    assert hierarchy('# 开发接手入口', 'unnumbered-entry')[0]
    assert not hierarchy('# 一. 开发接手入口', 'numbered-entry')[0]
    assert heading_anchors('# 一. Word与Excel\n## 1. 标题、编号\n## 1. 标题、编号\n```md\n# 假标题\n```') == {'一-word与excel', '1-标题编号', '1-标题编号-1'}
    assert '2-带代码-text' in heading_anchors('## 2. 带代码 `Text`')


def anchor_self_test():
    with TemporaryDirectory(prefix='md-anchor-probe-') as folder:
        base = Path(folder)
        doc = base / '说明.md'
        peer = base / '姊妹.md'
        peer.write_text('# 一. 项目档案\n', encoding='utf-8')
        valid = '# 一. 项目档案\n[自身](#一-项目档案) [姊妹](姊妹.md#一-项目档案) [编码](%E5%A7%8A%E5%A6%B9.md#一-项目档案)\n````md\n[示例](missing.md#不存在)\n````\n~~~\n[例](missing.md#不存在)\n~~~\n'
        doc.write_text(valid, encoding='utf-8')
        verify_local_anchors(doc)
        for target in ('#不存在', '姊妹.md#不存在', 'missing.md#不存在'):
            doc.write_text(valid + '[坏链接](' + target + ')\n', encoding='utf-8')
            try:
                verify_local_anchors(doc)
            except ValueError:
                pass
            else:
                raise AssertionError('Accepted broken anchor: ' + target)
