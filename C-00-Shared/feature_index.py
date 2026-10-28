"""Feature source ownership and portable Markdown navigation, derived from configuration."""
from pathlib import Path
import argparse
import copy
import importlib.util
import re
import sys
from unittest.mock import patch

from source_layout import addin_sources, project_data, resolve_source

sys.dont_write_bytecode = True
PRODUCTS = (('C-01-ExcelIterate', 'C1-Exceliterate_Configure.py', 'Excel-Iterate'),
            ('C-02-WordUpdater', 'C2-Wordupdater_Configure.py', 'WordUpdater'))
INDEX_NAME = 'B4-Code Index.md'
FEATURE_ROOT = 'AC-Original Code'


def configuration(root):
    name = next(name for directory, name, _ in PRODUCTS if directory == root.name)
    spec = importlib.util.spec_from_file_location('_feature_index_' + root.name, root / 'AA-Program' / name)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def commands(root):
    """Read the default UI, including hidden tabs and disabled placeholders."""
    config = configuration(root)
    tabs = [(config.TAB_LABEL, config.BUTTONS)]
    tabs += [(tab['label'], tab['buttons']) for tab in config.EXTRA_TABS]
    result = []
    def visit(items, labels, ids):
        for item in items:
            branch, identifiers = labels + [item['label']], ids + [item['id']]
            if item.get('children'):
                visit(item['children'], branch, identifiers)
            else:
                result.append(dict(item, labels=branch, identifiers=identifiers))
    for tab, buttons in tabs:
        for button in buttons:
            visit([button], [tab, button['group'][1]], [tab, button['group'][0]])
    return tabs, result


def code(text):
    pattern = r'//[^\n]*|/\*[\s\S]*?\*/|@"(?:""|[^"])*"|"(?:\\.|[^"\\])*"|\'(?:\\.|[^\'\\])*\''
    return re.sub(pattern, lambda m: ''.join('\n' if c == '\n' else ' ' for c in m[0]), text)


def callback(command):
    return command.get('callback') or 'OnOpenUrl_' + command['id']


def verify_product(root, project=None):
    root = Path(root).resolve()
    project = project if project is not None else project_data(root)
    _, leaves = commands(root)
    expected = {leaf['id'] for leaf in leaves}
    features = project.get('features', {})
    if set(features) != expected:
        raise ValueError(root.name + ': 功能映射与菜单叶子不一致：' + str(sorted(set(features) ^ expected)))
    paths = addin_sources(root, project)
    if len(paths) != len(set(paths)):
        raise ValueError(root.name + ': 编译源码重复')
    compiled = set(project['sources'])
    for old, references in project.get('source_aliases', {}).items():
        if not references or len(references) != len(set(references)) or any(ref not in compiled for ref in references):
            raise ValueError(root.name + ': 拆分源码读取别名无效：' + old)
    root_directory = root / FEATURE_ROOT
    if not root_directory.is_dir() or root_directory.is_symlink():
        raise ValueError('缺少功能源码目录：' + str(root_directory))
    actual = set()
    for file in root_directory.rglob('*'):
        if file.is_symlink() or not file.resolve().is_relative_to(root_directory.resolve()):
            raise ValueError('功能目录不得包含越界路径或链接：' + str(file))
        if file.is_file():
            if file.suffix != '.cs':
                raise ValueError('功能目录只放功能 C# 源码：' + str(file))
            actual.add(file.relative_to(root).as_posix())
        if any(not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.-]*', part)
               for part in file.relative_to(root_directory).parts):
            raise ValueError('功能路径必须采用稳定英文名称：' + str(file))
        if len(str(file)) >= 260:
            raise ValueError('功能路径超过当前构建链兼容长度：' + str(file))
    declared = {ref for ref in compiled if ref.startswith(FEATURE_ROOT + '/')}
    if actual != declared:
        raise ValueError(root.name + ': 功能目录与实际编译清单不一致：' + str(sorted(actual ^ declared)))
    texts = {ref: code(resolve_source(root, ref).read_text(encoding='utf-8-sig')) for ref in project['sources']}
    combined = '\n'.join(texts.values())
    seen_directories, seen_entries, hierarchy = set(), set(), {}
    for leaf in leaves:
        bid = leaf['id']
        feature = features[bid]
        directory, entry = feature['directory'], feature['entry']
        parts = Path(directory).parts
        if not directory.startswith(FEATURE_ROOT + '/') or len(parts) != len(leaf['labels']) + 1:
            raise ValueError('功能目录层级与菜单不一致：' + bid)
        if directory in seen_directories or entry in seen_entries:
            raise ValueError('多个末级功能共用目录或入口：' + bid)
        seen_directories.add(directory)
        seen_entries.add(entry)
        for depth in range(1, len(leaf['identifiers'])):
            key = tuple(leaf['identifiers'][:depth])
            prefix = '/'.join(parts[:depth + 1])
            if key in hierarchy and hierarchy[key] != prefix:
                raise ValueError('同一菜单节点映射到多个目录：' + bid)
            hierarchy[key] = prefix
        if entry not in compiled or Path(entry).parent.as_posix() != directory:
            raise ValueError('功能入口必须位于其末级目录并参与编译：' + bid)
        name = callback(leaf)
        signature = r'\bpublic\s+void\s+' + re.escape(name) + r'\s*\('
        if len(re.findall(signature, texts[entry])) != 1 or len(re.findall(signature, combined)) != 1:
            raise ValueError('按钮回调必须唯一且存在于实际入口：' + bid + '/' + name)
        image = r'\bpublic\s+IPictureDisp\s+GetImage_' + re.escape(bid) + r'\s*\('
        if len(re.findall(image, texts[entry])) != 1 or len(re.findall(image, combined)) != 1:
            raise ValueError('按钮图标回调必须唯一且存在于实际入口：' + bid)
        for reference in feature.get('implementation', []):
            if reference not in compiled:
                raise ValueError('功能实现依赖必须参与当前产品编译：' + bid + '/' + reference)
            resolve_source(root, reference)
        tool_sources = {tool['source'] for tool in project['tools']}
        for reference in feature.get('tests', []):
            if reference not in tool_sources:
                raise ValueError('功能验证入口必须来自工具清单：' + bid + '/' + reference)
            resolve_source(root, reference)
    return leaves


def literal(value):
    return chr(96) + value + chr(96)


def link(root, reference, label=None):
    source = resolve_source(root, reference)
    relative = source.relative_to(root.parent).as_posix()
    return '[' + (label or source.name) + '](<../' + relative + '>)'


def tree(root, leaves, features, tab):
    labels, branches = {}, set()
    for leaf in leaves:
        if leaf['labels'][0] != tab:
            continue
        parts = Path(features[leaf['id']]['directory']).parts[1:]
        for depth, label in enumerate(leaf['labels'], 1):
            key = parts[:depth]
            branches.add(key)
            labels[key] = label
    tab_name = next(key[0] for key in branches)
    for source in project_data(root)['sources']:
        if source.startswith(FEATURE_ROOT + '/' + tab_name + '/') and '/_Shared/' in source:
            parts = Path(source).parent.parts[1:]
            for depth in range(1, len(parts) + 1):
                branches.add(parts[:depth])
    children = {}
    for key in sorted(branches):
        children.setdefault(key[:-1], []).append(key)
    output = [FEATURE_ROOT + '/']
    def emit(parent, prefix):
        entries = children.get(parent, [])
        for position, key in enumerate(entries):
            last = position == len(entries) - 1
            label = labels.get(key, '同项目共用业务代码' if key[-1] == '_Shared' else '')
            output.append(prefix + ('└─ ' if last else '├─ ') + key[-1] + '/' + ('（' + label + '）' if label else ''))
            emit(key, prefix + ('   ' if last else '│  '))
    emit((), '')
    return '\n'.join(output)


def render(repository):
    output = ['**功能代码索引**', '',
              '本文件按两个加载项的默认菜单结构定位功能源码。先在树中找到中文功能，再打开表中的英文目录、按钮入口和核心实现；文件内的方法名可直接搜索。', '',
              '路径以仓库根目录为基准，链接相对于本文件，可随整个仓库移动。界面自定名称、排序或隐藏状态不改变源码目录，稳定控件 ID 用于核对身份。', '',
              '按钮入口负责接线；专属算法和弹窗位于对应末级目录，产品内部共用业务代码位于就近的 _Shared，跨产品实现只在 C-00-Shared 维护。'
              '加载框架、生成文件、测试、安装器、依赖及构建产物分别保留在 src、lib、build 等工程位置。复现需同时携带表列依赖及工程底座，不将末级目录视为独立项目。', '',
              '内容由配置器菜单与各产品 project.json 的 features 映射生成。更新命令：' + literal('python C-00-Shared/feature_index.py --write') + '；'
              '只读检查：' + literal('python C-00-Shared/feature_index.py') + '；验证入口另记录于各功能的 tests。', '']
    for number, (directory, _, title) in enumerate(PRODUCTS, 1):
        root = Path(repository) / directory
        if not root.is_dir():
            continue
        leaves = verify_product(root)
        project = project_data(root)
        features = project['features']
        output += ['# ' + ('一' if number == 1 else '二') + '. ' + title + ' 功能代码索引', '',
                   '项目根目录：' + literal(directory + '/') + '。当前收录 ' + str(len(leaves)) + ' 个末级入口，含禁用占位项。', '',
                   '工程底座：' + '、'.join(link(root, ref) for ref in
                     ('src/addin/ThisAddIn.cs', 'src/addin/Connect.cs', '@shared/src/addin/UiKit.cs', '@shared/src/addin/HostDialogService.cs')) + '。', '']
        tabs = list(dict.fromkeys(leaf['labels'][0] for leaf in leaves))
        for index, tab in enumerate(tabs, 1):
            output += ['## ' + str(index) + '. ' + tab + (' 主选项卡' if index == 1 else ' 选项卡'), '',
                       '### (1). 功能目录树', '', '~~~text', tree(root, leaves, features, tab), '~~~', '',
                       '### (2). 末级功能源码定位表', '',
                       '| 中文功能路径 | 稳定控件 ID | 功能目录（相对仓库根） | 入口文件与回调 | 核心实现与共用依赖（阅读顺序） |',
                       '| --- | --- | --- | --- | --- |']
            for leaf in leaves:
                if leaf['labels'][0] != tab:
                    continue
                bid = leaf['id']
                feature = features[bid]
                label = ' → '.join(leaf['labels'][1:])
                if leaf.get('enabled') is False:
                    label += '（禁用占位，尚未开放）'
                directory_path = directory + '/' + feature['directory']
                cell = '[' + directory_path + '/](<../' + directory_path + '/>)'
                entry = link(root, feature['entry']) + '；' + literal(callback(leaf) + '()')
                dependencies = ' → '.join(link(root, ref, Path(ref).name + ('（共享）' if ref.startswith('@shared/') else ''))
                                          for ref in feature.get('implementation', [])) or '逻辑直接位于入口文件。'
                output.append('| ' + ' | '.join((label, literal(bid), cell, entry, dependencies)) + ' |')
            output.append('')
    return '\n'.join(output)


def check(repository):
    repository = Path(repository)
    expected = render(repository)
    target = repository / 'C-04-Devlog' / INDEX_NAME
    if not target.is_file() or target.read_text(encoding='utf-8-sig') != expected:
        raise ValueError('B4 代码索引与实际功能映射不一致，请运行 feature_index.py --write')
    print('PASS: B4 feature index, physical source ownership, unique callbacks and implementation links')


def self_test(repository):
    count = 0
    for directory, _, _ in PRODUCTS:
        root = Path(repository) / directory
        if not root.is_dir():
            continue
        project = project_data(root)
        verify_product(root, project)
        count += 1
        bid = next(iter(project['features']))
        invalid = []
        missing = copy.deepcopy(project); del missing['features'][bid]; invalid.append(missing)
        wrong = copy.deepcopy(project); wrong['features'][bid]['entry'] = 'src/addin/Connect.cs'; invalid.append(wrong)
        missing_dependency = copy.deepcopy(project); missing_dependency['features'][bid]['implementation'].append('missing.cs'); invalid.append(missing_dependency)
        duplicate = copy.deepcopy(project); duplicate['sources'].append(duplicate['sources'][0]); invalid.append(duplicate)
        alias = copy.deepcopy(project); alias['source_aliases']['bad.cs'] = ['missing.cs']; invalid.append(alias)
        depth = copy.deepcopy(project); depth['features'][bid]['directory'] += '/Unexpected'; invalid.append(depth)
        for candidate in invalid:
            try:
                verify_product(root, candidate)
            except (ValueError, OSError):
                count += 1
            else:
                raise AssertionError('Invalid feature source mapping accepted')
    check(repository)
    target = Path(repository) / 'C-04-Devlog' / INDEX_NAME
    original_read = Path.read_text
    def stale_index(path, *args, **kwargs):
        text = original_read(path, *args, **kwargs)
        return text + '\n索引已过期。\n' if path == target else text
    with patch.object(Path, 'read_text', stale_index):
        try:
            check(repository)
        except ValueError:
            count += 1
        else:
            raise AssertionError('Stale B4 content accepted')
    print('FeatureIndex self-test: %d PASS, 0 FAIL' % count)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--write', action='store_true', help='Regenerate B4 explicitly; default is read-only')
    parser.add_argument('--self-test', action='store_true')
    args = parser.parse_args()
    root = Path(__file__).resolve().parent.parent
    if args.write:
        (root / 'C-04-Devlog' / INDEX_NAME).write_text(render(root), encoding='utf-8', newline='\n')
    elif args.self_test:
        self_test(root)
    else:
        check(root)


if __name__ == '__main__':
    main()
