"""Offline document composition, asset and failure regressions; no host or UI."""
from pathlib import Path
import argparse
import hashlib
import os
import struct
import sys
import tempfile
import zlib

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'src'))
from document_assets import (collect, manifest_bytes, markdown_targets, normalize_text,
                             project_documents, validate_png, validate_svg,
                             referenced_paths, verify_manifest)


def png(pixel=b'\x10\x20\x30', width=1, height=1, filter_byte=0):
    def chunk(name, data):
        return struct.pack('>I', len(data)) + name + data + struct.pack('>I', zlib.crc32(name + data) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(bytes([filter_byte]) + pixel)) + chunk(b'IEND', b''))


GUIDE = '''**使用说明**

<!-- agent-overview:start -->

三个产品共用一份使用说明。

<!-- agent-overview:end -->

<!-- scope:word:start -->
# 二. WordUpdater
Word 专属正文
<!-- scope:word:end -->

<!-- scope:excel:start -->
# 一. Excel-Iterate
![图](<../C-01-ExcelIterate/AB-Graphic/中文 空格(一).png> "标题")
[矢量](<../C-01-ExcelIterate/AB-Graphic/图示.svg>)
<!-- scope:excel:end -->
'''

B1 = '''**项目档案**

最后更新：2026-10-18 ｜ 当前版本：**0.9.9.1**

<!-- scope:shared:start -->
# 一. 接手概览

<!-- agent-overview:start -->

共享概览。

<!-- agent-overview:end -->
<!-- scope:shared:end -->

<!-- scope:excel:start -->
# 二. Excel-Iterate 项目档案
共享概览之外的项目档案。
<!-- scope:excel:end -->

<!-- scope:word:start -->
# 三. WordUpdater 项目档案
Word 项目档案。
<!-- scope:word:end -->
'''

B2 = '''**开发经验**

<!-- scope:shared:start -->
# 一. Word与Excel共享开发经验
![图](../C-01-ExcelIterate/AB-Graphic/副本.png)
<!-- scope:shared:end -->

<!-- scope:excel:start -->
# 二. Excel-Iterate 专属开发经验
Excel 专属正文
<!-- scope:excel:end -->

<!-- scope:word:start -->
# 三. WordUpdater 专属开发经验
Word 专属正文
<!-- scope:word:end -->
'''

B3 = '''**维护纪律**

<!-- scope:shared:start -->
# 一. Word与Excel共享维护纪律
共享纪律正文
<!-- scope:shared:end -->
'''


def fixture(parent):
    base = parent / 'C-01-ExcelIterate'
    for folder in ('lib', 'src', 'AA-Program', 'AB-Graphic'):
        (base / folder).mkdir(parents=True)
    image, duplicate, unused = [base / 'AB-Graphic' / name for name in ('中文 空格(一).png', '副本.png', '未引用.png')]
    raw = png(); image.write_bytes(raw); duplicate.write_bytes(raw); unused.write_bytes(png(b'\x01\x02\x03'))
    svg = base / 'AB-Graphic/图示.svg'
    svg.write_bytes(b'<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><path fill="url( #local )"/></svg>')
    docs = parent / 'C-04-Devlog'; docs.mkdir()
    (docs / 'A-UserGuide.md').write_text(GUIDE, encoding='utf-8-sig')
    (docs / 'B1-Project-Files.md').write_text(B1, encoding='utf-8')
    (docs / 'B2-DevExperience.md').write_text(B2, encoding='utf-8')
    (docs / 'B3-Uphold-Discipline.md').write_text(B3, encoding='utf-8')
    (docs / 'B0-Entry.md').write_text('# 一. 入口\n', encoding='utf-8')
    (base / 'AA-Program/C1-Exceliterate_Configure.py').write_text('# fixture\n', encoding='utf-8')
    (base / 'AA-Program/一键重建并打包.bat').write_bytes(b'chcp 65001\r\nset PYTHONUTF8=1\r\nset PYTHONIOENCODING=utf-8\r\n')
    (base / 'AA-Program/Excel_Iterate_Setup.exe').write_text('', encoding='utf-8')
    return base, docs, image, duplicate, unused, svg, raw


def self_test():
    checks = 0

    def check(condition, message):
        nonlocal checks
        if not condition:
            raise AssertionError(message)
        checks += 1
        print('PASS: ' + message)

    def reject(action, message):
        try:
            action()
        except (ValueError, OSError):
            check(True, message)
        else:
            raise AssertionError('Accepted: ' + message)

    with tempfile.TemporaryDirectory(prefix='文档资源-离线·') as directory:
        parent = Path(directory)
        base, docs, image, duplicate, unused, svg, raw = fixture(parent)
        bundle = collect(base)
        xml = manifest_bytes(bundle)
        documents = {item['name']: item for item in bundle['documents']}
        check(len(bundle['documents']) == 2, 'two composed documents')
        check(documents['UserGuide']['path'] == 'C-04-Devlog/A-UserGuide.md', 'guide keeps the C-04 logical path')
        check(documents['Readme']['path'] == 'C-04-Devlog/B-Devlog.md', 'developer log is a composed virtual file')
        check('Excel-Iterate' in documents['UserGuide']['content'].decode() and 'Word 专属正文' not in documents['UserGuide']['content'].decode(), 'guide keeps only the current project chapter')
        check('Excel 专属正文' in documents['Readme']['content'].decode() and 'Word 专属正文' not in documents['Readme']['content'].decode(), 'developer log keeps shared and current project chapters')
        check(len(documents['UserGuide']['assets']) == 2, 'guide collects its PNG and SVG')
        check(len(documents['Readme']['assets']) == 1, 'developer log collects only its own PNG')
        check(len(bundle['assets']) == 2, 'identical PNG bytes share one DLL resource')
        check(manifest_bytes(collect(base)) == xml, 'resource manifest ordering is deterministic')
        check(set(referenced_paths(base)) == {image, duplicate, svg}, 'only referenced assets enter the dependency closure')
        check(documents['UserGuide']['content'].endswith(os.linesep.encode()) or b'\r\n' in documents['UserGuide']['content'], 'compiler newline bytes preserved')
        resources = {item['resource']: item['content'] for item in bundle['documents'] + bundle['assets']}
        check(verify_manifest(xml, bundle, resources.__getitem__) == (2, 2), 'XML and actual resource bytes verify')
        reject(lambda: verify_manifest(xml.replace(b'version="1"', b'version="2"'), bundle, resources.__getitem__), 'wrong XML version rejected')
        reject(lambda: verify_manifest(xml + b'<extra/>', bundle, resources.__getitem__), 'malformed XML rejected')
        reject(lambda: verify_manifest(b'<!DOCTYPE x [<!ENTITY a "x">]>' + xml, bundle, resources.__getitem__), 'manifest DTD rejected')
        corrupt = dict(resources); corrupt[bundle['assets'][0]['resource']] = b'bad'
        reject(lambda: verify_manifest(xml, bundle, corrupt.__getitem__), 'stale embedded image bytes rejected')

        guide = docs / 'A-UserGuide.md'
        body = guide.read_text(encoding='utf-8-sig')
        guide.write_text(body + '\n<!-- delivery-audit:start -->[旧图](../C-01-ExcelIterate/AB-Graphic/不存在.png)<!-- delivery-audit:end -->\n', encoding='utf-8')
        check(len(collect(base)['assets']) == 2 and '不存在'.encode() not in normalize_text(guide.read_text(encoding='utf-8')).encode(), 'delivery audit excluded from resources and closure')
        for malformed in ('<!-- delivery-audit:start -->', '<!-- delivery-audit:end --><!-- delivery-audit:start -->',
                          '<!-- delivery-audit:start --><!-- delivery-audit:start --><!-- delivery-audit:end -->'):
            guide.write_text(body + malformed, encoding='utf-8'); reject(lambda: collect(base), 'invalid audit markers fail closed')
        guide.write_text(body, encoding='utf-8')
        for target in ('../C-01-ExcelIterate/AB-Graphic/不存在.png', '../../outside.png',
                       '/absolute.png', 'C:/absolute.png',
                       '../C-01-ExcelIterate/AB-Graphic/中文%20空格(一).png?x=1',
                       '../C-01-ExcelIterate/AB-Graphic/中文%20空格(一).png#x',
                       '../C-01-ExcelIterate/src/evil.png'):
            guide.write_text('[图](<' + target + '>)\n', encoding='utf-8'); reject(lambda: collect(base), 'missing/outside/disallowed asset rejected: ' + target)
        guide.write_text(body + '[remote](https://example.invalid/x.png)\n', encoding='utf-8')
        check(len(collect(base)['assets']) == 2, 'remote PNG is not fetched or packaged')
        guide.write_text('[图](../C-01-ExcelIterate/AB-Graphic/中文%20空格%28一%29.png)\n', encoding='utf-8')
        check(collect(base)['documents'][0]['assets'][0]['path'] == 'C-01-ExcelIterate/AB-Graphic/中文 空格(一).png', 'percent-encoded Chinese-space destination normalized')
        check(list(markdown_targets('[x](a(b).png "t") [y][id]\n[id]: <z z.svg>\n')) == ['a(b).png', 'z z.svg'], 'balanced parentheses, titles and named references parsed')
        guide.write_text(body, encoding='utf-8')

        check(validate_png(raw) == (1, 1), 'actual PNG payload decoded and validated')
        for damaged in (b'not png', raw[:-1], raw + b'trailing', raw[:30] + bytes([raw[30] ^ 1]) + raw[31:], png(b'\x01'), png(filter_byte=5), png(width=0)):
            reject(lambda damaged=damaged: validate_png(damaged), 'damaged PNG rejected')
        validate_svg(svg.read_bytes()); check(True, 'static SVG internal fragment URL accepted')
        for invalid in (b'<svg/>', b'<svg', b'<!DOCTYPE svg><svg xmlns="http://www.w3.org/2000/svg"/>',
                        b'<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>',
                        b'<svg xmlns="http://www.w3.org/2000/svg" onload="x"/>',
                        b'<svg xmlns="http://www.w3.org/2000/svg"><image href="https://x"/></svg>',
                        b'<svg xmlns="http://www.w3.org/2000/svg"><style>@import "x";</style></svg>'):
            reject(lambda invalid=invalid: validate_svg(invalid), 'unsafe/invalid SVG rejected')

        image.write_bytes(b'bad'); reject(lambda: collect(base), 'invalid source PNG blocks collection'); image.write_bytes(raw)
        svg.write_bytes(b'<svg/>'); reject(lambda: collect(base), 'invalid source SVG blocks collection'); svg.write_bytes(b'<svg xmlns="http://www.w3.org/2000/svg"/>')
        docs.joinpath('B2-DevExperience.md').unlink(); reject(lambda: collect(base), 'missing document blocks collection')

    print('DocumentAssetsProbe: %d PASS, 0 FAIL' % checks)
    return checks


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, help='also verify this add-in current source closure')
    parser.add_argument('--source-only', action='store_true', help='skip temporary failure fixtures')
    args = parser.parse_args()
    if not args.source_only:
        self_test()
    if args.root:
        bundle = collect(args.root)
        for document in bundle['documents']:
            print('PASS: %s %s; %d asset paths' % (document['name'], document['sha256'], len(document['assets'])))
        print('PASS: %d unique PNG/SVG resources; XML SHA256 %s' % (len(bundle['assets']), hashlib.sha256(manifest_bytes(bundle)).hexdigest()))


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, AssertionError, KeyError) as error:
        print('FAIL:', error); sys.exit(1)
