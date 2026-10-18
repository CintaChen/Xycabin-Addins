"""Offline document asset, fingerprint and failure regressions; no host or UI."""
from pathlib import Path
import argparse
import hashlib
import os
import struct
import sys
import tempfile
import xml.etree.ElementTree as ET
import zlib

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'src'))
from document_assets import collect, manifest_bytes, normalize_document, markdown_targets, validate_png, validate_svg, referenced_paths, verify_manifest


def png(pixel=b'\x10\x20\x30', width=1, height=1, filter_byte=0):
    def chunk(name, data): return struct.pack('>I', len(data)) + name + data + struct.pack('>I', zlib.crc32(name + data) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(bytes([filter_byte]) + pixel)) + chunk(b'IEND', b'')


def self_test():
    checks = 0
    def check(condition, message):
        nonlocal checks
        if not condition: raise AssertionError(message)
        checks += 1; print('PASS: ' + message)
    def reject(action, message):
        try: action()
        except (ValueError, OSError): check(True, message)
        else: raise AssertionError('Accepted: ' + message)
    with tempfile.TemporaryDirectory(prefix='文档资源-离线·') as directory:
        parent = Path(directory); base = parent / 'C-01-ExcelIterate'
        for folder in ('src/documentation', 'lib', 'icons', 'AP1-程序', 'AP2-美工', 'AP3-说明'):
            (base / folder).mkdir(parents=True)
        readme, guide = base / 'AP3-说明/开发日志.md', base / 'AP3-说明/使用说明.md'
        image, duplicate, unused = [base / 'AP2-美工' / name for name in ('中文 空格(一).png', '副本.png', '未引用.png')]
        raw = png(); image.write_bytes(raw); duplicate.write_bytes(raw); unused.write_bytes(png(b'\x01\x02\x03'))
        svg = base / 'src/documentation/图示.svg'; svg.write_bytes(b'<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><path fill="url( #local )"/></svg>')
        body = '# 一. 文档\n![图](<../AP2-美工/中文 空格(一).png> "标题")\n[同图](../AP2-美工/副本.png)\n[SVG](../src/documentation/图示.svg)\n'
        readme.write_text(body, encoding='utf-8-sig'); guide.write_text('用户概览\n', encoding='utf-8')
        (base / 'AP3-说明/AGENTS.md').write_text('# 一. 入口\n', encoding='utf-8')
        (base / 'project.json').write_text('{}', encoding='utf-8')
        (base / 'AP1-程序/Configure_program.py').write_text('# fixture\n', encoding='utf-8')
        (base / 'AP1-程序/一键重建并打包.bat').write_bytes(b'chcp 65001\r\nset PYTHONUTF8=1\r\nset PYTHONIOENCODING=utf-8\r\n')
        (parent / 'verify_shared.py').write_text('# fixture\n', encoding='utf-8')
        bundle = collect(base); xml = manifest_bytes(bundle)
        check(len(bundle['documents']) == 2 and len(bundle['documents'][0]['assets']) == 3, 'two documents, inline PNG, ordinary PNG and SVG')
        check(len(bundle['assets']) == 2, 'same PNG bytes use one DLL resource with two path aliases')
        check(manifest_bytes(collect(base)) == xml, 'resource manifest ordering is deterministic')
        check(set(referenced_paths(base)) == {image, duplicate, svg}, 'only referenced assets enter dependency closure')
        check(bundle['documents'][0]['content'] == body.replace('\n', os.linesep).encode('utf-8'), 'MD BOM stripped and actual compiler newline bytes preserved')
        old = base / 'build-old.tmp'; old.write_text(readme.read_text(encoding='utf-8-sig'), encoding='utf-8')
        check(old.read_bytes() == normalize_document(readme), 'new normalized bytes equal original write_text compiler input')
        resources = {item['resource']: item['content'] for item in bundle['documents'] + bundle['assets']}
        check(verify_manifest(xml, bundle, resources.__getitem__) == (2, 2), 'XML and all actual resource bytes verify')
        reject(lambda: verify_manifest(xml.replace(b'version="1"', b'version="2"'), bundle, resources.__getitem__), 'wrong XML version rejected')
        reject(lambda: verify_manifest(xml + b'<extra/>', bundle, resources.__getitem__), 'malformed XML rejected')
        reject(lambda: verify_manifest(b'<!DOCTYPE x [<!ENTITY a "x">]>' + xml, bundle, resources.__getitem__), 'manifest DTD rejected')
        corrupt = dict(resources); corrupt[bundle['assets'][0]['resource']] = b'bad'
        reject(lambda: verify_manifest(xml, bundle, corrupt.__getitem__), 'stale embedded image bytes rejected')
        guide.write_text('[图][photo]\n[photo]: <../AP2-美工/中文 空格(一).png>\n', encoding='utf-8')
        check(len(collect(base)['documents'][1]['assets']) == 1, 'reference-style image links collected')
        check(len(collect(base)['assets']) == 2, 'assets shared by both documents deduplicated')
        guide.write_text('用户概览\n', encoding='utf-8')
        for sample in ('```md\n![例](../AP2-美工/不存在.png)\n```\n', '~~~md\n[例](../src/documentation/不存在.svg)\n~~~\n', '````md\n```\n![例](../AP2-美工/不存在.png)\n````\n', '`![例](../AP2-美工/不存在.png)`\n', '``[例](../AP2-美工/不存在.png)``\n', '<!--\n![例](../AP2-美工/不存在.png)\n-->\n', '\\[例](../AP2-美工/不存在.png)\n', '\n    ![例](../AP2-美工/不存在.png)\n'):
            readme.write_text(body + sample, encoding='utf-8')
            check(len(collect(base)['documents'][0]['assets']) == 3, 'code/comment/escaped sample ignored: ' + sample.splitlines()[0])
        readme.write_text(body + '\n<!-- delivery-audit:start -->[旧图](../AP2-美工/不存在.png)<!-- delivery-audit:end -->\n', encoding='utf-8')
        check(len(collect(base)['assets']) == 2 and '不存在'.encode('utf-8') not in normalize_document(readme), 'delivery audit excluded from resources and closure')
        for malformed in ('<!-- delivery-audit:start -->', '<!-- delivery-audit:end --><!-- delivery-audit:start -->', '<!-- delivery-audit:start --><!-- delivery-audit:start --><!-- delivery-audit:end -->'):
            readme.write_text(body + malformed, encoding='utf-8'); reject(lambda: collect(base), 'invalid audit markers fail closed')
        for target in ('../AP2-美工/不存在.png', '../../outside.png', '../src/evil.png', '/absolute.png', 'C:/absolute.png', '../AP2-美工/中文%20空格(一).png?x=1', '../AP2-美工/中文%20空格(一).png#x'):
            readme.write_text('[图](<' + target + '>)', encoding='utf-8'); reject(lambda: collect(base), 'missing/outside/disallowed local asset rejected: ' + target)
        readme.write_text(body + '[remote](https://example.invalid/x.png)\n', encoding='utf-8')
        check(len(collect(base)['assets']) == 2, 'remote PNG is not fetched or packaged')
        readme.write_text('[图](../AP2-美工/中文%20空格%28一%29.png)\n', encoding='utf-8')
        check(collect(base)['documents'][0]['assets'][0]['path'] == 'AP2-美工/中文 空格(一).png', 'percent-encoded Chinese-space destination normalized')
        check(list(markdown_targets('[x](a(b).png "t") [y][id]\n[id]: <z z.svg>\n')) == ['a(b).png', 'z z.svg'], 'balanced parentheses, titles and named references parsed')
        check(validate_png(raw) == (1, 1), 'actual PNG payload decoded and validated')
        for damaged in (b'not png', raw[:-1], raw + b'trailing', raw[:30] + bytes([raw[30] ^ 1]) + raw[31:], png(b'\x01'), png(filter_byte=5), png(width=0)):
            reject(lambda damaged=damaged: validate_png(damaged), 'damaged PNG rejected')
        validate_svg(svg.read_bytes()); check(True, 'static SVG internal fragment URL accepted')
        for invalid in (b'<svg/>', b'<svg', b'<!DOCTYPE svg><svg xmlns="http://www.w3.org/2000/svg"/>', b'<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>', b'<svg xmlns="http://www.w3.org/2000/svg" onload="x"/>', b'<svg xmlns="http://www.w3.org/2000/svg"><image href="https://x"/></svg>', b'<svg xmlns="http://www.w3.org/2000/svg"><style>@import "x";</style></svg>'):
            reject(lambda invalid=invalid: validate_svg(invalid), 'unsafe/invalid SVG rejected')
        readme.write_text(body, encoding='utf-8')
        image.write_bytes(b'bad'); reject(lambda: collect(base), 'invalid source PNG blocks collection'); image.write_bytes(raw)
        svg.write_bytes(b'<svg/>'); reject(lambda: collect(base), 'invalid source SVG blocks collection'); svg.write_bytes(b'<svg xmlns="http://www.w3.org/2000/svg"/>')
        from tools import file_inventory, layout_probe
        files, _ = file_inventory.collect(parent)
        first = file_inventory.source_fingerprint(files)
        check(file_inventory.input_file(image, base) and not file_inventory.input_file(unused, base), 'inventory upgrades referenced PNG only to necessary input')
        image.write_bytes(png(b'\x33\x22\x11'))
        check(file_inventory.source_fingerprint(files) != first, 'PNG-only edit changes cleanup/source proof fingerprint')
        image.write_bytes(raw); before = file_inventory.source_fingerprint(files)
        unused.write_bytes(png(b'\x77\x88\x99'))
        check(file_inventory.source_fingerprint(files) == before, 'unreferenced screenshot bytes do not become build input')
        records, _, _, _ = file_inventory.scan(parent)
        check(next(item for item in records if item['path'] == str(image))['category'] == file_inventory.INPUT, 'read-only inventory classifies document PNG as input')
        import build_cache
        website = parent / 'C-03-Website·Xycabin'; (website / 'src').mkdir(parents=True); (website / 'AP1-程序').mkdir()
        (website / 'AP1-程序/Configure_Site.py').write_text('# fixture', encoding='utf-8')
        snapshot = parent / 'snapshot.html'; snapshot.write_text('<script id="installer-assets" type="application/json">{"packages":{}}</script>', encoding='utf-8')
        old_site, old_environment = build_cache.SITE, build_cache.environment
        try:
            build_cache.SITE, build_cache.environment = website, lambda: 'fixture-environment'
            key = build_cache.inputs(base, snapshot)
            image.write_bytes(png(b'\x99\x88\x77'))
            check(build_cache.inputs(base, snapshot) != key, 'PNG-only edit invalidates whole-package incremental reuse')
            image.write_bytes(raw); key = build_cache.inputs(base, snapshot)
            unused.write_bytes(png(b'\x11\x22\x33'))
            check(build_cache.inputs(base, snapshot) == key, 'unused screenshot does not invalidate whole package')
        finally: build_cache.SITE, build_cache.environment = old_site, old_environment
        layout_probe.verify(parent, source_only=True, folders=(base.name,))
        check(True, 'source-only layout accepts valid document assets')
        image.unlink()
        reject(lambda: layout_probe.verify(parent, source_only=True, folders=(base.name,)), 'source-only layout still blocks missing inline PNG')
    print('DocumentAssetsProbe: %d PASS, 0 FAIL' % checks)
    return checks


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, help='also verify this add-in current source closure')
    parser.add_argument('--source-only', action='store_true', help='skip temporary failure fixtures')
    args = parser.parse_args()
    if not args.source_only: self_test()
    if args.root:
        bundle = collect(args.root)
        for document in bundle['documents']:
            print('PASS: %s %s; %d asset paths' % (document['name'], document['sha256'], len(document['assets'])))
        print('PASS: %d unique PNG/SVG resources; XML SHA256 %s' % (len(bundle['assets']), hashlib.sha256(manifest_bytes(bundle)).hexdigest()))


if __name__ == '__main__':
    try: main()
    except (OSError, ValueError, AssertionError, KeyError) as error:
        print('FAIL:', error); sys.exit(1)
