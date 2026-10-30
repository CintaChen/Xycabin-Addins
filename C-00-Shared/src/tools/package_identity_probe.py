"""Offline regression: deployment filenames must not be used as CLR identities."""
from pathlib import Path
import argparse
import copy
import runpy
import shutil
import sys
import uuid
import xml.etree.ElementTree as ET

sys.dont_write_bytecode = True


def verify(project_root):
    root = Path(project_root).resolve()
    sys.path.insert(0, str(root))
    sys.path.insert(1, str(root.parent / 'C-00-Shared'))
    package = runpy.run_path(str(root / 'verify_package.py'), run_name='_identity_package')
    generator = runpy.run_path(str(root / 'gen_manifest.py'), run_name='_identity_generator')
    payload = package['PAYLOAD'][:3]
    records = package['assembly_metadata']([root / 'build' / name for name in payload])
    metadata = dict(zip(payload, records))
    checker = package['check_payload_identities']
    checks = 0

    def check(value, label):
        nonlocal checks
        if not value:
            raise AssertionError(label)
        checks += 1

    def fixture(data):
        app = ET.Element('{urn:schemas-microsoft-com:asm.v1}assembly')
        for name, record in data.items():
            dep = ET.SubElement(app, '{urn:schemas-microsoft-com:asm.v2}dependency')
            target = ET.SubElement(dep, '{urn:schemas-microsoft-com:asm.v2}dependentAssembly', {'codebase': name})
            identity = ET.fromstring(generator['identity_xml'](record['identity']))
            identity.tag = '{urn:schemas-microsoft-com:asm.v2}assemblyIdentity'
            target.append(identity)
        entry = ET.SubElement(app, '{urn:schemas-microsoft-com:vsta.v3}entryPoint', {'class': 'Example.ThisAddIn'})
        identity = ET.fromstring(generator['identity_xml'](data[payload[0]]['identity']))
        identity.tag = '{urn:schemas-microsoft-com:asm.v2}assemblyIdentity'
        entry.append(identity)
        return app

    def declaration(app, name):
        dep = next(n for n in app.iter() if n.get('codebase') == name)
        return next(iter(dep))

    def reject(app, data, label):
        try:
            checker(app, data)
        except ValueError:
            check(True, label)
        else:
            raise AssertionError('Invalid deployment accepted: ' + label)

    app = fixture(metadata)
    checker(app, metadata)
    check(True, 'actual DLL identities and references accepted')
    core_file = 'Interop.Microsoft.Office.Core.dll'
    core = metadata[core_file]['identity']
    for field, bad in (('name', 'Wrong.Assembly'), ('version', '0.0.0.0'),
                       ('publicKeyToken', '0000000000000000'), ('language', 'other'),
                       ('processorArchitecture', 'x86' if core['processorArchitecture'] != 'x86' else 'amd64')):
        changed = copy.deepcopy(app)
        declaration(changed, core_file).set(field, bad)
        reject(changed, metadata, 'incorrect ' + field + ' rejected despite unchanged DLL bytes')
    changed = copy.deepcopy(app)
    old = declaration(changed, core_file)
    old.attrib.update(name='Interop.Microsoft.Office.Core', version='2.8.0.0')
    old.attrib.pop('publicKeyToken', None)
    if core['name'] != 'Interop.Microsoft.Office.Core' or core['version'] != '2.8.0.0' or core.get('publicKeyToken'):
        reject(changed, metadata, 'reported legacy-Interop manifest / office-PIA payload mismatch reproduced')
    weak = copy.deepcopy(metadata)
    old_name = weak[core_file]['identity']['name']
    weak[core_file]['identity'].update(name='Interop.Microsoft.Office.Core', version='2.8.0.0')
    weak[core_file]['identity'].pop('publicKeyToken', None)
    for record in weak.values():
        for reference in record['references']:
            if reference['name'].casefold() == old_name.casefold():
                reference.update(name='Interop.Microsoft.Office.Core', version='2.8.0.0')
                reference.pop('publicKeyToken', None)
    checker(fixture(weak), weak)
    check('publicKeyToken' not in declaration(fixture(weak), core_file).attrib, 'unsigned interop identity stays unsigned')
    changed = copy.deepcopy(app)
    main = declaration(changed, payload[0])
    main.set('publicKeyToken', '0000000000000000')
    if not metadata[payload[0]]['identity'].get('publicKeyToken'):
        reject(changed, metadata, 'zero token cannot substitute for an unsigned main DLL')
    bad_reference = copy.deepcopy(metadata)
    used = next((r for r in bad_reference[payload[0]]['references'] if r['name'].casefold() == core['name'].casefold()), None)
    if used is not None:
        used['version'] = '0.0.0.0'
        reject(app, bad_reference, 'compiled AssemblyRef version mismatch rejected')
    changed = copy.deepcopy(app)
    entry = next(n for n in changed.iter() if n.get('class'))
    next(iter(entry)).set('version', '0.0.0.0')
    reject(changed, metadata, 'VSTO entry-point identity mismatch rejected')
    scratch = root / 'build' / ('manifest-fixture-' + uuid.uuid4().hex)
    scratch.mkdir()
    try:
        for name in payload:
            shutil.copyfile(root / 'build' / name, scratch / name)
        main = generator['main']
        old_build, old_args = main.__globals__['BUILD'], sys.argv
        try:
            main.__globals__['BUILD'] = str(scratch)
            sys.argv = ['gen_manifest.py']
            main()
        finally:
            main.__globals__['BUILD'], sys.argv = old_build, old_args
        checker(ET.parse(scratch / (payload[0] + '.manifest')).getroot(), metadata)
        check((scratch / package['PAYLOAD'][-1]).is_file(), 'real generation entry produces both manifests with actual identities')
    except Exception:
        print('GENERATOR_FIXTURE_RETAINED: ' + str(scratch))
        raise
    else:
        shutil.rmtree(scratch)
    print('PackageIdentityProbe: %d PASS; actual PE metadata, no installation or WPS' % checks)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--project-root', required=True, type=Path)
    try:
        verify(parser.parse_args().project_root)
    except (OSError, ValueError, AssertionError) as error:
        print('FAIL: ' + str(error))
        sys.exit(1)
