"""Run the shared generate/build/sign/package sequence; never install or stop WPS."""
import os
import argparse
import subprocess
import sys
sys.dont_write_bytecode = True
from project_support import ROOT, OUT, PROJECT, CONFIGURE_NAME, load_site, verify_layout


# rule=R-18
def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--mode', choices=('incremental', 'full'), default='full')
    parser.add_argument('--open-installer-only', action='store_true')
    parser.add_argument('--site-html')
    parser.add_argument('--installer-output')
    args = parser.parse_args()
    if args.open_installer_only:
        os.startfile(str(OUT / (PROJECT['assembly'] + '_Setup.exe')))
        return
    website = load_site()
    verify_layout(source_only=True)
    if args.site_html:
        website.validate_snapshot(args.site_html, ROOT, require_embedded=True)
    environment = dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONUTF8='1')
    for index, script in enumerate(('AA-Program/' + CONFIGURE_NAME, 'build.py', 'resign.py', 'build_installer.py'), 1):
        print('[%d/4] %s' % (index, script), flush=True)
        command = [sys.executable, str(ROOT / script)]
        if script in ('build.py', 'build_installer.py'):
            command += ['--mode', args.mode]
        if script == 'build.py' and args.site_html:
            command += ['--site-html', args.site_html]
        if script == 'build_installer.py' and args.installer_output:
            command += ['--output', args.installer_output]
        subprocess.run(command, cwd=ROOT, env=environment, check=True)
    verify_layout(source_only=bool(args.installer_output))
    print('完成：' + str(args.installer_output or (OUT / (PROJECT['assembly'] + '_Setup.exe'))))


if __name__ == '__main__':
    try:
        main()
    except (OSError, subprocess.CalledProcessError) as exc:
        print('FAIL:', exc)
        sys.exit(1)
