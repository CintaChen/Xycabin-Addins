"""Delegate the document gate to the repository-level verifier."""
from pathlib import Path
import subprocess
import sys

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
TARGET = ROOT.parent / 'verify_devlog.py'


def main():
    if '--self-test' in sys.argv:
        command = [sys.executable, str(TARGET), '--self-test']
    else:
        command = [sys.executable, str(TARGET)] + sys.argv[1:]
    raise SystemExit(subprocess.call(command))


if __name__ == '__main__':
    main()
