#!/usr/bin/env python3
"""Verify the separately delivered frozen ZIP, without extracting or changing it.
Usage: python tools/verify_frozen_archive.py /path/to/WebArch_ALPHA_1_0_3_HOME_1_1_CONGELATA.zip
Standard library only. Integrity checks are not a VBA compilation or Excel test.
"""
from pathlib import Path, PurePosixPath
import argparse
import hashlib
import json
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]
RECORD = ROOT / 'releases' / 'alpha-1.0.3-home1.1'

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('archive', type=Path)
    args = parser.parse_args()
    try:
        meta = json.loads((RECORD / 'artifact.json').read_text(encoding='utf-8'))
        if args.archive.stat().st_size != meta['package_bytes']:
            raise ValueError('Archive byte count differs from frozen package')
        raw = args.archive.read_bytes()
        if hashlib.sha256(raw).hexdigest() != meta['package_sha256']:
            raise ValueError('Archive SHA-256 differs from frozen package')
        expected = {}
        for line in (RECORD / 'COMPONENTS.sha256').read_text(encoding='ascii').splitlines():
            if not line.strip():
                continue
            digest, path = line.split('  ', 1)
            parts = PurePosixPath(path)
            if path in expected or parts.is_absolute() or '..' in parts.parts:
                raise ValueError('Invalid component path')
            expected[path] = digest
        if len(expected) != meta['production_component_count']:
            raise ValueError('Component count mismatch')
        prefix = meta['package_root'] + '/'
        with zipfile.ZipFile(args.archive) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ValueError('Duplicate ZIP entries')
            for name in names:
                p = PurePosixPath(name)
                if p.is_absolute() or '..' in p.parts or not name.startswith(prefix):
                    raise ValueError('Unexpected ZIP entry: ' + name)
            bad = archive.testzip()
            if bad:
                raise ValueError('ZIP CRC failure: ' + bad)
            actual = {name[len(prefix):] for name in names
                      if name[len(prefix):].startswith(('02_VBA_CORE_IMPORT_ORDER/', '03_SETUP_MANUALE/'))
                      and name.endswith(('.bas', '.cls'))}
            if actual != set(expected):
                raise ValueError('Production component set mismatch')
            for path, digest in expected.items():
                if hashlib.sha256(archive.read(prefix + path)).hexdigest() != digest:
                    raise ValueError('Component mismatch: ' + path)
            if hashlib.sha256(archive.read(prefix + '08_ALPHA_FREEZE/SOURCE_MANIFEST.json')).hexdigest() != meta['source_manifest_sha256']:
                raise ValueError('Source manifest mismatch')
            if hashlib.sha256(archive.read(prefix + '01_BASE_TEMPLATE/WebArch_PA0_BASE_TEMPLATE.xlsx')).hexdigest() != meta['template_sha256']:
                raise ValueError('Template mismatch')
        print('PASS: frozen ZIP, template, manifest and all 36 production components match.')
        print('No Excel/VBA execution performed. No files changed.')
        return 0
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as error:
        print('FAIL: ' + str(error), file=sys.stderr)
        return 1

if __name__ == '__main__':
    sys.exit(main())
