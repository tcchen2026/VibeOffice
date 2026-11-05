#!/usr/bin/env python3
"""Isolate syntax in the C2 hand-built precedence diagnostic, not the corpus."""
import hashlib
import json
import pathlib
import re
import sys
import zipfile

base = pathlib.Path('tmp/legacy/artifacts/C2-parens-office')
names = '--names' in sys.argv
out = base / ('name-controls' if names else 'bisect')
out.mkdir(exist_ok=True)
cases = [
    ('values-only', 'after', set(), False),
    ('arithmetic-only', 'after', set(range(1, 14)), False),
    ('references-only', 'after', {14, 15, 16, 18, 19, 20}, False),
    ('no-name-after', 'after', set(range(1, 21)) - {17}, False),
    ('name-only', 'after', set(), True),
    ('no-name-before', 'before', set(range(1, 21)) - {17}, False),
    ('union-after', 'after', {20}, False),
    ('union-before', 'before', {20}, False),
]
pairs = []
if names:
    cases = [
        ('rename-only', 'TestAreas', 'E1,E2'),
        ('qualified-union', 'Areas', "'Precedence'!E1,'Precedence'!E2"),
        ('qualified-absolute', 'Areas', "'Precedence'!$E$1,'Precedence'!$E$2"),
        ('unqualified-single', 'Areas', 'E1'),
        ('constant-name', 'Areas', '1'),
    ]
    for name, defined_name, formula in cases:
        with zipfile.ZipFile(base / 'bisect/name-only.xlsx') as z:
            parts = {p: z.read(p) for p in z.namelist()}
        parts['xl/workbook.xml'] = re.sub(rb'<definedName name="Areas">.*?</definedName>', ('<definedName name="' + defined_name + '">' + formula + '</definedName>').encode(), parts['xl/workbook.xml'])
        file = out / (name + '.xlsx')
        with zipfile.ZipFile(file, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for part, data in parts.items():
                z.writestr(part, data)
        pairs.append({'name': name, 'file': str(file.resolve()), 'sha256': hashlib.sha256(file.read_bytes()).hexdigest()})
    cases = []
for name, version, keep, keep_name in cases:
    with zipfile.ZipFile(base / ('precedence-' + version + '.xlsx')) as z:
        parts = {p: z.read(p) for p in z.namelist()}
    if not keep_name:
        parts['xl/workbook.xml'] = re.sub(rb'<definedNames>.*?</definedNames>', b'', parts['xl/workbook.xml'])
    def cell(m):
        row = int(re.match(rb'<c r="B(\d+)"', m[0])[1])
        return m[0] if row in keep else re.sub(rb'<f>.*?</f>', b'', m[0])
    parts['xl/worksheets/sheet1.xml'] = re.sub(rb'<c r="B\d+"[^>]*>.*?</c>', cell, parts['xl/worksheets/sheet1.xml'])
    file = out / (name + '.xlsx')
    with zipfile.ZipFile(file, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        for part, data in parts.items():
            z.writestr(part, data)
    pairs.append({'name': name, 'file': str(file.resolve()), 'sha256': hashlib.sha256(file.read_bytes()).hexdigest()})
(out / 'pairs.json').write_text(json.dumps(pairs, indent=2) + '\n')
(out / 'CHECKLIST.md').write_text('# Precedence diagnostic isolation\n\nEight cases maximum, OFFICE_MAX=2. Keep every value and style. Remove only named formulas or the named range to isolate the shared failure of both diagnostic versions. Original inputs remain unchanged. Every open must be repair-free.\n\n' + '\n'.join('- ' + p['name'] + ': ' + p['sha256'] for p in pairs) + '\n')
