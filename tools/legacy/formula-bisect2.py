#!/usr/bin/env python3
"""Second C2 Office round: empty external caches and isolated registered calls."""
import hashlib
import json
import pathlib
import re
import zipfile

base = pathlib.Path('tmp/legacy/artifacts/C2-office-2')
out = pathlib.Path('tmp/legacy/artifacts/C2-office-bisect-2').resolve()
out.mkdir(parents=True, exist_ok=True)
rows = json.loads((base / 'diff.json').read_text())
cases = []


def write(label, name, edit, description):
    source = next(r['file'] for r in rows if r['source'].endswith('__' + name + '.xls'))
    with zipfile.ZipFile(source) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    edit(parts)
    file = out / (label + '.xlsx')
    with zipfile.ZipFile(file, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, value in parts.items():
            z.writestr(n, value)
    cases.append({'name': label, 'file': str(file), 'sha256': hashlib.sha256(file.read_bytes()).hexdigest(), 'source': source, 'change': description})


def no_empty_cache(parts):
    for n in parts:
        if n.startswith('xl/externalLinks/') and n.endswith('.xml'):
            parts[n] = parts[n].replace(b'<sheetDataSet></sheetDataSet>', b'')


for name in ['FormulaRefs', 'external_named_function']:
    write(name + '-no-empty-cache-set', name, no_empty_cache, 'Omit only empty sheetDataSet elements; all links, formulas and names unchanged')


def one_formula(parts, keep, replacement=None):
    n = 'xl/worksheets/sheet1.xml'
    def cell(match):
        value = match[0]
        if match[1].decode() != keep:
            return re.sub(rb'<f(?:\s[^>]*)?>.*?</f>', b'', value)
        return re.sub(rb'<f(?:\s[^>]*)?>.*?</f>', ('<f>' + replacement + '</f>').encode(), value) if replacement else value
    parts[n] = re.sub(rb'<c r="([^"]+)"[^>]*>.*?</c>', cell, parts[n])


for cell in ['B46', 'C46', 'B51']:
    write('user-function-only-' + cell, 'user_defined_function', lambda p, c=cell: one_formula(p, c), 'Keep only formula ' + cell + '; retain all cell caches')
write('user-function-only-uppercase-SUM', 'user_defined_function', lambda p: one_formula(p, 'B46', 'SUM(B9:C42)'), 'Only B46, uppercase SUM')
write('user-function-only-xlfn-sumif', 'user_defined_function', lambda p: one_formula(p, 'C46', '_xlfn.sumif(D9:D42)'), 'Only C46, _xlfn.sumif prefix')
write('user-function-only-unknown-name', 'user_defined_function', lambda p: one_formula(p, 'C46', 'CUSTOMFN(D9:D42)'), 'Only C46, diagnostic CUSTOMFN name')
(out / 'pairs.json').write_text(json.dumps([{'name': c['name'], 'file': c['file']} for c in cases], indent=2) + '\n')
(out / 'hashes.json').write_text(json.dumps(cases, indent=2) + '\n')
(out / 'CHECKLIST.md').write_text('# I24.C2 formula bisection 2\n\nEight files, OFFICE_MAX=2. Compare against C2-office-2 and first bisection controls.\n\n' + '\n'.join('- ' + c['name'] + ': ' + c['change'] + '; SHA256 ' + c['sha256'] for c in cases) + '\n\nRecord jobs and repair/PDF status; do not accept dialogs.\n')
print(json.dumps({'out': str(out), 'files': len(cases)}))
