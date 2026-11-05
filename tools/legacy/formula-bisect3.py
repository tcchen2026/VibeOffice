#!/usr/bin/env python3
"""Isolate the remaining external-named-function failure after empty-cache fix."""
import hashlib
import json
import pathlib
import re
import zipfile

base = pathlib.Path('tmp/legacy/artifacts/C2-office-bisect-2/external_named_function-no-empty-cache-set.xlsx')
out = pathlib.Path('tmp/legacy/artifacts/C2-office-bisect-3').resolve()
out.mkdir(parents=True, exist_ok=True)
cases = []
with zipfile.ZipFile(base) as z:
    original = {n: z.read(n) for n in z.namelist()}
for mode in ['no-XLM-name', 'no-cell-formulas', 'no-name-or-formulas', 'prefixed-XLM-name']:
    parts = dict(original)
    if mode in ['no-XLM-name', 'no-name-or-formulas']:
        parts['xl/workbook.xml'] = re.sub(rb'<definedName name="Zellfarbe"[^>]*>.*?</definedName>', b'', parts['xl/workbook.xml'])
    if mode == 'prefixed-XLM-name':
        parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace(b'>GET.CELL(', b'>_xlfn.GET.CELL(')
    if mode in ['no-cell-formulas', 'no-name-or-formulas']:
        for n in parts:
            if n.startswith('xl/worksheets/') and n.endswith('.xml'):
                parts[n] = re.sub(rb'<f(?:\s[^>]*)?>.*?</f>', b'', parts[n])
    file = out / (mode + '.xlsx')
    with zipfile.ZipFile(file, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, value in parts.items():
            z.writestr(n, value)
    cases.append({'name': mode, 'file': str(file), 'sha256': hashlib.sha256(file.read_bytes()).hexdigest()})
(out / 'pairs.json').write_text(json.dumps([{'name': c['name'], 'file': c['file']} for c in cases], indent=2) + '\n')
(out / 'hashes.json').write_text(json.dumps(cases, indent=2) + '\n')
(out / 'CHECKLIST.md').write_text('# I24.C2 formula bisection 3\n\nFour files, OFFICE_MAX=2. Base is the still-failing external_named_function control with only its empty sheetDataSet omitted.\n\n' + '\n'.join('- ' + c['name'] + ': ' + c['sha256'] for c in cases) + '\n\nThe no-name-or-formulas control tests overlapping faults. Preserve all job results, never accept a dialog.\n')
print(json.dumps({'out': str(out), 'files': len(cases)}))
