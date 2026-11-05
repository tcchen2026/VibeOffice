#!/usr/bin/env python3
"""C2 Office failures: whole-file C1 controls and one-concept changes, eight jobs."""
import hashlib
import json
import pathlib
import re
import zipfile

base = pathlib.Path('tmp/legacy/artifacts/C2-office-2')
out = pathlib.Path('tmp/legacy/artifacts/C2-office-bisect-1').resolve()
out.mkdir(parents=True, exist_ok=True)
rows = json.loads((base / 'diff.json').read_text())
cases = []


def write(label, source, change=None, description='C1 whole-file control'):
    target = out / (label + '.xlsx')
    if change is None:
        target.write_bytes(pathlib.Path(source).read_bytes())
    else:
        with zipfile.ZipFile(source) as z:
            parts = {n: z.read(n) for n in z.namelist()}
        change(parts)
        with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as z:
            for n, data in parts.items():
                z.writestr(n, data)
    cases.append({'name': label, 'file': str(target), 'sha256': hashlib.sha256(target.read_bytes()).hexdigest(), 'source': str(source), 'change': description})


def replace(parts, part, old, new):
    parts[part] = parts[part].replace(old.encode(), new.encode())


selected = {name: next(r for r in rows if r['source'].endswith('__' + name + '.xls')) for name in ['FormulaRefs', 'external_named_function', 'user_defined_function']}
for name, row in selected.items():
    write(name + '-C1-control', row['original'])


def no_blank_link(parts):
    for n in ['xl/externalLinks/externalLink2.xml', 'xl/externalLinks/_rels/externalLink2.xml.rels']:
        del parts[n]
    replace(parts, 'xl/workbook.xml', '<externalReference r:id="rId7"/>', '')
    replace(parts, 'xl/_rels/workbook.xml.rels', '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink2.xml"/>', '')
    replace(parts, '[Content_Types].xml', '<Override PartName="/xl/externalLinks/externalLink2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>', '')
    replace(parts, 'xl/worksheets/sheet1.xml', '[3]', '[2]')


write('FormulaRefs-without-blank-link', selected['FormulaRefs']['file'], no_blank_link, 'Remove only the unused space-only external workbook, renumber its following formula index')


def no_xlm_name(parts):
    parts['xl/workbook.xml'] = re.sub(rb'<definedName name="Zellfarbe"[^>]*>.*?</definedName>', b'', parts['xl/workbook.xml'])


def no_cell_formulas(parts):
    for n in parts:
        if n.startswith('xl/worksheets/') and n.endswith('.xml'):
            parts[n] = re.sub(rb'<f(?:\s[^>]*)?>.*?</f>', b'', parts[n])


write('external-function-without-XLM-name', selected['external_named_function']['file'], no_xlm_name, 'Remove only the Zellfarbe defined name containing GET.CELL')
write('external-function-without-cell-formulas', selected['external_named_function']['file'], no_cell_formulas, 'Keep names and links, remove only worksheet formula elements')


def named_call(parts, prefix):
    replace(parts, 'xl/worksheets/sheet1.xml', '<f>Sum(', '<f>' + prefix + ('SUM' if not prefix else 'Sum') + '(')
    replace(parts, 'xl/worksheets/sheet1.xml', '<f>sumif(', '<f>' + prefix + ('SUMIF' if not prefix else 'sumif') + '(')


write('user-function-xll-prefix', selected['user_defined_function']['file'], lambda p: named_call(p, '_xll.'), 'Prefix both registered user-function names with _xll.')
write('user-function-uppercase', selected['user_defined_function']['file'], lambda p: named_call(p, ''), 'Uppercase the two user-function names, no other formula change')
(out / 'pairs.json').write_text(json.dumps([{'name': c['name'], 'file': c['file']} for c in cases], indent=2) + '\n')
(out / 'hashes.json').write_text(json.dumps(cases, indent=2) + '\n')
(out / 'CHECKLIST.md').write_text('# I24.C2 formula bisection 1\n\nEight files, OFFICE_MAX=2. All original failed inputs are retained in C2-office-2.\n\n' + '\n'.join('- ' + c['name'] + ': ' + c['change'] + '; SHA256 ' + c['sha256'] for c in cases) + '\n\nRecord repair status, dialogs, job IDs and PDF hashes. Do not accept repair dialogs.\n')
print(json.dumps({'files': len(cases), 'out': str(out)}))
