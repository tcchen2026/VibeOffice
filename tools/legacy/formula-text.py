#!/usr/bin/env python3
"""Compare a formula-renderer revision with prior packages and LO source exports.

Usage: python3 -I tools/legacy/formula-text.py BEFORE_RUN AFTER_RUN SAMPLES_JSON OUT
Only formula text may change; every other XML value/attribute and binary part is
compared. LO's explicit cell formulas are compared verbatim, without normalizing
away parentheses, quoting, case or function names. Shared followers without text
are counted separately. Inputs are read only.
"""
import collections
import json
import pathlib
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile

before, after, samples, output = map(pathlib.Path, sys.argv[1:5])
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def rows(folder):
    return {(r['file'], r['mutation']): r for r in map(json.loads, (folder / 'runs.jsonl').read_text().splitlines())}


def package(folder, row):
    return row.get('pkg') or folder / 'packages' / (row['sha256'] + ('-' + row['mutation'] if row['mutation'] else '') + '.' + row['type'])


def formulas(doc):
    found = {}
    for c in doc.iter(NS + 'c'):
        f = c.find(NS + 'f')
        if f is not None and f.text:
            found['cell:' + c.get('r')] = f.text
    for i, n in enumerate(doc.iter(NS + 'definedName')):
        text = n.get('refersTo', n.text)
        if text:
            found['name:' + str(i) + ':' + n.get('name', '')] = text
    return found


def without_formulas(doc):
    for n in doc.iter():
        if n.tag in (NS + 'f', NS + 'definedName'):
            if n.text:
                n.text = 'FORMULA'
            if 'refersTo' in n.attrib:
                n.set('refersTo', 'FORMULA')
    return ET.tostring(doc)


def cell_texts(file):
    found, empty = {}, collections.Counter()
    with zipfile.ZipFile(file) as z:
        rels = {n.get('Id'): posixpath.normpath(posixpath.join('xl', n.get('Target'))).lstrip('/') for n in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        for sheet in ET.fromstring(z.read('xl/workbook.xml')).find(NS + 'sheets'):
            for c in ET.fromstring(z.read(rels[sheet.get(REL + 'id')])).iter(NS + 'c'):
                f = c.find(NS + 'f')
                if f is None:
                    continue
                if not f.text:
                    empty[f.get('t', 'empty')] += 1
                else:
                    found[sheet.get('name'), c.get('r')] = f.text
    return found, dict(empty)


old_rows, new_rows = rows(before), rows(after)
counts = collections.Counter()
unexpected, stats, pairs, examples = [], [], {}, []
for key, row in new_rows.items():
    old = old_rows[key]
    if row['status'] != old['status'] or row.get('stats') != old.get('stats'):
        stats.append({'source': key, 'before': old.get('stats', old['status']), 'after': row.get('stats', row['status'])})
    if row['status'] != 'converted':
        continue
    counts['packages'] += 1
    with zipfile.ZipFile(package(before, old)) as a, zipfile.ZipFile(package(after, row)) as b:
        if set(a.namelist()) != set(b.namelist()):
            unexpected.append({'source': key, 'kind': 'part names'})
            continue
        for part in a.namelist():
            left, right = a.read(part), b.read(part)
            if left == right:
                continue
            counts['changedParts'] += 1
            if not part.endswith('.xml'):
                unexpected.append({'source': key, 'part': part})
                continue
            x, y = ET.fromstring(left), ET.fromstring(right)
            fa, fb = formulas(x), formulas(y)
            if without_formulas(x) != without_formulas(y) or fa.keys() != fb.keys():
                unexpected.append({'source': key, 'part': part})
            for location in fa.keys() & fb.keys():
                if fa[location] == fb[location]:
                    continue
                counts['changedFormulaTexts'] += 1
                if not row['mutation']:
                    counts['changedOriginalFormulaTexts'] += 1
                pair = pairs.setdefault((fa[location], fb[location]), {'before': fa[location], 'after': fb[location], 'count': 0, 'source': row['file'], 'part': part, 'location': location})
                pair['count'] += 1

lo_results, lo_pairs, lo_totals = [], [], collections.Counter()
for source in json.loads(samples.read_text()):
    old, new = old_rows[source, None], new_rows[source, None]
    lo_file = next((before.parent / 'references' / old['sha256'] / 'xlsx').glob('*.xlsx'))
    a, _ = cell_texts(package(before, old))
    b, _ = cell_texts(package(after, new))
    lo, empty = cell_texts(lo_file)
    count = collections.Counter()
    differences = []
    for place, text in b.items():
        if place not in lo:
            count['noExplicitLOText'] += 1
            continue
        count['compared'] += 1
        count['beforeExact'] += a.get(place) == lo[place]
        count['afterExact'] += text == lo[place]
        if text != lo[place]:
            count['different'] += 1
            lo_pairs.append({'source': source, 'sheet': place[0], 'cell': place[1], 'converted': text, 'libreoffice': lo[place]})
            if len(differences) < 12:
                differences.append({'sheet': place[0], 'cell': place[1], 'before': a.get(place), 'after': text, 'libreoffice': lo[place]})
        if '15228' in source and place[1] == 'E19':
            examples.append({'source': source, 'sheet': place[0], 'cell': place[1], 'before': a[place], 'after': text, 'libreoffice': lo[place]})
    lo_totals.update(count)
    lo_results.append({'source': source, 'libreoffice': str(lo_file), 'counts': dict(count), 'emptyLOFormulas': empty, 'differences': differences})

output.mkdir(parents=True, exist_ok=True)
report = {'counts': dict(counts), 'unexpectedPartChanges': unexpected, 'statusOrStatsChanges': stats, 'distinctChangedTexts': len(pairs), 'libreofficeTotals': dict(lo_totals), 'libreoffice': lo_results, 'examples': examples}
(output / 'report.json').write_text(json.dumps(report, indent=2, ensure_ascii=False) + '\n')
(output / 'formula-pairs.json').write_text(json.dumps(list(pairs.values()), ensure_ascii=False) + '\n')
(output / 'libreoffice-pairs.json').write_text(json.dumps(lo_pairs, ensure_ascii=False) + '\n')
print(json.dumps({k: v for k, v in report.items() if k not in ('libreoffice', 'examples')}, indent=2))
if unexpected or stats:
    raise SystemExit(1)
