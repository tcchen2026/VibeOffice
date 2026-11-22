#!/usr/bin/env python3
"""Compare converted XLS values with stored BIFF references and external LO output.

python3 -I tools/legacy/xls-agreement.py RUN_DIR [--styles]
No third-party reader is needed. Numeric comparisons allow serialization rounding.
"""
import collections
import json
import math
import pathlib
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
ERRORS = {0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA'}


def unescape(s):
    return re.sub(r'_x([0-9a-fA-F]{4})_', lambda m: chr(int(m[1], 16)), s)


def text(node):
    return unescape(''.join(t.text or '' for t in node.findall(NS + 't') + node.findall(NS + 'r/' + NS + 't')))


def canonical(node):
    if node is None:
        return None
    return [node.tag.split('}')[-1], sorted(node.attrib.items()), sorted((canonical(c) for c in node), key=str)]


def read(path, styles=False):
    with zipfile.ZipFile(path) as z:
        book = ET.fromstring(z.read('xl/workbook.xml'))
        rels = {r.get('Id'): posixpath.normpath(posixpath.join('xl', r.get('Target'))).lstrip('/') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        strings = [text(si) for si in ET.fromstring(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in z.namelist() else []
        style_doc = ET.fromstring(z.read('xl/styles.xml')) if styles else None
        xf = list(style_doc.find(NS + 'cellXfs')) if styles else []
        sections = {key: list(style_doc.find(NS + key)) for key in ['fonts', 'fills', 'borders']} if styles else {}
        numfmt = {int(n.get('numFmtId')): n.get('formatCode') for n in style_doc.findall(NS + 'numFmts/' + NS + 'numFmt')} if styles else {}
        rows = {}
        for s in book.find(NS + 'sheets'):
            sheet = ET.fromstring(z.read(rels[s.get(REL + 'id')]))
            cells = {}
            for c in sheet.findall('.//' + NS + 'sheetData/' + NS + 'row/' + NS + 'c'):
                kind, value = c.get('t', 'n'), c.findtext(NS + 'v')
                if kind == 's':
                    kind, value = 's', strings[int(value)]
                elif kind == 'inlineStr':
                    kind, value = 's', text(c.find(NS + 'is'))
                elif kind == 'str':
                    kind, value = 's', unescape(value or '')
                elif value is None:
                    kind, value = 'blank', None
                elif kind == 'n':
                    value = float(value)
                elif kind == 'b':
                    value = value == '1'
                row = {'type': kind, 'value': value, 'formula': c.find(NS + 'f') is not None}
                if styles:
                    f = xf[int(c.get('s', '0'))]
                    row['style'] = {key: canonical(sections[sec][int(f.get(attr, '0'))]) for key, sec, attr in [('font', 'fonts', 'fontId'), ('fill', 'fills', 'fillId'), ('border', 'borders', 'borderId')]}
                    row['style'].update({key: canonical(f.find(NS + key)) for key in ['alignment', 'protection']})
                    n = int(f.get('numFmtId', '0')); row['style']['numFmt'] = numfmt.get(n, n)
                cells[c.get('r')] = row
            rows[s.get('name')] = cells
        return rows


def address(r, c):
    out = ''
    c += 1
    while c:
        c, m = divmod(c - 1, 26)
        out = chr(65 + m) + out
    return out + str(r + 1)


def equal(a, b):
    # A missing blank and an explicitly stored blank have the same value.
    if a['type'] == 'blank':
        return b is None or b['type'] == 'blank'
    if b is None or a['type'] != b['type']:
        return False
    av, bv = a['value'], b['value']
    return math.isclose(av, bv, rel_tol=1e-14, abs_tol=1e-12) if a['type'] == 'n' else av == bv


def compare(a, b):
    counts = collections.Counter(); examples = []
    for name, cells in a.items():
        for addr, v in cells.items():
            other = b.get(name, {}).get(addr)
            category = 'formula' if v.get('formula') else 'value'
            counts[category + 'Total'] += 1
            if equal(v, other):
                counts[category + 'Agree'] += 1
            else:
                counts['difference:' + category + ':' + v['type'] + ':' + (other or {}).get('type', 'missing')] += 1
                if len(examples) < 8:
                    examples.append({'sheet': name, 'cell': addr, 'source': v, 'converted': other})
            if 'style' in v and other and 'style' in other:
                for k, style in v['style'].items():
                    if style != other['style'][k]:
                        counts['styleDiff:' + k] += 1
    return {'counts': dict(counts), 'examples': examples}


def main():
    run = pathlib.Path(sys.argv[1]); base = pathlib.Path('tmp/legacy/artifacts')
    listing = run / 'results.json'
    rows = json.loads(listing.read_text()) if listing.exists() else [json.loads(line) for line in (run / 'runs.jsonl').read_text().splitlines()]
    output = []
    for row in rows:
        if row.get('mutation') or row['status'] != 'converted' or row['type'] not in ['xlsx', 'xltx', 'xlsm']:
            continue
        sha = row.get('sourceHash', row.get('sha256')); path = row.get('pkg', run / 'packages' / (sha + '.' + row['type']))
        converted = read(path, '--styles' in sys.argv); result = {'file': row['file'], 'sha256': sha}
        source = base / 'B-complete' / 'references' / (sha + '.json')
        if source.exists():
            sheets = json.loads(source.read_text()).get('references', {}).get('sheets', [])
            original = {}
            for sheet in sheets:
                cells = {}
                for c in sheet['cells']:
                    kind, value = c['type'], c['value']
                    if kind == 'e': value = ERRORS.get(value, value)
                    if kind == 'b': value = bool(value)
                    if kind == 'empty': kind, value = 's', ''
                    cells[address(c['row'], c['col'])] = {'type': kind, 'value': value, 'formula': 'formula' in c}
                original[sheet['name']] = cells
            result['source'] = compare(original, converted)
            for name, cells in converted.items():
                for addr, value in cells.items():
                    value['formula'] = original.get(name, {}).get(addr, {}).get('formula', False)
        lo = next((base / 'references' / sha / 'xlsx').glob('*.xlsx'), None)
        if lo:
            result['libreoffice'] = compare(converted, read(lo, '--styles' in sys.argv))
        output.append(result)
    totals = {kind: dict(sum((collections.Counter(row.get(kind, {}).get('counts', {})) for row in output), collections.Counter())) for kind in ['source', 'libreoffice']}
    (run / 'agreement.json').write_text(json.dumps({'files': len(output), 'totals': totals, 'rows': output}, indent=2) + '\n')
    print(json.dumps({'files': len(output), 'totals': totals}, indent=2))


if __name__ == '__main__':
    main()
