"""Author the pivot source kinds absent from the pinned public corpus.

Usage: python3 tools/ledger/test/pivot-fixtures.py CORPUS_DIR OUT_DIR

These are structural fixtures, not Office-authored or Office-accepted samples.
The named source keeps the public workbook's data and cache. The consolidation
source has two numeric input areas and matching cached records/output cells.
All unmodified parts stay byte-identical to the pinned source.
"""
import argparse
import hashlib
import json
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

SOURCE = 'libreoffice__0bef51ea4c70__Pivot1_Row.xlsx'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
WORKBOOK = 'xl/workbook.xml'
CACHE = 'xl/pivotCache/pivotCacheDefinition1.xml'
RECORDS = 'xl/pivotCache/pivotCacheRecords1.xml'
PIVOT = 'xl/pivotTables/pivotTable1.xml'
SHEET1 = 'xl/worksheets/sheet1.xml'
SHEET2 = 'xl/worksheets/sheet2.xml'


def replace_one(text, pattern, replacement):
    result, count = re.subn(pattern, lambda _: replacement, text, flags=re.S)
    if count != 1:
        raise ValueError(f'Expected one source element for {pattern}; found {count}')
    return result


def element(text, tag, replacement):
    return replace_one(text, rf'<{tag}\b[^>]*(?:/>|>.*?</{tag}>)', replacement)


def column(index):
    result = ''
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def sheet_data(rows):
    result = []
    for r, values in sorted(rows.items()):
        cells = []
        for c, value in sorted(values.items()):
            ref = column(c) + str(r)
            if isinstance(value, str):
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>')
            else:
                cells.append(f'<c r="{ref}"><v>{value}</v></c>')
        result.append(f'<row r="{r}">' + ''.join(cells) + '</row>')
    return '<sheetData>' + ''.join(result) + '</sheetData>'


def named(parts):
    workbook = parts[WORKBOOK].decode()
    if '<definedNames' in workbook:
        raise ValueError('The pinned source unexpectedly has defined names')
    names = '<definedNames><definedName name="PivotSource">\'Sheet1\'!$A$1:$C$6</definedName></definedNames>'
    workbook = replace_one(workbook, r'(?=<calcPr\b)', names)
    cache = element(parts[CACHE].decode(), 'worksheetSource', '<worksheetSource name="PivotSource"/>')
    return {WORKBOOK: workbook.encode(), CACHE: cache.encode()}


def consolidation(parts):
    # Both input areas use row labels and the same two column labels. The
    # cached Row/Column/Value records list exactly those ten numeric values.
    inputs = [('X1', 50, 1), ('X2', 100, 2), ('X3', 150, 3), ('X4', 200, 4), ('X5', 75, 5)]
    rows = {1: {1: 'Item', 2: 'Sales', 3: 'Units', 5: 'Item', 6: 'Sales', 7: 'Units'}}
    for i, values in enumerate(inputs):
        start = 1 if i < 3 else 5
        row = i + 2 if i < 3 else i - 1
        rows.setdefault(row, {}).update({start + j: value for j, value in enumerate(values)})
    source = element(parts[SHEET1].decode(), 'sheetData', sheet_data(rows))
    source = element(source, 'dimension', '<dimension ref="A1:G4"/>')

    cache = element(parts[CACHE].decode(), 'cacheSource',
                    '<cacheSource type="consolidation"><consolidation autoPage="0"><rangeSets count="2">'
                    '<rangeSet ref="A1:C4" sheet="Sheet1"/><rangeSet ref="E1:G3" sheet="Sheet1"/>'
                    '</rangeSets></consolidation></cacheSource>')
    cache = replace_one(cache, r'\brecordCount="5"', 'recordCount="10"')
    labels = ''.join(f'<s v="{name}"/>' for name, _, _ in inputs)
    fields = ('<cacheFields count="3"><cacheField name="Row" numFmtId="0"><sharedItems count="5">' + labels +
              '</sharedItems></cacheField><cacheField name="Column" numFmtId="0"><sharedItems count="2">'
              '<s v="Sales"/><s v="Units"/></sharedItems></cacheField><cacheField name="Value" numFmtId="0">'
              '<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" '
              'minValue="1" maxValue="200"/></cacheField></cacheFields>')
    cache = element(cache, 'cacheFields', fields)
    records = []
    output = {3: {1: 'Row', 2: 'Column', 3: 'Sum of Value'}}
    items = []
    for i, (name, sales, units) in enumerate(inputs):
        for j, value in enumerate((sales, units)):
            records.append(f'<r><x v="{i}"/><x v="{j}"/><n v="{value}"/></r>')
            output[4 + i * 2 + j] = {1: name, 2: ('Sales', 'Units')[j], 3: value}
            items.append(f'<i><x v="{i}"/><x v="{j}"/></i>')
    records = f'<?xml version="1.0" encoding="UTF-8"?><pivotCacheRecords xmlns="{S}" count="10">' + ''.join(records) + '</pivotCacheRecords>'
    pivot = element(parts[PIVOT].decode(), 'location', '<location ref="A3:C13" firstHeaderRow="1" firstDataRow="1" firstDataCol="2"/>')
    fields = ('<pivotFields count="3">' + ''.join(
        '<pivotField axis="axisRow" compact="0" outline="0" subtotalTop="0" showAll="0" defaultSubtotal="0">'
        f'<items count="{count}">' + ''.join(f'<item x="{i}"/>' for i in range(count)) + '</items></pivotField>'
        for count in (5, 2)) + '<pivotField dataField="1" compact="0" outline="0" showAll="0"/></pivotFields>')
    pivot = element(pivot, 'pivotFields', fields)
    pivot = element(pivot, 'rowItems', '<rowItems count="10">' + ''.join(items) + '</rowItems>')
    pivot = element(pivot, 'dataFields', '<dataFields count="1"><dataField name="Sum of Value" fld="2" baseField="0" baseItem="0"/></dataFields>')
    target = element(parts[SHEET2].decode(), 'sheetData', sheet_data(output))
    target = element(target, 'dimension', '<dimension ref="A3:C13"/>')
    return {CACHE: cache.encode(), RECORDS: records.encode(), PIVOT: pivot.encode(), SHEET1: source.encode(), SHEET2: target.encode()}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('corpus', type=Path)
    parser.add_argument('out', type=Path)
    args = parser.parse_args()
    source = args.corpus / SOURCE
    args.out.mkdir(parents=True, exist_ok=True)
    if any(args.out.iterdir()):
        raise ValueError('Use an empty output directory; existing fixtures are never overwritten')
    results = []
    with zipfile.ZipFile(source) as original:
        parts = {entry.filename: original.read(entry) for entry in original.infolist()}
        for name, build in [('named-source.xlsx', named), ('consolidation-source.xlsx', consolidation)]:
            changes = build(parts)
            for data in changes.values():
                ET.fromstring(data)
            target = args.out / name
            with zipfile.ZipFile(target, 'x') as saved:
                saved.comment = original.comment
                for entry in original.infolist():
                    saved.writestr(entry, changes.get(entry.filename, parts[entry.filename]))
            results.append({'file': name, 'changedParts': sorted(changes), 'sha256': hashlib.sha256(target.read_bytes()).hexdigest(),
                            'authoredFixture': True, 'officeAcceptance': 'pending'})
    manifest = {'source': str(source), 'sourceSha256': hashlib.sha256(source.read_bytes()).hexdigest(),
                'purpose': 'Structural source-kind fixtures; no claim of Office-authored or Office-accepted output.', 'fixtures': results}
    (args.out / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    print(json.dumps(manifest, indent=2))


if __name__ == '__main__':
    main()
