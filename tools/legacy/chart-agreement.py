"""Compare chart data in converted packages with independently converted LO files.

python3 -I tools/legacy/chart-agreement.py RUN_DIRECTORY
Series are matched as a multiset within each sheet, independent of chart z-order.
The measured values are cached chart points, not recalculated worksheet cells.
"""
import collections
import json
import math
from pathlib import Path
import posixpath
import sys
import xml.etree.ElementTree as E
import zipfile

S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def read(file):
    result = {}
    with zipfile.ZipFile(file) as z:
        names = set(z.namelist())

        def relationships(part):
            name = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
            if name not in names:
                return {}
            return {r.get('Id'): posixpath.normpath(posixpath.join(posixpath.dirname(part), r.get('Target')))
                    .lstrip('/') for r in E.fromstring(z.read(name)) if r.get('TargetMode') != 'External'}

        book = E.fromstring(z.read('xl/workbook.xml'))
        sheets = relationships('xl/workbook.xml')
        for sheet in book.find(S + 'sheets'):
            part = sheets[sheet.get(R + 'id')]
            root = E.fromstring(z.read(part))
            drawing = root.find(S + 'drawing')
            if drawing is None:
                continue
            drawing_part = relationships(part)[drawing.get(R + 'id')]
            drawing_rels = relationships(drawing_part)
            charts = E.fromstring(z.read(drawing_part)).iter(C + 'chart')
            series = []
            for reference in charts:
                chart_part = drawing_rels[reference.get(R + 'id')]
                chart = E.fromstring(z.read(chart_part))
                for item in chart.iter(C + 'ser'):
                    source = item.find(C + 'val')
                    if source is None:
                        source = item.find(C + 'yVal')
                    if source is None:
                        continue
                    data = []
                    for point in source.iter(C + 'pt'):
                        value = point.findtext(C + 'v', '')
                        try:
                            number = float(value)
                            if math.isfinite(number):
                                value = format(number, '.12g')
                        except ValueError:
                            pass
                        data.append((int(point.get('idx')), value))
                    series.append(tuple(sorted(data)))
            result[sheet.get('name')] = series
    return result


def main():
    run = Path(sys.argv[1])
    files = [json.loads(line) for line in (run / 'runs.jsonl').read_text().splitlines()]
    rows = []
    for row in files:
        if row['status'] != 'converted' or row.get('mutation') or row.get('type') not in ['xlsx', 'xltx']:
            continue
        file = Path(row.get('pkg') or run / 'packages' / (row['sha256'] + '.' + row['type']))
        current = read(file)
        if not any(current.values()):
            continue
        reference = next((Path('tmp/legacy/artifacts/references') / row['sha256'] / 'xlsx').glob('*.xlsx'), None)
        if reference is None:
            rows.append({'file': row['file'], 'status': 'no-reference'})
            continue
        external = read(reference)
        matches = total = external_total = 0
        differences = []
        for sheet in current.keys() | external.keys():
            a = collections.Counter(current.get(sheet, []))
            b = collections.Counter(external.get(sheet, []))
            matches += sum((a & b).values())
            total += sum(a.values())
            external_total += sum(b.values())
            if a != b:
                differences.append({'sheet': sheet, 'currentOnly': list((a - b).elements()),
                                    'referenceOnly': list((b - a).elements())})
        rows.append({'file': row['file'], 'status': 'compared', 'series': total, 'referenceSeries': external_total,
                     'matchingSeries': matches, 'same': not differences, 'differences': differences})
    summary = {'files': len(rows), 'compared': sum(r['status'] == 'compared' for r in rows),
               'same': sum(r.get('same', False) for r in rows),
               'series': sum(r.get('series', 0) for r in rows),
               'referenceSeries': sum(r.get('referenceSeries', 0) for r in rows),
               'matchingSeries': sum(r.get('matchingSeries', 0) for r in rows)}
    (run / 'chart-agreement.json').write_text(json.dumps({'summary': summary, 'rows': rows}, indent=2))
    print(json.dumps(summary))


if __name__ == '__main__':
    main()
