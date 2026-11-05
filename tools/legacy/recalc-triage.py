#!/usr/bin/env python3
"""Separate recorded Ledger differences using the independent LO source exports."""
import collections
import json
import math
import pathlib
import runpy
import sys

run = pathlib.Path(sys.argv[1])
read = runpy.run_path('tools/legacy/xls-agreement.py')['read']
rows = [json.loads(s) for s in (run / 'runs.jsonl').read_text().splitlines() if s.strip()]
source = {r['file']: r for r in rows if not r['mutation']}
counts = collections.Counter()
functions = collections.defaultdict(collections.Counter)
details = []


def value(cell):
    if not cell:
        return None
    return {'e': cell['value']} if cell['type'] == 'e' else cell['value']


def same(a, b):
    if isinstance(a, (int, float)) and not isinstance(a, bool) and isinstance(b, (int, float)) and not isinstance(b, bool):
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-10)
    return type(a) is type(b) and a == b


for file in (run / 'recalc-final').glob('*.json'):
    if file.name == 'summary.json':
        continue
    r = json.loads(file.read_text())
    if not r.get('mismatches'):
        continue
    sha = source[r['file']]['sha256']
    lo = next((run.parent / 'references' / sha / 'xlsx').glob('*.xlsx'), None)
    cells = read(lo) if lo else {}
    for mismatch in r['mismatches']:
        cell = cells.get(mismatch['sheet'], {}).get(mismatch['cell'])
        other = value(cell)
        kind = 'noLOReference' if not lo else 'matchesLO' if same(mismatch.get('actual'), other) else 'sourceCacheMatchesLO' if same(mismatch['expected'], other) else 'threeDifferentResults'
        counts[kind] += 1
        for name in mismatch['functions']:
            functions[name][kind] += 1
        details.append({'source': r['file'], **mismatch, 'libreoffice': other, 'classification': kind})
result = {'differences': sum(counts.values()), 'classifications': dict(counts), 'functions': dict(functions), 'rows': details}
(run / 'recalc-triage.json').write_text(json.dumps(result, indent=2, ensure_ascii=False) + '\n')
print(json.dumps({k: v for k, v in result.items() if k != 'rows'}, indent=2))
