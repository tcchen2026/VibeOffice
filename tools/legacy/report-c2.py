#!/usr/bin/env python3
"""Final C2 counts, source-derived formula totals, and all Office job hashes."""
import collections
import hashlib
import json
import pathlib

base = pathlib.Path('tmp/legacy/artifacts')
run = base / 'C2-final'
rows = [json.loads(line) for line in (run / 'runs.jsonl').read_text().splitlines()]
originals = [r for r in rows if not r['mutation']]
converted = [r for r in originals if r['status'] == 'converted']
count = lambda values: dict(collections.Counter(values))
formulas = sum((collections.Counter(r['stats']['formulas']) for r in converted), collections.Counter())
jobs = []
for folder in sorted(base.glob('C2-office-*')):
    for file in sorted(folder.glob('results/**/result.json')):
        r = json.loads(file.read_text())
        pdf = file.with_name('office.pdf')
        jobs.append({'round': folder.name, 'result': str(file), 'job': r['id'], 'inputHash': r.get('inputSha256'), 'status': r['status'], 'repairFree': r.get('canOpenWithoutRepair'), 'pdfStatus': r.get('pdfStatus'), 'pdfHash': hashlib.sha256(pdf.read_bytes()).hexdigest() if pdf.exists() else None})
recalc = json.loads((run / 'recalc-final/summary.json').read_text())
report = {
    'cases': len(rows), 'statuses': count(r['status'] for r in rows),
    'originals': len(originals), 'originalStatuses': count(r['status'] for r in originals),
    'excelCorpus': sum(r['file'].startswith('legacy-excel/') for r in originals),
    'excelConverted': sum(r['file'].startswith('legacy-excel/') for r in converted),
    'versions': count(r['sniff']['version'] for r in converted), 'types': count(r['type'] for r in converted),
    'formulas': dict(formulas), 'formulaFiles': sum(r['stats']['formulas']['source'] > 0 for r in converted),
    'losses': count(l['id'] for r in converted for l in r['losses']),
    'formulaLosses': [{'source': r['file'], 'losses': [l for l in r['losses'] if l['id'] in ['legacy-formulas', 'legacy-defined-name', 'legacy-external-name']]} for r in converted if any(l['id'] in ['legacy-formulas', 'legacy-defined-name', 'legacy-external-name'] for l in r['losses'])],
    'agreement': json.loads((run / 'agreement.json').read_text())['totals'],
    'recalculation': {k: v for k, v in recalc.items() if k != 'functions'},
    'recalculationClasses': json.loads((run / 'recalc-triage.json').read_text())['classifications'],
    'largestConverted': max(converted, key=lambda r: r['size']),
    'largeSynthetic': json.loads((base / 'C2-large/conversion.json').read_text()),
    'officeProof': json.loads((run / 'office-proof.json').read_text()), 'officeHistory': jobs,
    'libreofficeReopen': json.loads((base / 'C2-lo-reopen/results.json').read_text()),
    'browser': json.loads((base / 'browser-C2/results.json').read_text())
}
(run / 'report.json').write_text(json.dumps(report, indent=2) + '\n')
print(json.dumps({k: v for k, v in report.items() if k not in ['formulaLosses', 'largestConverted', 'largeSynthetic', 'officeProof', 'officeHistory', 'browser', 'libreofficeReopen']}, indent=2))
