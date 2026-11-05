#!/usr/bin/env python3
"""Summarize a measured XLS phase, including the exact Office jobs and hashes."""
import collections
import hashlib
import json
import pathlib
import sys

run = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else 'tmp/legacy/artifacts/C1-final')
rows = [json.loads(s) for s in (run / 'runs.jsonl').read_text().splitlines()]
originals = [r for r in rows if not r['mutation']]
excel = [r for r in originals if r['file'].startswith('legacy-excel/')]
count = lambda values: dict(collections.Counter(values))
jobs = []
for folder in sorted(run.parent.glob('C1-office-*')):
    for p in sorted(folder.glob('results/**/result.json')):
        r = json.loads(p.read_text()); pdf = p.with_name('office.pdf')
        jobs.append({'round': folder.name, 'result': str(p), 'id': r['id'], 'inputHash': r.get('inputSha256'), 'status': r['status'], 'repairFree': r.get('canOpenWithoutRepair'), 'pdfHash': hashlib.sha256(pdf.read_bytes()).hexdigest() if pdf.exists() else None})
final_jobs = [j for j in jobs if j['round'] in ['C1-office-excel', 'C1-office-long-fixed']
              or (j['round'] == 'C1-office-shapes-fixed' and '/ppt-rotation-gradients/' in j['result'])
              or (j['round'] == 'C1-office-shapes-final' and '/doc-rotation-gradients/' in j['result'])]
report = {
    'allCases': len(rows), 'allStatuses': count(r['status'] for r in rows),
    'originals': len(originals), 'originalStatuses': count(r['status'] for r in originals),
    'excelCorpus': len(excel), 'excelStatuses': count(r['status'] for r in excel),
    'excelRefusals': count(r['code'] for r in excel if r['status'] == 'refused'),
    'convertedVersions': count(r['sniff']['version'] for r in originals if r['status'] == 'converted'),
    'convertedTypes': count(r['type'] for r in originals if r['status'] == 'converted'),
    'losses': count(l['id'] for r in originals for l in r.get('losses', [])),
    'notificationCounts': count(l['id'] for r in originals for l in r.get('losses', []) if l.get('notify')),
    'packageChecks': sum(r.get('packageCheck', {}).get('status') == 0 for r in rows),
    'agreement': json.loads((run / 'agreement.json').read_text())['totals'],
    'largestConverted': max((r for r in originals if r['status'] == 'converted'), key=lambda r: r['size']),
    'largeSynthetic': json.loads((run.parent / 'C1-large/conversion.json').read_text()),
    'office': jobs,
    'officeFinal': final_jobs,
    'officeFinalCounts': {'files': len(final_jobs), 'repairFreeWithPDF': sum(j['repairFree'] and bool(j['pdfHash']) for j in final_jobs)}
}
(run / 'report.json').write_text(json.dumps(report, indent=2) + '\n')
print(json.dumps({k: v for k, v in report.items() if k not in ['office', 'officeFinal', 'largestConverted', 'largeSynthetic']}, indent=2))
