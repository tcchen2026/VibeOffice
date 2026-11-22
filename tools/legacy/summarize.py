"""Summarize a legacy measure run and the resumable LibreOffice reference outputs."""
import argparse
from collections import Counter
import json
from pathlib import Path


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--corpus', default='/work/corpora')
    p.add_argument('--run', default='tmp/legacy/artifacts/A-accepted-cfb')
    p.add_argument('--references', default='tmp/legacy/artifacts/references')
    p.add_argument('--out', default='tmp/legacy/artifacts/A-report.json')
    args = p.parse_args()
    manifest = json.loads((Path(args.corpus) / 'legacy-manifest.json').read_text())
    rows = [json.loads(line) for line in (Path(args.run) / 'runs.jsonl').read_text().splitlines()]
    originals = {r['file']: r for r in rows if not r.get('mutation')}
    refs = {}
    for item in manifest:
        if item['sha256'] not in refs:
            f = Path(args.references) / item['sha256'] / 'result.json'
            refs[item['sha256']] = json.loads(f.read_text()) if f.exists() else None
    report = dict(documents=len(manifest), unique_documents=len(refs), cases=len(rows), states=dict(Counter(r['status'] for r in rows)),
                  lo_unique_complete=sum(v is not None for v in refs.values()), formats={}, source_references={})
    for kind in ('word', 'excel', 'powerpoint'):
        items = [e for e in manifest if e['format'] == kind]
        values = [originals[e['file']] for e in items]
        lo = Counter()
        for item in items:
            result = refs[item['sha256']]
            if not result: lo['pending'] += 1; continue
            for ext, v in result['conversions'].items(): lo[ext + ':' + v['status']] += 1
        largest = max(items, key=lambda e: e['size']); largest_run = originals[largest['file']]
        report['formats'][kind] = dict(documents=len(items), versions=dict(Counter(r.get('sniff', {}).get('version', 'unidentified') for r in values)),
            detected=dict(Counter(r.get('sniff', {}).get('format', 'unidentified') for r in values)),
            encrypted=sum(r.get('sniff', {}).get('encrypted') is True for r in values), macros=sum(r.get('features', {}).get('macros', False) for r in values),
            states=dict(Counter(r['status'] for r in values)), reasons=dict(Counter(r.get('code', '') for r in values)), lo=dict(lo),
            largest=dict(file=largest['file'], bytes=largest['size'], ms=largest_run.get('ms'), code=largest_run.get('code')))
    references = dict(files=0, word_text_documents=0, word_stored_page_counts=0, excel_workbooks=0,
                      excel_sheets=0, excel_cells=0, excel_formulas=0, presentations=0, slides=0, property_errors=0)
    for item in manifest:
        file = Path(args.run) / 'references' / (item['sha256'] + '.json')
        if not file.exists(): continue
        r = json.loads(file.read_text()); source = r['references']; references['files'] += 1
        references['property_errors'] += sum(k.endswith('Error') for k in source)
        if 'stories' in source:
            references['word_text_documents'] += 1
            summary = source.get('summary', [])
            if summary and summary[0]['values'].get('14', {}).get('value') is not None: references['word_stored_page_counts'] += 1
        if 'sheets' in source:
            references['excel_workbooks'] += 1; references['excel_sheets'] += len(source['sheets'])
            for sheet in source['sheets']:
                references['excel_cells'] += len(sheet['cells'])
                references['excel_formulas'] += sum('formula' in c for c in sheet['cells'])
        if 'slides' in source:
            references['presentations'] += 1; references['slides'] += len(source['slides'])
    report['source_references'] = references
    Path(args.out).write_text(json.dumps(report, indent=2) + '\n')
    print(json.dumps(report, indent=2))

if __name__ == '__main__': main()
