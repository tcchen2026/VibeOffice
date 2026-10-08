"""Validate every corpus attempt against its original, including failed/excluded accounting.

python3 tools/ooxml/validate.py RESULTS.jsonl ORIGINALS SAVED OUT.jsonl
  [--sdk /path/to/oxval.dll] [--dotnet dotnet] [--lo OUTDIR] [--limit N]
LibreOffice gets its own temporary profile and writes PDFs outside the repository.
"""
import argparse
from collections import Counter
import json
from pathlib import Path
import re
import select
import shutil
import subprocess
import tempfile
import sys
from package import Package

def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('results'); p.add_argument('originals'); p.add_argument('saved'); p.add_argument('out')
    p.add_argument('--sdk'); p.add_argument('--dotnet', default='dotnet'); p.add_argument('--lo'); p.add_argument('--limit', type=int)
    p.add_argument('--resume', action='store_true')
    args = p.parse_args()
    rows = [json.loads(l) for l in Path(args.results).read_text().splitlines() if l.strip()]
    # Each driver transaction can write more than its first save. Validate every
    # emitted state, including undo/redo and recovered drafts, with the same oracle.
    expanded = []
    for row in rows:
        expanded.append(row)
        for state in ('before', 'undo', 'redo', 'second', 'draft', 'recovered'):
            scenario = row.get('scenario', 'save') + '-' + state
            if (Path(args.saved) / scenario / row['file']).is_file():
                expanded.append(dict(row, scenario=scenario, _original_count=state in ('before', 'undo')))
    rows = expanded
    done = set()
    failed = 0
    if args.resume and Path(args.out).exists():
        for line in Path(args.out).read_text().splitlines():
            try:
                r = json.loads(line); done.add((r['file'], r['scenario'], r.get('sha256')))
                failed += r['status'] == 'failed'
            except (ValueError, KeyError):
                pass
    out = open(args.out, 'a' if args.resume else 'w')
    sdk = None
    with tempfile.TemporaryDirectory(prefix='vo-lo-profile-') as profile:
        for i, row in enumerate(rows[:args.limit]):
            key = (row['file'], row.get('scenario', 'save'), row.get('sha256'))
            if key in done:
                continue
            original = Path(args.originals) / row['file']
            saved = Path(args.saved) / row.get('scenario', 'save') / row['file']
            r = {k: row[k] for k in ('file', 'scenario', 'sha256') if k in row}
            if not saved.exists():
                r.update(status='excluded' if row.get('status') == 'excluded' else 'failed', error=row.get('error', 'Missing saved copy'))
            else:
                try:
                    a, b = Package(original), Package(saved)
                    r['originalIssues'], r['savedIssues'] = a.check(), b.check()
                    ac = Counter(json.dumps(e, sort_keys=True) for e in r['originalIssues'])
                    bc = Counter(json.dumps(e, sort_keys=True) for e in r['savedIssues'])
                    r['newIssues'] = [json.loads(e) for e, n in (bc - ac).items() for _ in range(n)]
                    r['status'] = 'failed' if r['newIssues'] else 'ok'
                    if args.sdk:
                        if sdk is None or sdk.poll() is not None:
                            sdk = subprocess.Popen([args.dotnet, args.sdk, '--stdin-pairs'], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True)
                        sdk.stdin.write(json.dumps(dict(original=str(original), saved=str(saved))) + '\n'); sdk.stdin.flush()
                        if not select.select([sdk.stdout], [], [], 120)[0]:
                            sdk.kill(); sdk.wait(); sdk = None
                            raise TimeoutError('SDK validator timed out')
                        line = sdk.stdout.readline()
                        if not line:
                            raise ValueError('SDK validator returned no report')
                        r['sdk'] = json.loads(line)
                        if r['sdk']['status'] != 'ok': r['status'] = 'failed'
                    if args.lo:
                        counts = []
                        for kind, file in [('original', original), ('saved', saved)]:
                            dest = Path(args.lo).resolve() / row.get('scenario', 'save') / kind
                            dest.mkdir(parents=True, exist_ok=True)
                            # A corpus symlink's basename is its identity. Stage real bytes so LO
                            # cannot derive another basename, and never accept an old PDF as success.
                            with tempfile.TemporaryDirectory(dir=profile) as job:
                                source = Path(job) / file.name
                                shutil.copyfile(file, source)
                                pdf = Path(job) / (file.stem + '.pdf')
                                cmd = ['soffice', '-env:UserInstallation=' + Path(profile).as_uri(), '--headless', '--convert-to', 'pdf', '--outdir', job, str(source)]
                                proc = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
                                if proc.returncode or not pdf.exists():
                                    raise ValueError('LibreOffice ' + kind + ': ' + proc.stdout + proc.stderr)
                                info = subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True, check=True).stdout
                                shutil.copyfile(pdf, dest / pdf.name)
                            m = re.search(r'^Pages:\s*(\d+)', info, re.M)
                            if not m: raise ValueError('pdfinfo returned no page count')
                            counts.append(int(m[1]))
                        r['pages'] = dict(original=counts[0], saved=counts[1])
                        expected = row.get('expectedPages', row.get('expectedSlides'))
                        if row.get('_original_count'): expected = counts[0]
                        if expected is None and row.get('scenario', 'save') in ('save', 'repeat'):
                            expected = counts[0]
                        if expected is not None and counts[1] != expected:
                            r['status'] = 'failed'; r['pageCountMismatch'] = True
                except Exception as e:
                    r.update(status='failed', error=str(e))
            failed += r['status'] == 'failed'
            out.write(json.dumps(r) + '\n'); out.flush()
            if (i + 1) % 100 == 0: print(f'{i + 1}/{len(rows)} checked', file=sys.stderr)
    out.close()
    if sdk is not None:
        sdk.stdin.close()
        sdk.wait(timeout=10)
    return int(failed != 0)

if __name__ == '__main__':
    sys.exit(main())
