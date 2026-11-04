"""Opt-in Office VM check; fixtures and results stay outside git. Standard-library Python only.

python3 tests/win_test_server/live.py --out tmp/win-test-server/live
Requires the SSH tunnel, the local POI corpus and the historical batch-7 repair controls.
"""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import time
import urllib.parse
import urllib.request
import zipfile


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--api', default='http://127.0.0.1:18765')
    parser.add_argument('--corpus', type=Path, default=Path('/home/john/corpora/poi/test-data'))
    parser.add_argument('--repair-cases', type=Path, default=Path('tmp/lossless/office-batch-7-isolation-1'))
    parser.add_argument('--out', type=Path, required=True)
    parser.add_argument('--case', action='append', help='Only named cases (repeatable); default runs all ten')
    args = parser.parse_args()
    args.out.mkdir(parents=True, exist_ok=True)
    cases = [
        ('word.docx', args.corpus / 'document/sample.docx', 'ok', 60),
        ('excel.xlsx', args.corpus / 'spreadsheet/sample.xlsx', 'ok', 60),
        ('powerpoint.pptx', args.corpus / 'xmldsign/hello-world-unsigned.pptx', 'ok', 60),
        ('excel-repair.xlsx', args.repair_cases / '01-excel-original.xlsx', 'rejected', 60),
        ('powerpoint-repair.pptx', args.repair_cases / '03-powerpoint-fixture-original.pptx', 'rejected', 60),
        ('powerpoint-control.pptx', args.repair_cases / '04-powerpoint-without-test-object.pptx', 'ok', 60),
        ('password.docx', args.corpus / 'document/bug53475-password-is-pass.docx', 'blocked', 30),
        ('deadline.docx', Path('tests/quire/lines/linecases.docx'), 'timeout', 5),
        ('after-failure.docx', args.corpus / 'document/sample.docx', 'ok', 60),
    ]
    for _, source, _, _ in cases:
        if not source.is_file():
            parser.error(f'Missing fixture: {source}')
    # Independent malformed-XML control for Word, leaving the real source untouched.
    corrupt = args.out / 'invalid.docx'
    with zipfile.ZipFile(cases[0][1]) as src, zipfile.ZipFile(corrupt, 'w', zipfile.ZIP_DEFLATED) as dst:
        for member in src.infolist():
            data = src.read(member)
            if member.filename == 'word/document.xml':
                data = b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><broken'
            dst.writestr(member, data)
    cases.insert(3, ('word-invalid.docx', corrupt, 'rejected', 30))
    if args.case:
        cases = [case for case in cases if case[0] in args.case]
        if len(cases) != len(set(args.case)):
            parser.error('Unknown --case')
    jobs = []
    for name, source, expect, timeout in cases:
        data = source.read_bytes()
        query = urllib.parse.urlencode({'filename': name, 'timeout': timeout})
        req = urllib.request.Request(args.api + '/jobs?' + query, data=data, headers={'Content-Type': 'application/octet-stream'})
        with urllib.request.urlopen(req, timeout=30) as response:
            job = json.load(response)
        jobs.append((name, source, expect, job['id'], hashlib.sha256(data).hexdigest()))
    results, failures = [], []
    for name, source, expect, ident, digest in jobs:
        deadline = time.monotonic() + 600
        while True:
            with urllib.request.urlopen(args.api + '/jobs/' + ident, timeout=20) as response:
                result = json.load(response)
            if result['status'] not in ('queued', 'running'):
                break
            if time.monotonic() >= deadline:
                raise TimeoutError(ident)
            time.sleep(.3)
        (args.out / (name + '.json')).write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
        passed = result['inputSha256'] == digest and result['inputUnchanged']
        if expect == 'ok':
            passed &= result['status'] == 'ok' and result['canOpenWithoutRepair'] is True and result['pdfStatus'] == 'ok'
        elif expect == 'rejected':
            passed &= result['status'] in ('repair_required', 'open_failed') and result['canOpenWithoutRepair'] is False
        else:
            passed &= result['status'] == expect and result['canOpenWithoutRepair'] is None
        pages = None
        if result.get('pdf'):
            data = urllib.request.urlopen(args.api + result['pdf'], timeout=30).read()
            passed &= data.startswith(b'%PDF-') and hashlib.sha256(data).hexdigest() == result['pdfSha256']
            pdf = args.out / (name + '.pdf')
            pdf.write_bytes(data)
            info = subprocess.check_output(['pdfinfo', str(pdf)], text=True)
            pages = int(next(line.split(':')[1] for line in info.splitlines() if line.startswith('Pages:')))
            passed &= pages > 0
        for dialog in result['dialogs']:
            if dialog.get('screenshot'):
                image = urllib.request.urlopen(args.api + dialog['screenshot'], timeout=30).read()
                (args.out / (name + '.png')).write_bytes(image)
        if expect == 'blocked':
            passed &= bool(result['dialogs']) and bool(result['dialogs'][0].get('screenshot'))
        if not passed:
            failures.append(name)
        results.append({'name': name, 'source': str(source.resolve()), 'expect': expect, 'passed': passed, 'pages': pages, **result})
        print(f'{name}: {result["status"]}, canOpenWithoutRepair={result["canOpenWithoutRepair"]}, pages={pages}, {"PASS" if passed else "FAIL"}', flush=True)
    (args.out / 'results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2) + '\n')
    if failures:
        raise SystemExit('Failed: ' + ', '.join(failures))


if __name__ == '__main__':
    main()
