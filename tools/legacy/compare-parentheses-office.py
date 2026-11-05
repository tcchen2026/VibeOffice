#!/usr/bin/env python3
"""Check C2 renderer Office opens and compare every PDF page with its control."""
import hashlib
import json
import pathlib
import subprocess

base = pathlib.Path('tmp/legacy/artifacts/C2-parens-office')
cache_file = base / 'raster-cache.json'
cache = json.loads(cache_file.read_text()) if cache_file.exists() else {}


def digest(file):
    return hashlib.sha256(pathlib.Path(file).read_bytes()).hexdigest()


def raster(file):
    key = digest(file)
    if key in cache:
        return cache[key]
    proc = subprocess.Popen(['pdftoppm', '-r', '40', '-gray', str(file)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    h, size = hashlib.sha256(), 0
    while data := proc.stdout.read(65536):
        h.update(data)
        size += len(data)
    error = proc.stderr.read().decode()
    if proc.wait() != 0:
        raise RuntimeError(error)
    cache[key] = {'sha256': h.hexdigest(), 'bytes': size}
    cache_file.write_text(json.dumps(cache, indent=2) + '\n')
    return cache[key]


def result(round_name, item):
    path = base / round_name / 'results' / item['name'] / 'file/result.json'
    r = json.loads(path.read_text())
    pdf = path.with_name('office.pdf')
    if not r.get('canOpenWithoutRepair') or r.get('pdfStatus') != 'ok':
        raise RuntimeError(str(path) + ': ' + r['status'])
    if digest(item['file']) != r['inputSha256']:
        raise RuntimeError('Office input changed: ' + item['name'])
    return {'job': r['id'], 'result': str(path), 'pdf': str(pdf), 'pdfHash': digest(pdf), 'inputHash': r['inputSha256']}


proof = json.loads((base / 'unchanged-proof.json').read_text())
comparisons, jobs = [], []
for item in json.loads((base / 'accepted/inputs.json').read_text()):
    r = result('accepted', item)
    jobs.append(r)
    old = item['previous']
    a, b = raster(old['pdf']), raster(r['pdf'])
    comparisons.append({'name': item['name'], 'beforePDF': old['pdf'], 'afterPDF': r['pdf'], 'before': a, 'after': b, 'identicalPixels': a == b})
    proof.append({**old, 'currentPackage': item['file'], 'currentHash': r['inputHash'], 'acceptedPackage': item['file'], 'acceptedHash': r['inputHash'], **r})

controls = [item for item in json.loads((base / 'controls/inputs.json').read_text()) if item['pair'] == '15228']
results = {item['name']: result('controls', item) for item in controls}
results.update({item['name']: result('qualified', item) for item in json.loads((base / 'qualified/inputs.json').read_text())})
jobs.extend(results.values())
for pair in ('15228', 'precedence'):
    old, new = results[pair + '-before'], results[pair + '-after']
    a, b = raster(old['pdf']), raster(new['pdf'])
    comparisons.append({'name': pair, 'beforePDF': old['pdf'], 'afterPDF': new['pdf'], 'before': a, 'after': b, 'identicalPixels': a == b})
    if pair == '15228':
        item = next(i for i in controls if i['name'] == '15228-after')
        proof.append({'source': item['source'], 'currentPackage': item['file'], 'currentHash': new['inputHash'], 'acceptedPackage': item['file'], 'acceptedHash': new['inputHash'], **new, 'diagnostic': False, 'partDifferences': []})

report = {'jobs': jobs, 'comparisons': comparisons, 'identicalPixels': all(r['identicalPixels'] for r in comparisons), 'proof': proof}
(base / 'verification.json').write_text(json.dumps(report, indent=2) + '\n')
print(json.dumps({'officeOpens': len(jobs), 'pdfPairs': len(comparisons), 'identicalPixels': report['identicalPixels'], 'converted': sum(not p['diagnostic'] for p in proof), 'printControls': sum(p['diagnostic'] for p in proof)}))
if not report['identicalPixels']:
    raise SystemExit(1)
