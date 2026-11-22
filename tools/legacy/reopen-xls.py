#!/usr/bin/env python3
"""Open our packages in LibreOffice and save separate XLSX copies, without UI."""
import hashlib
import json
import pathlib
import subprocess
import sys
import tempfile
import time

listing, output = sys.argv[1:3]
out = pathlib.Path(output).resolve()
out.mkdir(parents=True, exist_ok=True)
results = []
for i, row in enumerate(json.loads(pathlib.Path(listing).read_text())):
    source = pathlib.Path(row if isinstance(row, str) else row['file']).resolve()
    dest = out / str(i)
    dest.mkdir(exist_ok=True)
    start = time.monotonic()
    with tempfile.TemporaryDirectory(prefix='i24-c2-reopen-', dir='/dev/shm') as profile:
        try:
            result = subprocess.run(['soffice', '-env:UserInstallation=' + pathlib.Path(profile).as_uri(), '--headless', '--convert-to', 'xlsx', '--outdir', str(dest), str(source)], capture_output=True, timeout=120)
            (dest / 'libreoffice.log').write_bytes(result.stdout + result.stderr)
            target = dest / (source.stem + '.xlsx')
            info = {'source': str(source), 'sourceHash': hashlib.sha256(source.read_bytes()).hexdigest(), 'seconds': time.monotonic() - start, 'status': 'saved' if target.exists() and result.returncode == 0 else 'failed'}
            if target.exists():
                info.update(file=str(target), sha256=hashlib.sha256(target.read_bytes()).hexdigest())
        except subprocess.TimeoutExpired:
            info = {'source': str(source), 'status': 'timeout'}
    results.append(info)
    print(json.dumps(info), flush=True)
(out / 'results.json').write_text(json.dumps(results, indent=2) + '\n')
sys.exit(any(r['status'] != 'saved' for r in results))
