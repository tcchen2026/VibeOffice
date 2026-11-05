"""Bounded LibreOffice references for the pinned binary corpus (never reads LO code).
Run with python3 -I. Results are resumable and keyed by source SHA-256.
"""
import argparse
from concurrent.futures import ThreadPoolExecutor
import hashlib
import json
import os
from pathlib import Path
import signal
import subprocess
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET


def convert(item, root, output, timeout, profiles):
    file = root / item['file']
    digest = item['sha256']
    dest = output / digest
    dest.mkdir(parents=True, exist_ok=True)
    result_path = dest / 'result.json'
    if result_path.exists():
        return json.loads(result_path.read_text())
    result = dict(sha256=digest, source=item['file'], conversions={})
    kind = {'word': 'docx', 'excel': 'xlsx', 'powerpoint': 'pptx'}[item['format']]
    for extension in (kind, 'pdf'):
        folder = dest / extension
        folder.mkdir(exist_ok=True)
        # Keep profiles on a POSIX filesystem: shared host mounts may not support
        # LibreOffice's locks or Python's fd-relative temporary-directory cleanup.
        with tempfile.TemporaryDirectory(prefix='vibeoffice-i24-', dir=profiles) as profile:
            user = Path(profile) / 'user'
            user.mkdir()
            # Never run document macros or refresh links in a public test fixture.
            (user / 'registrymodifications.xcu').write_text('''<?xml version="1.0"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item><item oor:path="/org.openoffice.Office.Calc/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>2</value></prop></item><item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>2</value></prop></item></oor:items>''')
            args = ['soffice', '-env:UserInstallation=' + Path(profile).as_uri(), '--headless', '--norestore',
                    '--nodefault', '--nolockcheck', '--nofirststartwizard', '--convert-to', extension,
                    '--outdir', str(folder), str(file.resolve())]
            started = time.monotonic()
            scratch = Path(profile) / 'tmp'
            scratch.mkdir()
            proc = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, start_new_session=True,
                                    env={**os.environ, 'TMPDIR': str(scratch)})
            timed_out = False
            try: log, _ = proc.communicate(timeout=timeout)
            except subprocess.TimeoutExpired:
                timed_out = True
                os.killpg(proc.pid, signal.SIGKILL)
                log, _ = proc.communicate()
            (dest / (extension + '.log')).write_bytes(log)
            saved = folder / (file.stem + '.' + extension)
            result['conversions'][extension] = dict(status='converted' if saved.is_file() and saved.stat().st_size else 'timeout' if timed_out else 'refused',
                 seconds=round(time.monotonic() - started, 3), exit=proc.returncode, file=str(saved.relative_to(output)) if saved.is_file() else None)
            if saved.is_file():
                result['conversions'][extension]['sha256'] = hashlib.sha256(saved.read_bytes()).hexdigest()
                if extension != 'pdf':
                    try:
                        with zipfile.ZipFile(saved) as z:
                            parts = [n for n in z.namelist() if n == 'word/document.xml' or
                                (n.startswith('ppt/slides/slide') and n.endswith('.xml') and '/_rels/' not in n) or
                                (n.startswith('xl/worksheets/sheet') and n.endswith('.xml') and '/_rels/' not in n)]
                            texts = {}
                            for name in parts:
                                if z.getinfo(name).file_size > 128 * 1024 * 1024: raise ValueError('Reference XML exceeds limit')
                                tree = ET.fromstring(z.read(name))
                                texts[name] = [e.text or '' for e in tree.iter() if e.tag.rsplit('}', 1)[-1] in ('t', 'v')]
                            (dest / 'content.json').write_text(json.dumps(texts, ensure_ascii=False))
                    except Exception as e: result['content_error'] = type(e).__name__ + ': ' + str(e)
    result_path.write_text(json.dumps(result, indent=2) + '\n')
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--corpus', default='/work/corpora')
    parser.add_argument('--out', default='tmp/legacy/artifacts/references')
    parser.add_argument('--timeout', type=int, default=90)
    parser.add_argument('--workers', type=int, default=2)
    parser.add_argument('--profiles', default='/tmp')
    parser.add_argument('--limit', type=int)
    args = parser.parse_args()
    root, output = Path(args.corpus), Path(args.out).resolve()
    output.mkdir(parents=True, exist_ok=True)
    manifest = json.loads((root / 'legacy-manifest.json').read_text())
    by_hash = {e['sha256']: e for e in manifest}
    items = sorted(by_hash.values(), key=lambda e: e['size'])[:args.limit]
    with ThreadPoolExecutor(max_workers=args.workers) as pool, (output / 'runs.jsonl').open('a') as log:
        for i, result in enumerate(pool.map(lambda e: convert(e, root, output, args.timeout, args.profiles), items)):
            log.write(json.dumps(result) + '\n'); log.flush()
            print(f'{i + 1}/{len(items)} {result["source"]}: ' + ', '.join(f'{k} {v["status"]}' for k, v in result['conversions'].items()), flush=True)

if __name__ == '__main__': main()
