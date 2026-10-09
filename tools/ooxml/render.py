"""Render existing saved states once, with exact-content caching and failed-attempt accounting.

python3 tools/ooxml/render.py MANIFEST.jsonl OUTDIR [--jobs 4] [--resume]
Rows: app, file, scenario, original, saved, status (driver status), optional expectedPages.
Requires LibreOffice, pdfinfo, and a Python with the distribution's UNO bindings.
"""
import argparse
from collections import Counter
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
import hashlib
import json
import os
from pathlib import Path
import re
import select
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
import zipfile


def sha(path):
    with open(path, 'rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def payload(path):
    """Ignore ZIP compression/timestamps only when every member can be read exactly."""
    try:
        with zipfile.ZipFile(path) as archive:
            entries = archive.infolist()
            if len({e.filename for e in entries}) != len(entries) or sum(e.file_size for e in entries) > 512 * 1024 * 1024:
                raise ValueError('Use raw identity for duplicate names or large packages')
            digest = hashlib.sha256()
            for entry in sorted(entries, key=lambda e: e.filename):
                name = entry.filename.encode('utf-8')
                digest.update(len(name).to_bytes(8, 'big') + name + entry.file_size.to_bytes(8, 'big'))
                with archive.open(entry) as stream:
                    for chunk in iter(lambda: stream.read(1024 * 1024), b''):
                        digest.update(chunk)
            return 'opc:' + digest.hexdigest()
    except Exception:
        return 'file:' + sha(path)


class Worker:
    def __init__(self, args, output):
        self.args, self.output = args, output
        self.process = self.directory = self.log = None

    def read(self, timeout):
        if not select.select([self.process.stdout], [], [], timeout)[0]:
            raise TimeoutError('LibreOffice timed out after %s seconds' % timeout)
        line = self.process.stdout.readline()
        if not line:
            raise RuntimeError('LibreOffice worker exited; see ' + str(self.log.name))
        return json.loads(line)

    def start(self):
        self.directory = tempfile.TemporaryDirectory(prefix='vo-render-')
        profile = Path(self.directory.name) / 'profile'
        profile.mkdir()
        self.log = open(self.output / ('worker-' + Path(self.directory.name).name + '.log'), 'w')
        self.process = subprocess.Popen([
            self.args.uno_python, str(Path(__file__).with_name('render-worker.py')),
            str(profile), str(self.args.memory_mb),
        ], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=self.log, text=True, start_new_session=True)
        if not self.read(25).get('ready'):
            raise RuntimeError('LibreOffice worker did not become ready')

    def stop(self):
        if self.process:
            try:
                os.killpg(self.process.pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
            self.process.wait()
            self.process.stdin.close()
            self.process.stdout.close()
            self.process = None
        if self.log:
            self.log.close()
            self.log = None
        if self.directory:
            self.directory.cleanup()
            self.directory = None

    def render(self, source, key):
        attempts = []
        for _ in range(2):
            result = self.attempt(source, key)
            attempts.append(result)
            if result['status'] != 'ok' or result.get('workerRestart'):
                self.stop()
            if result['status'] == 'ok' or not result.get('workerFailure'):
                break
        # Keep a crash even if the same file works in a fresh worker. Timeouts
        # and ordinary document errors are not retried; each retires its worker.
        if len(attempts) > 1:
            result = dict(result, workerAttempts=attempts[:-1])
        return result

    def attempt(self, source, key):
        started = time.monotonic()
        try:
            if self.process is None:
                self.start()
            # Keep the filename for FILENAME fields; a fresh destination prevents stale-PDF passes.
            with tempfile.TemporaryDirectory(dir=self.directory.name, prefix='job-') as job:
                staged = Path(job) / source.name
                shutil.copyfile(source, staged)
                pdf = Path(job) / (source.stem + '.pdf')
                self.process.stdin.write(json.dumps({'source': str(staged), 'pdf': str(pdf)}) + '\n')
                self.process.stdin.flush()
                result = self.read(self.args.timeout)
                if result['status'] != 'ok':
                    return dict(result, seconds=round(time.monotonic() - started, 3))
                if not pdf.is_file():
                    raise ValueError('LibreOffice returned without a PDF')
                info = subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True, timeout=15, check=True)
                pages = re.search(r'^Pages:\s*(\d+)', info.stdout, re.M)
                if not pages or int(pages[1]) < 1:
                    raise ValueError('PDF has no readable page count')
                target = self.output / 'pdf' / (key + '.pdf')
                pdf_hash = sha(pdf)
                shutil.move(pdf, target)
                return dict(result, pages=int(pages[1]), pdf=str(target), pdfSha256=pdf_hash,
                            seconds=round(time.monotonic() - started, 3))
        except Exception as error:
            message = str(error)
            return dict(status='failed', error=message,
                        workerFailure=isinstance(error, (BrokenPipeError, ConnectionError)) or
                        'worker exited' in message or 'worker did not' in message,
                        seconds=round(time.monotonic() - started, 3))


def render_key(engine, name, fingerprint):
    return hashlib.sha256((json.dumps(engine, sort_keys=True) + '\0' + name + '\0' + fingerprint).encode()).hexdigest()


def reusable(item):
    return item['status'] == 'ok' and Path(item['pdf']).is_file() and sha(item['pdf']) == item['pdfSha256']


def import_successes(directory, engine, reason):
    """Explicitly reviewed harness-only changes may reuse verified successful PDFs."""
    directory = Path(directory).resolve()
    previous = json.loads((directory / 'run.json').read_text())['engine']
    changes = {k for k in previous.keys() | engine.keys() if previous.get(k) != engine.get(k)}
    if changes - {'converter', 'worker'}:
        raise ValueError('Reuse requires identical LibreOffice, fonts, limits and rendering policy')
    if not reason.strip():
        raise ValueError('Record why the reviewed code change leaves PDF generation unchanged')
    for line in (directory / 'cache.jsonl').read_text().splitlines():
        item = json.loads(line)
        if not reusable(item):
            continue
        source = Path(item['source'])
        if sha(source) != item['sourceSha256'] or payload(source) != item['payload']:
            raise ValueError('Prior render input changed: ' + str(source))
        if render_key(previous, source.name, item['payload']) != item['key']:
            raise ValueError('Prior render cache identity does not match its environment')
        yield dict(item, key=render_key(engine, source.name, item['payload']), reusedFrom=dict(
            run=str(directory), key=item['key'], engine=previous, reason=reason))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('manifest'); parser.add_argument('output')
    parser.add_argument('--jobs', type=int, default=4)
    parser.add_argument('--timeout', type=int, default=90)
    parser.add_argument('--memory-mb', type=int, default=6144)
    parser.add_argument('--uno-python', default='/usr/bin/python3')
    parser.add_argument('--resume', action='store_true')
    parser.add_argument('--reuse-successes', help='Prior run whose converter changes have been reviewed; successful PDFs only')
    parser.add_argument('--reuse-reason', default='', help='Explain why the reviewed change leaves rendering unchanged')
    args = parser.parse_args()
    if min(args.jobs, args.timeout, args.memory_mb) < 1:
        parser.error('jobs, timeout and memory must be positive')
    rows = [json.loads(line) for line in Path(args.manifest).read_text().splitlines() if line.strip()]
    output = Path(args.output).resolve(); output.mkdir(parents=True, exist_ok=True)
    report = output / 'results.jsonl'
    if report.exists() and not args.resume:
        parser.error('Output already has results; use --resume or a fresh directory')
    (output / 'pdf').mkdir(exist_ok=True)
    engine = dict(libreoffice=subprocess.check_output(['soffice', '--version'], text=True).strip(),
                  converter=sha(Path(__file__)), worker=sha(Path(__file__).with_name('render-worker.py')),
                  timeout=args.timeout, memoryMB=args.memory_mb,
                  fonts=hashlib.sha256(b'\n'.join(sorted(subprocess.check_output(['fc-list']).splitlines()))).hexdigest(),
                  policy='hidden, read-only, no document/link updates, no macros, abort interactions')
    metadata = dict(manifest=str(Path(args.manifest).resolve()), manifestSha256=sha(args.manifest),
                    attempts=len(rows), engine=engine)
    if args.reuse_successes:
        if not args.reuse_reason.strip():
            parser.error('--reuse-successes requires --reuse-reason')
        metadata['reuse'] = dict(run=str(Path(args.reuse_successes).resolve()), reason=args.reuse_reason,
                                metadataSha256=sha(Path(args.reuse_successes) / 'run.json'))
    meta = output / 'run.json'
    if args.resume and meta.exists() and json.loads(meta.read_text()) != metadata:
        parser.error('Manifest or rendering environment changed; use a fresh output directory')
    # Frozen manifests include raw input hashes. Check them once per path on
    # resume too, before trusting completed rows from a preceding invocation.
    verified = {}
    for row in rows:
        for kind in ('original', 'saved'):
            expected = row.get(kind + 'Sha256')
            if expected and row.get(kind):
                path = str(Path(row[kind]).resolve())
                if path not in verified:
                    verified[path] = sha(path)
                if verified[path] != expected:
                    parser.error('Manifest input changed: ' + row[kind])
    meta.write_text(json.dumps(metadata, indent=2) + '\n')
    lock, local, workers, cache, pending = threading.Lock(), threading.local(), [], {}, {}
    cache_path = output / 'cache.jsonl'
    if cache_path.exists():
        for line in cache_path.read_text().splitlines():
            item = json.loads(line)
            if item['status'] != 'ok' or reusable(item):
                cache[item['key']] = item
    cache_out = open(cache_path, 'a')
    if args.reuse_successes:
        for item in import_successes(args.reuse_successes, engine, args.reuse_reason):
            if item['key'] not in cache:
                cache[item['key']] = item
                cache_out.write(json.dumps(item) + '\n')
        cache_out.flush()

    def render(path):
        path = Path(path)
        fingerprint = payload(path)
        key = render_key(engine, path.name, fingerprint)
        with lock:
            if key in cache:
                return dict(cache[key], reused=True)
            owner = key not in pending
            if owner:
                pending[key] = Future()
            future = pending[key]
        if not owner:
            return dict(future.result(), reused=True)
        try:
            if not hasattr(local, 'worker'):
                local.worker = Worker(args, output)
                with lock:
                    workers.append(local.worker)
            result = dict(local.worker.render(path, key), key=key, payload=fingerprint,
                          source=str(path), sourceSha256=sha(path))
            with lock:
                cache[key] = result
                cache_out.write(json.dumps(result) + '\n'); cache_out.flush()
            future.set_result(result)
            return dict(result, reused=False)
        except BaseException as error:
            future.set_exception(error)
            raise

    def check(row):
        result = dict(row)
        result['driverStatus'] = row.get('status', 'ok')
        if not row.get('saved') or not Path(row['saved']).is_file():
            result.update(status='excluded' if result['driverStatus'] == 'excluded' else 'failed',
                          error=row.get('error', 'No saved artifact'))
            return result
        try:
            result['originalRender'] = render(row['original'])
            result['savedRender'] = render(row['saved'])
            a, b = result['originalRender'], result['savedRender']
            result['status'] = result['driverStatus'] if result['driverStatus'] != 'ok' else 'ok' if a['status'] == b['status'] == 'ok' else 'failed'
            if a['status'] == b['status'] == 'ok':
                expected = row.get('expectedPages', a['pages'])
                result['expectedPages'] = expected
                result['pageCountMatch'] = b['pages'] == expected
                if not result['pageCountMatch']:
                    result['status'] = 'failed'
                    result['error'] = 'Page count differs from the expected count'
        except Exception as error:
            result.update(status='failed', error=str(error))
        return result

    identity = lambda r: (r['app'], r['file'], r['scenario'])
    if len({identity(r) for r in rows}) != len(rows):
        parser.error('Duplicate app/file/scenario rows')
    done, totals = set(), Counter()
    if args.resume and report.exists():
        for line in report.read_text().splitlines():
            item = json.loads(line); done.add(identity(item)); totals[item['status']] += 1
    try:
        with open(report, 'a') as stream, ThreadPoolExecutor(max_workers=args.jobs) as pool:
            tasks = [pool.submit(check, row) for row in rows if identity(row) not in done]
            for future in as_completed(tasks):
                result = future.result()
                stream.write(json.dumps(result) + '\n'); stream.flush()
                totals[result['status']] += 1
                if sum(totals.values()) % 100 == 0:
                    print(json.dumps({'checked': sum(totals.values()), 'attempts': len(rows), 'status': totals,
                                      'uniqueRenders': len(cache)}), flush=True)
    finally:
        for worker in workers:
            worker.stop()
        cache_out.close()
    summary = dict(attempts=len(rows), completed=sum(totals.values()), status=totals, uniqueRenders=len(cache), engine=engine)
    (output / 'summary.json').write_text(json.dumps(summary, indent=2) + '\n')
    print(json.dumps(summary), flush=True)
    return int(bool(totals['failed']))


if __name__ == '__main__':
    sys.exit(main())
