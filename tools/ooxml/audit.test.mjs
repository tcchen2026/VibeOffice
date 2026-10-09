import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test('cell audit distinguishes a timeout during style access from an unreadable style', () => {
  const script = fileURLToPath(new URL('../ledger/test/loss-audit.py', import.meta.url));
  // Raise the real alarm in the reader's style accessor, without waiting for the
  // large-workbook timeout or depending on an openpyxl installation.
  const probe = `
import contextlib, io, json, runpy, signal, sys, tempfile, zipfile
from pathlib import Path
from types import SimpleNamespace
script, mode = sys.argv[1:]
class Cell:
    value, coordinate = 'text', 'A1'
    @property
    def font(self):
        if mode == 'timeout': signal.raise_signal(signal.SIGALRM)
        raise ValueError('Invalid style reference')
with tempfile.TemporaryDirectory() as tmp:
    root = Path(tmp)
    for folder in ('originals', 'saved'):
        (root / folder).mkdir()
        with zipfile.ZipFile(root / folder / 'sample.xlsx', 'w') as z:
            z.writestr('[Content_Types].xml', '<Types/>')
    def load(path, **kwargs):
        sheets = [] if Path(path).parent.name == 'originals' else [SimpleNamespace(title='Sheet1', _cells={(1, 1): Cell()})]
        return SimpleNamespace(worksheets=sheets)
    sys.modules['openpyxl'] = SimpleNamespace(load_workbook=load)
    output = root / 'report.json'
    sys.argv = [script, str(root / 'originals'), str(root / 'saved'), str(output)]
    with contextlib.redirect_stdout(io.StringIO()):
        try: runpy.run_path(script, run_name='__main__')
        except SystemExit as error:
            assert error.code == 1
    print(output.read_text())
`;
  for (const mode of ['timeout', 'invalid']) {
    const result = spawnSync('python3', ['-c', probe, script, mode], {
      encoding: 'utf8', timeout: 10000,
      env: { ...process.env, CELLS: '1', TIMEOUT: '40', MAXMB: '1e9', SCENARIO: '' }
    });
    assert.equal(result.status, 0, result.stderr);
    const report = JSON.parse(result.stdout);
    assert.equal(report.accounting.attempted, 1);
    assert.equal(report.accounting.failed, 1);
    assert.equal(report.cells.books, 0);
    assert.equal(report.cells.unreadable, mode === 'timeout' ? 0 : 1);
    assert.deepEqual(report.cells.timeouts || [], mode === 'timeout' ? ['sample'] : []);
    assert.equal(report.accounting.errors[0].error,
      mode === 'timeout' ? 'Cell comparison timeout' : 'Cannot compare cell style at A1');
  }
});

test('render reuse ignores ZIP compression but detects every changed member and malformed input', () => {
  const script = fileURLToPath(new URL('./render.py', import.meta.url));
  const probe = `
import importlib.util, pathlib, sys, tempfile, zipfile
spec = importlib.util.spec_from_file_location('render', sys.argv[1])
module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
with tempfile.TemporaryDirectory() as tmp:
    root = pathlib.Path(tmp)
    def write(name, compression, value=b'<document/>', extra=False):
        p = root / name
        with zipfile.ZipFile(p, 'w', compression=compression) as z:
            z.writestr('word/document.xml', value)
            z.writestr('word/media/image1.png', b'image bytes')
            if extra: z.writestr('unknown.xml', b'<extra/>')
        return p
    a = write('original.docx', zipfile.ZIP_STORED)
    b = write('repacked.docx', zipfile.ZIP_DEFLATED)
    assert module.sha(a) != module.sha(b)
    assert module.payload(a) == module.payload(b)
    assert module.payload(a) != module.payload(write('changed.docx', zipfile.ZIP_STORED, b'<changed/>'))
    assert module.payload(a) != module.payload(write('extra.docx', zipfile.ZIP_STORED, extra=True))
    bad = root / 'malformed.docx'; bad.write_bytes(b'not a ZIP')
    assert module.payload(bad) == 'file:' + module.sha(bad)
`;
  const result = spawnSync('python3', ['-c', probe, script], { encoding: 'utf8', timeout: 10000 });
  assert.equal(result.status, 0, result.stderr);
});

test('render workers retire on failure, recover a lost bridge once, and retain failed attempts', () => {
  const script = fileURLToPath(new URL('./render.py', import.meta.url));
  const probe = `
import importlib.util, sys
spec = importlib.util.spec_from_file_location('render', sys.argv[1])
module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
lost = dict(status='failed', error='Binary URP bridge already disposed', workerFailure=True)
bad = dict(status='failed', error='LibreOffice did not open the document')
ok = dict(status='ok', pages=1)
def run(replies):
    worker = module.Worker(None, None)
    pending = iter(replies); stopped = []
    worker.attempt = lambda *args: next(pending)
    worker.stop = lambda: stopped.append(True)
    result = worker.render(None, None)
    return result, len(stopped)
result, stops = run([lost, ok])
assert result['status'] == 'ok' and stops == 1
assert result['workerAttempts'] == [lost]
result, stops = run([lost, bad])
assert result['status'] == 'failed' and stops == 2
assert result['workerAttempts'] == [lost]
result, stops = run([lost, lost])
assert result['status'] == 'failed' and stops == 2
assert result['workerAttempts'] == [lost]
for failure in [bad, dict(status='failed', error='LibreOffice timed out after 90 seconds')]:
    result, stops = run([failure])
    assert result == failure and stops == 1
result, stops = run([dict(ok, workerRestart=True)])
assert result['status'] == 'ok' and stops == 1
`;
  const result = spawnSync('python3', ['-c', probe, script], { encoding: 'utf8', timeout: 10000 });
  assert.equal(result.status, 0, result.stderr);
});

test('reviewed render reuse verifies source and PDF bytes, rejects changed policy, and never imports failures', () => {
  const script = fileURLToPath(new URL('./render.py', import.meta.url));
  const probe = `
import importlib.util, json, pathlib, sys, tempfile
spec = importlib.util.spec_from_file_location('render', sys.argv[1])
module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
with tempfile.TemporaryDirectory() as tmp:
    root = pathlib.Path(tmp)
    source = root / 'a.docx'; source.write_bytes(b'input')
    pdf = root / 'a.pdf'; pdf.write_bytes(b'pdf')
    old = dict(converter='old', worker='old', fonts='same', policy='same')
    new = dict(old, converter='new', worker='new')
    (root / 'run.json').write_text(json.dumps(dict(engine=old)))
    item = dict(status='ok', pdf=str(pdf), pdfSha256=module.sha(pdf), source=str(source),
                sourceSha256=module.sha(source), payload=module.payload(source))
    item['key'] = module.render_key(old, source.name, item['payload'])
    (root / 'cache.jsonl').write_text(json.dumps(item) + '\\n' + json.dumps(dict(status='failed', error='lost bridge')) + '\\n')
    imported = list(module.import_successes(root, new, 'Worker recovery only'))
    assert len(imported) == 1 and imported[0]['key'] != item['key']
    assert imported[0]['reusedFrom']['key'] == item['key']
    for engine, reason in [(dict(new, fonts='changed'), 'reason'), (new, '')]:
        try: list(module.import_successes(root, engine, reason))
        except ValueError: pass
        else: raise AssertionError('Unreviewed/changed rendering was accepted')
    source.write_bytes(b'changed')
    try: list(module.import_successes(root, new, 'reason'))
    except ValueError: pass
    else: raise AssertionError('Changed source was accepted')
    source.write_bytes(b'input'); pdf.write_bytes(b'corrupt')
    assert list(module.import_successes(root, new, 'reason')) == []
`;
  const result = spawnSync('python3', ['-c', probe, script], { encoding: 'utf8', timeout: 10000 });
  assert.equal(result.status, 0, result.stderr);
});
