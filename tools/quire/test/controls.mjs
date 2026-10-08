// Focused browser checks against authored Word controls, with edit/undo/copy/draft artifacts.
// node tools/quire/test/controls.mjs ~/corpora/word /tmp/quire-controls-edits
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [corpus, output] = process.argv.slice(2);
if (!corpus || !output) throw new Error('Usage: controls.mjs CORPUS OUTPUT');
const cases = [
  ['libreoffice__dc4315b6159e__sdt-data-binding-char-style.docx', 'text'],
  ['libreoffice__d45ac3c2f95d__content-control-date-data-binding.docx', 'date'],
  ['libreoffice__d65cc38e46cb__sdt-run-checkbox.docx', 'checkbox'],
];
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), page = await openPage('quire', { persistence: false });
const reports = [];
try {
  for (const [name, type] of cases) {
    const data = fs.readFileSync(path.join(corpus, name)).toString('base64');
    const result = await page.evaluate(`(${check.toString()})(${JSON.stringify(data)}, ${JSON.stringify(type)})`);
    for (const [kind, data] of Object.entries(result.artifacts)) {
      fs.mkdirSync(path.join(output, kind), { recursive: true });
      fs.writeFileSync(path.join(output, kind, name), Buffer.from(data, 'base64'));
    }
    delete result.artifacts;
    reports.push({ file: name, scenario: 'edited', attempted: true, ...result });
    console.log(name + ': ' + result.status + ' (' + result.checks + ' checks)');
  }
  if (page.errors.length) throw new Error(page.errors.join('\n'));
  fs.writeFileSync(path.join(output, 'results.jsonl'), reports.map(r => JSON.stringify(r)).join('\n') + '\n');
} finally { await page.close(); stop(); }

async function check(data, type) {
  const { D, O, preserve: P, opc: K } = L, artifacts = {};
  let checks = 0;
  const assert = (ok, message) => { if (!ok) throw new Error(message); checks++; };
  const encode = async blob => { const b = new Uint8Array(await blob.arrayBuffer()); let s = ''; for (let i = 0; i < b.length; i += 32768) s += String.fromCharCode(...b.subarray(i, i + 32768)); return btoa(s); };
  const save = async (kind, doc) => { const blob = await L.docx.write(doc); artifacts[kind] = await encode(blob); return blob; };
  const read = async blob => (await L.docx.read(new Uint8Array(await blob.arrayBuffer()))).doc;
  const records = doc => [...P.controls(doc).records.values()];
  const document = (await L.docx.read(Uint8Array.from(atob(data), c => c.charCodeAt(0)))).doc;
  L.app.loadDoc(document, 'Controls', { saved: true });
  let record = records(document).find(r => r.control.type === type && r.complete);
  assert(record, 'Expected control not found');
  const key = record.control.key, beforeText = record.text, beforeXML = record.control.shell.xml;
  await save('before', document);
  let expected;
  D.tx('Edit control', () => {
    if (type === 'text') { expected = beforeText.slice(0, 1) + ' edited' + beforeText.slice(1); O.insertText(D.pos(record.start.pos.p, record.start.pos.o + 1), ' edited'); }
    else { P.setControlValue(key, type === 'date' ? '2027-03-14' : !record.control.checked); expected = P.controls(document).records.get(key).text; }
  });
  assert(P.controls(document).records.get(key).complete, 'Edit left an incomplete boundary');
  assert(!document.losses?.length, 'Value edit unexpectedly converted a control');
  const edited = await save('edited', document), reopened = await read(edited);
  assert(records(reopened).find(r => r.control.type === type)?.text === expected, 'The stored binding reverted the edited text on reopen');
  if (type === 'text') {
    assert(document.keep.boundUpdates && Object.values(document.keep.boundUpdates).some(v => v.value === expected), 'Missing XML store update');
    const copied = O.copyRange(record.start.pos, P.controls(document).records.get(key).end.pos);
    const imported = K.remapSources(copied, K.import(P.clipboard(copied)));
    const destination = D.newDoc(); L.app.loadDoc(destination, 'Pasted control', { saved: true });
    D.tx('Paste', () => O.insertBlocks(D.pos(destination.main.blocks[0], 0), imported));
    const copiedFile = await save('copied', destination), reread = await read(copiedFile);
    assert(records(reread).length === 1, 'Copy lost the control');
    assert(records(reread)[0].text === expected, 'Copy lost its edited XML store value');
    // Return to the source window without replacing its undo history.
    L.app.loadDoc(document, 'Controls', { saved: true });
    D.tx('Type again', () => O.insertText(D.pos(record.start.pos.p, record.start.pos.o + 1), '!'));
    D.undo();
    assert(P.controls(document).records.get(key).text === expected, 'Undo did not restore bound text');
  } else {
    assert(D.undo(), 'Missing undo transaction');
    record = P.controls(document).records.get(key);
    assert(record.text === beforeText && record.control.shell.xml === beforeXML, 'Undo did not restore original control');
    await save('undo', document);
    assert(D.redo(), 'Missing redo transaction');
  }
  const draft = await save('draft', document), recovered = await read(draft), final = await save('recovered', recovered);
  assert(records(await read(final)).find(r => r.control.type === type)?.text === expected, 'Draft recovery lost the control value');
  return { status: 'ok', checks, artifacts };
}
