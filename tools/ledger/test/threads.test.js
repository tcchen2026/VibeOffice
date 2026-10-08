/* Complete threads, their legacy links and persons survive edits and copies. */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} };
require('../../../public/ledger/js/ops.js');
const K = L.opc, O = L.ops;
const corpus = process.env.VO_CORPORA || path.join(os.homedir(), 'corpora');
const file = path.join(corpus, 'excel/libreoffice__accaee5c7a5a__threadedComment.xlsx');
const part = 'xl/threadedComments/threadedComment1.xml', person = 'xl/persons/person.xml';
const rootId = '{FB8EA27E-C1B6-4481-B034-193455CC7425}';
const output = process.env.THREAD_RESULTS, results = [];
const all = pkg => pkg.names.filter(n => /\/threadedComments\/.*\.xml$/.test(n)).flatMap(n => pkg.xml(n).children.filter(e => e.localName === 'threadedComment'));
const cm = wb => Array.from(wb.sheets[0].comments.values()).find(c => c.thread?.length);
async function open(t, bytes) {
  if (!fs.existsSync(file)) { t.skip('Fetch pinned corpus: ' + file); return; }
  const wb = await L.xlsxRead.read(bytes || fs.readFileSync(file)); return wb;
}
async function save(wb, scenario, original = file) {
  const bytes = await L.xlsxWrite.bytes(wb), pkg = await K.open(bytes);
  if (output) {
    const dir = path.join(output, scenario); fs.mkdirSync(dir, { recursive: true });
    const name = path.basename(original); fs.writeFileSync(path.join(dir, name), bytes);
    results.push({ original, file: name, scenario, status: 'ok' });
    fs.writeFileSync(path.join(output, 'results.jsonl'), results.map(r => JSON.stringify(r)).join('\n') + '\n');
  }
  return { bytes, pkg, losses: bytes.dropped };
}
test('all pinned threaded comment parts and persons keep their original bytes', async t => {
  for (const name of ['poi__880073bc86bc__LIBRE_OFFICE-128382-0.xlsx', 'poi__f8b0c349aff3__right-to-left.xlsx', 'poi__f58f64c0e194__64759.xlsx', path.basename(file)]) {
    const source = path.join(corpus, 'excel', name);
    if (!fs.existsSync(source)) { t.skip('Fetch pinned corpus'); return; }
    const wb = await L.xlsxRead.read(fs.readFileSync(source)), s = await save(wb, 'save', source);
    for (const n of wb.pkg.names.filter(n => /\/(threadedComments|persons)\/.*\.xml$/.test(n))) assert.deepEqual(s.pkg.bytes(n), wb.pkg.bytes(n), n);
    assert.equal(all(s.pkg).length, wb.sheets.reduce((n, sh) => n + [...sh.comments.values()].reduce((m, cm) => m + (cm.thread?.length || 0), 0), 0));
  }
});
test('row shifts and deletion move every reply and keep legacy links; undo restores raw threads', async t => {
  const wb = await open(t); if (!wb) return;
  const sh = wb.sheets[0]; O.insertLines(sh, 'r', 0, 2);
  let s = await save(wb, 'rows');
  assert.deepEqual(all(s.pkg).map(e => e.getAttribute('ref')), ['A4', 'A4']);
  const note = s.pkg.xml('xl/comments1.xml').getElementsByTagName('comment')[0];
  assert.equal(note.getAttribute('ref'), 'A4'); assert.equal(note.getAttributeNS('http://schemas.microsoft.com/office/spreadsheetml/2014/revision', 'uid'), rootId);
  assert.equal(all(s.pkg)[1].getAttribute('parentId'), rootId);
  wb.undo.undo(); s = await save(wb, 'rows-undo'); assert.deepEqual(s.pkg.bytes(part), wb.pkg.bytes(part));
  wb.undo.redo(); s = await save(wb, 'rows-redo'); assert.equal(all(s.pkg)[1].getAttribute('ref'), 'A4');
  O.insertLines(sh, 'r', 3, -1); s = await save(wb, 'delete'); assert.equal(all(s.pkg).length, 0);
  wb.undo.undo(); s = await save(wb, 'delete-undo'); assert.equal(all(s.pkg).length, 2);
});
test('editing a displayed note converts its thread with an undoable notice', async t => {
  const wb = await open(t); if (!wb) return;
  const c = cm(wb), sh = wb.sheets[0];
  O.tx(wb, 'Edit Comment', () => O.setComment(sh, c.r, c.c, { ...c, text: 'Changed note', runs: undefined }));
  let s = await save(wb, 'note-edit'); assert.equal(all(s.pkg).length, 0);
  assert(s.losses.some(e => e.id.startsWith('thread-conversion:')));
  const reread = await L.xlsxRead.read(s.bytes); assert.equal([...reread.sheets[0].comments.values()][0].text, 'Changed note');
  assert.equal([...reread.sheets[0].comments.values()][0].author, c.thread[0].author);
  wb.undo.undo(); s = await save(wb, 'note-edit-undo'); assert.deepEqual(s.pkg.bytes(part), wb.pkg.bytes(part));
  assert(!s.losses.some(e => e.id.startsWith('thread-conversion:')));
});
test('copied sheets remap comment definitions, parent links and legacy author identities together', async t => {
  const wb = await open(t); if (!wb) return;
  O.copySheet(wb, wb.sheets[0], wb.sheets.length);
  let s = await save(wb, 'sheet-copy'), threads = all(s.pkg);
  assert.equal(threads.length, 4); assert.equal(new Set(threads.map(t => t.getAttribute('id'))).size, 4);
  const reread = await L.xlsxRead.read(s.bytes), copy = [...reread.sheets[1].comments.values()].find(c => c.thread);
  assert.equal(copy.thread[1].parent, copy.thread[0].id); assert.notEqual(copy.thread[0].id, rootId);
  assert.equal(copy.author, 'tc=' + copy.thread[0].id);
  const ids = s.pkg.names.filter(n => /^xl\/comments\d+\.xml$/.test(n)).flatMap(n => s.pkg.xml(n).getElementsByTagName('comment').map(e => e.getAttributeNS('http://schemas.microsoft.com/office/spreadsheetml/2014/revision', 'uid')));
  assert.equal(new Set(ids).size, ids.length, 'Ordinary copied notes also have fresh revision IDs');
  assert.deepEqual(s.pkg.bytes(person), wb.pkg.bytes(person));
  wb.undo.undo(); s = await save(wb, 'sheet-copy-undo'); assert.equal(all(s.pkg).length, 2);
});

async function mentionFixture() {
  const zip = await L.zip.read(fs.readFileSync(file)), files = [];
  const ns = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
  const personId = '{00000000-0000-4000-8000-000000000001}', mentionId = '{00000000-0000-4000-8000-000000000002}';
  for (const [name, entry] of zip) {
    if (name === 'xl/comments1.xml' || name === 'xl/drawings/vmlDrawing1.vml') continue;
    let data = await entry.bytes();
    if (name === part) {
      const xml = new TextDecoder().decode(data), root = K.parse(xml), first = root.children[0];
      const kept = K.mergeBag(K.raw(first), { ['{' + ns + '}text']: '<text xmlns="' + ns + '">Hello Ada</text>', __append: '<mentions xmlns="' + ns + '"><mention mentionpersonId="' + personId + '" mentionId="' + mentionId + '" startIndex="6" length="3"/></mentions>' });
      const pos = L.xmlTree.source.get(first); data = K.patch(xml, [{ start: pos.start, end: pos.end, value: kept }]);
    } else if (name === person) data = K.mergeBag(new TextDecoder().decode(data), { __append: '<person xmlns="' + ns + '" displayName="Ada" id="' + personId + '" userId="ada@example.invalid" providerId="None"/>' });
    else if (name === 'xl/worksheets/sheet1.xml') data = K.mergeBag(new TextDecoder().decode(data), { ['{' + K.NS.s + '}legacyDrawing']: [] });
    else if (name === 'xl/worksheets/_rels/sheet1.xml.rels') data = K.mergeBag(new TextDecoder().decode(data), { ['{' + K.NS.pkg + '}Relationship']: K.parse(new TextDecoder().decode(data)).children.filter(e => !/\/(comments|vmlDrawing)$/.test(e.getAttribute('Type'))).map(K.raw) });
    else if (name === '[Content_Types].xml') data = new TextDecoder().decode(data).replace(/<Override\b[^>]*PartName="\/xl\/comments1.xml"[^>]*\/>/, '');
    files.push({ name, data });
  }
  const bytes = new Uint8Array(await (await L.zip.write(files)).arrayBuffer());
  let source = file;
  if (output) { source = path.join(output, 'mentions-no-legacy.xlsx'); fs.writeFileSync(source, bytes); }
  return { bytes, source, personId, mentionId };
}
test('threads without legacy notes load; cross-workbook paste retains mentions and adds their persons', async t => {
  if (!fs.existsSync(file)) { t.skip('Fetch pinned corpus'); return; }
  const f = await mentionFixture(), wb = await open(t, f.bytes), c = cm(wb);
  assert(c); assert.equal(c.thread.length, 2); assert.equal(c.text, 'Hello Ada\n\nA reply');
  let s = await save(wb, 'no-legacy', f.source); assert.deepEqual(s.pkg.bytes(part), wb.pkg.bytes(part));
  const dest = new L.model.Workbook(); dest.addSheet('Copied');
  const copied = L.threads.copy(c, dest);
  O.tx(dest, 'Paste Comments', () => O.setComment(dest.sheets[0], 7, 3, copied));
  s = await save(dest, 'cross-workbook', f.source);
  let threads = all(s.pkg); assert.equal(threads.length, 2);
  assert(threads.every(e => e.getAttribute('ref') === 'D8'));
  const mention = threads[0].getElementsByTagName('mention')[0];
  assert.equal(mention.getAttribute('mentionpersonId'), f.personId); assert.notEqual(mention.getAttribute('mentionId'), f.mentionId);
  assert.equal(mention.getAttribute('startIndex'), '6'); assert.equal(mention.getAttribute('length'), '3');
  const persons = s.pkg.names.filter(n => /\/persons\//.test(n)).flatMap(n => s.pkg.xml(n).children.map(e => e.getAttribute('id')));
  assert(persons.includes(f.personId)); assert(persons.includes(threads[0].getAttribute('personId')));
  const reopened = await L.xlsxRead.read(s.bytes); assert.equal(cm(reopened).thread[1].parent, cm(reopened).thread[0].id);
  O.tx(wb, 'Copy Comment', () => O.setComment(wb.sheets[0], 8, 4, L.threads.copy(c, wb)));
  s = await save(wb, 'same-workbook', f.source); threads = all(s.pkg);
  assert.equal(new Set(threads.map(e => e.getAttribute('id'))).size, 4);
  assert.equal(new Set(threads.flatMap(e => e.getElementsByTagName('mention').map(m => m.getAttribute('mentionId')))).size, 2);
});
