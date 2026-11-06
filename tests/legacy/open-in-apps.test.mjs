// An Office 97-2003 file opens in its app through the app's own Open: converted (common/legacy/import.js) to the
// OOXML package Office would write, unsaved so that Save asks for a new name, and what wasn't converted is told.
// These files used to be refused. node --test tests/legacy/open-in-apps.test.mjs
import { after, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const corpus = process.env.VO_CORPORA || '/work/corpora';
const file = (dir, name) => path.join(corpus, dir, name);
let stop;
after(() => stop?.());

async function open(app, src, name, probe) {
  stop = stop || await ensureServer(process.env.BASE);
  const page = await openPage(app, { persistence: false, automated: true });
  try {
    const b64 = fs.readFileSync(src).toString('base64');
    return await page.evaluate(`(async () => {
      const msgs = []; L.ui.msg = (text) => { msgs.push(String(text)); return Promise.resolve(); };
      const b = Uint8Array.from(atob(${JSON.stringify(b64)}), (c) => c.charCodeAt(0));
      await L.app.openFile(new File([b], ${JSON.stringify(name)}));
      return Object.assign(${probe}, { msgs });
    })()`);
  } finally { await page.close(); }
}

const doc = file('legacy-word', 'poi__e38d8be7d96c__cap.stanford.edu_profiles_viewbiosketch_facultyid=4009&name=m_maciver.doc');
test('Quire opens a Word 97-2003 document', { skip: !fs.existsSync(doc) && 'no corpus' }, async () => {
  const r = await open('quire', doc, 'm_maciver.doc', `({ name: L.app.fileName, saved: L.app.saved, text: document.querySelector('#pages').textContent, tables: document.querySelectorAll('#pages table').length })`);
  assert.equal(r.name, 'm_maciver');
  assert.equal(r.saved, false, 'Save asks for a .docx name');
  assert.match(r.text, /BIOGRAPHICAL SKETCH/);
  assert.ok(r.tables > 0, 'its tables are tables (I24 E2)');
  assert.doesNotMatch(r.msgs.join('\n'), /Tables were replaced/);
});

const xls = file('legacy-excel', 'poi__193e60ac0218__15228.xls');
test('Ledger opens an Excel 97-2003 workbook, formulas and all', { skip: !fs.existsSync(xls) && 'no corpus' }, async () => {
  const r = await open('ledger', xls, '15228.xls', `(() => { const b = L.app.books[L.app.cur]; return { type: b.type, sheets: L.grid.wb.sheets.length, f: (L.grid.wb.sheets[0].get(18, 4) || {}).f }; })()`);
  assert.equal(r.type, 'xls', 'Save asks for an .xlsx name');
  assert.equal(r.sheets, 18);
  assert.equal(r.f, 'YEAR(RunDat)*12+MONTH(RunDat)-YEAR(ValDat)*12-MONTH(ValDat)');
  assert.match(r.msgs.join('\n'), /converted from the Excel 97-2003 workbook format/);
});

const ppt = file('legacy-powerpoint', 'poi__50d56a23bc3e__slide_master.ppt');
test('Lectern opens a PowerPoint 97-2003 presentation', { skip: !fs.existsSync(ppt) && 'no corpus' }, async () => {
  const r = await open('lectern', ppt, 'slide_master.ppt', `({ name: L.app.fileName, saved: L.app.saved, slides: L.pres.slides.length, title: L.pres.slides[0].shapes.map((s) => s.tx ? s.tx.ps.map((p) => p.rs.map((x) => x.t).join('')).join(' ') : '').join(' ') })`);
  assert.equal(r.name, 'slide_master');
  assert.equal(r.saved, false, 'Save asks for a .pptx name');
  assert.equal(r.slides, 3);
  assert.match(r.title, /Slide Masters/);
});

test('a file of another kind is refused with what it is', { skip: !fs.existsSync(xls) && 'no corpus' }, async () => {
  const r = await open('quire', xls, 'book.doc', `({ name: L.app.fileName })`);
  assert.match(r.msgs.join('\n'), /This is an Excel 97-2003 workbook, which Quire does not open/);
});

// A file with a password to open asks for it through the app's Password box, as Office does, and keeps it for Save
// (Office keeps a password to open on Save As .docx/.xlsx/.pptx); a file encrypted with Office's default password
// opens without asking and is saved without one. Passwords: Apache POI's tests (tests/legacy/encrypted-fixtures.mjs).
async function openWith(app, src, name, answers, probe) {
  stop = stop || await ensureServer(process.env.BASE);
  const page = await openPage(app, { persistence: false, automated: true });
  try {
    const b64 = fs.readFileSync(src).toString('base64');
    const r = await page.evaluate(`(async () => {
      const msgs = [], asked = [], answers = ${JSON.stringify(answers)};
      L.ui.msg = (text) => { msgs.push(String(text)); return Promise.resolve(); };
      L.ui.password = (file, wrong) => { asked.push(wrong ? 'again' : 'ask'); return Promise.resolve(answers.length ? answers.shift() : null); };
      const b = Uint8Array.from(atob(${JSON.stringify(b64)}), (c) => c.charCodeAt(0));
      await L.app.openFile(new File([b], ${JSON.stringify(name)}));
      return Object.assign(${probe}, { msgs, asked });
    })()`);
    assert.deepEqual(page.errors.filter((e) => !/cancelled|password is incorrect/.test(e)), []);   // the apps log what they refused
    return r;
  } finally { await page.close(); }
}
const encXls = file('legacy-excel', 'poi__47d331bcc24c__password.xls'), defaultXls = file('legacy-excel', 'poi__dae995b60ef4__50833.xls');
const encDoc = file('legacy-word', 'poi__7878f791ecc8__password_tika_binaryrc4.doc'), encPpt = file('legacy-powerpoint', 'poi__3303dbfa2b1f__Password_Protected-hello.ppt');
const ledgerState = `(() => { const b = L.app.books[L.app.cur]; const s = L.grid.wb.sheets[0]; return { password: b.password || null, type: b.type, text: [...s.rows.values ? s.rows.values() : s.rows].flatMap((row) => row ? Object.values(row.cells || {}).map((c) => c && c.v) : []).join(' ') }; })()`;

test('Ledger asks for the password to open an Excel 97-2003 workbook and keeps it for Save', { skip: !fs.existsSync(encXls) && 'no corpus' }, async () => {
  const r = await openWith('ledger', encXls, 'password.xls', ['password'], ledgerState);
  assert.deepEqual(r.asked, ['ask']);
  assert.equal(r.password, 'password');
  assert.equal(r.type, 'xls');
  assert.match(r.text, /A ZIP bomb is a variant/);
  const wrong = await openWith('ledger', encXls, 'password.xls', ['nope'], `({ books: L.app.books.filter((b) => /password/.test(b.name)).length })`);
  assert.equal(wrong.books, 0, 'a wrong password cancels the open, as in Excel');
  assert.match(wrong.msgs.join('\n'), /password you supplied is not correct/);
  const cancel = await openWith('ledger', encXls, 'password.xls', [], `({ books: L.app.books.filter((b) => /password/.test(b.name)).length })`);
  assert.deepEqual([cancel.books, cancel.msgs], [0, []]);
});

test('a workbook encrypted with Office\'s default password opens without asking and is saved without one', { skip: !fs.existsSync(defaultXls) && 'no corpus' }, async () => {
  const r = await openWith('ledger', defaultXls, '50833.xls', [], ledgerState);
  assert.deepEqual(r.asked, []);
  assert.equal(r.password, null);
  assert.match(r.text, /test cell value/);
});

test('Quire asks for the password to open a Word 97-2003 document and keeps it for Save', { skip: !fs.existsSync(encDoc) && 'no corpus' }, async () => {
  const state = `({ name: L.app.fileName, password: L.D.doc.openPassword || null, text: document.querySelector('#pages').textContent })`;
  const r = await openWith('quire', encDoc, 'tika.doc', ['tika'], state);
  assert.deepEqual(r.asked, ['ask']);
  assert.equal(r.password, 'tika');
  assert.match(r.text, /This is an encrypted Word 2007 File\./);
  const wrong = await openWith('quire', encDoc, 'tika.doc', ['nope'], state);
  assert.notEqual(wrong.name, 'tika');
  assert.match(wrong.msgs.join('\n'), /The password is incorrect\. Quire cannot open the document/);
});

test('Lectern asks again after a wrong password to open a PowerPoint 97-2003 presentation and keeps the right one', { skip: !fs.existsSync(encPpt) && 'no corpus' }, async () => {
  const state = `({ name: L.app.fileName, password: L.pres.password || null, text: L.pres.slides.flatMap((sl) => sl.shapes).map((s) => s.tx ? s.tx.ps.map((p) => p.rs.map((x) => x.t).join('')).join(' ') : '').join(' ') })`;
  const r = await openWith('lectern', encPpt, 'hello.ppt', ['nope', 'hello'], state);
  assert.deepEqual(r.asked, ['ask', 'again']);
  assert.equal(r.password, 'hello');
  assert.match(r.text, /This is a protected document/);
  const cancel = await openWith('lectern', encPpt, 'hello.ppt', [], state);
  assert.notEqual(cancel.name, 'hello');
  assert.deepEqual(cancel.msgs, []);
});
