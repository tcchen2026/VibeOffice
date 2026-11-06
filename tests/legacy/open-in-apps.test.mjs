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
  const r = await open('quire', doc, 'm_maciver.doc', `({ name: L.app.fileName, saved: L.app.saved, text: document.querySelector('#pages').textContent })`);
  assert.equal(r.name, 'm_maciver');
  assert.equal(r.saved, false, 'Save asks for a .docx name');
  assert.match(r.text, /BIOGRAPHICAL SKETCH/);
  assert.match(r.msgs.join('\n'), /converted from the Word 97-2003 document format[\s\S]*Tables were replaced/);
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
  assert.match(r.msgs.join('\n'), /Excel 97-2003 workbook, which Quire does not open/);
});
