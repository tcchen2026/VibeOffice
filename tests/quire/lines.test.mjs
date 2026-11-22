// Quire's lines against Word's: each line as tall as what is on it (public/quire/js/portions.js), checked on a real
// document Word laid out. The reference is Word's PDF of a Word-saved walkthrough (19 pages; tools/wordlines.py made
// lines/walkthrough-word.json from it); the document itself stays out of the repository (its text is not ours to
// publish): put it in /work/corpora/word-ref/fluent-walkthrough.docx, or name its folder in QUIRE_WORD_REF, else this
// part is skipped. Pages must break where Word's do, and between two lines that follow each other in both, Quire's
// distance must be Word's within BOUND points (a line too tall or too short shows as a step; tests/lib/lines.mjs).
//   node --test tests/quire/lines.test.mjs        (REPORT=1 prints each page's steps)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';
import { ROWS, match, pagesAndSteps } from '../lib/lines.mjs';

const BOUND = 0.75;
const refDir = process.env.QUIRE_WORD_REF || '/work/corpora/word-ref';
const walkthrough = path.join(refDir, 'fluent-walkthrough.docx');

/** open a .docx in Quire and return its rows (tests/lib/lines.mjs) */
async function quireRows(file) {
  const page = await openPage('quire', { persistence: false, automated: true });
  try {
    const b64 = fs.readFileSync(file).toString('base64');
    const rows = await page.evaluate(`(async () => {
      VO.opened = () => {}; L.layout.view = 'print'; L.layout.setZoom(1);
      const bytes = Uint8Array.from(atob(${JSON.stringify(b64)}), (c) => c.charCodeAt(0));
      const { doc } = await L.docx.read(bytes.buffer);
      L.app.loadDoc(doc, 'doc', { saved: true });
      await document.fonts.ready; L.layout.render(); await new Promise((r) => setTimeout(r, 500));
      return ${ROWS};
    })()`);
    return { rows, errors: page.errors };
  } finally { await page.close(); }
}

test('The walkthrough: pages break where Word\'s do, and its lines sit where Word puts them', { skip: !fs.existsSync(walkthrough) && `no ${walkthrough}` }, async () => {
  const ref = JSON.parse(fs.readFileSync(new URL('lines/walkthrough-word.json', import.meta.url), 'utf8'));
  const stop = await ensureServer();
  try {
    const { rows, errors } = await quireRows(walkthrough);
    const pages = pagesAndSteps(ref, match(ref, rows));
    if (process.env.REPORT) for (const p of pages) console.log(JSON.stringify(p));
    assert.equal(rows.pages.length, ref.pages.length, 'as many pages as Word');
    for (const p of pages) if (p.matched) assert.equal(p.startsOn, p.page, `page ${p.page} starts with the line Word's starts with`);
    for (const p of pages) for (const s of p.steps) assert.ok(Math.abs(s.step) <= BOUND, `page ${p.page} row ${s.row}: ${s.step}pt farther from Word's than the line above it`);
    assert.deepEqual(errors, []);
  } finally { stop(); }
});

test('Line cases: a line is as tall as what it holds, as LibreOffice (and Word) draw it', async () => {
  const lo = JSON.parse(fs.readFileSync(new URL('lines/linecases-lo.json', import.meta.url), 'utf8')).cases;
  const stop = await ensureServer();
  try {
    const { rows, errors } = await quireRows(new URL('lines/linecases.docx', import.meta.url).pathname);
    const tops = {};
    rows.pages.forEach((page, p) => page.forEach((r) => { const m = /^M(\d+)$/.exec(r[4]); if (m) tops[m[1]] = [p, r[2]]; }));
    const got = lo.map((c, i) => ({ ...c, quire: tops[i] && tops[i + 1] && tops[i][0] === tops[i + 1][0] ? Math.round((tops[i + 1][1] - tops[i][1]) * 100) / 100 : null }));
    if (process.env.REPORT) for (const c of got) console.log(`${c.asWord ? ' ' : '?'} ${String(c.quire).padStart(6)} ${String(c.pt).padStart(6)}  ${c.case}`);
    for (const c of got.filter((c) => c.asWord)) assert.ok(c.quire != null && Math.abs(c.quire - c.pt) <= 0.6, `${c.case}: ${c.quire}pt to the next marker, LibreOffice ${c.pt}pt`);
    assert.deepEqual(errors, []);
  } finally { stop(); }
});
