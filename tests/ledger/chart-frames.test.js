// What a save keeps of chart sheets and chart frames that Ledger doesn't edit (I24 C4 findings: a converted .xls lost its
// chart sheets' footers and its charts' grouping locks). Before the fix every chart sheet in the corpus lost its header
// and footer text (14), code name (48) and protection (8), and every chart frame its locks, creation ID, title, alt text
// and hyperlink. node --test tests/ledger/chart-frames.test.js
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path');
const L = require('./load.js');
const K = L.opc;
const corpus = process.env.EXCEL_CORPUS || '/work/corpora/excel';

async function roundTrip(t, name) {
  const input = path.join(corpus, name);
  if (!fs.existsSync(input)) { t.skip('Fetch pinned Excel corpus'); return null; }
  const bytes = fs.readFileSync(input);
  const wb = await L.xlsxRead.read(bytes);
  return { before: await K.open(bytes), after: await K.open(await L.xlsxWrite.bytes(wb)) };
}
const all = (el, name) => Array.from(el.getElementsByTagName('*')).filter((e) => e.localName === name);
const own = (e) => (e ? Object.fromEntries(Array.from(e.attributes).filter((a) => a.prefix !== 'xmlns' && a.name !== 'xmlns').map((a) => [a.name, a.value])) : null);
const pick = (o, keys) => (o ? Object.fromEntries(keys.filter((k) => o[k] != null).map((k) => [k, o[k]])) : null);
/** each chart sheet by name: its header and footer texts, sheetPr's code name and published flag, its protection */
function chartSheets(pkg) {
  const book = pkg.xml('xl/workbook.xml'), rels = pkg.rels('xl/workbook.xml'), out = {};
  for (const sh of all(book, 'sheet')) {
    const rel = rels.find((r) => r.id === sh.getAttribute('r:id'));
    if (!rel || !/chartsheet$/.test(rel.type)) continue;
    const doc = pkg.xml(rel.part);
    const hf = all(doc, 'headerFooter')[0];
    out[sh.getAttribute('name')] = {
      text: hf ? Array.from(hf.children).map((e) => e.localName + ':' + e.textContent).filter((x) => !/:$/.test(x)) : [],
      sheetPr: pick(own(all(doc, 'sheetPr')[0]), ['codeName', 'published']),
      protection: own(all(doc, 'sheetProtection')[0]),
    };
  }
  return out;
}
/** each chart frame by name: its macro, title, alt text, locks, creation ID and whether it links somewhere */
function chartFrames(pkg) {
  const out = {};
  for (const part of pkg.names.filter((p) => /^xl\/drawings\/drawing\d+\.xml$/.test(p))) {
    for (const fr of all(pkg.xml(part), 'graphicFrame')) {
      if (!all(fr, 'chart').length) continue;
      const nv = all(fr, 'cNvPr')[0], cid = all(nv, 'creationId')[0];
      out[nv.getAttribute('name')] = { macro: fr.getAttribute('macro'), cNvPr: pick(own(nv), ['title', 'descr', 'hidden']), locks: own(all(fr, 'graphicFrameLocks')[0]),
        creationId: cid ? cid.getAttribute('id') : null, link: all(nv, 'hlinkClick').length > 0 };
    }
  }
  return out;
}

test('chart sheets keep their headers and footers, code names and protection', async (t) => {
  const r = await roundTrip(t, 'poi__68aaa05ba444__57181.xlsm');
  if (!r) return;
  const before = chartSheets(r.before);
  assert.ok(Object.values(before).some((s) => s.text.length && s.sheetPr.codeName && s.protection), 'the workbook has what is checked');
  assert.deepEqual(chartSheets(r.after), before);
});

test('chart frames keep their locks, creation IDs, titles and hyperlinks', async (t) => {
  for (const name of ['poi__68aaa05ba444__57181.xlsm', 'libreoffice__aa146114c1fc__chart_hyperlink.xlsx', 'poi__49132d1375df__simple-monthly-budget.xlsx']) {
    const r = await roundTrip(t, name);
    if (!r) return;
    const before = chartFrames(r.before), after = chartFrames(r.after);
    for (const f of Object.values(before)) if (f.locks && !Object.keys(f.locks).length) f.locks = null;   // an empty lock element says nothing
    for (const f of Object.values(after)) if (f.locks && !Object.keys(f.locks).length) f.locks = null;
    assert.deepEqual(after, before, name);
  }
});
