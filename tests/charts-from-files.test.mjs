// Charts read from files, in all three apps, through the shared reader and drawing (common/chart-read.js, chart-draw.js;
// Ledger and Quire drew them with the simple model before: one value axis, no markers, blanks as zero, the longest
// category range). Checked against Excel, Word and PowerPoint 2016's prints of the same files (2026-11-07).
// node --test tests/charts-from-files.test.mjs
import { after, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { ensureServer, openPage } from '../tools/ooxml/cdp.mjs';

let stop;
after(() => stop?.());
const CORPORA = '/work/corpora';

// opens a corpus file in an app and returns probe(charts) for the charts it read, in reading order
async function charts(t, app, file, probe) {
  const path = `${CORPORA}/${file}`;
  if (!fs.existsSync(path)) { t.skip('Fetch the pinned corpora'); return null; }
  stop = stop || await ensureServer(process.env.BASE);
  const page = await openPage(app, { persistence: false, automated: true });
  try {
    const b64 = fs.readFileSync(path).toString('base64');
    const r = await page.evaluate(`(async () => {
      await L.app.openFile(new File([Uint8Array.from(atob(${JSON.stringify(b64)}), (c) => c.charCodeAt(0))], ${JSON.stringify(file.split('/').pop())}));
      await new Promise((r) => setTimeout(r, 400));
      const list = [];
      if (${JSON.stringify(app)} === 'ledger') { for (const sh of L.grid.wb.sheets) for (const d of sh.drawings) if (d.kind === 'chart' && d.chart) list.push({ chart: d.chart, svg: () => L.xchart.render(d.chart, L.grid.wb, sh, 480, 300, 1) }); }
      else if (${JSON.stringify(app)} === 'quire') { for (const p of L.D.allParas(L.D.doc, { all: true })) for (const it of p.runs) if (it.t === 'chart' && it.chart) list.push({ chart: it.chart, svg: () => L.chart.render(it.chart, it.w, it.h, null) }); }
      else for (const sl of L.pres.slides) for (const sh of sl.shapes) if (sh.chart) list.push({ chart: sh.chart, svg: () => L.chart.render(sh.chart, sh.w, sh.h, L.model.design(L.pres, sl)) });
      return (${probe})(list);
    })()`);
    assert.deepEqual(page.errors, []);
    return r;
  } finally { await page.close(); }
}

test('Ledger reads a chart with a secondary axis as Excel shows it', async (t) => {
  // Excel 2016: columns on the left axis, the line on its own right-hand axis
  const r = await charts(t, 'ledger', 'excel/poi__cbc221851f0e__57362.xlsx', `(list) => list.map(({ chart: c, svg }) => ({ v: c.v, sec: c.groups.some((g) => g.sec), right: !!(c.ax.v2 && !c.ax.v2.del), drawn: svg().querySelectorAll('text').length > 10 }))`);
  if (r) assert.deepEqual(r, [{ v: 2, sec: true, right: true, drawn: true }]);
});

test('a deleted secondary axis leaves its series on the primary scale, and the axis shows its own categories', async (t) => {
  // Word 2016 (tdf138773): the line scaled with the columns; 13 categories from 2000-01, not the line's 15-point range
  const r = await charts(t, 'quire', 'word/libreoffice__ce69c730ae7d__tdf138773.docx', `(list) => list.map(({ chart: c }) => ({ sec: c.groups.some((g) => g.sec), n: c.cats.length, first: c.cats[0], border: !!c.border }))`);
  if (r) assert.deepEqual(r, [{ sec: false, n: 13, first: '2000-01', border: false }]);
});

test('a chart area without formatting is white with a grey border in Word, and percentages are formatted', async (t) => {
  // Word 2016 (the SDK's chart.docx): #868686 borders around every chart; the pie's labels 11%, 28%, 30%, 18%, 13%
  const r = await charts(t, 'quire', 'word/openxml-sdk__31493658a65a__chart.docx', `(list) => ({
    borders: list.every(({ chart: c }) => c.border && c.border.c === '#868686'),
    pie: list.filter(({ chart: c }) => c.kind === 'pie').map(({ svg }) => [...svg().querySelectorAll('text')].map((e) => e.textContent).filter((x) => /%$/.test(x)).sort()).flat() })`);
  if (r) assert.deepEqual(r, { borders: true, pie: ['11%', '13%', '18%', '28%', '30%'] });
});

test('a title without text is "Chart Title" unless the file names the one series', async (t) => {
  // PowerPoint 2016 (chart-theme-override): its default title, not "Series 1"
  const r = await charts(t, 'lectern', 'powerpoint/libreoffice__46fc8495fece__chart-theme-override.pptx', `(list) => list.map(({ chart: c }) => c.title)`);
  if (r) assert.deepEqual(r, ['Chart Title']);
});
