// Issue #22: merged cells keep their perimeter cells' border segments, not just the anchor's style.
// Excel.Test.xlsx is the issue attachment; merge-cases.xlsx contains four small independent controls;
// merge-2003.xml is a boxed merge as XML Spreadsheet 2003 stores it, on the top-left cell only.
// Expectations were checked in Excel 2016 through the win10 API (office2016.json), including partial
// edges and a merge without an anchor cell. These are canvas checks, with no mouse/keyboard workflow.
//   node --test tests/from-issues/22/borders.test.mjs
//   node tools/older.mjs HEAD -- tests/from-issues/22/borders.test.mjs
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { ensureServer, openPage } from '../../../tools/ooxml/cdp.mjs';

let page, stop;
before(async () => {
  stop = await ensureServer();
  page = await openPage('ledger', { persistence: false, automated: true });
});
after(async () => { try { await page?.close(); } finally { stop?.(); } });

async function check(file, segments, views = [{}], setup = '') {
  const bytes = fs.readFileSync(new URL(file, import.meta.url)).toString('base64');
  const got = await page.evaluate(`(async () => {
    const wb = await L[${JSON.stringify(file.endsWith('.xml') ? 'xmlss' : 'xlsxRead')}].read(Uint8Array.from(atob(${JSON.stringify(bytes)}), c => c.charCodeAt(0)));
    const sh = wb.sheets[0], RD = L.render;
    ${setup}
    const snapshot = () => JSON.stringify([wb.styles.list, sh.rows, sh.cols, sh.merges, sh.cf]);
    const before = snapshot(), out = [];
    for (const opts of ${JSON.stringify(views)}) {
      sh.view.freeze = opts.freeze || null;
      const canvas = document.createElement('canvas');
      const view = { sh, zoom: 1, scrollR: 0, scrollC: 0, w: 3000, h: 700, dpr: 1, ...opts };
      const fr = RD.draw(canvas, view, { hideSel: true });
      const ctx = canvas.getContext('2d'), samples = [];
      for (const [r, c, side, color] of ${JSON.stringify(segments)}) {
        const p = RD.paneOf(fr, r, c);
        if (!p || r > p.rEnd || c > p.cEnd) continue;
        const x1 = RD.colX(fr, p, c), x2 = RD.colX(fr, p, c + 1);
        const y1 = RD.rowY(fr, p, r), y2 = RD.rowY(fr, p, r + 1);
        const x = side === 'l' ? x1 - 0.5 : side === 'r' ? x2 - 0.5 : (x1 + x2) / 2;
        const y = side === 't' ? y1 - 0.5 : side === 'b' ? y2 - 0.5 : (y1 + y2) / 2;
        if (x < p.x || x >= p.x + p.w || y < p.y || y >= p.y + p.h) continue;
        const rgb = [...ctx.getImageData(Math.floor(x * view.dpr), Math.floor(y * view.dpr), 1, 1).data].slice(0, 3);
        samples.push({ cell: L.formula.cellName(r, c), side, rgb, color });
      }
      out.push({ opts, samples });
    }
    return { out, unchanged: before === snapshot() };
  })()`);
  assert.equal(got.unchanged, true, 'painting must not rewrite the stored styles or perimeter cells');
  for (const { opts, samples } of got.out) {
    assert.ok(samples.length, 'each view must check visible borders');
    for (const { cell, side, rgb, color } of samples) assert.deepEqual(rgb, color,
      `${file}: ${cell} ${side}, view ${JSON.stringify(opts)}`);
  }
  assert.deepEqual(page.errors, []);
  assert.deepEqual(page.requests, []);
  return got.out.reduce((n, v) => n + v.samples.length, 0);
}

const black = [0, 0, 0], white = [255, 255, 255];

test('Issue 22: the reported workbook keeps its table outlines at every zoom, when scrolled and in frozen panes', async () => {
  const segments = [];
  // Excel's first table has six horizontal rules and four right edges, including the missing Y edge.
  for (let r = 0; r < 6; r++) {
    for (let c = 0; c < 25; c++) segments.push([r, c, 'b', black]);
    for (const c of [4, 12, 18, 24]) segments.push([r, c, 'r', black]);
  }
  // The second table: header underline, the missing row-13 bottom, and six vertical dividers/outlines.
  for (let c = 0; c < 25; c++) for (const r of [9, 12]) segments.push([r, c, 'b', black]);
  for (let r = 8; r < 13; r++) for (const c of [4, 8, 12, 16, 20, 24]) segments.push([r, c, 'r', black]);
  // No border or gridline inside a merge.
  segments.push([3, 7, 'r', white], [10, 6, 'b', white], [11, 23, 'b', white]);
  await check('Excel.Test.xlsx', segments, [
    { zoom: 0.8 }, { zoom: 1 }, { zoom: 1.5, dpr: 2 },
    { scrollR: 2, scrollC: 14 }, { scrollR: 11, scrollC: 23 },
    { freeze: { r: 1, c: 2 }, scrollR: 2, scrollC: 14 },
  ]);
});

test('Issue 22: partial edges and mixed colors stay on their perimeter segments; absent anchors still draw', async () => {
  const segments = [
    // An anchor-only border is not expanded into a full rectangle by Excel.
    [2, 1, 't', black], [2, 2, 't', white], [2, 3, 't', white],
    [2, 1, 'l', black], [3, 1, 'l', white], [4, 1, 'l', white],
    [2, 1, 'b', white], [2, 1, 'r', white], [4, 1, 'b', white], [4, 3, 'r', white],
    // The three colored top segments must retain their own styles.
    [7, 1, 't', [255, 0, 0]], [7, 2, 't', [0, 128, 0]], [7, 3, 't', [0, 0, 255]],
    [8, 1, 'l', black], [8, 3, 'r', black], [9, 2, 'b', [0, 0, 255]],
    // This merge has no stored top-left cell at all.
    [12, 2, 't', white], [13, 3, 'r', black], [14, 1, 'b', black], [14, 2, 'b', black], [14, 3, 'b', black],
    // A diagonal spans the merge; the outline's partial vertical segments stay partial.
    [17, 2, 't', black], [17, 1, 'l', black], [18, 1, 'l', white], [18, 3, 'r', white], [19, 2, 'b', black],
  ];
  await check('merge-cases.xlsx', segments);
});

test('Issue 22: hidden anchor rows and columns do not suppress the remaining visible perimeter', async () => {
  await check('merge-cases.xlsx', [[8, 3, 'r', black], [9, 2, 'b', [0, 0, 255]]], [{}], `
    sh.rows[7].hidden = true;
    sh.cols[1] = { hidden: true };
    L.layout.invalidate(sh);
  `);
});

test('Issue 22: an XML Spreadsheet 2003 merge, whose format only its top-left cell can hold, keeps its whole outline', async () => {
  const segments = [[2, 1, 'r', white], [3, 2, 'b', white]];
  for (let c = 1; c <= 3; c++) segments.push([2, c, 't', black], [4, c, 'b', black]);
  for (let r = 2; r <= 4; r++) segments.push([r, 1, 'l', black], [r, 3, 'r', black]);
  await check('merge-2003.xml', segments);
});

test('Issue 22: print pages keep the last row and right edge of the reported tables', async () => {
  const bytes = fs.readFileSync(new URL('Excel.Test.xlsx', import.meta.url)).toString('base64');
  const got = await page.evaluate(`(async () => {
    const wb = await L.xlsxRead.read(Uint8Array.from(atob(${JSON.stringify(bytes)}), c => c.charCodeAt(0)));
    const sh = wb.sheets[0], g = L.layout.geo(sh), pages = L.print.allPages([sh]), out = [];
    for (const [i, item] of pages.entries()) {
      const { pg, pag } = item, k = 2, zoom = pag.scale * k * 0.75;
      const cv = await L.print.renderPage(sh, pg, { page: i + 1, pages: pages.length, pag }, k);
      const ctx = cv.getContext('2d');
      const X = c => pag.margins.l * 72 * k + Math.round((g.cols.pos(c) - g.cols.pos(pg.c1)) * zoom);
      const Y = r => pag.margins.t * 72 * k + Math.round((g.rows.pos(r) - g.rows.pos(pg.r1)) * zoom);
      const dark = (x, y) => {
        const a = ctx.getImageData(Math.floor(x) - 1, Math.floor(y) - 1, 3, 3).data;
        for (let j = 0; j < a.length; j += 4) if (Math.max(a[j], a[j + 1], a[j + 2]) < 80) return true;
        return false;
      };
      if (pg.r1 <= 12 && pg.r2 >= 12) for (let c = pg.c1; c <= Math.min(24, pg.c2); c++)
        out.push({ page: i + 1, edge: L.formula.cellName(12, c) + ' bottom', dark: dark((X(c) + X(c + 1)) / 2, Y(13) - 0.5) });
      if (pg.c1 <= 24 && pg.c2 >= 24) for (const r of [0, 1, 2, 3, 4, 5, 8, 9, 10, 11, 12])
        out.push({ page: i + 1, edge: L.formula.cellName(r, 24) + ' right', dark: dark(X(25) - 0.5, (Y(r) + Y(r + 1)) / 2) });
    }
    return { pages: pages.length, samples: out };
  })()`);
  assert.equal(got.pages, 3, 'the same three pages as Excel');
  assert.equal(got.samples.length, 36, '25 bottom segments and 11 right segments');
  for (const s of got.samples) assert.equal(s.dark, true, `printed page ${s.page}: ${s.edge}`);
  assert.deepEqual(page.errors, []);
  assert.deepEqual(page.requests, []);
});
