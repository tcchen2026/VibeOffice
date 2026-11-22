// Text bent by a warp (bodyPr's prstTxWarp; LettersArt), drawn from DrawingML's 40 text warp definitions
// (common/presets.js) as PowerPoint 2016 bends text, measured in its PDF of tools/shapes/warp-deck.py's 228 cells
// (2026-11-07; common/geometry.js): the warp fills the whole shape; paths in pairs are bands whose lines' ink is
// stretched to fill them, a path alone carries its lines at their size, 0.19 em above the baseline of the line on the
// inside of its curve, the others stacked outward; lines are dealt out as evenly as they go, the first slots taking one
// more. Lectern drew such text straight, or (made as WordArt) as one line along a hand-drawn curve; Quire drew a text
// box's warped text straight.
// node --test tests/text-warps.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../tools/ooxml/cdp.mjs';

const corpora = process.env.VO_CORPORA || '/work/corpora';
const deck = path.join(corpora, 'powerpoint', 'libreoffice__9995a58494f5__tdf116350-texteffects.pptx');
const doc = path.join(corpora, 'word', 'libreoffice__104b70f8d31c__tdf125885_WordArt.docx');
let page, stop;
before(async () => {
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

/** L.geom.warp on made-up lines: each a run of n glyphs 10 wide, ink 7 up and 2 down, size 10 */
const bend = (warp, w, h, lines) => page.evaluate(`(() => {
  const line = ([n, algn = 'ctr']) => ({ glyphs: Array.from({ length: n }, (_, i) => ({ x: i * 10, w: 10 })), width: n * 10, ink: [0, -7, n * 10, 2], size: 10, pitch: 12, algn });
  const out = L.geom.warp(${JSON.stringify(warp)}, null, ${w}, ${h}, ${JSON.stringify(lines)}.map(line));
  /* where each glyph's ink corners and baseline middle go: [top left, bottom right, baseline middle] */
  const at = (m, x, y) => [Math.round((m[0] * x + m[2] * y + m[4]) * 10) / 10, Math.round((m[1] * x + m[3] * y + m[5]) * 10) / 10];
  return out && out.map((ln) => ln.map((pieces) => { const a = pieces[0].m, b = pieces[pieces.length - 1].m; return [at(a, 0, -7), at(b, 10, 2), at(pieces[Math.floor(pieces.length / 2)].m, 5, 0)]; }));
})()`);

test('a band stretches its lines\' ink to fill it: Plain Text fills the shape', async () => {
  const r = await bend('textPlain', 200, 100, [[4]]);
  assert.deepEqual(r[0][0][0], [0, 0], 'the first glyph\'s ink at the top left');
  assert.deepEqual(r[0][3][1], [200, 100], 'the last glyph\'s ink at the bottom right');
});

test('a path alone carries its line at its size, the path 0.19 em above the baseline', async () => {
  const [[...glyphs]] = await bend('textArchUp', 160, 110, [[3]]);
  // Arch Up's path is the top half of the ellipse in the shape: the middle glyph stands on its top, 1.9 below it
  assert.deepEqual(glyphs[1][2], [80, 1.9]);
  const [[[, , start]]] = await bend('textArchUp', 160, 110, [[3, 'l']]);
  assert.ok(start[0] < 10 && start[1] > 40, 'aligned left, the text starts where the path does: ' + start);
});

test('lines past the path stack outward: the last on Arch Up, the first on Arch Down', async () => {
  const up = await bend('textArchUp', 160, 110, [[3], [3]]), down = await bend('textArchDown', 160, 110, [[3], [3]]);
  assert.deepEqual(up[1][1][2], [80, 1.9], 'Arch Up: the second line on the path');
  assert.deepEqual(up[0][1][2], [80, -10.1], 'and the first a pitch above it');
  assert.deepEqual(down[0][1][2], [80, 111.9], 'Arch Down: the first line on the path');
  assert.deepEqual(down[1][1][2], [80, 123.9], 'and the second a pitch below');
});

test('lines are dealt out among the bands and paths as evenly as they go', async () => {
  // Deflate-Inflate has two bands: two lines each; Button three paths: two lines on the first, then one each
  const di = await bend('textDeflateInflate', 160, 110, [[2], [2], [2], [2]]);
  assert.ok(di[1][0][1][1] < 50 && di[2][0][0][1] > 40, 'lines 1-2 in the top band, 3-4 in the bottom one');
  const bt = await bend('textButton', 160, 110, [[2], [2], [2], [2]]);
  assert.deepEqual(bt.map((ln) => Math.round(ln[1][2][1])), [-10, 2, 57, 112], 'the middle line on the middle path, the last on the bottom arc');
  assert.equal(await bend('textNoSuchWarp', 160, 110, [[1]]), null);
});

test('Lectern: a warped text body is drawn bent and kept as text; its own LettersArt bends the same way', async (t) => {
  if (!fs.existsSync(deck)) return t.skip('Fetch pinned corpus: ' + deck);
  const r = await page.evaluate(`(async () => {
    const pres = await L.pptx.read(Uint8Array.from(atob(${JSON.stringify(fs.readFileSync(deck).toString('base64'))}), (c) => c.charCodeAt(0)));
    const s = pres.slides[0], el = L.render.slide(pres, s, { mode: 'thumb', index: 0 });
    document.body.appendChild(el);
    const warped = s.shapes.filter((x) => x.tx && x.tx.warp);
    const glyphs = el.querySelector('[data-sid="' + warped[0].id + '"] svg.warp');
    const wa = { id: 'wa', type: 'wordart', x: 0, y: 0, w: 200, h: 80, wa: { text: 'Bent', font: 'Arial', warp: 'textArchUp', fill: { t: 'solid', c: '#FF0000' } } };
    const waEl = L.render.shape(wa, { pres, slide: s, design: L.model.design(pres, s), mode: 'thumb' }, {});
    el.remove();
    return { types: [...new Set(warped.map((x) => x.type))], warps: warped.length, text: L.txt.plain(warped[0].tx), drawn: glyphs && [...glyphs.querySelectorAll('text')].map((x) => x.textContent).join(''), wa: [...waEl.querySelectorAll('svg.warp text')].map((x) => x.textContent).join('') };
  })()`);
  assert.deepEqual(page.errors, []);
  assert.deepEqual(r.types, ['text']);
  assert.equal(r.warps, 20);
  assert.equal(r.text, 'FontEffectArcUp');
  assert.equal(r.drawn, 'FontEffectArcUp', 'each letter drawn on the arch');
  assert.equal(r.wa, 'Bent');
});

test('Quire: a text box\'s warped text is drawn bent from its laid-out lines', async (t) => {
  if (!fs.existsSync(doc)) return t.skip('Fetch pinned corpus: ' + doc);
  const quire = await openPage('quire', { persistence: false, automated: true });
  try {
    const r = await quire.evaluate(`(async () => {
      await L.app.openFile(new File([Uint8Array.from(atob(${JSON.stringify(fs.readFileSync(doc).toString('base64'))}), (c) => c.charCodeAt(0))], 'w.docx'));
      await new Promise((r) => setTimeout(r, 1500)); await document.fonts.ready;
      const box = document.querySelector('.shp.warped'), svg = box && box.querySelector(':scope > svg.warp');
      return svg && { chars: [...new Set([...svg.querySelectorAll('text')].map((x) => x.textContent))].sort().join(''), hidden: getComputedStyle(box.querySelector('.tbx')).opacity };
    })()`);
    assert.deepEqual(quire.errors, []);
    // Double Wave 1, its two right-aligned lines bent (Word 2016's PDF); the text under it kept for editing
    assert.ok(r, 'the text box is drawn bent');
    assert.equal(r.chars, [...new Set('abcTransformThisisalongersecondline.')].sort().join(''), 'its letters, each in slices');
    assert.equal(r.hidden, '0');
  } finally { await quire.close(); }
});
