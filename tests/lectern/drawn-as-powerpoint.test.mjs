// Small things Lectern drew unlike PowerPoint 2016, found by comparing its slides with PowerPoint's PDFs of the
// PowerPoint-made decks in the corpus (tools/lectern-pictures.mjs, tools/lectern-vs-powerpoint.py, 2026-11-07):
// - an empty placeholder's fill shows (a title bar left empty: LibreOffice's empty-title-placeholder);
// - an outline whose gradient stops are all clear is no line (tdf149588_transparentSolidFill drew a pink curve);
// - numCol="1" on a slide undoes its layout's two columns (tdf140852 was set in two columns);
// - SmartArt saved without its drawing, in a layout Lectern has, is laid out by Lectern's engine (an org chart of
//   nine was one box).
// - a gradient of two colours blends in linear light along an S-curve (tdf105739's red-to-green slide was #805828 at
//   its centre, PowerPoint's #BA813A), any other in sRGB (POI's backgrounds); a theme's gradient fill is a gradient
//   (tdf152186 drew them flat) and its effect style's shadow is drawn;
// - a placeholder without geometry takes its layout's, with its holes (custgeom-placeholder hid the shapes behind).
// node --test tests/lectern/drawn-as-powerpoint.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const dir = path.join(process.env.VO_CORPORA || '/work/corpora', 'powerpoint');
const file = (name) => path.join(dir, fs.existsSync(dir) ? fs.readdirSync(dir).find((f) => f.includes(name)) || name : name);
let page, stop;
before(async () => {
  if (!fs.existsSync(dir)) return;
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

/** runs probe(el, pres) on a slide of a corpus deck (the first), drawn as in a thumbnail */
async function onSlide(name, probe, slide = 1) {
  const r = await page.evaluate(`(async () => {
    const pres = await L.pptx.read(Uint8Array.from(atob(${JSON.stringify(fs.readFileSync(file(name)).toString('base64'))}), (c) => c.charCodeAt(0)));
    const el = L.render.slide(pres, pres.slides[${slide - 1}], { mode: 'thumb', index: ${slide - 1} });
    el.style.cssText += ';position:fixed;left:0;top:0';
    document.body.appendChild(el);
    await document.fonts.ready;
    try { return (${probe})(el, pres); } finally { el.remove(); }
  })()`);
  assert.deepEqual(page.errors, []);
  return r;
}

test("an empty title placeholder shows its master's fill", async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + dir);
  const fills = await onSlide('empty-title-placeholder', `(el, pres) => {
    const title = pres.slides[0].shapes.find((s) => s.ph && s.ph.type === 'title');
    return [...el.querySelector('[data-sid="' + title.id + '"]').querySelectorAll('svg.geo path')].map((p) => p.getAttribute('fill')).filter((f) => f && f !== 'none').length;
  }`);
  assert.ok(fills > 0, 'the title bar is drawn');
});

test('an outline with clear gradient stops is not drawn', async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + dir);
  const visible = await onSlide('tdf149588_transparentSolidFill', `(el) => [...el.querySelectorAll('svg.geo path')]
    .filter((p) => p.getAttribute('stroke') && p.getAttribute('stroke') !== 'none' && p.getAttribute('stroke') !== 'transparent' && !/,\\s*0\\)$/.test(p.getAttribute('stroke'))).length`);
  assert.equal(visible, 0);
});

test('numCol="1" on a slide undoes its layout\'s columns', async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + dir);
  const columns = await onSlide('tdf140852', `(el) => [...el.querySelectorAll('.txi')].map((x) => getComputedStyle(x).columnCount).filter((c) => c !== 'auto')`);
  assert.deepEqual(columns, []);
});

test('SmartArt saved without its drawing, in a layout Lectern has, is laid out with every item', async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + dir);
  // the Open XML SDK's organization chart has no drawing part: PowerPoint 2016 draws its nine people; Lectern drew "Top"
  const texts = await onSlide('9722227b7781__SmartArt_OrgChart1_3D', `(el) => [...el.querySelectorAll('.t-group .tx')].map((x) => x.textContent.trim()).filter(Boolean).sort()`);
  assert.deepEqual(texts, ['Assistant', 'Bottom-Left', 'Bottom-Left', 'Bottom-Right', 'Bottom-Right', 'Mid-Left', 'Mid-Right', 'Temp', 'Top']);
});

test('gradients: two colours blended as Office blends them, theme gradients and shadows drawn', async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + dir);
  const bg = await onSlide('tdf105739', `(el) => getComputedStyle(el.querySelector('.sl-bg')).backgroundImage`);
  // halfway and a quarter of the way (PowerPoint's PDF: #BA813A and #EE4920; a straight blend #805828 and #BF2C14)
  assert.match(bg, /rgb\(188, 128, 56\) 50%/);
  assert.match(bg, /rgb\(238, 70, 28\) 25%/);
  // four stops: straight from one to the next, as PowerPoint draws them
  const four = await onSlide('poi__b08b5fe3f751__backgrounds', `(el) => getComputedStyle(el.querySelector('.sl-bg')).backgroundImage.match(/%/g).length`, 2);
  assert.equal(four, 4);
  const styled = await onSlide('tdf152186', `(el, pres) => pres.slides[0].shapes.map((s) => s.fill && s.fill.t)`);
  assert.deepEqual(styled, ['solid', 'grad', 'grad']);
  // effectRef idx 2 of an Office 2007 theme: a soft shadow below
  const shadow = await onSlide('tdf114845_rotateShape', `(el, pres) => { const s = pres.slides[0].shapes.find((x) => x.shadow && !x.ph); return s && [s.shadow.dy > 0, s.shadow.blur > 0]; }`);
  assert.deepEqual(shadow, [true, true]);
});

test('a placeholder without geometry takes its layout\'s custom shape, holes and all', async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + dir);
  const r = await onSlide('custgeom-placeholder', `(el, pres) => { const s = pres.slides[0].shapes[0]; const p = el.querySelector('[data-sid="' + s.id + '"] svg.geo path'); return [s.geom, p && p.getAttribute('fill-rule')]; }`);
  assert.deepEqual(r, ['custom', 'evenodd']);
});

test("a slide's colour mapping holds for its master's background; a chart keeps its own", async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + dir);
  // LibreOffice's chart_pt_color_bg1: the slide maps bg1 to dk1 (a black slide); its doughnut's point in bg1 at 85%
  // keeps the chart's own mapping (bg1 light): PowerPoint 2016 draws it light grey on black, Lectern drew it the
  // other way round
  const r = await onSlide('chart_pt_color_bg1', `(el) => [getComputedStyle(el.querySelector('.sl-bg')).backgroundColor, [...el.querySelectorAll('svg path')].map((p) => (p.getAttribute('fill') || '').toUpperCase()).filter((f) => f === '#D9D9D9' || f === '#262626').sort()]`);
  assert.deepEqual(r, ['rgb(0, 0, 0)', ['#D9D9D9']]);
});
