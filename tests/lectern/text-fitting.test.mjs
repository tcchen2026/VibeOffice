// Paragraph spacing and Shrink text on overflow (normAutofit) as PowerPoint 2016 draws them, measured in its PDFs of
// LibreOffice's TextFittingComparisonWithMSO decks (2026-11-07, pdftotext -bbox and the PDF's font sizes):
// - the space after a paragraph and the space before the next add up (6 + 6 pt: lines 55.2 pt apart, not 49.2);
// - shrunk font sizes are rounded to whole points (36 pt at 92.5% is drawn at 33 pt, at 62.5% at 23 pt);
// - the line spacing reduction comes off the percentage (90% less 20% is 70%: 19.32 pt lines of 23 pt text);
// - spacing and indents given in points are not shrunk; a bullet doesn't make its line taller.
// Lectern collapsed the two spaces into the larger, drew 33.3 pt text 72% spaced, shrank the 10 pt spacing and the
// indents, and grew each bulleted line to the bullet's size: shrunk text ran past the bottom of its placeholder.
// Over 134 slides of 8 decks with shrunk text, the difference from PowerPoint's pictures fell by 41%.
// node --test tests/lectern/text-fitting.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const dir = path.join(process.env.VO_CORPORA || '/work/corpora', 'powerpoint');
const files = { one: path.join(dir, 'libreoffice__7e0264e3b0b1__TextFittingComparisonWithMSO_1.pptx'), three: path.join(dir, 'libreoffice__ed8120aa56c4__TextFittingComparisonWithMSO_3.pptx') };
let page, stop;
before(async () => {
  if (!Object.values(files).every((f) => fs.existsSync(f))) return;
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

/** each text box's paragraphs on a slide: font size, line height, top, left padding, and the line tops of each */
async function paragraphs(file, slide) {
  const r = await page.evaluate(`(async () => {
    const pres = await L.pptx.read(Uint8Array.from(atob(${JSON.stringify(fs.readFileSync(file).toString('base64'))}), (c) => c.charCodeAt(0)));
    const el = L.render.slide(pres, pres.slides[${slide - 1}], { mode: 'thumb', index: ${slide - 1} });
    el.style.cssText += ';position:fixed;left:0;top:0';
    document.body.appendChild(el);
    await document.fonts.ready;
    const out = [...el.querySelectorAll('.tx')].map((tx) => [...tx.querySelectorAll('p.pa')].map((p) => {
      const cs = getComputedStyle(p), range = document.createRange();
      range.selectNodeContents(p);
      const lines = [...new Set([...range.getClientRects()].map((q) => Math.round(q.top * 100) / 100))];
      return { size: parseFloat(cs.fontSize), lh: parseFloat(cs.lineHeight), top: p.getBoundingClientRect().top, pad: parseFloat(cs.paddingLeft), lines };
    }));
    el.remove();
    return out;
  })()`);
  assert.deepEqual(page.errors, []);
  return r;
}
const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 0.15, `${msg}: ${a} (PowerPoint ${b})`);

test('the space after a paragraph and before the next add up', async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + dir);
  const [box] = await paragraphs(files.three, 1);
  near(box[1].top - box[0].top, 43.2 + 6 + 6, 'one line of 36 pt Arial, 6 pt after and 6 pt before');
});

test('shrunk text: whole-point sizes, line spacing reduced off the percentage, point spacing and indents kept', async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + dir);
  const shrunk = (await paragraphs(files.three, 1))[1];
  assert.equal(shrunk[0].size, 33, '36 pt at 92.5%');
  near(shrunk[1].top - shrunk[0].top, 33 * 1.2 * 0.8 + 12, 'a line at 100% less 20%, and 6 + 6 pt');
  const [box] = await paragraphs(files.one, 10);
  assert.equal(box[0].size, 23, '36 pt at 62.5%');
  near(box[0].lh, 23 * 1.2 * 0.7, 'lines at 90% less 20%');
  near(box[1].top - box[0].top, 23 * 1.2 * 0.7 + 10, 'one line and the 10 pt before, a bulleted line no taller');
  near(box[2].lines[2] - box[2].lines[1], 23 * 1.2 * 0.7, 'the lines of a wrapped paragraph');
  const unshrunk = (await paragraphs(files.one, 1)).flat().find((p) => p.pad > 0);
  assert.equal(box[0].pad, unshrunk.pad, 'the indent is the same as unshrunk text\'s');
});
