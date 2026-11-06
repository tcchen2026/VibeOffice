// The text of a SmartArt drawing's shape sits in the box the drawing gives it (dsp:txXfrm), turned by that box's own
// rotation on top of the shape's. LibreOffice's smartart-rotation2.pptx turns three shapes upside down (rot 180) with
// their text boxes turned back (rot 180): PowerPoint 2016 draws "Text", "Was", "Mirrored" upright at the top of each
// shape (2026-11-07); Lectern drew them upside down at the bottom.
// node --test tests/lectern/smartart-text-box.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const file = path.join(process.env.VO_CORPORA || '/work/corpora', 'powerpoint', 'libreoffice__2a857bcf3b03__smartart-rotation2.pptx');
let page, stop;
before(async () => {
  if (!fs.existsSync(file)) return;
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

test('text turned back inside an upside-down SmartArt shape reads upright at its top', async (t) => {
  if (!page) return t.skip('Fetch pinned corpus: ' + file);
  const r = await page.evaluate(`(async () => {
    const pres = await L.pptx.read(Uint8Array.from(atob(${JSON.stringify(fs.readFileSync(file).toString('base64'))}), (c) => c.charCodeAt(0)));
    const el = L.render.slide(pres, pres.slides[0], { mode: 'thumb', index: 0 });
    el.style.cssText += ';position:fixed;left:0;top:0';
    document.body.appendChild(el);
    const out = {};
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    for (let n; (n = walker.nextNode());) {
      const word = n.data.trim();
      if (!['Text', 'Was', 'Mirrored'].includes(word)) continue;
      const at = (i) => { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 1); return r.getBoundingClientRect(); };
      const first = at(n.data.indexOf(word)), last = at(n.data.indexOf(word) + word.length - 1);
      const shape = n.parentElement.closest('.shape, [data-id]') || n.parentElement;
      let box = shape; while (box && !box.querySelector('svg.geo')) box = box.parentElement;
      const b = (box || shape).getBoundingClientRect();
      out[word] = { upright: first.left < last.left && first.top <= last.top + 1, top: first.top + first.height / 2 < b.top + b.height / 2 };
    }
    el.remove();
    return out;
  })()`);
  assert.deepEqual(page.errors, []);
  const ok = { upright: true, top: true };
  assert.deepEqual(r, { Text: ok, Was: ok, Mirrored: ok });
});
