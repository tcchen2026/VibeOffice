// Every DrawingML preset shape is drawn as PowerPoint 2016 draws it: the outline and fill of each of the 175 presets
// at its default adjustments, 110 × 70 pt, against where PowerPoint paints it in its PDF of tools/shapes/deck.mjs's deck
// (tests/shapes/powerpoint.json, tools/shapes/masks.py, 2026-11-07). The shapes come from DrawingML's own definitions
// (common/presets.js, evaluated in common/geometry.js); the hand-made drawings before them differed by more than 5% in
// 51 presets (tabs drawn as rectangles, arcs without their wedge, curved and circular arrows, callouts, gears, stars).
// Also: adjust handles from the definitions land where they are dragged, and arcs in freeforms use DrawingML's angles.
// node --test tests/shapes.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { ensureServer, openPage } from '../tools/ooxml/cdp.mjs';

const REF = JSON.parse(fs.readFileSync(new URL('./shapes/powerpoint.json', import.meta.url), 'utf8'));
let page, stop;
before(async () => {
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

test('each preset shape paints where PowerPoint paints it', async () => {
  const off = await page.evaluate(`(() => {
    const R = ${JSON.stringify(REF)}, out = [];
    const c = document.createElement('canvas'); c.width = R.cols; c.height = R.rows;
    const g = c.getContext('2d', { willReadFrequently: true });
    const mask = (b64) => { const s = atob(b64), m = new Uint8Array(R.cols * R.rows); for (let j = 0; j < m.length; j++) m[j] = (s.charCodeAt(j >> 3) >> (7 - (j & 7))) & 1; return m; };
    const grow = (m) => { const o = new Uint8Array(m.length); for (let y = 0; y < R.rows; y++) for (let x = 0; x < R.cols; x++) { if (!m[y * R.cols + x]) continue; for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) { const X = x + dx, Y = y + dy; if (X >= 0 && Y >= 0 && X < R.cols && Y < R.rows) o[Y * R.cols + X] = 1; } } return o; };
    for (const [name, [b64, ox, oy]] of Object.entries(R.shapes)) {
      g.setTransform(1, 0, 0, 1, 0, 0); g.fillStyle = '#fff'; g.fillRect(0, 0, R.cols, R.rows);
      g.setTransform(R.scale, 0, 0, R.scale, R.margin * R.scale + ox, R.margin * R.scale + oy);
      for (const p of L.geom.parts({ geom: name, w: R.w, h: R.h }, R.w, R.h)) {
        const path = new Path2D(p.d);
        if (p.fill !== 'none') { g.fillStyle = '#9DC3E6'; g.fill(path, p.rule || 'nonzero'); }
        if (p.stroke) { g.strokeStyle = '#000'; g.lineWidth = 1; g.stroke(path); }
      }
      const px = g.getImageData(0, 0, R.cols, R.rows).data, ours = new Uint8Array(R.cols * R.rows);
      for (let j = 0; j < ours.length; j++) ours[j] = (px[j * 4] * 0.299 + px[j * 4 + 1] * 0.587 + px[j * 4 + 2] * 0.114) < 235 ? 1 : 0;
      const theirs = mask(b64), og = grow(ours), tg = grow(theirs);
      let union = 0, apart = 0;
      for (let j = 0; j < ours.length; j++) { union += ours[j] | theirs[j]; apart += (ours[j] && !tg[j]) || (theirs[j] && !og[j]) ? 1 : 0; }
      const share = Math.round((1000 * apart) / Math.max(1, union)) / 10;
      if (share > 3) out.push(name + ' ' + share + '%');
    }
    return out;
  })()`);
  assert.deepEqual(page.errors, []);
  assert.deepEqual(off, [], 'presets drawn apart from PowerPoint (share of their painted pixels)');
});

test('adjust handles from the definitions follow the pointer', async () => {
  const r = await page.evaluate(`(() => {
    const out = {};
    for (const [geom, i, dx, dy] of [['roundRect', 0, 6, 0], ['rightArrow', 0, -10, 8], ['rightArrow', 1, -10, 8], ['star24', 0, 0, 6], ['wedgeRectCallout', 0, 30, 25]]) {
      const sh = { geom, w: 120, h: 80 }, hd = L.geom.handles(sh)[i], a = L.geom.adj(sh), p = hd.pos(120, 80, a);
      const moved = Object.assign({}, a, hd.set(p[0] + dx, p[1] + dy, 120, 80, a)), q = hd.pos(120, 80, moved);
      out[geom + i] = [Math.round(q[0] - p[0]), Math.round(q[1] - p[1])];
    }
    /* a polar handle (the block arc's angle and thickness) goes where it is put on its path */
    const sh = { geom: 'blockArc', w: 120, h: 80 }, hd = L.geom.handles(sh)[0], a = L.geom.adj(sh);
    const goal = hd.pos(120, 80, Object.assign({}, a, { adj1: 12600000, adj3: 15000 })), q = hd.pos(120, 80, Object.assign({}, a, hd.set(goal[0], goal[1], 120, 80, a)));
    out.blockArc = Math.round(Math.hypot(q[0] - goal[0], q[1] - goal[1]) * 10) / 10;
    return out;
  })()`);
  // a handle tied to one direction moves only that way; the callout's tail goes anywhere
  assert.deepEqual(r, { roundRect0: [6, 0], rightArrow0: [0, 8], rightArrow1: [-10, 0], star240: [0, 6], wedgeRectCallout0: [30, 25], blockArc: 0 });
});

test("a freeform's elliptical arc starts and ends at DrawingML's angles", async () => {
  // a quarter of a 100 × 50 ellipse from 0° to 45°: the visual angle 45° is where the ellipse meets the diagonal,
  // at (100 / √(1 + 4)), not at the parametric 45° (100 cos 45°, 50 sin 45°)
  const d = await page.evaluate(`L.geom.parts({ geom: 'custom', path: { paths: [{ w: 200, h: 100, cmds: [['M', 200, 50], ['A', 100, 50, 0, 45]] }] } }, 200, 100)[0].d`);
  const end = d.trim().split(/[\s,]+/).slice(-2).map(Number);
  assert.ok(Math.abs(end[0] - (100 + 100 / Math.sqrt(5))) < 0.05 && Math.abs(end[1] - (50 + 100 / Math.sqrt(5))) < 0.05, d);
});

test("every preset's text rectangle is the right way out", async () => {
  // the definitions put the pie's in t="ir" r="it"; PowerPoint 2016 sets its text from il, it to ir, ib
  // (LibreOffice's tdf149918_pie_textarea: the text starts 49.8 pt in and 33.2 pt down a 340.2 × 226.8 pt pie)
  const r = await page.evaluate(`(() => {
    const bad = Object.keys(L.presetShapes).filter((k) => { const t = L.geom.textRect({ geom: k, w: 200, h: 100 }); return !(t[2] >= t[0] && t[3] >= t[1]); });
    return { bad, pie: L.geom.textRect({ geom: 'pie', w: 340.2, h: 226.8, adj: { adj1: 20345271, adj2: 5405194 } }).map((v) => Math.round(v * 10) / 10) };
  })()`);
  assert.deepEqual(r, { bad: [], pie: [49.8, 33.2, 290.4, 193.6] });
});
