// The SmartArt layout engine (public/common/smartart.js): every layout, for 1 to 7 items, gives finite shapes
// inside the box, one node per item it shows, and one text size per role within the limits.
//
//   node --test tools/smartart.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
globalThis.L = {};
require('../public/common/smartart.js');
const SA = globalThis.L.smartart;
const BOX = { x: 72, y: 144, w: 576, h: 252 };
const measure = (t, size) => t.length * size * 0.5;

const items = (n, kids = 0) => Array.from({ length: n }, (_, i) => ({ id: 'n' + i, text: 'Step ' + (i + 1), kids: Array.from({ length: kids }, (_, j) => ({ id: `n${i}k${j}`, text: 'Detail ' + j, kids: [] })) }));
const shown = (lay, list) => (lay.id === 'radial' ? [list[0], ...(list[0].kids.length ? list[0].kids : list.slice(1))] : lay.tree ? SA.outline(list).map((o) => o.item) : list);

for (const lay of SA.LAYOUTS) {
  test(`${lay.name}: shapes inside the box for 1–7 items`, () => {
    for (let n = 1; n <= 7; n++) {
      const list = lay.tree || lay.id === 'radial' ? [{ id: 'top', text: 'Top', kids: items(n - 1 || 1) }] : items(n, lay.levels > 1 ? 2 : 0);
      const { shapes } = SA.layout(lay.id, list, BOX, { measure });
      for (const s of shapes) {
        for (const k of ['x', 'y', 'w', 'h']) assert.ok(Number.isFinite(s[k]), `${lay.id} n=${n} ${s.role}.${k}`);
        if (s.rot) continue;   // rotated arrows of a cycle stay near their place
        assert.ok(s.x >= BOX.x - 0.5 && s.y >= BOX.y - 0.5 && s.x + s.w <= BOX.x + BOX.w + 0.5 && s.y + s.h <= BOX.y + BOX.h + 0.5, `${lay.id} n=${n} ${s.role} ${JSON.stringify([s.x, s.y, s.w, s.h])} outside`);
      }
      const nodes = shapes.filter((s) => s.role === 'node');
      assert.deepEqual(nodes.map((s) => s.itemId).sort(), shown(lay, list).map((it) => it.id).sort(), `${lay.id} n=${n}: one node per item`);
      const sizes = new Set(nodes.map((s) => s.fontSize));
      assert.equal(sizes.size, 1, `${lay.id} n=${n}: one text size for the items`);
      const [size] = sizes;
      assert.ok(size >= 6 && size <= 40, `${lay.id} n=${n}: size ${size}`);
    }
  });
}

test('longer text gets a smaller size, never below the minimum', () => {
  const short = SA.layout('process', items(3), BOX, { measure }).shapes.find((s) => s.role === 'node').fontSize;
  const long = SA.layout('process', items(3).map((it) => ({ ...it, text: 'A much longer description of this step in the plan' })), BOX, { measure }).shapes.find((s) => s.role === 'node').fontSize;
  assert.ok(long < short, `${long} < ${short}`);
  assert.ok(long >= 6);
});

test('blank diagrams and the outline', () => {
  assert.equal(SA.blank('orgChart')[0].kids.length, 3);
  assert.equal(SA.blank('process').length, 3);
  const o = SA.outline([{ id: 'a', text: '', kids: [{ id: 'b', text: '', kids: [] }] }, { id: 'c', text: '', kids: [] }]);
  assert.deepEqual(o.map((x) => [x.item.id, x.level]), [['a', 0], ['b', 1], ['c', 0]]);
});
