// Word 2016's PDF clips inline pictures at gutter midpoints, not at text margins.
// Measured controls: office2016.json. Test painting/hit targets, not the implementation's CSS.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open, pause } from '../../lib/input.mjs';
import { makeColumnsFixture } from './fixtures.mjs';

test('Issue 5: oversized inline pictures clip between columns without resizing or clipping floats/AutoShapes', async t => {
  const page = await open('quire', { width: 1500, height: 1400 });
  try {
    await page.evaluate(`VO.opened = () => {}; window.makeColumnsFixture = ${makeColumnsFixture};`);
    const load = async kind => {
      await page.evaluate(`(async () => {
        L.ed.clearObjSel(); L.layout.view = 'print'; L.layout.setZoom(1);
        const doc = await makeColumnsFixture(${JSON.stringify(kind)});
        L.app.loadDoc(doc, 'Column control', {saved: true});
        await document.fonts.ready;
        await Promise.all([...document.images].map(i => i.decode().catch(() => {})));
        L.layout.render(); L.layout.pages[0].el.scrollIntoView({block: 'start'});
      })()`);
      await pause(200);
    };
    // These page coordinates are measured independently from Word's PDF clip rectangles.
    const bounds = { two: [0, 306, 612], unequal: [0, 201, 365.04, 612], single: [0, 612],
      floating: [0, 306, 612], rotated: [0, 306, 612], shapes: [0, 306, 612] };
    for (const kind of Object.keys(bounds)) await t.test(kind, async () => {
      await load(kind);
      for (const zoom of kind === 'two' ? [1, 0.75, 1.5] : [1]) {
        await page.evaluate(`L.layout.setZoom(${zoom}); L.layout.pages[0].el.scrollIntoView({block: 'start'});`);
        await pause(100);
        const got = await page.evaluate(`(() => {
          const pg = L.layout.pages[0], origin = pg.el.getBoundingClientRect(), z = L.layout.scale() * 96 / 72;
          const bounds = ${JSON.stringify(bounds[kind])}, rotated = ${kind === 'rotated'};
          const hit = (object, x, y) => document.elementsFromPoint(origin.left + x * z, origin.top + y * z).some(e => e === object || object.contains(e));
          const drawings = [...pg.body.querySelectorAll('.img[data-o], .shp[data-o]')].filter(e => !e.closest('.flt'));
          const clips = drawings.flatMap((object, i) => {
            const r = object.getBoundingClientRect(), cx = ((r.left+r.right)/2-origin.left)/z, cy = ((r.top+r.bottom)/2-origin.top)/z;
            const probe = x => hit(object, x, cy + (rotated ? Math.tan(Math.PI/6) * (x-cx) : 0));
            const points = [];
            if (i) points.push({edge:'left',inside:probe(bounds[i]+2),outside:probe(bounds[i]-2)});
            if (i+1 < drawings.length) points.push({edge:'right',inside:probe(bounds[i+1]-2),outside:probe(bounds[i+1]+2)});
            return points;
          });
          const floats = [...pg.el.querySelectorAll('.flt .img')].map(e => {
            const r=e.getBoundingClientRect(); return hit(e, 400, ((r.top+r.bottom)/2-origin.top)/z);
          });
          const first = drawings[0], r = first.getBoundingClientRect();
          const dimensions = L.D.doc.main.blocks.filter(p => p.runs?.[0]?.t === 'img' || p.runs?.[0]?.t === 'shape').map(p => [p.runs[0].w,p.runs[0].h]);
          return {clips, floats, dimensions, margin:hit(first,60,((r.top+r.bottom)/2-origin.top)/z), pages:L.layout.pages.length};
        })()`);
        assert.equal(got.pages, 1, `${kind}, zoom ${zoom}: pagination`);
        for (const clip of got.clips) {
          assert.equal(clip.inside, true, `${kind}, zoom ${zoom}: visible inside ${clip.edge} boundary`);
          assert.equal(clip.outside, kind === 'shapes', `${kind}, zoom ${zoom}: only pictures clip outside ${clip.edge} boundary`);
        }
        for (const visible of got.floats) assert.equal(visible, true, 'floats may cross columns');
        if (kind === 'floating') assert.equal(got.floats.length, 2, 'both page-layer and square-wrapped floats');
        if (kind !== 'rotated') assert.equal(got.margin, true, 'negative indent may extend into the outer margin');
        assert.deepEqual(got.dimensions, kind === 'floating' ? [[320,60],[300,60],[300,60],[320,60]] : Array(bounds[kind].length-1).fill([320,60]), 'model sizes stay intact');
      }
    });
    await t.test('click visible picture and neighbouring column; changing columns removes the clip', async () => {
      await load('two');
      const point = async x => page.evaluate(`(() => {const p=L.layout.pages[0].el.getBoundingClientRect(),z=L.layout.scale()*96/72;return {x:p.left+${x}*z,y:p.top+100*z};})()`);
      await page.click(await point(180));
      assert.equal(await page.evaluate('L.ed.objSel?.it?.alt'), 'Column 1');
      await page.click(await point(330));
      assert.equal(await page.evaluate('L.ed.objSel?.it?.alt'), 'Column 2');
      await page.evaluate(`L.ed.clearObjSel(); L.D.doc.sect.cols={n:1}; L.layout.render();`);
      await pause(100);
      await page.click(await point(330));
      assert.equal(await page.evaluate('L.ed.objSel?.it?.alt'), 'Column 1', 'same picture extends beyond the former gutter');
    });
    page.clean();
  } finally { await page.close(); }
});
