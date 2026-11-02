// What PowerPoint 2016 needs from our IntelliArt parts: presentation links that name the layout part (or it draws
// nothing, and a hierarchy closes PowerPoint), and layout nodes numbered once per definition (or a few shapes are drawn
// across the whole diagram). Both were found with the Office 2016 PDF exports (docs/smartart.md).
// node --test tests/lectern/smartart-definitions.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

let page, stop;
before(async () => {
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

async function parts() {
  await L.app.newPresentation();
  const make = (layout, lines) => {
    L.ui.exec('newSlide');
    const sl = L.ed.slide(), body = sl.shapes.find((x) => x.ph && (x.ph.type || 'body') === 'body');
    L.diagram.make(L.pres, sl, body, layout, lines, {});
  };
  make('process', ['Plan', 'Build', 'Test', 'Launch']);
  make('radial', ['Core', '\tPart 1', '\tPart 2']);
  make('orgChart', ['Director', '\tSales', '\t\tNorth', '\tProduct']);
  const pkg = await L.opc.open(new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer()));
  const out = [];
  for (const name of pkg.names) {
    const m = /^ppt\/slides\/slide\d+\.xml$/.exec(name);
    if (!m) continue;
    const rels = pkg.rels(name), target = (type) => rels.find((r) => r.type.endsWith('/' + type))?.part;
    if (!target('diagramData')) continue;
    const data = pkg.text(target('diagramData')), layout = pkg.text(target('diagramLayout'));
    out.push({
      uniqueId: /uniqueId="([^"]*)"/.exec(layout)[1],
      presIds: [...new Set([...data.matchAll(/type="pres(?:Of|ParOf)"[^>]*presId="([^"]*)"/g)].map((x) => x[1]))],
      names: [...layout.matchAll(/<dgm:layoutNode name="([^"]*)"/g)].map((x) => x[1]),
    });
  }
  return out;
}
let result;
const run = async () => {
  result ||= await page.evaluate(`(${parts.toString()})()`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

test('presentation links name the layout part they were made for', async () => {
  const diagrams = await run();
  assert.equal(diagrams.length, 3);
  for (const d of diagrams) assert.deepEqual(d.presIds, [d.uniqueId]);
});

test('layout nodes are numbered once per definition, not per item count', async () => {
  for (const d of await run()) {
    assert.equal(new Set(d.names).size, d.names.length, d.uniqueId + ': a name repeats');
    assert.deepEqual(d.names.filter((n) => /^k\d+[a-z]\d+$/.test(n)), [], d.uniqueId);
  }
});
