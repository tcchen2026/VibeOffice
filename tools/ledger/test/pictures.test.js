const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
require('../../../public/ledger/js/layout.js');
require('../../../public/ledger/js/ops.js');
const K = L.opc, O = L.ops, output = process.env.PICTURE_RESULTS, results = [];
const file = path.join(process.env.VO_CORPORA || path.join(os.homedir(), 'corpora'), 'excel', 'openxml-sdk__97ec4a7d7ff8__Photo Formats - CGM-O12-XL-Pictures.xlsx');
const elements = el => [el, ...el.getElementsByTagName('*')];
const pictures = pkg => pkg.names.filter(n => /^xl\/drawings\/drawing\d+\.xml$/.test(n)).sort().flatMap(n => elements(pkg.xml(n)).filter(e => e.localName === 'pic'));
const child = (el, tag) => el.children.find(e => e.localName === tag);
// Compare expanded property content, excluding only image relationship tokens.
function semantic(el) {
  if (!el) return null;
  return [el.namespaceURI, el.localName, Array.from(el.attributes).filter(a => a.name !== 'xmlns' && a.prefix !== 'xmlns' && !/\/relationships$/.test(a.namespaceURI || '')).map(a => [a.namespaceURI || '', a.localName, a.value]).sort(), el.children.map(semantic)];
}
const properties = el => ['blipFill', 'spPr', 'style'].map(n => semantic(child(K.parse(K.transitionalXML(K.raw(el))), n)));
async function open(t) {
  if (!fs.existsSync(file)) { t.skip('Fetch pinned corpus'); return; }
  return L.xlsxRead.read(fs.readFileSync(file));
}
async function save(wb, scenario) {
  const bytes = await L.xlsxWrite.bytes(wb), pkg = await K.open(bytes);
  if (output) {
    fs.mkdirSync(path.join(output, scenario), { recursive: true }); fs.writeFileSync(path.join(output, scenario, path.basename(file)), bytes);
    results.push({ file: path.basename(file), original: file, scenario, status: 'ok' });
    fs.writeFileSync(path.join(output, 'results.jsonl'), results.map(r => JSON.stringify(r)).join('\n') + '\n');
  }
  return { bytes, pkg };
}
test('ordinary picture effects, recoloring and authored transforms survive cells, history and recovery', async t => {
  const wb = await open(t); if (!wb) return;
  const original = pictures(wb.pkg).map(properties); assert.equal(original.length, 4);
  assert.deepEqual(pictures((await save(wb, 'save')).pkg).map(properties), original);
  O.tx(wb, 'Type', () => O.enter(wb.sheets[0], 0, 0, 'Edited'));
  const edited = await save(wb, 'text'); assert.deepEqual(pictures(edited.pkg).map(properties), original);
  wb.undo.undo(); assert.deepEqual(pictures((await save(wb, 'text-undo')).pkg).map(properties), original);
  wb.undo.redo(); const redo = await save(wb, 'text-redo');
  const recovered = await L.xlsxRead.read(redo.bytes);
  assert.deepEqual(pictures((await save(recovered, 'recovered')).pkg).map(properties), original);
});
test('picture geometry, crop, metadata and copies replace only their own properties', async t => {
  const wb = await open(t); if (!wb) return;
  const sh = wb.sheets[0], d = sh.drawings[0], before = properties(pictures(wb.pkg)[0]);
  const moved = JSON.parse(JSON.stringify(d)); moved.anchor.from.cOff += 9; moved.anchor.to.cOff += 9;
  moved.descr = 'Retained picture'; moved.crop = { l: 5, t: 0, r: 0, b: 0 };
  O.tx(wb, 'Format picture', () => O.setDrawings(sh, [moved]));
  const edited = await save(wb, 'picture-edit'), pic = pictures(edited.pkg)[0];
  assert.equal(elements(pic).find(e => e.localName === 'cNvPr').getAttribute('descr'), moved.descr);
  assert.equal(elements(pic).find(e => e.localName === 'srcRect').getAttribute('l'), '5000');
  const oldPr = child(K.parse(K.transitionalXML(K.raw(pictures(wb.pkg)[0]))), 'spPr');
  const newPr = child(pic, 'spPr');
  assert.deepEqual(newPr.children.filter(e => e.localName !== 'xfrm').map(semantic), oldPr.children.filter(e => e.localName !== 'xfrm').map(semantic));
  assert.deepEqual(semantic(child(child(pic, 'blipFill'), 'blip')), before[0][3].find(e => e[1] === 'blip'));
  const reread = await L.xlsxRead.read(edited.bytes); assert.deepEqual(reread.sheets[0].drawings[0].anchor, moved.anchor);
  wb.undo.undo(); assert.deepEqual(properties(pictures((await save(wb, 'picture-edit-undo')).pkg)[0]), before);
  wb.undo.redo(); await save(wb, 'picture-edit-redo');
  const copy = L.sheetObjects.copy(sh.drawings[0]); copy.id = d.id + 1; copy.anchor.from.r += 5; copy.anchor.to.r += 5;
  O.tx(wb, 'Copy picture', () => O.setDrawings(sh, sh.drawings.concat(copy)));
  const copied = await save(wb, 'picture-copy'); assert.equal(pictures(copied.pkg).length, 5);
  const part = copied.pkg.xml(sh.extra.ooxmlDrawing), ids = elements(part).filter(e => e.localName === 'cNvPr').map(e => e.getAttribute('id'));
  assert.equal(new Set(ids).size, ids.length);
  O.tx(wb, 'Delete picture', () => O.setDrawings(sh, [])); assert.equal(pictures((await save(wb, 'picture-delete')).pkg).length, 3);
});
