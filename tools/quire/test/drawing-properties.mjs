// Property edits, table alternatives and original watermarks through the real writer/history/draft hooks.
// node tools/quire/test/drawing-properties.mjs CORPUS OUTPUT
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [corpus, output] = process.argv.slice(2);
if (!corpus || !output) throw new Error('Usage: drawing-properties.mjs CORPUS OUTPUT');
const cases = [
  ['libreoffice__e6eafd728bd0__picture-effects-preservation.docx', 'crop'],
  ['libreoffice__9c4d4a754d6f__shape-effect-preservation.docx', 'shadow'],
  ['libreoffice__a58731a6fb6d__textWatermark.docx', 'watermark'],
  ['libreoffice__fa852c560adb__pictureWatermark.docx', 'watermark'],
  ['table-row-alternatives.docx', 'row'],
  ['table-cell-alternatives.docx', 'cell'],
  ['table-empty-row-alternative.docx', 'empty-row'],
  ['table-empty-cell-alternative.docx', 'empty-cell'],
];
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), page = await openPage('quire', { persistence: false }), reports = [];
try {
  for (const [file, feature] of cases) {
    const data = /(?:row|cell)$/.test(feature) ? null : fs.readFileSync(path.join(corpus, file)).toString('base64');
    try {
      const result = await page.evaluate(`(${check.toString()})(${JSON.stringify(data)}, ${JSON.stringify(feature)})`);
      for (const [scenario, bytes] of Object.entries(result.artifacts)) {
        const folder = path.join(output, scenario); fs.mkdirSync(folder, { recursive: true });
        fs.writeFileSync(path.join(folder, file), Buffer.from(bytes, 'base64'));
        if (scenario !== 'original') reports.push({ file, scenario, attempted: true, status: 'ok', checks: result.checks });
      }
      console.log(file + ': ' + result.checks + ' checks');
    } catch (e) { reports.push({ file, scenario: feature, attempted: true, status: 'failed', error: e.message }); console.error(file + ': ' + e.message); process.exitCode = 1; }
  }
  if (page.errors.length) { console.error(page.errors); process.exitCode = 1; }
} finally {
  fs.writeFileSync(path.join(output, 'results.jsonl'), reports.map(r => JSON.stringify(r)).join('\n') + '\n');
  await page.close(); stop();
}

async function check(data, feature) {
  const { D, O, opc: K } = L, artifacts = {};
  let checks = 0;
  const assert = (ok, message) => { if (!ok) throw new Error(message); checks++; };
  const encode = async blob => { const b = new Uint8Array(await blob.arrayBuffer()); let s = ''; for (let i = 0; i < b.length; i += 32768) s += String.fromCharCode(...b.subarray(i, i + 32768)); return btoa(s); };
  if (!data) {
    const doc = D.newDoc(), table = D.simpleTable(3, 3, 432);
    table.rows.forEach((r, ri) => r.cells.forEach((c, ci) => c.blocks[0].runs.push(D.text('cell-' + ri + '-' + ci))));
    doc.main.blocks = [D.para([D.text('outside')]), table];
    const pkg = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
    let xml = pkg.text(pkg.main), tree = K.parse(xml);
    const rows = Array.from(tree.getElementsByTagName('*')).filter(e => e.localName === 'tr');
    const elements = feature.endsWith('row') ? rows.slice(0, 2) : Array.from(rows[0].children).filter(e => e.localName === 'tc').slice(0, 2);
    const first = L.xmlTree.source.get(elements[0]), last = L.xmlTree.source.get(elements.at(-1));
    const content = xml.slice(first.start, last.end);
    xml = K.patch(xml, [{ start: first.start, end: last.end, value: `<mc:AlternateContent><mc:Choice Requires="w14">${feature.startsWith('empty') ? '' : content}</mc:Choice><mc:Fallback>${content}</mc:Fallback></mc:AlternateContent>` }]);
    data = await encode(await L.zip.write(pkg.names.map(name => ({ name, data: name === pkg.main ? xml : pkg.bytes(name) }))));
  }
  artifacts.original = data;
  const doc = (await L.docx.read(Uint8Array.from(atob(data), c => c.charCodeAt(0)))).doc;
  L.app.loadDoc(doc, 'Preserved properties', { saved: true });
  const save = async (name, d = doc, blob) => {
    blob ||= await L.docx.write(d); artifacts[name] = await encode(blob);
    return (await L.docx.read(new Uint8Array(await blob.arrayBuffer()))).doc;
  };
  const count = (d, tag) => d.pkg.names.filter(n => /word\/(document|header\d+|footer\d+)\.xml$/.test(n)).reduce((n, part) => n + Array.from(d.pkg.xml(part)?.getElementsByTagName('*') || []).filter(e => e.localName === tag).length, 0);
  const watermarks = d => d.pkg.names.filter(n => /word\/(header|footer).*\.xml$/.test(n)).flatMap(part => Array.from(d.pkg.xml(part)?.getElementsByTagName('*') || []).filter(e => /WaterMark/i.test(e.getAttribute('id') || '')).map(e => { const p = L.xmlTree.source.get(e); return p.text.slice(p.start, p.end); })).sort();
  const before = await save('before');
  if (['crop', 'shadow'].includes(feature)) {
    const objects = []; D.walk(doc.main, p => { if (p.t === 'p') for (const it of p.runs) if (Object.hasOwn(it.keep?.opaque?.drawing || {}, feature)) objects.push({ p, it }); });
    assert(objects.length, 'No editable drawing with the requested property');
    const { p, it } = objects[0], expected = feature === 'crop' ? { l: 0.15, t: 0.05, r: 0, b: 0 } : { c: '#123456', a: 0.6, dx: 5, dy: 3, blur: 2 };
    D.tx('Edit one drawing property', () => { D.touch(p); it[feature] = expected; });
    const edited = await save('edited');
    for (const tag of ['glow', 'reflection', 'softEdge', 'scene3d', 'sp3d', 'extLst']) assert(count(edited, tag) === count(before, tag), tag + ' changed with another property');
    assert(!doc.losses?.some(e => /will be saved as ordinary content|^Compatibility .* was converted/.test(e.what)), 'Drawing property edit converted its frame: ' + JSON.stringify(doc.losses));
    const reopened = []; D.walk(edited.main, p => { if (p.t === 'p') for (const i of p.runs) if (i.keep?.opaque && i.t === it.t) reopened.push(i); });
    assert(reopened.some(i => feature === 'crop' ? Math.abs(i.crop?.l - expected.l) < 0.0001 : Math.abs(i.shadow?.dx - expected.dx) < 0.01), 'The changed drawing property did not reopen');
  } else if (feature === 'watermark') {
    assert(watermarks(before).length > 0, 'Missing original watermark');
    assert(JSON.stringify(watermarks(before)) === JSON.stringify(watermarks(doc)), 'Unedited watermark XML changed');
    D.tx('Replace watermark', () => { D.touchKey(doc, 'watermark'); doc.watermark = { type: 'text', text: 'REVISED', color: 'C0C0C0', layout: 'diagonal' }; });
    const edited = await save('edited');
    assert(edited.watermark?.text === 'REVISED', 'The replaced watermark did not reopen');
    assert(watermarks(edited).every(x => x.includes('REVISED')), 'An original watermark was resurrected');
  } else if (feature.startsWith('empty')) {
    assert(count(before, 'AlternateContent') === 1, 'Empty choice was dropped');
    const table = doc.main.blocks.find(b => b.t === 'tbl');
    const p = table.rows[0].cells[0].blocks[0];
    D.tx('Edit beside empty alternative', () => O.insertText(D.pos(p, 0), 'edited '));
    const edited = await save('edited');
    assert(count(edited, 'AlternateContent') === 1, 'Editing beside an empty choice dropped its fallback');
  } else {
    assert(count(before, 'AlternateContent') === 1, 'Table alternative was not retained once');
    const table = doc.main.blocks.find(b => b.t === 'tbl');
    const record = feature === 'row' ? table.rows[0].trPr.ac : table.rows[0].cells[0].tcPr.ac;
    assert(record?.count === 2, 'Alternative members were not attached to rows/cells');
    D.tx('Delete one alternative member', () => { D.touchTbl(table); if (feature === 'row') table.rows.splice(1, 1); else table.rows[0].cells.splice(1, 1); });
    const edited = await save('edited');
    assert(count(edited, 'AlternateContent') === 0, 'Deleting a member resurrected its original wrapper');
    assert(!edited.pkg.text(edited.pkg.main).includes(feature === 'row' ? 'cell-1-0' : 'cell-0-1'), 'Deleted table content returned');
  }
  assert(D.undo(), 'Missing undo');
  const undone = await save('undo');
  if (feature === 'watermark') assert(JSON.stringify(watermarks(undone)) === JSON.stringify(watermarks(before)), 'Undo did not restore original watermarks');
  else if (/(?:row|cell)$/.test(feature)) assert(count(undone, 'AlternateContent') === 1, 'Undo did not restore table alternative');
  assert(D.redo(), 'Missing redo'); await save('redo');
  const recovered = await save('draft', doc, await window.__corpusHooks.snapshot(doc));
  const again = await save('recovered', recovered);
  if (feature === 'watermark') assert(again.watermark?.text === 'REVISED', 'Draft lost watermark edit');
  return { checks, artifacts };
}
