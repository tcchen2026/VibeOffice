// Actual list commands and imported picture-bullet identities.
// node tools/quire/test/picture-bullets.mjs CORPUS OUTPUT
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [corpus, output] = process.argv.slice(2);
if (!corpus || !output) throw new Error('Usage: picture-bullets.mjs CORPUS OUTPUT');
const file = 'openxml-sdk__9be9c0d1e02b__bullet-picture.docx';
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), page = await openPage('quire', { persistence: false });
let reports = [];
try {
  const data = fs.readFileSync(path.join(corpus, file)).toString('base64');
  const result = await page.evaluate(`(${check.toString()})(${JSON.stringify(data)})`);
  for (const [scenario, bytes] of Object.entries(result.artifacts)) {
    const dir = path.join(output, scenario); fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, file), Buffer.from(bytes, 'base64'));
    if (scenario !== 'original') reports.push({ file, scenario, attempted: true, status: 'ok', checks: result.checks });
  }
  const shot = await page.send('Page.captureScreenshot', { format: 'png' });
  fs.writeFileSync(path.join(output, 'picture-bullets.png'), Buffer.from(shot.data, 'base64'));
  console.log(result.checks + ' picture-bullet checks');
  if (page.errors.length) throw new Error(page.errors.join('\n'));
} catch (error) {
  reports.push({ file, scenario: 'picture-bullets', attempted: true, status: 'failed', error: error.message });
  console.error(error.message); process.exitCode = 1;
} finally {
  fs.writeFileSync(path.join(output, 'results.jsonl'), reports.map(r => JSON.stringify(r)).join('\n') + '\n');
  fs.writeFileSync(path.join(output, 'browser.json'), JSON.stringify({ errors: page.errors, requests: page.requests }, null, 2));
  await page.close(); stop();
}

async function check(data) {
  const { D, opc: K } = L, artifacts = { original: data }; let checks = 0;
  const assert = (ok, message) => { if (!ok) throw new Error(message); checks++; };
  const bytes = value => Uint8Array.from(atob(value), c => c.charCodeAt(0));
  const encode = async blob => {
    const a = new Uint8Array(await blob.arrayBuffer()); let s = '';
    for (let i = 0; i < a.length; i += 32768) s += String.fromCharCode(...a.subarray(i, i + 32768));
    return btoa(s);
  };
  let doc = (await L.docx.read(bytes(data))).doc;
  L.app.loadDoc(doc, 'Picture bullets', { saved: true });
  const part = pkg => pkg.rels(pkg.main).find(r => /\/numbering$/.test(r.type)).part;
  const elements = (pkg, name) => Array.from(pkg.xml(part(pkg)).getElementsByTagName('*')).filter(e => e.localName === name);
  const attr = (el, name) => Array.from(el.attributes).find(a => a.localName === name)?.value;
  const content = e => [e.namespaceURI, e.localName, Array.from(e.attributes).filter(a => a.name !== 'xmlns' && a.prefix !== 'xmlns' && a.localName !== 'Ignorable').map(a => [a.prefix ? e.lookupNamespaceURI(a.prefix) : '', a.localName, a.value]).sort(), Array.from(e.children).map(content)];
  const pictures = pkg => JSON.stringify(elements(pkg, 'numPicBullet').map(content));
  const original = pictures(doc.pkg), source = doc.pkg;
  const save = async (scenario, blob) => {
    blob ||= await L.docx.write(doc); artifacts[scenario] = await encode(blob);
    return K.open(new Uint8Array(await blob.arrayBuffer()));
  };
  const kept = pkg => {
    assert(pictures(pkg) === original, 'Picture definitions changed');
    for (const rel of source.rels(part(source))) if (!rel.external) for (const name of K.subtree(source, rel.part).parts) {
      const a = source.bytes(name), b = pkg.bytes(name);
      assert(b?.length === a.length && a.every((v, i) => v === b[i]), 'Picture dependency changed: ' + name);
    }
  };
  kept(await save('before'));
  let paragraph, info;
  D.walk(doc.main, p => { if (paragraph || p.t !== 'p') return; const n = L.lists.info(p); if (n && D.numLevelDef(doc, n.numId, n.lvl)?.pictureKeep) { paragraph = p; info = n; } });
  assert(!!paragraph, 'The public fixture has no displayed picture-bullet list');
  const select = () => { const p = D.pos(paragraph, 0); L.ed.sel = { a: p, f: p }; };
  const customize = value => {
    select(); L.dlg.customizeList(D.numDef(doc, info.numId).abs, info.lvl);
    if (value === 'align') document.querySelector('#cl-al').value = 'right';
    else { document.querySelector('#cl-style').value = 'decimal'; document.querySelector('#cl-text').value = '{1}.'; }
    document.querySelector('.dlg .dlg-foot .primary').click();
  };
  customize('align'); kept(await save('alignment'));
  assert(D.numLevelDef(doc, info.numId, info.lvl).picture, 'Alignment edit removed the preview');
  customize('symbol'); const changed = await save('symbol');
  assert(!D.numLevelDef(doc, info.numId, info.lvl).picture, 'Changed symbol still shows the picture');
  assert(doc.losses.some(e => e.id.startsWith('picture-bullet-edit:')), 'Missing symbol conversion notice');
  assert(elements(changed, 'lvlPicBulletId').length < elements(source, 'lvlPicBulletId').length, 'Changed symbol still references the picture');
  assert(D.undo(), 'Missing symbol undo'); kept(await save('undo'));
  assert(!doc.losses.some(e => e.id.startsWith('picture-bullet-edit:')), 'Undo retained a stale conversion notice');
  assert(D.redo(), 'Missing symbol redo'); await save('redo');
  const convertedDraft = await window.__corpusHooks.snapshot(doc), convertedMetadata = window.__corpusHooks.docInfo(doc);
  await save('symbol-draft', convertedDraft);
  assert(D.undo(), 'Missing restore for draft');
  const snapshot = await window.__corpusHooks.snapshot(doc), metadata = window.__corpusHooks.docInfo(doc);
  await save('draft', snapshot); VO.opened = () => {};
  await window.__corpusHooks.open(new File([snapshot], 'bullets.docx'), { draft: true, name: metadata.name, saved: true, lossState: metadata.lossState });
  doc = window.__corpusHooks.current(); kept(await save('recovered'));
  const foreign = (await L.docx.read(bytes(data))).doc, blocks = L.clone(foreign.main.blocks);
  D.tx('Import numbered content', () => { D.touchKey(doc, 'numbering'); D.touchKey(doc.main, 'blocks'); L.app.mergeNumbering(foreign, blocks); doc.main.blocks.push(...blocks); });
  const imported = await save('imported'), definitions = elements(imported, 'numPicBullet').map(e => attr(e, 'numPicBulletId'));
  assert(new Set(definitions).size === definitions.length, 'Imported bullet definitions collide');
  const ids = elements(imported, 'lvlPicBulletId').map(e => attr(e, 'val'));
  assert(ids.every(id => definitions.includes(id)), 'A level references a missing picture');
  assert(new Set(ids).size > new Set(elements(source, 'lvlPicBulletId').map(e => attr(e, 'val'))).size, 'Imported list reused the source definition identity');
  assert(D.undo(), 'Missing import undo'); kept(await save('import-undo'));
  D.tx('Remove list definitions', () => {
    D.touchKey(doc, 'numbering'); D.touchKey(doc, 'styles');
    doc.numbering.abs = {}; doc.numbering.nums = {};
    for (const style of Object.values(doc.styles)) if (style.pPr) delete style.pPr.num;
    D.walk(doc.main, p => { if (p.t === 'p') { D.touch(p); delete p.pPr.num; } });
  });
  const unused = await save('unused'); kept(unused);
  assert(elements(unused, 'lvlPicBulletId').length === 0, 'Removed lists left picture references');
  await window.__corpusHooks.open(new File([convertedDraft], 'bullets.docx'), { draft: true, name: convertedMetadata.name, saved: true, lossState: convertedMetadata.lossState });
  doc = window.__corpusHooks.current();
  const converted = await save('symbol-recovered');
  assert(doc.losses.some(e => e.id.startsWith('picture-bullet-edit:')), 'Draft recovery lost the unacknowledged symbol conversion');
  assert(elements(converted, 'lvlPicBulletId').length === elements(changed, 'lvlPicBulletId').length, 'Draft recovery resurrected the converted picture');
  return { artifacts, checks };
}
