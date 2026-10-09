// Numbering metadata through real list commands, copies, history and drafts.
// node tools/quire/test/numbering.mjs CORPUS OUTPUT
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [corpus, output, only] = process.argv.slice(2);
if (!corpus || !output) throw new Error('Usage: numbering.mjs CORPUS OUTPUT');
const cases = [
  ['poi__1c99b0be78c7__Numbering.docx', 'custom'],
  ['libreoffice__7efd6aa50382__tdf128245.docx', 'style'],
].filter(([, kind]) => !only || kind === only);
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), page = await openPage('quire', { persistence: false }), reports = [];
try {
  for (const [file, kind] of cases) {
    try {
      const result = await page.evaluate(`(${check.toString()})(${JSON.stringify(fs.readFileSync(path.join(corpus, file)).toString('base64'))}, ${JSON.stringify(kind)})`);
      for (const [scenario, bytes] of Object.entries(result.artifacts)) {
        const dir = path.join(output, scenario); fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(path.join(dir, file), Buffer.from(bytes, 'base64'));
        if (scenario !== 'original') reports.push({ file, scenario, attempted: true, status: 'ok', checks: result.checks });
      }
      console.log(file + ': ' + result.checks + ' checks');
    } catch (error) {
      reports.push({ file, scenario: kind, attempted: true, status: 'failed', error: error.message });
      console.error(file + ': ' + error.message); process.exitCode = 1;
    }
  }
  const shot = await page.send('Page.captureScreenshot', { format: 'png' });
  fs.writeFileSync(path.join(output, 'numbering.png'), Buffer.from(shot.data, 'base64'));
  if (page.errors.length) { console.error(page.errors.join('\n')); process.exitCode = 1; }
} finally {
  fs.writeFileSync(path.join(output, 'results.jsonl'), reports.map(r => JSON.stringify(r)).join('\n') + '\n');
  fs.writeFileSync(path.join(output, 'browser.json'), JSON.stringify({ errors: page.errors, requests: page.requests }, null, 2));
  await page.close(); stop();
}

async function check(data, kind) {
  const { D, opc: K } = L, artifacts = { original: data }; let checks = 0;
  const assert = (ok, message) => { if (!ok) throw new Error(message); checks++; };
  const bytes = value => Uint8Array.from(atob(value), c => c.charCodeAt(0));
  const encode = async blob => {
    const a = new Uint8Array(await blob.arrayBuffer()); let s = '';
    for (let i = 0; i < a.length; i += 32768) s += String.fromCharCode(...a.subarray(i, i + 32768));
    return btoa(s);
  };
  let doc = (await L.docx.read(bytes(data))).doc;
  L.app.loadDoc(doc, 'Numbering', { saved: true });
  const source = doc.pkg;
  const root = pkg => pkg.xml(pkg.rels(pkg.main).find(r => /\/numbering$/.test(r.type)).part);
  const attr = (el, name) => Array.from(el.attributes).find(a => a.localName === name)?.value;
  const elements = (pkg, name) => Array.from(root(pkg).getElementsByTagName('*')).filter(e => e.localName === name);
  // Property-level preservation has separate checks. Here every numbering
  // wrapper, identity, absent default, legacy setting and MC branch is compared.
  const content = (e, ignoreAlignment = false) => {
    if (['pPr', 'rPr'].includes(e.localName) || ignoreAlignment && e.localName === 'lvlJc') return null;
    const children = Array.from(e.children).map(c => content(c, ignoreAlignment)).filter(Boolean);
    if (e.localName === 'numbering') children.sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
    return [e.namespaceURI, e.localName, Array.from(e.attributes).filter(a => a.name !== 'xmlns' && a.prefix !== 'xmlns' && a.localName !== 'Ignorable').map(a => [a.prefix ? e.lookupNamespaceURI(a.prefix) : '', a.localName, a.value]).sort(), children];
  };
  const signature = (pkg, ignoreAlignment = false) => JSON.stringify(content(root(pkg), ignoreAlignment));
  const save = async (scenario, blob) => {
    blob ||= await L.docx.write(doc); artifacts[scenario] = await encode(blob);
    return K.open(new Uint8Array(await blob.arrayBuffer()));
  };
  assert(signature(await save('before')) === signature(source), 'Unchanged numbering metadata differs');
  const find = () => {
    let found;
    D.walk(doc.main, p => {
      if (found || p.t !== 'p') return;
      const info = L.lists.info(p), abs = info && D.numDef(doc, info.numId)?.abs;
      if (abs && (kind === 'custom' ? D.numLevelDef(doc, info.numId, info.lvl)?.custFmt : abs.styleLink)) found = { p, info, abs };
    });
    assert(!!found, 'Fixture has no editable ' + kind + ' list');
    return found;
  };
  const select = p => { const pos = D.pos(p, 0); L.ed.sel = { a: pos, f: pos }; };
  let current = find(); select(current.p);
  const format = D.numLevelDef(doc, current.info.numId, current.info.lvl).custFmt;
  const dialog = change => {
    select(current.p); L.dlg.customizeList(current.abs, current.info.lvl);
    const style = document.querySelector('#cl-style');
    if (kind === 'custom') assert(style.value === 'custom', 'List dialog lost the custom format selection');
    if (change === 'alignment') document.querySelector('#cl-al').value = 'right';
    else { style.value = 'upperRoman'; document.querySelector('#cl-text').value = '{1}.'; }
    document.querySelector('.dlg .dlg-foot .primary').click();
  };
  dialog('alignment');
  const aligned = await save('alignment');
  if (kind === 'custom') {
    assert(signature(aligned, true) === signature(source, true), 'Alignment changed unrelated numbering metadata');
    assert(D.numLevelDef(doc, current.info.numId, current.info.lvl).custFmt === format, 'Alignment replaced the custom number format');
    current.abs = D.numDef(doc, current.info.numId).abs;
    dialog('format'); const converted = await save('format');
    const abstract = elements(converted, 'abstractNum').find(e => attr(e, 'abstractNumId') === String(current.abs.id));
    const level = Array.from(abstract.children).find(e => e.localName === 'lvl' && attr(e, 'ilvl') === String(current.info.lvl));
    const formats = Array.from(level.getElementsByTagName('*')).filter(e => e.localName === 'numFmt');
    assert(formats.length === 2 && formats.every(e => attr(e, 'val') === 'upperRoman' && !attr(e, 'format')), 'Chosen number format did not replace both compatibility branches');
    assert(D.undo(), 'Missing format undo');
    assert(signature(await save('format-undo'), true) === signature(source, true), 'Undo lost original metadata');
    assert(D.redo(), 'Missing format redo'); await save('format-redo');
    assert(D.undo(), 'Missing restore before recovery');
  } else {
    const changed = D.numDef(doc, current.info.numId).abs;
    assert(!changed.styleLink, 'Editing a delegated list did not materialize its levels');
    assert(doc.losses.some(e => e.id.startsWith('numbering-style:')), 'Missing numbering-style conversion notice');
    assert(D.undo(), 'Missing style undo');
    assert(signature(await save('style-undo')) === signature(source), 'Style undo did not restore its original definition');
    assert(!doc.losses.some(e => e.id.startsWith('numbering-style:')), 'Undo retained a stale conversion notice');
    assert(D.redo(), 'Missing style redo'); await save('style-redo');
    assert(D.undo(), 'Missing style restore');
  }
  const snapshot = await window.__corpusHooks.snapshot(doc), metadata = window.__corpusHooks.docInfo(doc);
  await save('draft', snapshot); VO.opened = () => {};
  await window.__corpusHooks.open(new File([snapshot], 'numbering.docx'), { draft: true, name: metadata.name, saved: true, lossState: metadata.lossState });
  doc = window.__corpusHooks.current();
  const recovered = await save('recovered');
  assert(signature(recovered, kind === 'custom') === signature(source, kind === 'custom'), 'Draft recovery changed numbering metadata');
  current = find(); select(current.p);
  const originalIds = Object.keys(doc.numbering.abs), originalNumIds = Object.keys(doc.numbering.nums);
  L.lists.apply(current.abs, { newList: true });
  const copied = await save('copied');
  const newIds = elements(copied, 'abstractNum').map(e => attr(e, 'abstractNumId'));
  const newNums = elements(copied, 'num').map(e => attr(e, 'numId'));
  assert(new Set(newIds).size === newIds.length && originalIds.every(id => newIds.includes(id)) && newIds.length === originalIds.length + 1, 'New list changed or collided with an original abstract ID');
  assert(new Set(newNums).size === newNums.length && originalNumIds.every(id => newNums.includes(id)) && newNums.length === originalNumIds.length + 1, 'New list changed or collided with a concrete ID');
  assert(D.undo(), 'Missing copy undo'); await save('copy-undo');
  // Reusing one imported model also checks that two copy cohorts stay distinct.
  const foreign = (await L.docx.read(bytes(data))).doc;
  for (let i = 1; i <= 2; i++) {
    const blocks = L.clone(foreign.main.blocks);
    D.tx('Import numbered content', () => { D.touchKey(doc, 'numbering'); D.touchKey(doc.main, 'blocks'); L.app.mergeNumbering(foreign, blocks); doc.main.blocks.push(...blocks); });
    const imported = await save('import-' + i), definitions = elements(imported, 'abstractNum').map(e => attr(e, 'abstractNumId'));
    assert(new Set(definitions).size === definitions.length, 'Imported definitions collide');
    assert(elements(imported, 'abstractNumId').every(e => definitions.includes(attr(e, 'val'))), 'Imported concrete list has a missing abstract definition');
  }
  return { artifacts, checks };
}
