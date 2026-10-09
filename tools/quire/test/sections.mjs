// Section settings through the actual reader/writer, history and draft hooks.
// node tools/quire/test/sections.mjs CORPUS OUTPUT [--fixture]
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [corpus, output, ...flags] = process.argv.slice(2);
if (!corpus || !output) throw new Error('Usage: sections.mjs CORPUS OUTPUT [--fixture]');
const files = ['libreoffice__25b1ff42f98c__tdf149089.docx',
  'libreoffice__36354c47c353__floattable-nested-3tables.docx',
  'libreoffice__4d38e2895a98__defaultStyle.docx', 'section-settings.docx'].filter(f => !flags.includes('--fixture') || f === 'section-settings.docx');
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), page = await openPage('quire', { persistence: false }), reports = [];
try {
  for (const file of files) {
    try {
      const data = file === 'section-settings.docx' ? null : fs.readFileSync(path.join(corpus, file)).toString('base64');
      const result = await page.evaluate(`(${check.toString()})(${JSON.stringify(data)})`);
      for (const [scenario, bytes] of Object.entries(result.artifacts)) {
        const dir = path.join(output, scenario); fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(path.join(dir, file), Buffer.from(bytes, 'base64'));
        if (scenario !== 'original') reports.push({ file, scenario, attempted: true, status: 'ok', checks: result.checks });
      }
      console.log(file + ': ' + result.checks + ' checks');
    } catch (error) {
      reports.push({ file, scenario: 'section', attempted: true, status: 'failed', error: error.message });
      console.error(file + ': ' + error.message); process.exitCode = 1;
    }
  }
  const shot = await page.send('Page.captureScreenshot', { format: 'png' });
  fs.writeFileSync(path.join(output, 'sections.png'), Buffer.from(shot.data, 'base64'));
  if (page.errors.length) { console.error(page.errors); process.exitCode = 1; }
} finally {
  fs.writeFileSync(path.join(output, 'results.jsonl'), reports.map(r => JSON.stringify(r)).join('\n') + '\n');
  fs.writeFileSync(path.join(output, 'browser.json'), JSON.stringify({ errors: page.errors, requests: page.requests }, null, 2));
  await page.close(); stop();
}

async function check(data) {
  const { D, opc: K } = L, artifacts = {}; let checks = 0;
  const fixture = !data;
  const assert = (ok, message) => { if (!ok) throw new Error(message); checks++; };
  const encode = async blob => {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let text = '';
    for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768));
    return btoa(text);
  };
  const sections = pkg => Array.from(pkg.xml(pkg.main).getElementsByTagName('*')).filter(e => e.localName === 'sectPr' && e.parentNode?.localName !== 'sectPrChange');
  if (!data) {
    const doc = D.newDoc(); doc.main.blocks = [D.para([D.text('First section')]), D.para([D.text('Final section')])];
    doc.main.blocks[0].sect = D.defaultSect();
    doc.hf.test = { id: 'test', kind: 'hdr', blocks: [D.para([D.text('Preserved header')])] };
    doc.sect.refs.hdr.default = doc.main.blocks[0].sect.refs.hdr.default = 'test';
    const pkg = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
    const xml = pkg.text(pkg.main), replacement = '<w:pgNumType w:start="1" w:chapStyle="2" w:chapSep="hyphen"/><w:noEndnote/><w:rtlGutter/>' +
      '<mc:AlternateContent><mc:Choice Requires="w14"><w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="4096"/></mc:Choice><mc:Fallback><w:docGrid w:linePitch="360"/></mc:Fallback></mc:AlternateContent>';
    const edits = sections(pkg).map((e, i) => {
      const loc = L.xmlTree.source.get(e);
      const raw = K.merge(K.raw(e), { ['{' + K.NS.w + '}cols']: [], ['{' + K.NS.w + '}docGrid']: [] }, 'w:CT_SectPr');
      const revision = `<w:sectPrChange w:id="${10 + i}" w:author="Section review"><w:sectPr><w:rtlGutter/></w:sectPr></w:sectPrChange>`;
      return { start: loc.start, end: loc.end, value: raw.replace('</w:sectPr>', replacement + revision + '</w:sectPr>') };
    });
    data = await encode(await L.zip.write(pkg.names.map(name => ({ name, data: name === pkg.main ? K.patch(xml, edits) : pkg.bytes(name) }))));
  }
  artifacts.original = data;
  const doc = (await L.docx.read(Uint8Array.from(atob(data), c => c.charCodeAt(0)))).doc;
  L.app.loadDoc(doc, 'Section preservation', { saved: true });
  const hasGridRepair = b => b.t === 'tbl' && (b.keep?.gridRepair || b.rows.some(r => r.cells.some(c => c.blocks.some(hasGridRepair))));
  const repairedGrid = doc.main.blocks.some(hasGridRepair), repairedStyle = Object.values(doc.styles).some(s => s.keep?.defaultStyleRepair);
  const repairLoss = (d, name) => (d.losses || []).filter(e => e.id.startsWith(name + ':'));
  assert(!repairLoss(doc, 'table-grid').length && !repairLoss(doc, 'default-styles').length, 'Recovery notice was emitted before a save was prepared');
  const content = (el, pkg) => [el.namespaceURI, el.localName,
    Array.from(el.attributes).filter(a => a.name !== 'xmlns' && a.prefix !== 'xmlns' && !(a.localName === 'Ignorable' && el.lookupNamespaceURI(a.prefix) === K.NS.mc)).map(a => {
      const ns = a.prefix ? el.lookupNamespaceURI(a.prefix) : '';
      const value = ns === K.NS.r ? pkg.rels(pkg.main).find(r => r.id === a.value)?.part || a.value : a.value;
      return [ns, a.localName, value];
    }).sort(), Array.from(el.children).map(c => content(c, pkg))];
  const sectionValues = pkg => JSON.stringify(sections(pkg).map(e => content(e, pkg)));
  const original = sectionValues(doc.pkg);
  const save = async (scenario, value = doc, blob) => {
    blob ||= await L.docx.write(value); artifacts[scenario] = await encode(blob);
    assert(!value.losses?.some(e => e.id.startsWith('property:')), 'Section preservation fell back to conversion');
    return (await L.docx.read(new Uint8Array(await blob.arrayBuffer()))).doc;
  };
  const before = await save('before');
  assert(sectionValues(before.pkg) === original, 'An unedited section changed');
  assert(!!repairLoss(doc, 'table-grid').length === repairedGrid, 'Table recovery notice does not match the written tables');
  assert(!!repairLoss(doc, 'default-styles').length === repairedStyle, 'Style recovery notice does not match the written styles');
  if (repairedGrid) {
    D.tx('Delete table with recovered columns', () => { D.touchKey(doc.main, 'blocks'); doc.main.blocks = doc.main.blocks.filter(b => !hasGridRepair(b)); });
    await save('delete-repaired-table');
    assert(!repairLoss(doc, 'table-grid').length, 'Deleted recovered tables left a stale save notice');
    assert(D.undo(), 'Missing table deletion undo'); await save('restore-repaired-table');
    assert(repairLoss(doc, 'table-grid').length === 1, 'Undo did not restore the grouped table recovery notice');
  }
  D.tx('Change top margin', () => { D.touchKey(doc, 'sect'); doc.sect.mt += 9; });
  const expected = sections(doc.pkg).map(e => content(e, doc.pkg));
  const margin = expected.at(-1)[3].find(e => e[1] === 'pgMar');
  margin[2].find(a => a[1] === 'top')[2] = String(Math.round(doc.sect.mt * 20));
  const edited = await save('edited');
  assert(sectionValues(edited.pkg) === JSON.stringify(expected), 'A margin edit changed another section property');
  assert(D.undo(), 'Missing undo'); const undo = await save('undo');
  assert(sectionValues(undo.pkg) === original, 'Undo did not restore the section');
  assert(D.redo(), 'Missing redo'); const redo = await save('redo');
  assert(sectionValues(redo.pkg) === JSON.stringify(expected), 'Redo did not restore the margin edit');
  const blob = await window.__corpusHooks.snapshot(doc), info = window.__corpusHooks.docInfo(doc);
  await save('draft', doc, blob);
  VO.opened = () => {};
  await window.__corpusHooks.open(new File([blob], 'sections.docx'), { draft: true, name: info.name, saved: true, lossState: info.lossState });
  const recovered = await save('recovered', window.__corpusHooks.current());
  assert(sectionValues(recovered.pkg) === JSON.stringify(expected), 'Draft recovery changed section settings');
  assert(!!repairLoss(window.__corpusHooks.current(), 'table-grid').length === repairedGrid, 'Draft recovery lost the unacknowledged table notice');
  assert(!!repairLoss(window.__corpusHooks.current(), 'default-styles').length === repairedStyle, 'Draft recovery lost the unacknowledged style notice');
  if (fixture) {
    let current = window.__corpusHooks.current();
    const atEnd = () => { const p = current.main.blocks.at(-1), pos = D.pos(p, Math.min(5, D.plen(p))); L.ed.sel = { a: pos, f: pos }; };
    const clickOK = () => document.querySelector('.dlg .dlg-foot .primary').click();
    const revisionIds = pkg => Array.from(pkg.xml(pkg.main).getElementsByTagName('*')).filter(e => e.localName === 'sectPrChange').map(e => Array.from(e.attributes).find(a => a.localName === 'id').value);
    const uniqueSections = (saved, count) => {
      const ids = revisionIds(saved.pkg);
      assert(sections(saved.pkg).length === count, 'Unexpected section count');
      assert(ids.length === count && new Set(ids).size === count, 'Copied sections share revision identities: ' + ids.join(', '));
    };
    L.mailmerge.data = { fields: ['Name'], rows: [['First'], ['Second']], included: new Set([0, 1]) };
    uniqueSections(await save('merged-letters', L.mailmerge.buildMerged()), 4);
    atEnd(); L.fields.insertGenerated('INDEX', { columns: 2 });
    uniqueSections(await save('index-columns', current), 4);
    // Generated fields also enqueue page-number refresh; return to the saved source.
    current = (await L.docx.read(Uint8Array.from(atob(artifacts.recovered), c => c.charCodeAt(0)))).doc;
    L.app.loadDoc(current, 'Section commands', { saved: true });
    atEnd();
    L.dlg.pageNumbers('format');
    document.querySelector('#pf-start').value = '5'; clickOK();
    assert(current.sect.pgNum.start === 5 && current.sect.pgNum.chapStyle === '2' && current.sect.pgNum.chapSep === 'hyphen', 'Page Number Format changed chapter settings');
    const numbering = await save('number-dialog', current);
    assert(numbering.sect.pgNum.chapStyle === '2' && numbering.sect.pgNum.chapSep === 'hyphen', 'Chapter settings changed in saved numbering');
    assert(D.undo(), 'Missing numbering undo');
    assert(current.sect.pgNum.start === 1, 'Numbering undo lost the original start');
    assert(D.redo(), 'Missing numbering redo');
    await save('number-redo', current);
    const beforeSplit = sectionValues((await save('before-split', current)).pkg);
    atEnd(); L.app.insertSectionBreak('continuous');
    const split = await save('split', current); uniqueSections(split, 3);
    const splitIds = revisionIds(split.pkg);
    assert(splitIds.includes('10') && splitIds.includes('11'), 'Original revision identities changed');
    assert(D.undo(), 'Missing section undo');
    assert(sectionValues((await save('split-undo', current)).pkg) === beforeSplit, 'Section undo changed original settings');
    assert(D.redo(), 'Missing section redo');
    assert(JSON.stringify(revisionIds((await save('split-redo', current)).pkg)) === JSON.stringify(splitIds), 'Redo changed copied identities');
    const draft = await window.__corpusHooks.snapshot(current), metadata = window.__corpusHooks.docInfo(current);
    await save('split-draft', current, draft);
    await window.__corpusHooks.open(new File([draft], 'sections.docx'), { draft: true, name: metadata.name, saved: true, lossState: metadata.lossState });
    current = window.__corpusHooks.current();
    const restored = await save('split-recovered', current); uniqueSections(restored, 3);
    assert(JSON.stringify(revisionIds(restored.pkg)) === JSON.stringify(splitIds), 'Draft recovery changed copied identities');
    // Both commands can add a section or reuse an existing boundary.
    for (const name of ['columns', 'pageSetup']) {
      const count = D.sections(current).length;
      atEnd(); L.ed.enter();
      for (const attempt of ['new', 'existing']) {
        atEnd(); L.dlg[name]();
        document.querySelector(name === 'columns' ? '#cd-apply' : '#ps-apply').value = 'forward';
        clickOK();
        uniqueSections(await save(name + '-' + attempt, current), count + 1);
      }
    }
  }
  return { artifacts, checks };
}
