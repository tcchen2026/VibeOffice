// Design edits, detached clipboard transfer and suite draft recovery.
// node tools/lectern/test/designs.mjs OUTPUT_DIR [SCENARIO_REGEX]
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [out, filter = ''] = process.argv.slice(2);
if (!out) throw Error('Usage: designs.mjs OUTPUT_DIR [SCENARIO_REGEX]');
const corpus = path.join(process.env.VO_CORPORA || path.join(os.homedir(), 'corpora'), 'powerpoint');
const cases = ['color', 'font', 'background', 'geometry', 'style', 'copy-design', 'new-layout', 'delete-slide', 'draft', 'footer', 'title-align'].map(edit => ({
  file: 'libreoffice__07eb8b76b2cf__master-slides.pptx', edit, scenario: edit,
}));
cases.push(...['decoration-move', 'decoration-delete', 'decoration-reorder', 'decoration-add', 'layout-decoration-move'].map(edit => ({ file: 'unread-decorations.pptx', edit, scenario: edit, fixture: true })));
const selected = cases.filter(c => new RegExp(filter).test(c.scenario));
fs.mkdirSync(out, { recursive: true });
if (selected.some(c => c.fixture)) execFileSync('python3', [fileURLToPath(new URL('./design-fixtures.py', import.meta.url)), path.join(out, 'fixtures'), '--corpus', corpus]);
const results = path.join(out, 'results.jsonl');
const rows = fs.existsSync(results) ? fs.readFileSync(results, 'utf8').trim().split('\n').filter(Boolean).map(JSON.parse).filter(r => !selected.some(c => c.scenario === r.scenario)) : [];
const stop = await ensureServer(), page = await openPage('lectern', { persistence: false });
try {
  for (const c of selected) {
    const row = { ...c, attempted: true }; page.errors.length = 0;
    try {
      const input = path.join(c.fixture ? path.join(out, 'fixtures') : corpus, c.file);
      const result = await page.evaluate(`(${run.toString()})(${JSON.stringify({ ...c, data: fs.readFileSync(input).toString('base64') })})`);
      for (const [state, data] of Object.entries(result.artifacts)) {
        const dir = path.join(out, state === 'saved' ? c.scenario : c.scenario + '-' + state);
        fs.mkdirSync(dir, { recursive: true }); fs.writeFileSync(path.join(dir, c.file), Buffer.from(data, 'base64'));
      }
      delete result.artifacts; Object.assign(row, result);
      if (page.errors.length) throw Error(page.errors.join('\n'));
      if (c.edit === 'draft' || c.edit === 'paste') {
        const shot = await page.send('Page.captureScreenshot', { format: 'png' });
        fs.writeFileSync(path.join(out, c.scenario + '.png'), Buffer.from(shot.data, 'base64'));
      }
    } catch (error) { Object.assign(row, { status: 'failed', error: error.stack }); }
    rows.push(row); fs.writeFileSync(results, rows.map(r => JSON.stringify(r)).join('\n') + '\n');
    console.log(c.scenario + ': ' + row.status + (row.error ? ' ' + row.error : ''));
  }
} finally { await page.close(); stop(); }
if (rows.some(r => r.status !== 'ok')) process.exitCode = 1;

async function run(o) {
  const K = L.opc, M = L.model, A = L.app, H = L.hist, E = L.ed;
  const check = (ok, why) => { if (!ok) throw Error(why); };
  const all = e => e ? [e, ...e.getElementsByTagName('*')] : [];
  const expanded = e => e && JSON.stringify([e.namespaceURI, e.localName, Array.from(e.attributes).filter(a => a.name !== 'xmlns' && a.prefix !== 'xmlns').map(a => [a.namespaceURI, a.localName, a.value]).sort(), Array.from(e.children).map(expanded), e.children.length ? '' : e.textContent]);
  const count = pkg => {
    const result = { masters: 0, layouts: 0 };
    for (const name of pkg.names.filter(n => /\.xml$/.test(n))) {
      const tag = pkg.xml(name)?.localName;
      if (tag === 'sldMaster') result.masters++;
      if (tag === 'sldLayout') result.layouts++;
    }
    return result;
  };
  VO.opened = () => {}; H.clear();
  await __corpusHooks.open(new File([Uint8Array.from(atob(o.data), c => c.charCodeAt(0))], o.file));
  const pres = L.pres, original = pres.pkg, before = count(original), artifacts = {}, slide = pres.slides[0], design = M.design(pres, slide);
  E.goto(0, { force: true }); A.view = 'normal'; A.focusArea = 'editor'; H.clear();
  const originalLayouts = original.names.filter(n => /\/slideLayouts\/[^/]+\.xml$/.test(n));
  const masterPart = design.keep.part, themePart = design.keep.theme;
  const record = o.edit.startsWith('layout-') ? design.keep.layoutParts[0] : null;
  const decorationPart = record?.part || masterPart;
  const decorationList = record ? design.layoutDecos[record.lkey] : design.deco;
  const decoration = decorationList.find(s => s.name === 'Review decoration 1');
  let expected = { ...before }, edited = true;
  async function save(state, blob) {
    blob ||= await L.pptx.write(L.pres);
    const pkg = await K.open(new Uint8Array(await blob.arrayBuffer()));
    check(JSON.stringify(count(pkg)) === JSON.stringify(expected), 'Design count changed: ' + JSON.stringify(count(pkg)));
    for (const part of pkg.names.filter(n => /\/slideLayouts\/[^/]+\.xml$/.test(n) && !originalLayouts.includes(n))) {
      const ids = all(pkg.xml(part)).filter(e => e.namespaceURI === 'http://schemas.openxmlformats.org/presentationml/2006/main' && e.localName === 'cNvPr').map(e => e.getAttribute('id'));
      check(new Set(ids).size === ids.length, 'New layout has duplicate shape IDs: ' + part);
    }
    for (const part of originalLayouts.filter(p => !edited || !['footer', 'title-align'].includes(o.edit) && !(o.edit.startsWith('layout-') && p === decorationPart))) check(pkg.text(part) === original.text(part), 'Unrelated layout changed: ' + part);
    if (o.fixture) {
      const children = pkg => [...all(pkg.xml(decorationPart)).find(e => e.localName === 'spTree').children];
      const label = e => all(e).find(e => e.localName === 'cNvPr')?.getAttribute('name');
      const old = children(original), current = children(pkg);
      const unread = old.find(e => label(e) === 'Unread decoration');
      check(unread && expanded(unread) === expanded(current.find(e => label(e) === 'Unread decoration')), 'Unread decoration changed or disappeared');
      for (const e of old.filter(e => !['Review decoration 1', 'Review decoration 3'].includes(label(e)))) check(expanded(e) === expanded(current.find(n => label(n) === label(e) && n.localName === e.localName)), 'Unrelated design child changed: ' + label(e));
      const names = old.map(label), actual = current.map(label);
      if (edited && o.edit === 'decoration-delete') names.splice(names.indexOf('Review decoration 1'), 1);
      if (edited && o.edit === 'decoration-reorder') { const a = names.indexOf('Review decoration 1'), b = names.indexOf('Review decoration 3'); [names[a], names[b]] = [names[b], names[a]]; }
      if (edited && o.edit === 'decoration-add') names.push('Added decoration');
      check(JSON.stringify(names) === JSON.stringify(actual), 'Unrelated stacking order changed');
      if (edited && o.edit.endsWith('-move')) check(Math.abs(K.getBox(K.raw(current.find(e => label(e) === decoration.name))).x - decoration.x) < .01, 'Decoration move missing');
    }
    if (!edited || ['color', 'font', 'copy-design', 'delete-slide', 'draft'].includes(o.edit)) check(pkg.text(masterPart) === original.text(masterPart), 'Unrelated master changed');
    if (edited && ['color', 'font', 'copy-design', 'draft'].includes(o.edit)) {
      const theme = pkg.rels(masterPart).find(r => /\/theme$/.test(r.type));
      if (o.edit !== 'copy-design') {
        const tree = pkg.xml(theme.part), old = original.xml(themePart);
        check(expanded(all(tree).find(e => e.localName === 'fmtScheme')) === expanded(all(old).find(e => e.localName === 'fmtScheme')), 'Theme effects were rebuilt');
        if (o.edit === 'font') check(all(tree).find(e => e.localName === 'minorFont').getElementsByTagName('a:latin')[0].getAttribute('typeface') === 'Courier New', 'Font edit missing');
        else check(all(all(tree).find(e => e.localName === 'accent1'))[1].getAttribute('val') === 'AA2244', 'Color edit missing');
      }
    }
    if (edited && o.edit === 'footer') {
      const foot = all(pkg.xml(masterPart)).find(e => e.localName === 'ph' && ['ftr', 'dt', 'sldNum'].includes(e.getAttribute('type')))?.parentNode?.parentNode?.parentNode;
      check(all(foot).some(e => e.localName === 'latin' && e.getAttribute('typeface') === 'Courier New'), 'Footer font edit missing');
    }
    if (edited && o.edit === 'geometry') {
      const shape = all(pkg.xml(masterPart)).find(e => e.localName === 'cNvPr' && e.getAttribute('name') === design.deco[0].name)?.parentNode?.parentNode;
      check(Math.abs(K.getBox(K.raw(shape)).x - design.deco[0].x) < .01, 'Master decoration move missing');
    }
    if (edited && o.edit === 'style') {
      const style = all(pkg.xml(masterPart)).find(e => e.localName === 'titleStyle');
      check(all(style).some(e => e.localName === 'defRPr' && e.getAttribute('b') === '1'), 'Style edit missing');
    }
    const data = new Uint8Array(await blob.arrayBuffer()); let text = '';
    for (let i = 0; i < data.length; i += 32768) text += String.fromCharCode(...data.subarray(i, i + 32768));
    artifacts[state] = btoa(text); return pkg;
  }
  H.push('Design ' + o.edit);
  if (o.fixture) {
    check(decoration, 'Fixture decoration missing');
    const i = decorationList.indexOf(decoration), j = decorationList.findIndex(s => s.name === 'Review decoration 3');
    if (o.edit.endsWith('-move')) M.translate(decoration, 12, 8);
    else if (o.edit === 'decoration-delete') decorationList.splice(i, 1);
    else if (o.edit === 'decoration-reorder') [decorationList[i], decorationList[j]] = [decorationList[j], decorationList[i]];
    else { const copy = M.dup(decoration); copy.name = 'Added decoration'; decorationList.push(copy); }
  } else if (o.edit === 'font') design.fonts.minor = 'Courier New';
  else if (o.edit === 'background') design.bg = { t: 'solid', c: '#225588' };
  else if (o.edit === 'geometry') { check(design.deco.length, 'Missing decoration'); M.translate(design.deco[0], 12, 8); }
  else if (o.edit === 'footer') design.footer.rPr.font = 'Courier New';
  else if (o.edit === 'title-align') design.ctrTitleAlgn = 'r';
  else if (o.edit === 'style') design.tx.title[0].rPr.b = true;
  else if (o.edit === 'copy-design') {
    const copy = K.duplicate(design); copy.id = L.uid('dsn'); copy.colors.accent1 = '#AA2244'; pres.designs[copy.id] = copy; slide.design = copy.id;
    expected.masters++; expected.layouts += design.keep.layoutParts.length;
  } else if (o.edit === 'new-layout') {
    const layout = M.LAYOUTS.find(l => !design.keep.layouts[l.key]);
    check(layout, 'Fixture already has every modeled layout');
    M.insertSlides(pres, 1, [M.newSlide(pres, layout.key, design.id)]); expected.layouts++;
  } else if (o.edit === 'delete-slide') { pres.slides.splice(0, 1); M.gcDesigns(pres); }
  else design.colors.accent1 = '#AA2244';
  const after = { ...expected };
  await save('saved'); check(H.doUndo(), 'Undo missing'); expected = before; edited = false; await save('undo');
  check(H.doRedo(), 'Redo missing'); expected = after; edited = true; await save('redo');
  if (o.edit === 'draft') {
    const info = __corpusHooks.docInfo(L.pres), blob = await __corpusHooks.snapshot(L.pres); await save('draft', blob);
    H.clear(); await __corpusHooks.open(new File([blob], info.draftName || info.name), { draft: true, name: info.name, saved: true, lossState: info.lossState });
    await save('recovered');
  }
  return { status: 'ok', artifacts, before, after, expectedSlides: L.pres.slides.length, losses: L.pres.losses };
}
