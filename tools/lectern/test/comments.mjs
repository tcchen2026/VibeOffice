// Comment ownership, independently checked author/anchor references, history and suite drafts.
// node tools/lectern/test/comments.mjs OUTPUT_DIR FIXTURE_DIR
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [out, fixtures] = process.argv.slice(2);
if (!out || !fixtures) throw Error('Usage: comments.mjs OUTPUT_DIR FIXTURE_DIR');
const corpus = path.join(process.env.VO_CORPORA || path.join(os.homedir(), 'corpora'), 'powerpoint');
const legacy = 'poi__788ee515a5cd__45545_Comment.pptx';
const cases = [
  ...['copy', 'delete', 'paste', 'draft'].map(scenario => ({ file: legacy, scenario })),
  { file: 'openxml-sdk__5085cad8803f__[HC]viewPr-PresentationViewProperties-showComments-1.pptx', scenario: 'strict-copy' },
  ...['save', 'copy', 'copy-shape', 'delete-shape', 'copy-delete-shape', 'text', 'paste', 'draft'].map(scenario => ({ file: 'modern-comments.pptx', scenario: 'modern-' + scenario })),
  { file: 'modern-comments-ordered.pptx', scenario: 'modern-order', edit: 'save' },
  { file: 'modern-comments-no-creation.pptx', scenario: 'modern-creation', edit: 'save' },
];
fs.mkdirSync(out, { recursive: true });
const match = new RegExp(process.argv[4] || ''), selected = cases.filter(c => match.test(c.scenario));
const results = path.join(out, 'results.jsonl');
const rows = fs.existsSync(results) ? fs.readFileSync(results, 'utf8').trim().split('\n').filter(Boolean).map(JSON.parse).filter(r => !selected.some(c => c.scenario === r.scenario)) : [];
const stop = await ensureServer(), page = await openPage('lectern', { persistence: false });
try {
  for (const c of selected) {
    const row = { ...c, attempted: true }; page.errors.length = 0;
    try {
      const input = path.join(c.file.startsWith('modern-comments') ? fixtures : corpus, c.file);
      const result = await page.evaluate(`(${run.toString()})(${JSON.stringify({ ...c, data: fs.readFileSync(input).toString('base64') })})`);
      for (const [state, data] of Object.entries(result.artifacts)) {
        const dir = path.join(out, state === 'saved' ? c.scenario : c.scenario + '-' + state);
        fs.mkdirSync(dir, { recursive: true }); fs.writeFileSync(path.join(dir, c.file), Buffer.from(data, 'base64'));
      }
      delete result.artifacts; Object.assign(row, result);
      if (page.errors.length) throw Error(page.errors.join('\n'));
      if (c.scenario === 'modern-draft') {
        const shot = await page.send('Page.captureScreenshot', { format: 'png' });
        fs.writeFileSync(path.join(out, 'comment-draft.png'), Buffer.from(shot.data, 'base64'));
      }
    } catch (error) { Object.assign(row, { status: 'failed', error: error.stack }); }
    rows.push(row); fs.writeFileSync(path.join(out, 'results.jsonl'), rows.map(r => JSON.stringify(r)).join('\n') + '\n');
    console.log(c.scenario + ': ' + row.status + (row.error ? ' ' + row.error : ''));
  }
} finally { await page.close(); stop(); }
if (rows.some(r => r.status !== 'ok')) process.exitCode = 1;

async function run(o) {
  const K = L.opc, N = K.NS, H = L.hist, A = L.app, M = L.model;
  const all = e => e ? [e, ...e.getElementsByTagName('*')] : [];
  const kids = (e, name) => Array.from(e?.children || []).filter(c => !name || c.localName === name);
  const check = (ok, msg) => { if (!ok) throw Error(msg); };
  const scenario = o.edit || o.scenario.replace(/^(modern|strict)-/, ''), modern = o.scenario.startsWith('modern-');
  const artifacts = {}, bytes = Uint8Array.from(atob(o.data), c => c.charCodeAt(0));
  VO.opened = () => {}; H.clear(); await __corpusHooks.open(new File([bytes], o.file));
  const source = L.pres, originals = source.pkg, initial = source.slides.length;
  const counted = () => L.pres.slides.reduce((n, s) => n + (s.keep?.comments || []).length, 0);
  let expected = counted();
  const originalCounts = expected;
  async function inspect(blob) {
    const pkg = await K.open(new Uint8Array(await blob.arrayBuffer())), main = pkg.xml(pkg.main);
    const mr = pkg.rels(pkg.main), ids = kids(kids(main, 'sldIdLst')[0]);
    const authors = new Map(), used = new Set(); let count = 0, modernCount = 0;
    for (const r of mr.filter(r => /\/(commentAuthors|authors)$/.test(r.type))) for (const a of kids(pkg.xml(r.part))) authors.set(a.getAttribute('id'), a);
    for (const id of ids) {
      const part = mr.find(r => r.id === id.getAttributeNS(N.rel, 'id')).part;
      const slide = pkg.xml(part), rels = pkg.rels(part);
      const shapes = new Map(all(slide).filter(e => e.localName === 'cNvPr').map(e => [e.getAttribute('id'), e]));
      for (const rel of rels.filter(r => r.type.endsWith('/comments'))) {
        count++;
        const root = pkg.xml(rel.part), isModern = root.namespaceURI.includes('2018/8');
        if (isModern) {
          modernCount++;
          check(all(slide).some(e => e.localName === 'commentRel' && e.getAttributeNS(N.rel, 'id') === rel.id), 'Modern slide relationship missing');
          if (o.scenario === 'modern-order') {
            const original = originals.xml(owner.keep.part);
            for (const place of ['', 'cSld']) {
              const uris = root => kids(kids(place ? kids(root, place)[0] : root, 'extLst')[0]).map(e => e.getAttribute('uri'));
              check(JSON.stringify(uris(slide)) === JSON.stringify(uris(original)), 'Comment extensions changed sibling order');
            }
          }
        }
        for (const c of all(root).filter(e => ['cm', 'reply'].includes(e.localName))) {
          const aid = c.getAttribute('authorId'), key = isModern ? c.getAttribute('id') : aid + ':' + c.getAttribute('idx');
          check(!used.has(key), 'Comment identity repeated: ' + key); used.add(key);
          check(authors.has(aid), 'Comment author missing: ' + aid);
          if (!isModern) check(+authors.get(aid).getAttribute('lastIdx') >= +c.getAttribute('idx'), 'Author counter is behind copied comment');
          for (const aid of (c.getAttribute('assignedTo') || '').split(/\s+/).filter(Boolean)) check(authors.has(aid), 'Task assignee missing');
        }
        for (const e of all(root)) {
          if (e.localName === 'sldMk') {
            check(e.getAttribute('sldId') === id.getAttribute('id'), 'Comment anchors another slide');
            check(e.hasAttribute('cId'), 'Comment slide marker is missing its creation ID');
            check(all(slide).some(c => c.localName === 'creationId' && c.getAttribute('val') === e.getAttribute('cId')), 'Slide creation identity mismatch');
          }
          if (/^(sp|grpSp|graphicFrame|cxnSp|pic|ink)Mk$/.test(e.localName)) {
            const shape = shapes.get(e.getAttribute('id')); check(shape, 'Comment shape missing');
            if (e.hasAttribute('creationId')) check(all(shape).some(c => c.localName === 'creationId' && c.getAttribute('id') === e.getAttribute('creationId')), 'Shape creation identity mismatch');
          }
        }
      }
    }
    check(count === expected, 'Comment parts: expected ' + expected + ', got ' + count);
    if (modern) check(modernCount > 0, 'Modern threads lost');
    return pkg;
  }
  async function save(state, blob) {
    blob ||= await L.pptx.write(L.pres); const pkg = await inspect(blob);
    const data = new Uint8Array(await blob.arrayBuffer()); let text = '';
    for (let i = 0; i < data.length; i += 32768) text += String.fromCharCode(...data.subarray(i, i + 32768));
    artifacts[state] = btoa(text); return pkg;
  }
  A.view = 'normal'; A.focusArea = 'editor'; H.clear();
  const owner = source.slides.find(s => (s.keep?.comments || []).some(r => !modern || r.type.includes('/2018/')));
  check(owner, 'Source commented slide missing');
  if (scenario === 'paste') {
    A.view = 'sorter'; A.focusArea = 'slides'; A.slideSel = new Set([source.slides.indexOf(owner)]);
    const packet = JSON.parse(JSON.stringify(L.preserve.clipboard(A.copySlides(false))));
    // Import through a package with no source presentation/slide parts or author relationship.
    const clip = L.preserve.pasteboard(packet);
    check(clip.slides[0].keep.source !== source.pkg.id, 'Clipboard source was not detached');
    A.pasteItem(clip); expected += owner.keep.comments.length;
  } else if (scenario !== 'save') {
    H.push('Comment ' + scenario);
    if (['copy', 'draft', 'copy-delete-shape'].includes(scenario)) {
      const copy = M.dupSlide(owner);
      if (scenario === 'copy-delete-shape') copy.shapes = copy.shapes.filter(s => !s.keep?.commentAnchor);
      M.insertSlides(source, 1, [copy]); expected += owner.keep.comments.length;
    }
    else if (scenario === 'delete') { source.slides.splice(source.slides.indexOf(owner), 1); expected -= owner.keep.comments.length; }
    else {
      const shape = owner.shapes.find(s => s.keep?.commentAnchor); check(shape, 'Anchored shape missing');
      if (scenario === 'copy-shape') owner.shapes.push(M.dup(shape));
      else if (scenario === 'delete-shape') owner.shapes.splice(owner.shapes.indexOf(shape), 1);
      else shape.tx.ps[0].rs[0].t += ' edited';
    }
  }
  const edited = expected, expectedSlides = L.pres.slides.length;
  const saved = await save('saved');
  if (scenario === 'save') {
    if (o.scenario !== 'modern-creation') for (const r of owner.keep.comments) check(originals.text(r.part) === saved.text(r.part), 'Unchanged comments rewritten');
    for (const r of originals.rels(originals.main).filter(r => /\/(commentAuthors|authors)$/.test(r.type))) check(originals.text(r.part) === saved.text(r.part), 'Unchanged author records rewritten');
  } else {
    check(H.doUndo(), 'Undo missing'); expected = originalCounts; await save('undo');
    check(L.pres.slides.length === initial, 'Undo slide count');
    check(H.doRedo(), 'Redo missing'); expected = edited; await save('redo');
  }
  if (scenario === 'draft') {
    const info = __corpusHooks.docInfo(L.pres), blob = await __corpusHooks.snapshot(L.pres);
    await save('draft', blob); H.clear();
    await __corpusHooks.open(new File([blob], info.draftName || info.name), { draft: true, name: info.name, saved: true, lossState: info.lossState });
    await save('recovered');
  }
  if (['delete-shape', 'copy-delete-shape', 'text'].includes(scenario)) check(L.pres.losses.some(e => e.id.startsWith('comment-anchor:')), 'Anchor conversion was not reported');
  return { status: 'ok', expectedSlides, artifacts, losses: L.pres.losses };
}
