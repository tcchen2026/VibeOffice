// Build a bounded manual Office review from authored inputs and explicit edits.
// BASE can point at an isolated previous writer for the required before/after diff.
// node tools/ooxml/office-batch.mjs MANIFEST.json OUTPUT_DIR
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { ensureServer, openPage } from './cdp.mjs';
const [manifest, output] = process.argv.slice(2);
if (!manifest || !output) throw new Error('Usage: office-batch.mjs MANIFEST.json OUTPUT_DIR');
const cases = JSON.parse(fs.readFileSync(manifest, 'utf8'));
if (!cases.length || cases.length > 15) throw new Error('An Office batch has 1–15 files');
if (fs.existsSync(output) && fs.readdirSync(output).length) throw new Error('Use an empty output directory');
fs.mkdirSync(output, { recursive: true });
const helper = fs.readFileSync(new URL('./scenarios.js', import.meta.url), 'utf8');
const stop = await ensureServer(), reports = [], pages = new Map();
try {
  for (const c of cases) {
    const row = { ...c, attempted: true, officeAcceptance: 'pending' };
    try {
      if (!pages.has(c.app)) { const page = await openPage(c.app, { persistence: false }); await page.evaluate(helper); pages.set(c.app, page); }
      const page = pages.get(c.app); page.errors.length = 0;
      const bytes = fs.readFileSync(c.source); row.sourceHash = createHash('sha256').update(bytes).digest('hex');
      const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(c)}, ${JSON.stringify(bytes.toString('base64'))})`);
      if (page.errors.length) throw new Error(page.errors.join('\n'));
      const saved = Buffer.from(result.data, 'base64'); delete result.data;
      fs.writeFileSync(path.join(output, c.file), saved);
      Object.assign(row, result, { status: 'ok', savedHash: createHash('sha256').update(saved).digest('hex') });
      console.log(c.file + ': saved');
    } catch (error) { Object.assign(row, { status: 'failed', error: error.stack }); console.error(c.file + ': ' + error.message); process.exitCode = 1; }
    reports.push(row);
    fs.writeFileSync(path.join(output, 'manifest.json'), JSON.stringify(reports, null, 2) + '\n');
  }
  const lines = ['# Office batch', '', 'Open each file in its Office application. Record ✓, **repair prompt**, or **looks wrong** and any details. For macro files, review project presence without running the code. Original links point to the authored inputs outside this small batch.', '', '| File | Feature / edit | What to look for | Result |', '|---|---|---|---|'];
  for (const c of reports) lines.push(`| [${c.file}](${c.file}) · [original](${c.source.replaceAll(' ', '%20')}) | ${c.feature} | ${c.lookFor} | ${c.status === 'ok' ? 'Pending' : 'Generation failed: ' + c.error.split('\n')[0]} |`);
  fs.writeFileSync(path.join(output, 'CHECKLIST.md'), lines.join('\n') + '\n');
} finally { for (const page of pages.values()) await page.close(); stop(); }

async function save(c, data) {
  const { opc: K } = L, bytes = Uint8Array.from(atob(data), x => x.charCodeAt(0));
  const read = async bytes => c.app === 'quire' ? (await L.docx.read(bytes)).doc : c.app === 'ledger' ? L.xlsxRead.read(bytes) : L.pptx.read(bytes);
  let doc = await read(bytes), changed = [];
  if (c.app === 'quire') {
    const { D, O, preserve: P } = L;
    L.app.loadDoc(doc, 'Office review', { saved: true });
    if (['numbering-align', 'numbering-copy'].includes(c.edit)) {
      let selected;
      D.walk(doc.main, p => {
        if (selected || p.t !== 'p') return;
        const info = L.lists.info(p), abs = info && D.numDef(doc, info.numId)?.abs;
        const level = info && D.numLevelDef(doc, info.numId, info.lvl);
        if (abs && (c.numbering === 'picture' ? level?.picture : c.numbering === 'custom' ? level?.custFmt : abs.styleLink)) selected = { p, info, abs };
      });
      if (!selected) throw new Error('Missing ' + c.numbering + ' list');
      const pos = D.pos(selected.p, 0); L.ed.sel = { a: pos, f: pos };
      if (c.edit === 'numbering-copy') L.lists.apply(selected.abs, { newList: true });
      else {
        L.dlg.customizeList(selected.abs, selected.info.lvl);
        document.querySelector('#cl-al').value = 'right';
        document.querySelector('.dlg .dlg-foot .primary').click();
        if (c.copyAfter) L.lists.apply(D.numDef(doc, L.lists.info(selected.p).numId).abs, { newList: true });
      }
      changed.push(c.edit === 'numbering-copy' ? 'Started an independent list at the selected paragraph' : 'Right-aligned the selected list level');
      if (c.copyAfter) changed.push('Copied the edited list as an independent definition');
    } else if (c.edit === 'index-columns') {
      const p = doc.main.blocks[1], pos = D.pos(p, Math.min(5, D.plen(p))); L.ed.sel = { a: pos, f: pos };
      L.fields.insertGenerated('INDEX \\c "2"', { columns: 2 });
      changed.push('Inserted a two-column index in the second section');
    } else if (c.edit === 'section-review') {
      D.tx('Change top margin', () => { D.touchKey(doc, 'sect'); doc.sect.mt += 9; });
      const p = doc.main.blocks.at(-1), pos = D.pos(p, Math.min(5, D.plen(p))); L.ed.sel = { a: pos, f: pos };
      L.dlg.pageNumbers('format'); document.querySelector('#pf-start').value = '5';
      document.querySelector('.dlg .dlg-foot .primary').click();
      L.app.insertSectionBreak('continuous');
      changed.push('Increased the last section top margin by 9 points, restarted numbering at 5 and inserted a continuous section');
    } else if (c.edit === 'binding') {
      const record = [...P.controls(doc).records.values()].find(r => r.complete && r.control.binding && r.control.type === 'text');
      if (!record) throw new Error('No editable bound text control');
      D.tx('Edit bound value', () => O.insertText(D.pos(record.start.pos.p, record.start.pos.o + 1), ' EDITED'));
      changed.push('Inserted EDITED into the bound text');
    } else if (c.edit === 'paste') {
      const blocks = D.cloneBlocks(doc.main.blocks), imported = await P.hydrateClipboard(K.remapSources(blocks, K.import(P.clipboard(blocks))));
      doc = D.newDoc(); L.app.loadDoc(doc, 'Pasted object', { saved: true });
      D.tx('Paste object', () => O.insertBlocks(D.pos(doc.main.blocks[0], 0), imported));
      changed.push('Pasted all source blocks into a new document');
    } else if (['picture-properties', 'shape-shadow', 'drawing-anchor'].includes(c.edit)) {
      let owner, object;
      D.walk(doc.main, p => { if (p.t === 'p' && !object) { const it = p.runs.find(it => c.edit === 'shape-shadow' ? it.t === 'shape' : it.t === 'img'); if (it) { owner = p; object = it; } } });
      if (!object) throw new Error('No drawing for ' + c.edit);
      D.tx('Edit drawing properties', () => {
        D.touch(owner);
        if (c.edit === 'picture-properties') { object.crop = { l: 0.15, t: 0.05, r: 0, b: 0 }; object.border = { val: 'single', sz: 1.5, color: '123456' }; object.alt = 'Edited picture'; }
        else if (c.edit === 'shape-shadow') object.shadow = { c: '#123456', a: 0.6, dx: 5, dy: 3, blur: 2 };
        else object.float = { ...(object.float || {}), wrap: 'square', posH: { rel: 'column', off: 12 }, posV: { rel: 'paragraph', off: 10 }, allowOverlap: false };
      });
      changed.push(c.edit === 'picture-properties' ? 'Cropped the picture and changed its border and alternative text' : c.edit === 'shape-shadow' ? 'Changed the outer shadow' : 'Changed the drawing anchor and wrapping');
    } else if (c.edit === 'watermark') {
      D.tx('Replace watermark', () => { D.touchKey(doc, 'watermark'); doc.watermark = { type: 'text', text: 'REVISED', color: 'C0C0C0', layout: 'diagonal' }; });
      changed.push('Replaced the watermark with REVISED');
    } else if (c.edit === 'wordart') {
      let owner, shape;
      D.walk(doc.main, p => { if (p.t === 'p' && !shape) { const s = p.runs.find(r => r.wordart); if (s) { owner = p; shape = s; } } });
      if (!shape) throw new Error('No editable WordArt');
      D.tx('Edit WordArt', () => { D.touch(owner); shape.wordart.text += ' EDITED'; });
      changed.push('Appended EDITED to WordArt');
    } else if (c.edit === 'revision-link') {
      let owner, run;
      D.walk(doc.main, p => { if (p.t === 'p' && !run) { const r = p.runs.find(r => r.t === 'text' && (r.rPr.ins || r.rPr.del)); if (r) { owner = p; run = r; } } });
      if (!run) throw new Error('No tracked text');
      D.tx('Link tracked text', () => { D.touch(owner); run.rPr.link = { url: 'https://example.invalid/review' }; });
      changed.push('Added a hyperlink to tracked text');
    } else if (c.edit && c.edit !== 'save') changed = corpusScenario(c.app, doc, c.edit).changed;
  } else if (c.app === 'ledger' && c.steps) {
    const O = L.ops;
    for (const step of c.steps) {
      const table = step.table == null ? null : doc.tables[step.table];
      const sh = table?.sheet || (typeof step.sheet === 'string' ? doc.sheetByName(step.sheet) : doc.sheets[step.sheet || 0]);
      if (!sh) throw new Error('Missing worksheet for ' + step.kind);
      if (step.kind === 'lines') O.insertLines(sh, step.axis, table ? table.ref[step.axis + '1'] + step.offset : step.at, step.count);
      else if (step.kind === 'value') O.tx(doc, 'Review cell edit', () => O.put(sh, step.row, step.col, { ...sh.get(step.row, step.col), v: step.value, f: undefined }));
      else if (step.kind === 'copy-sheet') { O.copySheet(doc, sh, doc.sheets.length); if (step.deleteSource) O.deleteSheet(doc, sh); }
      else if (step.kind === 'copy-control') {
        const source = sh.drawings.find(d => d.objectKeep); if (!source) throw new Error('Missing control');
        for (let i = 0; i < (step.count || 1); i++) {
          const copy = L.sheetObjects.copy(source); copy.id = Math.max(...sh.drawings.map(d => d.id || 0)) + 1;
          if (copy.anchor.from) { copy.anchor.from.r += 5 * (i + 1); copy.anchor.to.r += 5 * (i + 1); }
          O.tx(doc, 'Copy control', () => O.setDrawings(sh, sh.drawings.concat(copy)));
        }
      }
      else if (step.kind === 'delete-sheet') O.deleteSheet(doc, sh);
      else if (step.kind === 'note') O.tx(doc, 'Review note', () => O.setComment(sh, step.row, step.col, { r: step.row, c: step.col, author: 'Review', text: step.text }));
      else if (step.kind === 'thread-text') {
        const cm = [...sh.comments.values()].find(cm => cm.thread?.some(t => !t.parent));
        if (!cm) throw new Error('Missing threaded comment');
        const root = cm.thread.find(t => !t.parent), text = step.text ?? (root.text + (step.append || ' EDITED'));
        O.tx(doc, 'Review thread root', () => O.setComment(sh, cm.r, cm.c, { ...cm, text, runs: undefined }));
      }
      else if (step.kind === 'bar') {
        const cf = L.clone(sh.cf), rule = cf.flatMap(c => c.rules).find(r => r.bar);
        if (!rule) throw new Error('Missing data bar');
        Object.assign(rule.bar, step.properties); O.tx(doc, 'Review data bar', () => O.setCF(sh, cf));
      } else throw new Error('Unknown review step ' + step.kind);
      changed.push(step);
    }
  } else if (c.app === 'lectern' && c.steps) {
    const M = L.model;
    L.pres = doc; L.hist.clear();
    const all = e => e ? [e, ...Array.from(e.children).flatMap(all)] : [];
    const sourceShape = sh => {
      const f = sh.keep?.identity, ref = f?.ids.find(r => r.definition && r.kind === 'shape');
      if (!ref || !doc.pkg) return '';
      const cnv = all(doc.pkg.xml(f.part)).find(e => e.localName === 'cNvPr' && e.getAttribute('id') === String(ref.id));
      return cnv ? K.raw(cnv.parentNode.parentNode) : '';
    };
    for (const step of c.steps) {
      const slide = doc.slides[step.slide || 0];
      if (!slide) throw new Error('Missing slide for ' + step.kind);
      let shape, list;
      const select = shapes => {
        for (const sh of shapes) {
          if (!shape && (!step.frame || sh.keep?.frame) && (!step.needle || sourceShape(sh).includes(step.needle))) { shape = sh; list = shapes; }
          if (!shape && sh.kids) select(sh.kids);
        }
      };
      if (['move-shape', 'copy-shape', 'paste-shape', 'shape-properties', 'bold'].includes(step.kind)) {
        select(slide.shapes); if (!shape) throw new Error('No matching shape for ' + step.kind);
      }
      if (step.kind === 'copy-slide') {
        const copy = M.dupSlide(slide);
        if (M.insertSlides) M.insertSlides(doc, step.at ?? 1, [copy]); else doc.slides.splice(step.at ?? 1, 0, copy);
      } else if (step.kind === 'insert-slide') {
        const added = M.newSlide(doc, 'blank', slide.design);
        if (M.insertSlides) M.insertSlides(doc, step.at ?? 1, [added]); else doc.slides.splice(step.at ?? 1, 0, added);
      } else if (step.kind === 'delete-slide') doc.slides.splice(step.slide || 0, 1);
      else if (step.kind === 'move-shape') M.translate(shape, step.dx || 12, step.dy || 8);
      else if (step.kind === 'copy-shape') list.push(M.dup(shape));
      else if (step.kind === 'paste-shape') {
        const item = L.preserve.pasteboard(JSON.parse(JSON.stringify(L.preserve.clipboard({ kind: 'shapes', shapes: [L.clone(shape)], text: '' }))));
        for (const sh of item.shapes) { M.translate(sh, 18, 12); slide.shapes.push(sh); }
      } else if (step.kind === 'shape-properties') Object.assign(shape, step.properties);
      else if (step.kind === 'bold') {
        const run = shape.tx?.ps.flatMap(p => p.rs).find(r => r.t); if (!run) throw new Error('No text run'); run.b = !run.b;
      } else if (step.kind === 'design-copy') {
        const copy = K.duplicate(doc.designs[slide.design]); copy.id = L.uid('dsn'); copy.colors.accent1 = '#AA2244';
        doc.designs[copy.id] = copy; slide.design = copy.id;
      } else if (step.kind === 'design-decoration') {
        const design = M.design(doc, slide), layout = step.layout && design.keep.layoutParts[0];
        const list = layout ? design.layoutDecos[layout.lkey] : design.deco;
        const decoration = list.find(s => s.name === (step.name || 'Review decoration 1'));
        if (!decoration) throw new Error('Missing design decoration');
        M.translate(decoration, step.dx || 12, step.dy || 8);
      } else if (step.kind === 'ungroup-frame') {
        M.walk(slide.shapes, sh => { if (!shape && (step.type === 'chart' ? sh.type === 'chart' : sh.keep?.frame?.label === 'Embedded object')) shape = sh; return true; });
        if (!shape) throw new Error('Missing graphic frame');
        L.ed.goto(doc.slides.indexOf(slide), { force: true }); L.app.view = 'normal'; L.app.focusArea = 'editor';
        const list = L.ed.find(shape.id).list;
        const group = { id: L.uid('group'), type: 'group', name: 'Rotated group', x: shape.x, y: shape.y, w: shape.w, h: shape.h, rot: 45, kids: [shape] };
        list.splice(list.indexOf(shape), 1, group); L.ed.select([group.id]); L.hist.clear(); L.app.ungroup();
      } else if (step.kind === 'transition') Object.assign(slide.trans, step.properties);
      else if (step.kind === 'animation') {
        if (!slide.anims.length) throw new Error('No modeled animation'); slide.anims[0].dur += 200;
      } else throw new Error('Unknown review step ' + step.kind);
      changed.push(step);
    }
  } else if (c.edit && c.edit !== 'save') changed = corpusScenario(c.app, doc, c.edit).changed;
  const format = c.file.split('.').pop();
  const blob = await (c.app === 'quire' ? L.docx.write(doc, { format }) : c.app === 'ledger' ? L.xlsxWrite.write(doc, { type: format }) : L.pptx.write(doc, { format }));
  const saved = new Uint8Array(await blob.arrayBuffer());
  await read(saved);
  let encoded = ''; for (let i = 0; i < saved.length; i += 32768) encoded += String.fromCharCode(...saved.subarray(i, i + 32768));
  return { data: btoa(encoded), changed, losses: doc.losses || [], ...(c.app === 'lectern' ? { expectedSlides: doc.slides.length } : {}) };
}
