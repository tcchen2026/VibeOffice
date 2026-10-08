/* Test scenarios, shared by the Node workbook driver and browser corpus drivers.
 * Mutations use the app's history/operations, not a reconstructed test model. */
(function (root) {
  const L = root.L;
  const skip = message => { throw Object.assign(new Error(message), { excluded: true }); };
  const walkShapes = (shapes, out = []) => { for (const s of shapes) { out.push(s); if (s.kids) walkShapes(s.kids, out); } return out; };
  root.corpusScenario = function (app, doc, scenario) {
    if (['save', 'repeat', 'draft'].includes(scenario)) return { changed: [] };
    if (app === 'quire') {
      const D = L.D;
      D.doc = doc; D.resetHistory();
      let p;
      const objectScenario = ['geometry', 'delete-object', 'duplicate'].includes(scenario);
      D.walk(doc.main, b => { if (!p && b.t === 'p' && b.runs.some(r => objectScenario ? ['img', 'shape', 'group', 'chart'].includes(r.t) : r.t === 'text')) p = b; });
      if (!p && scenario === 'duplicate') D.walk(doc.main, b => { if (!p && b.t === 'p' && b.runs.some(r => r.t === 'text')) p = b; });
      if (!p) skip('No editable paragraph');
      const object = p.runs.find(r => ['img', 'shape', 'group', 'chart'].includes(r.t));
      if (['geometry', 'delete-object'].includes(scenario) && !object) skip('Selected paragraph has no drawing');
      if (!['text', 'format', 'geometry', 'duplicate', 'delete-object'].includes(scenario)) skip('Scenario does not apply to Quire');
      const applied = D.tx('Corpus ' + scenario, () => {
        D.touch(p);
        if (scenario === 'text') L.O.insertText(D.pos(p, D.plen(p)), ' LOSSLESS-CHECK');
        if (scenario === 'format') { const r = p.runs.find(r => r.t === 'text'); r.rPr.b = !r.rPr.b; }
        if (scenario === 'geometry') object.w = (object.w || 36) + 6;
        if (scenario === 'delete-object') p.runs.splice(p.runs.indexOf(object), 1);
        if (scenario === 'duplicate') { D.touchList(doc.main); doc.main.blocks.push(...D.cloneBlocks([p])); }
      });
      if (applied === false) skip('Content-control lock prevents this edit');
      return { changed: [{ paragraph: p.id, kind: scenario }], undo: () => D.undo(), redo: () => D.redo() };
    }
    if (app === 'ledger') {
      const sh = doc.sheets.find(s => s.kind === 'worksheet');
      if (!sh) skip('No worksheet');
      const O = L.ops;
      if (!['text', 'format', 'insert-row', 'delete-row', 'insert-column', 'delete-column', 'rename-sheet', 'delete-sheet', 'geometry', 'duplicate', 'delete-object'].includes(scenario)) skip('Scenario does not apply to Ledger');
      if (['geometry', 'duplicate', 'delete-object'].includes(scenario) && !sh.drawings.length) skip('No worksheet drawing');
      if (scenario === 'delete-sheet' && doc.sheets.length < 2) skip('Cannot delete the only sheet');
      O.tx(doc, 'Corpus ' + scenario, () => {
        if (scenario === 'text') O.enter(sh, 0, 0, 'LOSSLESS-CHECK');
        if (scenario === 'format') O.format(sh, [{ r1: 0, c1: 0, r2: 0, c2: 0 }], { font: { b: true } });
        if (scenario === 'insert-row' || scenario === 'delete-row') O.insertLines(sh, 'r', 1, scenario === 'insert-row' ? 1 : -1);
        if (scenario === 'insert-column' || scenario === 'delete-column') O.insertLines(sh, 'c', 1, scenario === 'insert-column' ? 1 : -1);
        if (scenario === 'rename-sheet') O.renameSheet(doc, sh, doc.nextSheetName('Lossless'));
        if (scenario === 'delete-sheet') O.deleteSheet(doc, sh);
        if (scenario === 'delete-object') O.setDrawings(sh, sh.drawings.slice(1));
        if (scenario === 'duplicate') O.setDrawings(sh, sh.drawings.concat(L.clone(sh.drawings[0])));
        if (scenario === 'geometry') {
          const drawings = L.clone(sh.drawings);
          const anchor = drawings[0].anchor;
          if (anchor.from) anchor.from.cOff = (anchor.from.cOff || 0) + 6; else anchor.x = (anchor.x || 0) + 6;
          O.setDrawings(sh, drawings);
        }
      });
      return { changed: [{ sheet: sh.name, kind: scenario }], undo: () => doc.undo.undo(), redo: () => doc.undo.redo() };
    }
    const textShape = slide => walkShapes(slide.shapes).find(sh => sh.tx?.ps?.some(p => p.rs?.some(r => r.t)));
    const needsText = ['text', 'format'].includes(scenario);
    const needsShape = ['geometry', 'duplicate', 'delete-object'].includes(scenario);
    const s = needsText ? doc.slides.find(textShape) : needsShape ? doc.slides.find(s => s.shapes.length) : doc.slides[0];
    if (!s) skip(needsText ? 'No editable slide text' : needsShape ? 'No slide shape' : 'No slides');
    const text = textShape(s);
    if (needsText && !text) skip('No editable text');
    if (needsShape && !s.shapes.length) skip('No slide shape');
    if (scenario === 'delete-slide' && doc.slides.length < 2) skip('Cannot delete the only slide');
    if (!['text', 'format', 'insert-slide', 'delete-slide', 'geometry', 'duplicate', 'delete-object'].includes(scenario)) skip('Scenario does not apply to Lectern');
    L.pres = doc; L.hist.clear(); L.hist.push('Corpus ' + scenario);
    if (scenario === 'text' || scenario === 'format') {
      const run = text.tx.ps.flatMap(p => p.rs || []).find(r => r.t);
      if (scenario === 'text') run.t += ' LOSSLESS-CHECK'; else run.b = !run.b;
    }
    if (scenario === 'geometry') s.shapes[0].x += 6;
    if (scenario === 'duplicate') s.shapes.push(L.model.dup(s.shapes[0]));
    if (scenario === 'delete-object') s.shapes.shift();
    if (scenario === 'insert-slide') L.model.insertSlides(doc, 1, [L.model.newSlide(doc, 'blank', s.design)]);
    if (scenario === 'delete-slide') doc.slides.splice(1, 1);
    return { changed: [{ slide: s.id, shape: text?.id, kind: scenario }], expectedSlides: doc.slides.length, undo: () => L.hist.doUndo(), redo: () => L.hist.doRedo() };
  };
})(typeof window === 'undefined' ? globalThis : window);
