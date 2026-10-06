/* Quire — command registry. Every menu item, toolbar button and shortcut runs one of these. */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, ui = L.ui;
  const A = (L.app = L.app || {});
  const C = (id, def) => ui.cmd(id, def);
  const doc = () => D.doc;
  const dlg = (name, ...args) => () => { if (L.dlg && L.dlg[name]) return L.dlg[name](...args); ui.msg('This dialog is not available.'); };
  const pane = (name) => () => { if (L.panes) { A.toggleTask(true); L.panes.task.show(name); } };
  const viewing = () => A.view === 'reading' || A.view === 'preview';
  const canEdit = () => !!D.doc && !E.readOnly && !viewing();
  const hasSel = () => !!E.sel && !E.collapsed();
  const inTable = () => !!(L.tables && E.sel && L.tables.inTable());
  const objSel = () => !!E.objSel;
  const floatSel = () => !!(E.objSel && E.objSel.float);
  const run = () => E.curRun();
  const para = () => E.curPara();
  const tbl = (fn, ...args) => () => L.tables && L.tables[fn](...args);

  /* ================= lists ================= */
  const LS = (L.lists = {});
  /** the list a paragraph belongs to: {numId, lvl, fmt} or null */
  LS.info = (p) => {
    const d = doc();
    const pp = D.pProps(d, p);
    if (!pp.num || !pp.num.id || pp.num.id === '0') return null;
    const lv = D.numLevelDef(d, pp.num.id, pp.num.lvl || 0);
    if (!lv) return null;
    return { numId: pp.num.id, lvl: pp.num.lvl || 0, fmt: lv.fmt, lv };
  };
  /** toolbar Bullets / Numbering: toggle, continuing an adjacent list of the same kind */
  LS.toggle = function (kind) {
    if (!E.sel) return;
    const d = doc();
    const paras = E.selectedParas();
    const first = paras[0];
    const cur = LS.info(first);
    const isKind = (inf) => inf && (kind === 'bullet' ? inf.fmt === 'bullet' : inf.fmt !== 'bullet');
    if (isKind(cur)) { LS.remove(paras); return; }
    /* continue the list right above when it has the same kind */
    const prev = D.prevPara(d, first);
    const pinf = prev && O.containerOf(prev) === O.containerOf(first) ? LS.info(prev) : null;
    E.edit(kind === 'bullet' ? 'Bullets' : 'Numbering', () => {
      let numId;
      if (isKind(pinf)) numId = pinf.numId;
      else {
        D.touchKey(d, 'numbering');
        const last = A.lastList && A.lastList[kind];
        const abs = last ? L.clone(last) : kind === 'bullet' ? D.makeBulletAbs('•', 'Symbol', '') : D.makeNumberAbs('decimal', '%1.');
        numId = D.addNum(d, abs);
      }
      const lvl = isKind(pinf) ? pinf.lvl : 0;
      O.setParaProps(paras, (pPr) => { pPr.num = { id: numId, lvl: pPr.num && pPr.num.id !== '0' ? pPr.num.lvl || 0 : lvl }; delete pPr.ind; });
      return E.sel;
    });
  };
  LS.remove = function (paras) {
    paras = paras || E.selectedParas();
    const d = doc();
    E.edit('Remove Numbering', () => {
      O.setParaProps(paras, (pPr, p) => {
        const sp = D.styleProps(d, pPr.style || 'Normal').pPr;
        if (sp.num && sp.num.id && sp.num.id !== '0') pPr.num = { id: '0', lvl: 0 }; else delete pPr.num;
        if (/^List(Bullet|Number)/.test(pPr.style || '')) delete pPr.style;
        void p;
        delete pPr.ind;
      });
      return E.sel;
    });
  };
  /** apply a list definition (from the Bullets and Numbering dialog). opts: {restart, kind} */
  LS.apply = function (abs, opts) {
    opts = opts || {};
    const d = doc();
    const paras = E.selectedParas();
    const first = paras[0];
    const prev = D.prevPara(d, first);
    const pinf = prev ? LS.info(prev) : null;
    E.edit('Bullets and Numbering', () => {
      D.touchKey(d, 'numbering');
      let numId;
      const cur = LS.info(first);
      if (cur && !opts.restart && !opts.newList) {
        /* redefine the list the paragraphs are in */
        const nd = D.numDef(d, cur.numId);
        const absNew = L.clone(abs);
        absNew.id = nd.num.abs;
        d.numbering.abs[nd.num.abs] = absNew;
        numId = cur.numId;
      } else if (opts.continuePrev && pinf) numId = pinf.numId;
      else numId = D.addNum(d, L.clone(abs));
      if (opts.restart) { const n = d.numbering.nums[numId]; n.ov = { 0: { start: opts.start || 1 } }; }
      O.setParaProps(paras, (pPr) => { pPr.num = { id: numId, lvl: pPr.num && pPr.num.id !== '0' && pPr.num.id ? pPr.num.lvl || 0 : 0 }; delete pPr.ind; });
      return E.sel;
    });
    A.lastList = A.lastList || {};
    A.lastList[abs.levels && abs.levels[0] && abs.levels[0].fmt === 'bullet' ? 'bullet' : 'number'] = L.clone(abs);
  };
  /** change list level of selected list paragraphs */
  LS.shift = function (delta) {
    const d = doc();
    const paras = E.selectedParas().filter((p) => LS.info(p));
    if (!paras.length) return false;
    E.edit('Numbering', () => {
      for (const p of paras) { const inf = LS.info(p); D.touch(p); p.pPr.num = { id: inf.numId, lvl: L.clamp(inf.lvl + delta, 0, 8) }; delete p.pPr.ind; }
      return E.sel;
    });
    void d;
    return true;
  };
  LS.restartAt = function (start) {
    const d = doc();
    const p = E.sel.f.p;
    const inf = LS.info(p);
    if (!inf) return;
    E.edit('Restart Numbering', () => {
      D.touchKey(d, 'numbering');
      const absId = d.numbering.nums[inf.numId].abs;
      const nid = D.addNumInstance(d, absId, true);
      d.numbering.nums[nid].ov = { [inf.lvl]: { start: start || 1 } };
      /* this paragraph and the rest of the same list switch to the new instance */
      let q = p;
      while (q) { const qi = LS.info(q); if (qi && qi.numId === inf.numId) { D.touch(q); q.pPr.num = { id: nid, lvl: qi.lvl }; } else if (D.plen(q) && !qi) break; q = D.nextPara(d, q); }
      return E.sel;
    });
  };
  LS.continuePrev = function () {
    const d = doc();
    const p = E.sel.f.p;
    const inf = LS.info(p);
    if (!inf) return;
    let q = D.prevPara(d, p), target = null;
    while (q) { const qi = LS.info(q); if (qi && qi.numId !== inf.numId && qi.fmt !== 'bullet') { target = qi; break; } q = D.prevPara(d, q); }
    if (!target) { A.status('There is no earlier list to continue.'); return; }
    E.edit('Continue Numbering', () => {
      let r = p;
      while (r) { const ri = LS.info(r); if (ri && ri.numId === inf.numId) { D.touch(r); r.pPr.num = { id: target.numId, lvl: ri.lvl }; } r = D.nextPara(d, r); }
      return E.sel;
    });
  };

  /* ================= helpers ================= */
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  function stepSize(dir, one) {
    const cur = E.uniformRun('sz') || run().sz || 12;
    let v;
    if (one) v = cur + dir;
    else if (dir > 0) v = SIZES.find((s) => s > cur) || Math.round(cur / 10 + 1) * 10;
    else v = SIZES.slice().reverse().find((s) => s < cur) || Math.max(1, cur - 1);
    A.setFontSize(v);
  }
  function setAlign(jc) {
    E.formatPara((pPr) => { const cur = pPr.jc; pPr.jc = cur === jc && jc !== 'left' ? 'left' : jc; }, 'Align');
  }
  const alignIs = (jc) => { const v = para().jc || 'left'; return jc === 'left' ? v === 'left' || v === 'start' : jc === 'both' ? v === 'both' || v === 'distribute' : v === jc; };
  function indent(delta) {
    if (LS.info(E.sel.f.p) && LS.shift(delta > 0 ? 1 : -1)) return;
    const defTab = doc().settings.defTab || 36;
    E.formatPara((pPr, p) => {
      const cur = D.pProps(doc(), p).ind || {};
      const l = cur.l || 0;
      const nl = delta > 0 ? (Math.floor(l / defTab + 1e-6) + 1) * defTab : Math.max(0, (Math.ceil(l / defTab - 1e-6) - 1) * defTab);
      pPr.ind = Object.assign({}, pPr.ind, { l: nl });
    }, delta > 0 ? 'Increase Indent' : 'Decrease Indent');
  }
  function hanging(delta) {
    const defTab = doc().settings.defTab || 36;
    E.formatPara((pPr, p) => {
      const cur = D.pProps(doc(), p).ind || {};
      const l = Math.max(0, (cur.l || 0) + delta * defTab);
      const fl = l > 0 ? -Math.min(l, delta > 0 ? Math.max(defTab, -(cur.fl || 0) + defTab) : Math.max(0, -(cur.fl || 0) - defTab)) : 0;
      pPr.ind = Object.assign({}, pPr.ind, { l, fl: delta > 0 ? -Math.max(defTab, -(cur.fl || 0)) : fl });
    }, 'Hanging Indent');
  }
  function lineSpacing(mult) { E.formatPara((pPr) => { pPr.sp = Object.assign({}, pPr.sp, { line: mult, rule: 'auto' }); }, 'Line Spacing'); }
  function applyStyle(id) {
    const d = doc();
    if (!d.styles[id]) {
      const b = D.builtinStyles()[id];
      if (!b) return;
      E.edit('Style', () => { D.touchKey(d, 'styles'); d.styles[id] = b; D.stylesChanged(); return E.sel; });
    }
    const [a, b] = E.range();
    E.edit('Style', () => { O.applyStyle(a, b, id); return E.sel; });
  }
  A.applyStyle = applyStyle;
  function insertFieldNow(instr, label) {
    const res = L.fields ? L.fields.evaluate(instr, E.sel.f) : '';
    E.edit(label || 'Insert Field', () => {
      const pos = E.deleteSelection();
      return O.insertField(pos, instr, res == null ? '' : String(res), O.inheritRPr(pos));
    });
    if (L.fields && /^\s*(PAGE|NUMPAGES|SECTIONPAGES|SECTION)\b/i.test(instr)) LY.refreshDynamic();
  }
  A.insertFieldNow = insertFieldNow;
  const typeChar = (ch) => () => E.typeText(ch);
  function clip() {
    try {
      if (document.execCommand('copy')) return true;
    } catch (e) { /* not allowed */ }
    return false;
  }
  A.clipItems = [];
  L.bus.on('clip', (c) => { A.clipItems.unshift(c); A.clipItems.length = Math.min(A.clipItems.length, 24); if (L.panes && L.panes.task.current === 'clipboard') L.panes.task.show('clipboard'); });
  async function pasteCmd() {
    if (!E.sel) return;
    if (E.clip) { E.pasteBlocks(E.clip.blocks); return; }
    try {
      if (navigator.clipboard && navigator.clipboard.readText) {
        const t = await navigator.clipboard.readText();
        if (t) { E.pasteText(t); return; }
      }
    } catch (e) { /* reading the clipboard is often blocked */ }
    ui.msg('Your browser does not let pages read the clipboard from a menu.\nPress Ctrl+V to paste.', { icon: 'info' });
  }
  A.paste = pasteCmd;
  const selectedObjIs = (pred) => { const it = L.drawing && L.drawing.selectedItem(); return !!it && pred(it); };

  /* ================= File ================= */
  C('new', { label: '&New...', menuLabel: '&New...', icon: 'new', key: 'Ctrl+N', tip: 'New Blank Document', run: (arg, ev) => { if (ev && ev.currentTarget && ev.currentTarget.closest && ev.currentTarget.closest('.toolbar')) A.newDocument(); else pane('new')(); } });
  C('newBlank', { label: '&Blank document', icon: 'new', run: () => A.newDocument() });
  C('open', { label: '&Open...', icon: 'open', key: 'Ctrl+O', run: () => A.openDialog() });
  C('close', { label: '&Close', icon: 'close', run: () => A.closeDocument() });
  C('save', { label: '&Save', icon: 'save', key: 'Ctrl+S', enabled: () => !!D.doc, run: () => A.save() });
  C('saveAs', { label: 'Save &As...', key: 'F12', run: () => A.saveAs() });
  C('saveWeb', { label: 'Save as Web Pa&ge...', icon: 'webPage', run: () => A.exportAs('html') });
  C('fileSearch', { label: 'File Searc&h...', icon: 'find', run: pane('search') });
  C('webPreview', { label: 'We&b Page Preview', icon: 'webPage', run: () => A.setView('web') });
  C('pageSetup', { label: 'Page Set&up...', icon: 'pageSetup', run: dlg('pageSetup') });
  C('printPreview', { label: 'Print Pre&view', icon: 'preview', key: 'Ctrl+F2', checked: () => A.view === 'preview', run: () => A.setView(A.view === 'preview' ? A.prevView || 'print' : 'preview') });
  C('print', { label: '&Print...', icon: 'print', key: 'Ctrl+P', tip: 'Print', run: (arg, ev) => { if (ev && ev.currentTarget && ev.currentTarget.closest && ev.currentTarget.closest('.toolbar')) A.printPDF(); else dlg('print')(); } });
  C('exportPDF', { label: '&PDF Document...', icon: 'print', run: () => A.printPDF() });
  C('sendMail', { label: '&Mail Recipient (as Attachment)...', icon: 'mail', tip: 'E-mail', run: async () => { await A.exportAs('docx'); A.status('Saved a copy to attach to your e-mail message.', 6000); } });
  C('permission', { label: '&Permission', icon: 'permission', tip: 'Permission (Unrestricted Access)', run: pane('protect') });
  C('properties', { label: 'Propert&ies', icon: 'properties', run: dlg('properties') });
  C('exitApp', { label: 'E&xit', run: async () => { const n = A.docs ? A.docs.length : 1; for (let k = 0; k < n; k++) { const before = A.docs.length; await A.closeDocument(); if (A.docs.length === before && before > 1) return; } } });

  /* ================= Edit ================= */
  C('undo', { get label() { return D.canUndo() ? '&Undo ' + D.undoLabel() : "&Can't Undo"; }, tip: 'Undo', icon: 'undo', key: 'Ctrl+Z', enabled: () => D.canUndo() && !viewing(), run: () => { D.undo(); } });
  C('redo', { get label() { return D.canRedo() ? '&Redo ' + D.redoLabel() : "&Repeat"; }, tip: 'Redo', icon: 'redo', key: 'Ctrl+Y', enabled: () => D.canRedo() && !viewing(), run: () => { D.redo(); } });
  C('cut', { label: 'Cu&t', icon: 'cut', key: 'Ctrl+X', enabled: () => canEdit() && (hasSel() || objSel()), run: () => { E.focus(); if (!clip()) { const d = E.copyData(); if (d) { E.clip = d; L.bus.emit('clip', d); } } E.edit('Cut', () => E.deleteSelection()); } });
  C('copy', { label: '&Copy', icon: 'copy', key: 'Ctrl+C', enabled: () => hasSel() || objSel(), run: () => { E.focus(); const d = E.copyData(); if (d) { E.clip = d; L.bus.emit('clip', d); } clip(); } });
  C('officeClipboard', { label: 'Office Clip&board...', icon: 'clipboard', run: pane('clipboard') });
  C('paste', { label: '&Paste', icon: 'paste', key: 'Ctrl+V', enabled: () => canEdit(), run: () => pasteCmd() });
  C('pasteSpecial', { label: 'Paste &Special...', enabled: () => canEdit(), run: dlg('pasteSpecial') });
  C('pasteHyperlink', { label: 'Paste as &Hyperlink', enabled: () => canEdit() && !!E.clip, run: () => { const t = E.clip.text.trim(); if (!t) return; const url = /^(https?:|mailto:|www\.)/i.test(t) ? (/^www\./i.test(t) ? 'http://' + t : t) : null; E.edit('Paste as Hyperlink', () => { const pos = E.deleteSelection(); return O.insertText(pos, t, Object.assign(O.inheritRPr(pos), { style: 'Hyperlink', link: url ? { url } : { anchor: '' } })); }); } });
  C('clearFormats', { label: '&Formats', enabled: canEdit, run: () => { const [a, b] = E.range(); E.edit('Clear Formats', () => { if (!D.eqPos(a, b)) O.formatRange(a, b, (r) => { const k = {}; if (r.link) k.link = r.link; if (r.ins) k.ins = r.ins; if (r.del) k.del = r.del; return k; }); O.setParaProps(E.selectedParas(), (pPr) => { const k = {}; if (pPr.num) k.num = pPr.num; if (pPr.sect) k.sect = pPr.sect; return k; }); return E.sel; }); } });
  C('clearContents', { label: '&Contents', key: 'Del', enabled: () => canEdit() && hasSel(), run: () => E.edit('Clear', () => E.deleteSelection()) });
  C('selectAll', { label: 'Select A&ll', key: 'Ctrl+A', icon: 'selectAll', run: () => { E.focus(); E.selectAll(); } });
  C('find', { label: '&Find...', icon: 'find', key: 'Ctrl+F', run: dlg('find', 'find') });
  C('replace', { label: 'R&eplace...', icon: 'replace', key: 'Ctrl+H', enabled: () => !viewing(), run: dlg('find', 'replace') });
  C('goTo', { label: '&Go To...', key: 'Ctrl+G', run: dlg('find', 'goto') });
  C('links', { label: 'Lin&ks...', enabled: () => false, run: () => {} });
  C('editObject', { get label() { const it = L.drawing && L.drawing.selectedItem(); return it && it.t === 'chart' ? 'Chart &Object' : it && it.wordart ? 'WordArt &Object' : '&Object'; }, enabled: () => selectedObjIs((it) => it.t === 'chart' || !!it.wordart), run: () => A.formatObject(true) });

  /* ================= View ================= */
  C('viewNormal', { label: '&Normal', icon: 'normalView', key: 'Ctrl+Alt+N', radio: true, checked: () => A.view === 'normal', run: () => A.setView('normal') });
  C('viewWeb', { label: '&Web Layout', icon: 'webView', radio: true, checked: () => A.view === 'web', run: () => A.setView('web') });
  C('viewPrint', { label: '&Print Layout', icon: 'printView', key: 'Ctrl+Alt+P', radio: true, checked: () => A.view === 'print', run: () => A.setView('print') });
  C('viewReading', { label: 'R&eading Layout', icon: 'read', tip: 'Read', radio: true, checked: () => A.view === 'reading', run: () => A.setView(A.view === 'reading' ? A.prevView || 'print' : 'reading') });
  C('viewOutline', { label: '&Outline', icon: 'outlineView', key: 'Ctrl+Alt+O', radio: true, checked: () => A.view === 'outline', run: () => A.setView('outline') });
  C('taskPane', { label: 'Tas&k Pane', icon: 'taskpane', key: 'Ctrl+F1', checked: () => A.taskOpen(), run: () => A.toggleTask() });
  const TBNAMES = { standard: 'Standard', formatting: 'Formatting', drawing: 'Drawing', tables: 'Tables and Borders', reviewing: 'Reviewing', picture: 'Picture', outlining: 'Outlining', headerFooter: 'Header and Footer', mailMerge: 'Mail Merge', wordart: 'WordArt' };
  for (const k in TBNAMES) C('tb_' + k, { label: TBNAMES[k], tip: TBNAMES[k], checked: () => A.tbShown(k), run: () => A.toggleToolbar(k) });
  C('ruler', { label: '&Ruler', icon: 'ruler', checked: () => !!A.opts.ruler, run: () => { A.opts.ruler = !A.opts.ruler; A.saveOpts(); A.updateRulers(); } });
  C('docMap', { label: '&Document Map', icon: 'docMap', checked: () => !!A.opts.docMap, run: () => A.toggleDocMap() });
  C('thumbnails', { label: 'Thu&mbnails', icon: 'thumbnails', checked: () => !!A.opts.thumbs, run: () => A.toggleThumbs() });
  C('headerFooter', { label: '&Header and Footer', icon: 'headerFooter', checked: () => !!LY.hfEdit, run: () => (LY.hfEdit ? A.closeHeaderFooter() : A.editHeaderFooter('hdr')) });
  C('footnotes', { label: 'Foot&notes', enabled: () => Object.keys(doc().fn).length + Object.keys(doc().en).length > 0, run: () => { const k = Object.keys(doc().fn)[0]; if (k) A.gotoNote('fn:' + k); else { const e = Object.keys(doc().en)[0]; if (e) A.gotoNote('en:' + e); } } });
  C('markup', { label: '&Markup', checked: () => LY.opts.markup !== 'final' && LY.opts.markup !== 'original' && LY.opts.showComments, run: () => { const on = !(LY.opts.markup !== 'final' && LY.opts.markup !== 'original' && LY.opts.showComments); LY.opts.markup = on ? 'markup' : 'final'; LY.opts.showComments = on; LY.render(); E.restoreDom(true); } });
  C('fullScreen', { label: 'F&ull Screen', icon: 'fullScreen', checked: () => L.$('#app').classList.contains('fullscreen'), run: () => A.toggleFullScreen() });
  C('zoomDlg', { label: '&Zoom...', icon: 'zoom', run: dlg('zoom') });
  C('showMarks', { label: 'Show/Hide ¶', icon: 'showMarks', key: 'Ctrl+*', tip: 'Show/Hide ¶', checked: () => !!LY.opts.showMarks, run: () => { LY.opts.showMarks = !LY.opts.showMarks; LY.root.classList.toggle('marks', LY.opts.showMarks); LY.render(); E.restoreDom(true); } });
  C('gridlines', { label: 'Show &Gridlines', icon: 'grid', checked: () => LY.root && LY.root.classList.contains('gridlines'), run: () => { LY.root.classList.toggle('gridlines'); A.opts.gridlines = LY.root.classList.contains('gridlines'); A.saveOpts(); } });

  /* ================= Insert ================= */
  C('insertBreak', { label: '&Break...', enabled: canEdit, run: dlg('breakDlg') });
  C('pageNumbers', { label: 'Page N&umbers...', enabled: canEdit, run: dlg('pageNumbers') });
  C('dateTime', { label: 'Date and &Time...', icon: 'date', enabled: canEdit, run: dlg('dateTime') });
  C('insertField', { label: '&Field...', enabled: canEdit, run: dlg('field') });
  C('insertSymbol', { label: '&Symbol...', icon: 'symbol', enabled: canEdit, run: dlg('symbol') });
  C('insertComment', { label: 'Co&mment', icon: 'comment', key: 'Ctrl+Alt+M', tip: 'Insert Comment', enabled: () => !!D.doc && !viewing(), run: () => L.review && L.review.insertComment() });
  C('insertFootnote', { label: '&Footnote...', icon: 'footnote', enabled: canEdit, run: dlg('footnote') });
  C('insertFootnoteNow', { label: 'Insert Footnote', enabled: canEdit, run: () => L.notes && L.notes.insert('fn') });
  C('insertEndnoteNow', { label: 'Insert Endnote', enabled: canEdit, run: () => L.notes && L.notes.insert('en') });
  C('insertCaption', { label: '&Caption...', enabled: canEdit, run: dlg('caption') });
  C('crossReference', { label: 'Cross-&reference...', enabled: canEdit, run: dlg('crossRef') });
  C('indexTables', { label: 'In&dex and Tables...', enabled: canEdit, run: dlg('indexTables') });
  C('insertClipArt', { label: '&Clip Art...', icon: 'clipart', tip: 'Insert Clip Art', enabled: canEdit, run: pane('clipart') });
  C('insertPicture', { label: '&From File...', icon: 'picture', tip: 'Insert Picture', enabled: canEdit, run: () => A.insertPictureDialog() });
  C('newDrawing', { label: '&New Drawing', icon: 'drawing', enabled: canEdit, run: () => A.openAutoShapes() });
  C('insertWordArt', { label: '&WordArt...', icon: 'wordart', tip: 'Insert WordArt', enabled: canEdit, run: dlg('wordart') });
  C('insertChart', { label: 'C&hart', icon: 'chart', enabled: canEdit, run: () => A.insertChart() });
  C('insertExcel', { label: 'Insert Spreadsheet', icon: 'excel', tip: 'Insert Spreadsheet', enabled: canEdit, run: () => L.tables.insert(4, 5, { worksheet: true }) });
  C('insertDiagram', { label: 'Dia&gram...', icon: 'diagram', tip: 'Insert Diagram or Organization Chart', enabled: canEdit, run: dlg('diagram') });
  C('textBox', { label: 'Te&xt Box', icon: 'textbox', tip: 'Text Box', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'rect', textbox: true }) });
  C('insertFile', { label: 'Fil&e...', enabled: canEdit, run: () => A.insertFileDialog() });
  C('insertObject', { label: '&Object...', enabled: canEdit, run: dlg('insertObject') });
  C('bookmark', { label: 'Boo&kmark...', key: 'Ctrl+Shift+F5', run: dlg('bookmark') });
  C('hyperlink', { label: 'Hyperl&ink...', icon: 'hyperlink', key: 'Ctrl+K', tip: 'Insert Hyperlink', enabled: canEdit, run: dlg('hyperlink') });
  C('removeHyperlink', { label: '&Remove Hyperlink', enabled: canEdit, run: () => A.removeHyperlink() });

  /* ================= Format ================= */
  C('fontDlg', { label: '&Font...', icon: 'font', key: 'Ctrl+D', enabled: canEdit, run: dlg('font') });
  C('paragraphDlg', { label: '&Paragraph...', icon: 'paragraph', enabled: canEdit, run: dlg('paragraph') });
  C('bulletsDlg', { label: 'Bullets and &Numbering...', icon: 'numbering', enabled: canEdit, run: dlg('bullets') });
  C('bordersDlg', { label: '&Borders and Shading...', icon: 'bordersOut', enabled: canEdit, run: dlg('borders') });
  C('columnsDlg', { label: '&Columns...', icon: 'columns', enabled: canEdit, run: dlg('columns') });
  C('tabsDlg', { label: '&Tabs...', enabled: canEdit, run: dlg('tabs') });
  C('dropCap', { label: '&Drop Cap...', icon: 'dropCap', enabled: canEdit, run: dlg('dropCap') });
  C('textDirection', { label: 'Te&xt Direction...', icon: 'textDir', tip: 'Change Text Direction', enabled: canEdit, run: () => (inTable() ? L.tables.cycleTextDirection() : dlg('textDirection')()) });
  C('changeCase', { label: 'Change Cas&e...', icon: 'changeCase', enabled: canEdit, run: dlg('changeCase') });
  C('changeCaseCycle', { label: 'Change Case', key: 'Shift+F3', enabled: () => canEdit(), run: () => A.cycleCase() });
  C('printedWatermark', { label: 'Printed &Watermark...', enabled: canEdit, run: dlg('watermark') });
  C('themeDlg', { label: 'T&heme...', enabled: canEdit, run: dlg('theme') });
  C('autoFormat', { label: '&AutoFormat...', enabled: canEdit, run: dlg('autoFormatDoc') });
  C('stylesPane', { label: '&Styles and Formatting...', icon: 'stylesPane', tip: 'Styles and Formatting', key: 'Ctrl+Alt+Shift+S', run: pane('styles') });
  C('revealFormatting', { label: 'Re&veal Formatting...', icon: 'reveal', key: 'Shift+F1', run: pane('reveal') });
  C('formatObject', { get label() { const it = L.drawing && L.drawing.selectedItem(); return it ? (it.t === 'img' ? 'P&icture...' : it.tb ? 'Text B&ox...' : it.wordart ? 'Word&Art...' : it.t === 'chart' ? 'Ch&art...' : 'AutoSha&pe...') : 'Obj&ect...'; }, icon: 'formatPicture', tip: 'Format Object', enabled: () => objSel(), run: () => A.formatObject() });
  /* character */
  C('bold', { label: 'Bold', icon: 'bold', key: 'Ctrl+B', enabled: canEdit, checked: () => !!run().b, run: () => E.toggleRun('b', undefined, 'Bold') });
  C('italic', { label: 'Italic', icon: 'italic', key: 'Ctrl+I', enabled: canEdit, checked: () => !!run().i, run: () => E.toggleRun('i', undefined, 'Italic') });
  C('underline', { label: 'Underline', icon: 'underline', key: 'Ctrl+U', enabled: canEdit, checked: () => !!run().u && run().u !== 'none', run: () => { const u = run().u; E.formatRun({ u: u && u !== 'none' ? 'none' : 'single' }, 'Underline'); } });
  C('dblUnderline', { label: 'Double Underline', enabled: canEdit, run: () => { const u = run().u; E.formatRun({ u: u === 'double' ? 'none' : 'double' }, 'Underline'); } });
  C('wordUnderline', { label: 'Word Underline', enabled: canEdit, run: () => { const u = run().u; E.formatRun({ u: u === 'words' ? 'none' : 'words' }, 'Underline'); } });
  C('strike', { label: 'Strikethrough', icon: 'strike', enabled: canEdit, checked: () => !!run().strike, run: () => E.toggleRun('strike', undefined, 'Strikethrough') });
  C('subscript', { label: 'Subscript', icon: 'subscript', key: 'Ctrl+=', enabled: canEdit, checked: () => run().vert === 'subscript', run: () => E.formatRun({ vert: run().vert === 'subscript' ? undefined : 'subscript' }, 'Subscript') });
  C('superscript', { label: 'Superscript', icon: 'superscript', key: 'Ctrl+Shift++', enabled: canEdit, checked: () => run().vert === 'superscript', run: () => E.formatRun({ vert: run().vert === 'superscript' ? undefined : 'superscript' }, 'Superscript') });
  C('allCaps', { label: 'All Caps', enabled: canEdit, checked: () => !!run().caps, run: () => E.toggleRun('caps', undefined, 'All Caps') });
  C('smallCaps', { label: 'Small Caps', enabled: canEdit, checked: () => !!run().smallCaps, run: () => E.toggleRun('smallCaps', undefined, 'Small Caps') });
  C('hiddenText', { label: 'Hidden', enabled: canEdit, checked: () => !!run().hidden, run: () => E.toggleRun('hidden', undefined, 'Hidden') });
  C('resetChar', { label: 'Reset Character Formatting', key: 'Ctrl+Space', enabled: canEdit, run: () => { if (E.collapsed()) { E.pending = { b: false, i: false, u: undefined, strike: undefined, vert: undefined, caps: undefined, smallCaps: undefined, color: undefined, hl: undefined, font: undefined, sz: undefined }; L.bus.emit('sel', E.sel); return; } const [a, b] = E.range(); E.edit('Reset Character', () => { O.formatRange(a, b, (r) => { const k = {}; for (const x of ['link', 'ins', 'del', 'style']) if (r[x]) k[x] = r[x]; if (k.style && !/^(Hyperlink|FootnoteReference|EndnoteReference|CommentReference)$/.test(k.style)) delete k.style; return k; }); return E.sel; }); } });
  C('resetPara', { label: 'Reset Paragraph Formatting', key: 'Ctrl+Q', enabled: canEdit, run: () => E.formatPara((pPr) => { const k = {}; for (const x of ['style', 'num', 'sect', 'frame']) if (pPr[x]) k[x] = pPr[x]; return k; }, 'Reset Paragraph') });
  C('growFont', { label: 'Grow Font', icon: 'growFont', key: 'Ctrl+Shift+>', enabled: canEdit, run: () => stepSize(1) });
  C('shrinkFont', { label: 'Shrink Font', icon: 'shrinkFont', key: 'Ctrl+Shift+<', enabled: canEdit, run: () => stepSize(-1) });
  C('growFont1', { label: 'Grow Font 1 Point', key: 'Ctrl+]', enabled: canEdit, run: () => stepSize(1, true) });
  C('shrinkFont1', { label: 'Shrink Font 1 Point', key: 'Ctrl+[', enabled: canEdit, run: () => stepSize(-1, true) });
  C('focusFont', { label: 'Font', run: () => { const i = L.$('#tb-font'); if (i) i.focus(); } });
  C('focusSize', { label: 'Font Size', run: () => { const i = L.$('#tb-size'); if (i) i.focus(); } });
  C('focusStyle', { label: 'Style', run: () => { const i = L.$('#tb-style'); if (i) i.focus(); } });
  C('copyFormat', { label: 'Copy Formatting', run: () => A.startPainter(false) });
  C('pasteFormat', { label: 'Paste Formatting', enabled: canEdit, run: () => { if (A.painter) A.applyPainter(); } });
  C('painter', { label: 'Format Painter', icon: 'painter', tip: 'Format Painter', enabled: canEdit, checked: () => !!A.painter, run: (arg, ev) => { if (A.painter) { A.painter = null; L.$('#scroller').classList.remove('painting'); return; } A.startPainter(ev && ev.detail === 2); } });
  C('fontColorBtn', { label: 'Font Color', icon: 'fontColor', tip: 'Font Color', enabled: canEdit, swatch: () => '#' + A.lastFontColor, run: () => A.applyFontColor(A.lastFontColor) });
  C('highlightBtn', { label: 'Highlight', icon: 'highlight', tip: 'Highlight', enabled: canEdit, swatch: () => L.R.HIGHLIGHT[A.lastHighlight] || '#ffff00', run: () => A.applyHighlight(A.lastHighlight) });
  C('borderBtn', { get label() { const b = A.BORDER_ITEMS.find((x) => x[0] === A.lastBorder); return b ? b[2] : 'Outside Border'; }, get icon() { const b = A.BORDER_ITEMS.find((x) => x[0] === A.lastBorder); return b ? b[1] : 'bordersOut'; }, enabled: canEdit, run: () => A.applyBorder(A.lastBorder) });
  /* paragraph */
  C('alignLeft', { label: 'Align Left', icon: 'alignLeft', key: 'Ctrl+L', enabled: canEdit, checked: () => alignIs('left'), run: () => setAlign('left') });
  C('alignCenter', { label: 'Center', icon: 'alignCenter', key: 'Ctrl+E', enabled: canEdit, checked: () => alignIs('center'), run: () => setAlign('center') });
  C('alignRight', { label: 'Align Right', icon: 'alignRight', key: 'Ctrl+R', enabled: canEdit, checked: () => alignIs('right'), run: () => setAlign('right') });
  C('justify', { label: 'Justify', icon: 'justify', key: 'Ctrl+J', enabled: canEdit, checked: () => alignIs('both'), run: () => setAlign('both') });
  C('single', { label: 'Single Spacing', key: 'Ctrl+1', enabled: canEdit, run: () => lineSpacing(1) });
  C('onePointFive', { label: '1.5 Line Spacing', key: 'Ctrl+5', enabled: canEdit, run: () => lineSpacing(1.5) });
  C('double', { label: 'Double Spacing', key: 'Ctrl+2', enabled: canEdit, run: () => lineSpacing(2) });
  C('spaceBefore', { label: 'Space Before', key: 'Ctrl+0', enabled: canEdit, run: () => E.formatPara((pPr, p) => { const cur = (D.pProps(doc(), p).sp || {}).b || 0; pPr.sp = Object.assign({}, pPr.sp, { b: cur >= 12 ? 0 : 12, bAuto: undefined }); }, 'Space Before') });
  C('numbering', { label: 'Numbering', icon: 'numbering', tip: 'Numbering', enabled: canEdit, checked: () => { const i = E.sel && LS.info(E.sel.f.p); return !!i && i.fmt !== 'bullet'; }, run: () => LS.toggle('number') });
  C('bullets', { label: 'Bullets', icon: 'bullets', tip: 'Bullets', enabled: canEdit, checked: () => { const i = E.sel && LS.info(E.sel.f.p); return !!i && i.fmt === 'bullet'; }, run: () => LS.toggle('bullet') });
  C('incIndent', { label: 'Increase Indent', icon: 'incIndent', key: 'Ctrl+M', enabled: canEdit, run: () => indent(1) });
  C('decIndent', { label: 'Decrease Indent', icon: 'decIndent', key: 'Ctrl+Shift+M', enabled: canEdit, run: () => indent(-1) });
  C('hangIndent', { label: 'Hanging Indent', key: 'Ctrl+T', enabled: canEdit, run: () => hanging(1) });
  C('unhangIndent', { label: 'Reduce Hanging Indent', key: 'Ctrl+Shift+T', enabled: canEdit, run: () => hanging(-1) });
  C('restartNumbering', { label: '&Restart Numbering', enabled: () => !!(E.sel && LS.info(E.sel.f.p)), run: () => LS.restartAt(1) });
  C('continueNumbering', { label: '&Continue Numbering', enabled: () => !!(E.sel && LS.info(E.sel.f.p)), run: () => LS.continuePrev() });
  /* styles */
  C('styleNormal', { label: 'Normal', key: 'Ctrl+Shift+N', enabled: canEdit, run: () => applyStyle('Normal') });
  C('styleH1', { label: 'Heading 1', key: 'Ctrl+Alt+1', enabled: canEdit, run: () => applyStyle('Heading1') });
  C('styleH2', { label: 'Heading 2', key: 'Ctrl+Alt+2', enabled: canEdit, run: () => applyStyle('Heading2') });
  C('styleH3', { label: 'Heading 3', key: 'Ctrl+Alt+3', enabled: canEdit, run: () => applyStyle('Heading3') });
  C('styleListBullet', { label: 'List Bullet', key: 'Ctrl+Shift+L', enabled: canEdit, run: () => applyStyle('ListBullet') });

  /* ================= Insert shortcuts ================= */
  C('insertDateField', { label: 'Date Field', key: 'Alt+Shift+D', enabled: canEdit, run: () => insertFieldNow('DATE \\@ "M/d/yyyy"', 'Insert Date') });
  C('insertTimeField', { label: 'Time Field', key: 'Alt+Shift+T', enabled: canEdit, run: () => insertFieldNow('TIME \\@ "h:mm am/pm"', 'Insert Time') });
  C('insertPageField', { label: 'Page Field', key: 'Alt+Shift+P', enabled: canEdit, run: () => insertFieldNow('PAGE', 'Insert Page Number') });
  C('symCopyright', { label: '©', enabled: canEdit, run: typeChar('©') });
  C('symRegistered', { label: '®', enabled: canEdit, run: typeChar('®') });
  C('symTrademark', { label: '™', enabled: canEdit, run: typeChar('™') });
  C('symEllipsis', { label: '…', enabled: canEdit, run: typeChar('…') });
  C('optHyphen', { label: 'Optional Hyphen', enabled: canEdit, run: typeChar('­') });
  C('nbHyphen', { label: 'Nonbreaking Hyphen', enabled: canEdit, run: typeChar('‑') });
  C('emDash', { label: 'Em Dash', enabled: canEdit, run: typeChar('—') });
  C('enDash', { label: 'En Dash', enabled: canEdit, run: typeChar('–') });

  /* ================= fields ================= */
  C('updateField', { label: '&Update Field', key: 'F9', enabled: canEdit, run: () => L.fields && L.fields.updateSelection() });
  C('editField', { label: '&Edit Field...', enabled: () => !!(L.fields && L.fields.at(E.sel && E.sel.f)), run: () => L.dlg.field(L.fields.at(E.sel.f)) });
  C('toggleCodes', { label: '&Toggle Field Codes', key: 'Shift+F9', checked: () => !!LY.opts.fieldCodes, run: () => { LY.opts.fieldCodes = !LY.opts.fieldCodes; LY.render(); E.restoreDom(true); } });
  C('insertFieldBraces', { label: 'Insert Field', key: 'Ctrl+F9', enabled: canEdit, run: () => L.dlg.field(null, { codes: true }) });
  C('unlinkField', { label: 'Unlink Fields', key: 'Ctrl+Shift+F9', enabled: canEdit, run: () => L.fields && L.fields.unlinkSelection() });
  C('lockField', { label: 'Lock Fields', key: 'Ctrl+F11', enabled: canEdit, run: () => L.fields && L.fields.lockSelection(true) });
  C('unlockField', { label: 'Unlock Fields', key: 'Ctrl+Shift+F11', enabled: canEdit, run: () => L.fields && L.fields.lockSelection(false) });

  /* ================= Tools ================= */
  C('spelling', { label: '&Spelling and Grammar...', icon: 'spell', key: 'F7', tip: 'Spelling and Grammar', run: () => L.spell && L.spell.check() });
  C('research', { label: '&Research...', icon: 'research', tip: 'Research', run: () => { if (L.panes) { L.panes.researchQuery = E.sel && !E.collapsed() ? O.textRange(...E.range(), ' ').trim().slice(0, 80) : (E.sel && O.wordAt(E.sel.f) ? O.textRange(...O.wordAt(E.sel.f)) : ''); } pane('research')(); } });
  C('setLanguage', { label: '&Set Language...', run: dlg('language') });
  C('thesaurus', { label: '&Thesaurus...', key: 'Shift+F7', run: () => { if (L.panes) { L.panes.researchService = 'thesaurus'; const w = E.sel && (E.collapsed() ? O.wordAt(E.sel.f) : E.range()); L.panes.researchQuery = w ? O.textRange(w[0], w[1], ' ').trim() : ''; } pane('research')(); } });
  C('hyphenation', { label: '&Hyphenation...', run: dlg('hyphenation') });
  C('wordCount', { label: '&Word Count...', icon: 'wordCount', run: dlg('wordCount') });
  C('autoSummarize', { label: 'A&utoSummarize...', icon: 'summary', enabled: canEdit, run: dlg('autoSummarize') });
  C('trackChanges', { label: '&Track Changes', icon: 'trackChanges', key: 'Ctrl+Shift+E', tip: 'Track Changes', enabled: () => !!D.doc && !E.readOnly, checked: () => !!(D.doc && doc().settings.track), run: () => { if (L.review) L.review.toggleTracking(); } });
  C('compareDocs', { label: 'Compare and Merge &Documents...', enabled: canEdit, run: () => L.review && L.review.compareDialog() });
  C('protectDoc', { label: '&Protect Document...', icon: 'permission', run: pane('protect') });
  C('mailMerge', { label: '&Mail Merge...', icon: 'mailMerge', run: pane('mailmerge') });
  C('envelopes', { label: '&Envelopes and Labels...', icon: 'envelope', run: dlg('envelopes') });
  C('letterWizard', { label: '&Letter Wizard...', run: dlg('letterWizard') });
  C('macros', { label: '&Macros...', key: 'Alt+F8', run: dlg('macros') });
  C('templates', { label: 'Templates and Add-&Ins...', run: dlg('templates') });
  C('autocorrectDlg', { label: '&AutoCorrect Options...', run: dlg('autocorrect') });
  C('customize', { label: '&Customize...', run: dlg('customize') });
  C('optionsDlg', { label: '&Options...', run: dlg('options') });

  /* ================= Table ================= */
  C('drawTable', { label: 'Dra&w Table', icon: 'pen', tip: 'Draw Table', enabled: canEdit, checked: () => !!(L.tables && L.tables.tool === 'draw'), run: () => L.tables.setTool(L.tables.tool === 'draw' ? null : 'draw') });
  C('eraser', { label: 'Eraser', icon: 'eraser', tip: 'Eraser', enabled: canEdit, checked: () => !!(L.tables && L.tables.tool === 'erase'), run: () => L.tables.setTool(L.tables.tool === 'erase' ? null : 'erase') });
  C('insertTable', { label: '&Table...', icon: 'table', enabled: canEdit, run: dlg('insertTable') });
  C('colLeft', { label: 'Columns to the &Left', icon: 'colLeft', enabled: () => canEdit() && inTable(), run: tbl('insertCols', true) });
  C('colRight', { label: 'Columns to the &Right', icon: 'colRight', enabled: () => canEdit() && inTable(), run: tbl('insertCols', false) });
  C('rowAbove', { label: 'Rows &Above', icon: 'rowAbove', enabled: () => canEdit() && inTable(), run: tbl('insertRows', true) });
  C('rowBelow', { label: 'Rows &Below', icon: 'rowBelow', enabled: () => canEdit() && inTable(), run: tbl('insertRows', false) });
  C('insertCells', { label: 'C&ells...', enabled: () => canEdit() && inTable(), run: dlg('insertCells') });
  C('deleteTable', { label: '&Table', icon: 'deleteTable', enabled: () => canEdit() && inTable(), run: tbl('deleteTable') });
  C('deleteCols', { label: '&Columns', icon: 'deleteCol', enabled: () => canEdit() && inTable(), run: tbl('deleteCols') });
  C('deleteRows', { label: '&Rows', icon: 'deleteRow', enabled: () => canEdit() && inTable(), run: tbl('deleteRows') });
  C('deleteCells', { label: 'C&ells...', enabled: () => canEdit() && inTable(), run: dlg('deleteCells') });
  C('selectTable', { label: '&Table', enabled: inTable, run: tbl('select', 'table') });
  C('selectCol', { label: '&Column', enabled: inTable, run: tbl('select', 'col') });
  C('selectRow', { label: '&Row', enabled: inTable, run: tbl('select', 'row') });
  C('selectCell', { label: 'C&ell', enabled: inTable, run: tbl('select', 'cell') });
  C('mergeCells', { label: 'Mer&ge Cells', icon: 'mergeCells', tip: 'Merge Cells', enabled: () => canEdit() && inTable() && !!E.cellSel, run: tbl('merge') });
  C('splitCells', { label: 'S&plit Cells...', icon: 'splitCells', tip: 'Split Cells', enabled: () => canEdit() && inTable(), run: dlg('splitCells') });
  C('splitTable', { label: 'Spli&t Table', enabled: () => canEdit() && inTable(), run: tbl('splitTable') });
  C('tableAutoFormat', { label: 'Table Auto&Format...', icon: 'tableAutoFormat', tip: 'Table AutoFormat', enabled: () => canEdit() && inTable(), run: dlg('tableAutoFormat') });
  C('autofitContents', { label: 'AutoFit to &Contents', enabled: () => canEdit() && inTable(), run: tbl('autofit', 'contents') });
  C('autofitWindow', { label: 'AutoFit to &Window', enabled: () => canEdit() && inTable(), run: tbl('autofit', 'window') });
  C('fixedWidth', { label: '&Fixed Column Width', enabled: () => canEdit() && inTable(), run: tbl('autofit', 'fixed') });
  C('distRows', { label: 'Distribute Rows Evenl&y', icon: 'distRows', tip: 'Distribute Rows Evenly', enabled: () => canEdit() && inTable(), run: tbl('distribute', 'rows') });
  C('distCols', { label: 'Distribute Columns Evenl&y', icon: 'distCols', tip: 'Distribute Columns Evenly', enabled: () => canEdit() && inTable(), run: tbl('distribute', 'cols') });
  C('headingRows', { label: '&Heading Rows Repeat', enabled: () => canEdit() && inTable(), checked: () => inTable() && L.tables.isHeaderRow(), run: tbl('toggleHeaderRow') });
  C('textToTable', { label: 'Te&xt to Table...', enabled: () => canEdit() && hasSel() && !inTable(), run: dlg('textToTable') });
  C('tableToText', { label: 'Ta&ble to Text...', enabled: () => canEdit() && inTable(), run: dlg('tableToText') });
  C('sortDlg', { label: '&Sort...', icon: 'sortAsc', enabled: canEdit, run: dlg('sort') });
  C('sortAsc', { label: 'Sort Ascending', icon: 'sortAsc', enabled: canEdit, run: () => L.tables.quickSort(false) });
  C('sortDesc', { label: 'Sort Descending', icon: 'sortDesc', enabled: canEdit, run: () => L.tables.quickSort(true) });
  C('autoSum', { label: 'AutoSum', icon: 'autoSum', enabled: () => canEdit() && inTable(), run: () => L.tables.autoSum() });
  C('formulaDlg', { label: 'F&ormula...', enabled: canEdit, run: dlg('formula') });
  C('tableProps', { label: 'Table P&roperties...', icon: 'tableProps', enabled: () => inTable(), run: dlg('tableProps') });
  C('cellTop', { label: 'Align Top', icon: 'alignTop', enabled: () => canEdit() && inTable(), checked: () => inTable() && L.tables.cellVAlign() === 'top', run: tbl('setVAlign', 'top') });
  C('cellMiddle', { label: 'Align Center', icon: 'alignMiddle', enabled: () => canEdit() && inTable(), checked: () => inTable() && L.tables.cellVAlign() === 'center', run: tbl('setVAlign', 'center') });
  C('cellBottom', { label: 'Align Bottom', icon: 'alignBottom', enabled: () => canEdit() && inTable(), checked: () => inTable() && L.tables.cellVAlign() === 'bottom', run: tbl('setVAlign', 'bottom') });

  /* ================= Window & Help ================= */
  C('newWindow', { label: '&New Window', enabled: () => false, run: () => {} });
  C('arrangeAll', { label: '&Arrange All', enabled: () => false, run: () => {} });
  C('splitWindow', { label: '&Split', checked: () => !!A.splitOn, run: () => A.toggleSplit() });
  C('help', { label: 'Quire &Help', menuLabel: 'Quire &Help', icon: 'help', key: 'F1', tip: 'Help', run: pane('help') });
  C('showAssistant', { label: 'Show the &Assistant', run: () => ui.toast('The Assistant is resting. Type a question in the Help box at the top right instead.', 4000) });
  C('gettingStarted', { label: '&Getting Started', run: pane('getting-started') });
  C('keyboardHelp', { label: '&Keyboard Shortcuts', run: () => { if (L.panes) { L.panes.helpTopic = 'keys'; pane('help')(); } } });
  C('about', { label: '&About Quire', run: dlg('about') });

  /* ================= Reviewing ================= */
  const MK = { markupFinalMarkup: ['Final Showing Markup', 'markup'], markupFinal: ['Final', 'final'], markupOriginalMarkup: ['Original Showing Markup', 'origMarkup'], markupOriginal: ['Original', 'original'] };
  for (const k in MK) C(k, { label: MK[k][0], radio: true, checked: () => LY.opts.markup === MK[k][1], run: () => { LY.opts.markup = MK[k][1]; LY.render(); E.restoreDom(true); } });
  C('showComments', { label: '&Comments', checked: () => LY.opts.showComments !== false, run: () => { LY.opts.showComments = !LY.opts.showComments; LY.render(); E.restoreDom(true); } });
  C('showInsDel', { label: '&Insertions and Deletions', checked: () => LY.opts.markup !== 'final', run: () => { LY.opts.markup = LY.opts.markup === 'final' ? 'markup' : 'final'; LY.render(); E.restoreDom(true); } });
  C('showFormatting', { label: '&Formatting', checked: () => !!LY.opts.showFmtChanges, run: () => { LY.opts.showFmtChanges = !LY.opts.showFmtChanges; LY.render(); E.restoreDom(true); } });
  C('reviewersMenu', { label: '&Reviewers', run: () => L.review && L.review.reviewersDialog() });
  C('balloonsToggle', { label: '&Balloons', checked: () => LY.opts.balloons !== false, run: () => { LY.opts.balloons = !LY.opts.balloons; LY.render(); E.restoreDom(true); } });
  C('reviewingPane', { label: 'Reviewing &Pane', icon: 'reviewingPane', tip: 'Reviewing Pane', checked: () => !L.$('#revpane').hidden, run: () => L.review && L.review.togglePane() });
  C('prevChange', { label: 'Previous', icon: 'prevChange', tip: 'Previous', run: () => L.review && L.review.step(-1) });
  C('nextChange', { label: 'Next', icon: 'nextChange', tip: 'Next', run: () => L.review && L.review.step(1) });
  C('acceptChange', { label: '&Accept Change', icon: 'accept', tip: 'Accept Change', enabled: canEdit, run: () => L.review && L.review.acceptReject(true, 'selection') });
  C('acceptShown', { label: 'Accept All Changes &Shown', enabled: canEdit, run: () => L.review && L.review.acceptReject(true, 'all') });
  C('acceptAll', { label: 'Accept All Changes in &Document', enabled: canEdit, run: () => L.review && L.review.acceptReject(true, 'all') });
  C('rejectChange', { label: '&Reject Change/Delete Comment', icon: 'reject', tip: 'Reject Change/Delete Comment', enabled: canEdit, run: () => L.review && L.review.acceptReject(false, 'selection') });
  C('rejectShown', { label: 'Reject All Changes Shown', enabled: canEdit, run: () => L.review && L.review.acceptReject(false, 'all') });
  C('rejectAll', { label: 'Reject All Changes in Document', enabled: canEdit, run: () => L.review && L.review.acceptReject(false, 'all') });
  C('deleteAllComments', { label: 'Delete All Comments Shown', enabled: () => canEdit() && Object.keys(doc().comments).length > 0, run: () => L.review && L.review.deleteAllComments() });

  /* ================= Picture & drawing ================= */
  const imgSel = () => canEdit() && selectedObjIs((it) => it.t === 'img');
  C('picMoreContrast', { label: 'More Contrast', icon: 'moreContrast', enabled: imgSel, run: () => A.picAdjust('contrast', 0.1) });
  C('picLessContrast', { label: 'Less Contrast', icon: 'lessContrast', enabled: imgSel, run: () => A.picAdjust('contrast', -0.1) });
  C('picMoreBright', { label: 'More Brightness', icon: 'moreBright', enabled: imgSel, run: () => A.picAdjust('bright', 0.1) });
  C('picLessBright', { label: 'Less Brightness', icon: 'lessBright', enabled: imgSel, run: () => A.picAdjust('bright', -0.1) });
  C('picCrop', { label: 'Crop', icon: 'crop', enabled: imgSel, run: () => L.dlg.formatObject('picture') });
  C('picRotate', { label: 'Rotate Left 90°', icon: 'rotateL', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected('Rotate', (it) => { it.rot = (((it.rot || 0) - 90) % 360 + 360) % 360 || undefined; }) });
  C('picReset', { label: 'Reset Picture', icon: 'resetPicture', enabled: imgSel, run: () => L.drawing.modifySelected('Reset Picture', (it) => { delete it.bright; delete it.contrast; delete it.gray; delete it.bw; delete it.washout; delete it.crop; delete it.rot; if (it.natW) { it.w = it.natW; it.h = it.natH; } }) });
  C('selectObjects', { label: 'Select Objects', icon: 'select', checked: () => !!A.selectObjectsMode, run: () => { A.selectObjectsMode = !A.selectObjectsMode; L.$('#scroller').classList.toggle('selobj', A.selectObjectsMode); } });
  C('drawLine', { label: '&Line', icon: 'line', tip: 'Line', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'line', line: true }) });
  C('drawArrow', { label: '&Arrow', icon: 'arrowTool', tip: 'Arrow', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'line', line: true, arrow: true }) });
  C('drawDblArrow', { label: '&Double Arrow', icon: 'doubleArrow', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'line', line: true, dblArrow: true }) });
  C('drawRect', { label: '&Rectangle', icon: 'rect', tip: 'Rectangle', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'rect' }) });
  C('drawOval', { label: '&Oval', icon: 'oval', tip: 'Oval', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'ellipse' }) });
  C('fillColor', { label: 'Fill Color', icon: 'fillColor', tip: 'Fill Color', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected('Fill Color', (it) => { it.fill = { t: 'solid', c: A.lastFill || '#CCFFCC', a: 1 }; }) });
  C('lineColor', { label: 'Line Color', icon: 'lineColor', tip: 'Line Color', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected('Line Color', (it) => { it.line = Object.assign({ w: 0.75, dash: 'solid' }, it.line && it.line.t !== 'none' ? it.line : {}, { c: A.lastLine || '#000000' }); }) });
  const order = (label, fn) => ({ label, enabled: () => canEdit() && floatSel(), run: () => L.drawing.modifySelected(ui.stripAmp(label), fn) });
  C('bringFront', Object.assign(order('Bring to Fro&nt', (it) => { it.float.z = A.maxZ() + 1024; it.float.behind = false; if (it.float.wrap === 'behind') it.float.wrap = 'none'; }), { icon: 'bringFront' }));
  C('sendBack', Object.assign(order('Send to Bac&k', (it) => { it.float.z = Math.max(0, A.minZ() - 1024); }), { icon: 'sendBack' }));
  C('bringForward', Object.assign(order('Bring &Forward', (it) => { it.float.z = (it.float.z || 0) + 1024 * 1024; }), { icon: 'bringForward' }));
  C('sendBackward', Object.assign(order('Send Back&ward', (it) => { it.float.z = Math.max(0, (it.float.z || 0) - 1024 * 1024); }), { icon: 'sendBackward' }));
  C('bringFrontText', order('Bring in Front of &Text', (it) => { it.float.behind = false; if (it.float.wrap === 'behind') it.float.wrap = 'none'; }));
  C('sendBehindText', order('Send Behind T&ext', (it) => { it.float.behind = true; it.float.wrap = 'behind'; }));
  C('groupObj', { label: '&Group', icon: 'group', enabled: () => false, run: () => {} });
  C('ungroupObj', { label: '&Ungroup', icon: 'ungroup', enabled: () => canEdit() && selectedObjIs((it) => it.t === 'group'), run: () => A.ungroup() });
  for (const [k, dir, lbl] of [['nudgeUp', 'ArrowUp', '&Up'], ['nudgeDown', 'ArrowDown', '&Down'], ['nudgeLeft', 'ArrowLeft', '&Left'], ['nudgeRight', 'ArrowRight', '&Right']]) C(k, { label: lbl, enabled: () => canEdit() && floatSel(), run: () => L.drawing.nudge(dir, 1) });
  C('rotateLeft', { label: 'Rotate &Left 90°', icon: 'rotateL', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected('Rotate', (it) => { it.rot = (((it.rot || 0) - 90) % 360 + 360) % 360 || undefined; }) });
  C('rotateRight', { label: 'Rotate &Right 90°', icon: 'rotateR', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected('Rotate', (it) => { it.rot = ((it.rot || 0) + 90) % 360 || undefined; }) });
  C('flipH', { label: 'Flip &Horizontal', icon: 'flipH', enabled: () => canEdit() && selectedObjIs((it) => it.t === 'shape' || it.t === 'img'), run: () => L.drawing.modifySelected('Flip', (it) => { it.flipH = !it.flipH || undefined; }) });
  C('flipV', { label: 'Flip &Vertical', icon: 'flipV', enabled: () => canEdit() && selectedObjIs((it) => it.t === 'shape' || it.t === 'img'), run: () => L.drawing.modifySelected('Flip', (it) => { it.flipV = !it.flipV || undefined; }) });
  C('setDefaults', { label: 'Set AutoShape &Defaults', enabled: () => selectedObjIs((it) => it.t === 'shape'), run: () => { const it = L.drawing.selectedItem(); A.shapeDefaults = { fill: L.clone(it.fill), line: L.clone(it.line), shadow: it.shadow ? L.clone(it.shadow) : undefined }; A.status('New AutoShapes will use these settings.'); } });
  C('wordartEdit', { label: 'Edit Te&xt...', enabled: () => canEdit() && selectedObjIs((it) => !!it.wordart), run: () => L.dlg.wordart(true) });
  C('wordartGallery', { label: 'WordArt Gallery', icon: 'waGallery', enabled: () => canEdit() && selectedObjIs((it) => !!it.wordart), run: () => L.dlg.wordart('gallery') });

  /* ================= Outlining ================= */
  const ol = (fn, ...a) => () => L.outline && L.outline[fn](...a);
  C('promoteH1', { label: 'Promote to Heading 1', icon: 'promoteH1', enabled: canEdit, run: ol('setLevel', 1) });
  C('promote', { label: 'Promote', icon: 'promote', key: 'Alt+Shift+Left', enabled: canEdit, run: ol('shift', -1) });
  C('demote', { label: 'Demote', icon: 'demote', key: 'Alt+Shift+Right', enabled: canEdit, run: ol('shift', 1) });
  C('demoteBody', { label: 'Demote to Body Text', icon: 'demoteBody', enabled: canEdit, run: ol('setLevel', 0) });
  C('moveUp', { label: 'Move Up', icon: 'moveUp', key: 'Alt+Shift+Up', enabled: canEdit, run: ol('move', -1) });
  C('moveDown', { label: 'Move Down', icon: 'moveDown', key: 'Alt+Shift+Down', enabled: canEdit, run: ol('move', 1) });
  C('expand', { label: 'Expand', icon: 'expand', enabled: () => A.view === 'outline', run: ol('expand', true) });
  C('collapse', { label: 'Collapse', icon: 'collapse', enabled: () => A.view === 'outline', run: ol('expand', false) });
  C('olFirstLine', { label: 'Show First Line Only', icon: 'firstLine', enabled: () => A.view === 'outline', checked: () => !!(L.outline && L.outline.firstLine), run: ol('toggle', 'firstLine') });
  C('olFormatting', { label: 'Show Text Formatting', icon: 'showFormatting', enabled: () => A.view === 'outline', checked: () => !(L.outline && L.outline.plain), run: ol('toggle', 'plain') });
  C('updateTOC', { label: 'Update TOC', icon: 'updateTOC', tip: 'Update TOC', enabled: canEdit, run: () => L.fields && L.fields.updateAllOfType('TOC') });
  C('gotoTOC', { label: 'Go to TOC', icon: 'gotoTOC', tip: 'Go to TOC', run: () => L.fields && L.fields.gotoFirst('TOC') });

  /* ================= Header and Footer ================= */
  const hfOnly = () => !!LY.hfEdit;
  C('hfPageNum', { label: 'Insert Page Number', icon: 'pageNum', enabled: hfOnly, run: () => insertFieldNow('PAGE', 'Insert Page Number') });
  C('hfNumPages', { label: 'Insert Number of Pages', icon: 'numPages', enabled: hfOnly, run: () => insertFieldNow('NUMPAGES', 'Insert Number of Pages') });
  C('hfFormatPageNum', { label: 'Format Page Number', icon: 'formatPageNum', enabled: () => !!D.doc, run: () => L.dlg.pageNumbers('format') });
  C('hfDate', { label: 'Insert Date', icon: 'date', enabled: hfOnly, run: () => insertFieldNow('DATE \\@ "M/d/yyyy"', 'Insert Date') });
  C('hfTime', { label: 'Insert Time', icon: 'time', enabled: hfOnly, run: () => insertFieldNow('TIME \\@ "h:mm:ss am/pm"', 'Insert Time') });
  C('hfShowText', { label: 'Show/Hide Document Text', icon: 'showDocText', enabled: hfOnly, checked: () => LY.root && LY.root.classList.contains('hidebody'), run: () => LY.root.classList.toggle('hidebody') });
  C('hfLinkPrev', { label: 'Link to Previous', icon: 'linkPrev', enabled: () => hfOnly() && A.hfSectionIndex() > 0, checked: () => hfOnly() && A.hfLinked(), run: () => A.hfToggleLink() });
  C('hfSwitch', { label: 'Switch Between Header and Footer', icon: 'hfSwitch', enabled: hfOnly, run: () => A.hfSwitch() });
  C('hfPrev', { label: 'Show Previous', icon: 'showPrev', enabled: hfOnly, run: () => A.hfStep(-1) });
  C('hfNext', { label: 'Show Next', icon: 'showNext', enabled: hfOnly, run: () => A.hfStep(1) });
  C('hfClose', { label: 'Close Header and Footer', enabled: hfOnly, run: () => A.closeHeaderFooter() });

  /* ================= Mail Merge toolbar ================= */
  const mm = (fn, ...a) => () => L.mailmerge && L.mailmerge[fn](...a);
  const hasData = () => !!(L.mailmerge && L.mailmerge.data);
  C('mmMainDoc', { label: 'Main document setup', icon: 'mmMain', run: pane('mailmerge') });
  C('mmOpenSource', { label: 'Open Data Source', icon: 'open', run: mm('openSource') });
  C('mmRecipients', { label: 'Mail Merge Recipients', icon: 'mmRecipients', enabled: hasData, run: mm('recipientsDialog') });
  C('mmAddress', { label: 'Insert Address Block', icon: 'mmAddress', enabled: () => canEdit() && hasData(), run: mm('insertBlock', 'address') });
  C('mmGreeting', { label: 'Insert Greeting Line', icon: 'mmGreeting', enabled: () => canEdit() && hasData(), run: mm('insertBlock', 'greeting') });
  C('mmInsertField', { label: 'Insert Merge Fields', icon: 'mmField', enabled: () => canEdit() && hasData(), run: mm('fieldDialog') });
  C('mmViewData', { label: 'View Merged Data', icon: 'mmView', enabled: hasData, checked: () => !!(L.mailmerge && L.mailmerge.preview), run: mm('togglePreview') });
  C('mmFirst', { label: 'First Record', icon: 'first', enabled: hasData, run: mm('go', 0) });
  C('mmPrev', { label: 'Previous Record', icon: 'showPrev', enabled: hasData, run: mm('step', -1) });
  C('mmNext', { label: 'Next Record', icon: 'showNext', enabled: hasData, run: mm('step', 1) });
  C('mmLast', { label: 'Last Record', icon: 'last', enabled: hasData, run: mm('go', -1) });
  C('mmMergeNew', { label: 'Merge to New Document', icon: 'mmMergeNew', enabled: hasData, run: mm('mergeToNew') });
  C('mmMergePrint', { label: 'Merge to Printer', icon: 'mmMergePrint', enabled: hasData, run: mm('mergeToPDF') });

  /* ================= navigation ================= */
  C('browsePrev', { label: 'Previous', icon: 'browseUp', tip: 'Previous Page', run: () => A.browse(-1) });
  C('browseNext', { label: 'Next', icon: 'browseDown', tip: 'Next Page', run: () => A.browse(1) });
  C('extendSel', { label: 'Extend Selection', key: 'F8', run: () => A.extendStep() });
  C('goBack', { label: 'Go Back', key: 'Shift+F5', run: () => A.goBack() });

  /* ================= helpers used by commands ================= */
  A.maxZ = () => { let m = 251659264; D.walk(doc().main, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.float && it.float.z > m) m = it.float.z; }); return m; };
  A.minZ = () => { let m = 251659264; D.walk(doc().main, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.float && it.float.z < m) m = it.float.z; }); return m; };
  A.toggleFullScreen = function () {
    const app = L.$('#app');
    const on = !app.classList.contains('fullscreen');
    app.classList.toggle('fullscreen', on);
    let btn = L.$('.closefs');
    if (on && !btn) { btn = L.h('button', { class: 'btn closefs', type: 'button', text: 'Close Full Screen', onclick: () => A.toggleFullScreen() }); document.body.appendChild(btn); }
    if (!on && btn) btn.remove();
    LY.fitSizer();
    if (L.rulers) L.rulers.draw();
  };
  A.cycleCase = function () {
    if (!E.sel) return;
    let [a, b] = E.range();
    if (D.eqPos(a, b)) { const w = O.wordAt(a); if (!w) return; [a, b] = w; }
    const t = O.textRange(a, b, '\n');
    const mode = t === t.toUpperCase() && t !== t.toLowerCase() ? 'lower' : t === t.toLowerCase() ? 'title' : 'upper';
    E.edit('Change Case', () => { O.changeCase(a, b, mode); return { a, f: b }; });
  };
  A.removeHyperlink = function () {
    if (!E.sel) return;
    let [a, b] = E.range();
    if (D.eqPos(a, b)) {
      /* the whole link around the caret */
      const p = a.p;
      let pos = 0, s = -1, e = -1;
      for (const it of p.runs) { const n = D.ilen(it); if (it.rPr && it.rPr.link && pos <= a.o && a.o <= pos + n) { if (s < 0) s = pos; e = pos + n; } else if (s >= 0 && pos > a.o) break; pos += n; }
      if (s < 0) return;
      a = D.pos(p, s); b = D.pos(p, e);
    }
    E.edit('Remove Hyperlink', () => { O.formatRange(a, b, (r) => { delete r.link; if (r.style === 'Hyperlink' || r.style === 'FollowedHyperlink') delete r.style; return r; }); return E.sel; });
  };
  A.formatObject = function (editContent) {
    const it = L.drawing && L.drawing.selectedItem();
    if (!it) return;
    if (editContent && it.t === 'chart' && L.chart && L.chart.edit) { A.editChart(); return; }
    if (editContent && it.wordart) { L.dlg.wordart(true); return; }
    if (L.dlg && L.dlg.formatObject) L.dlg.formatObject();
  };
  A.insertChart = function () {
    const c = L.chart.sample ? L.chart.sample('bar') : null;
    if (!c) return;
    const it = { t: 'chart', w: 360, h: 216, chart: c };
    E.edit('Insert Chart', () => O.insertItem(E.deleteSelection(), it));
    setTimeout(() => { const os = A.selectItemAt(E.sel.f.p, E.sel.f.o - 1); if (os) A.editChart(); }, 30);
  };
  A.selectItemAt = function (p, o) {
    const f = LY.fragFor(p, o);
    const el = f && f.querySelector(`[data-o="${o}"]`);
    if (el) { E.selectInlineObject(D.pos(p, o), el); return E.objSel; }
    return null;
  };
  A.editChart = function () {
    const os = E.objSel;
    const it = L.drawing.selectedItem();
    if (!os || !it || it.t !== 'chart') return;
    const shape = { chart: L.clone(it.chart), w: it.w, h: it.h };
    L.chart.edit(shape, (res) => { if (!res) return; L.drawing.modifySelected('Chart', (x) => { x.chart = res.chart || shape.chart; }); });
  };
  A.ungroup = function () {
    const os = E.objSel;
    const it = L.drawing.selectedItem();
    if (!os || !it || it.t !== 'group') return;
    E.edit('Ungroup', () => {
      D.touch(os.p);
      const i = os.p.runs.indexOf(it);
      const baseH = it.float ? it.float.posH : null, baseV = it.float ? it.float.posV : null;
      const kids = (it.kids || []).map((k) => {
        const c = L.clone(k);
        c.float = L.clone(it.float || { posH: { rel: 'column', off: 0 }, posV: { rel: 'paragraph', off: 0 }, wrap: 'none', dist: { t: 0, b: 0, l: 9, r: 9 }, z: 251659264 });
        c.float.posH = { rel: baseH ? baseH.rel : 'column', off: L.round(((baseH && baseH.off) || 0) + (k.x || 0), 2) };
        c.float.posV = { rel: baseV ? baseV.rel : 'paragraph', off: L.round(((baseV && baseV.off) || 0) + (k.y || 0), 2) };
        delete c.x; delete c.y;
        return c;
      });
      os.p.runs.splice(i, 1, ...kids);
      return D.pos(os.p, os.o);
    });
    E.objSel = null;
    L.drawing.clearSelection();
  };
  A.insertFileDialog = async function () {
    const files = await L.pickFiles('.docx,.docm,.dotx,.txt,.htm,.html');
    const f = files[0];
    if (!f) return;
    try {
      let blocks;
      if (/\.txt$/i.test(f.name)) blocks = A.docFromText(await L.readAsText(f)).main.blocks;
      else if (/\.html?$/i.test(f.name)) blocks = await L.htmlio.htmlToBlocks(await L.readAsText(f));
      else {
        const res = await L.docx.read(await L.readAsArrayBuffer(f));
        const src = res.doc;
        A.mergeStylesFrom(src);
        blocks = src.main.blocks;
        /* carry notes and comments along */
        const d = doc();
        E.edit('Insert File', () => {
          D.touchKey(d, 'fn'); D.touchKey(d, 'en'); D.touchKey(d, 'comments'); D.touchKey(d, 'numbering');
          for (const k in src.fn) if (!d.fn[k]) d.fn[k] = src.fn[k];
          for (const k in src.en) if (!d.en[k]) d.en[k] = src.en[k];
          for (const k in src.comments) if (!d.comments[k]) d.comments[k] = src.comments[k];
          A.mergeNumbering(src, blocks);
          return E.sel;
        });
      }
      E.pasteBlocks(blocks);
    } catch (e) { ui.msg('The file could not be inserted. ' + (e.message || ''), { icon: 'error' }); }
  };
  /** copy styles missing from the current document */
  A.mergeStylesFrom = function (src, overwrite) {
    const d = doc();
    E.edit('Copy Styles', () => {
      D.touchKey(d, 'styles');
      for (const k in src.styles) if (overwrite || !d.styles[k]) d.styles[k] = L.clone(src.styles[k]);
      D.stylesChanged();
      return E.sel;
    });
  };
  /** remap list ids of blocks coming from another document */
  A.mergeNumbering = function (src, blocks) {
    const d = doc();
    const map = {};
    const absMap = {};
    for (const nid in src.numbering.nums) {
      const n = src.numbering.nums[nid];
      if (absMap[n.abs] == null) { const abs = L.clone(src.numbering.abs[n.abs]); let aid = 0; while (d.numbering.abs[aid] != null) aid++; abs.id = aid; d.numbering.abs[aid] = abs; absMap[n.abs] = aid; }
      let k = 1; while (d.numbering.nums[k] != null) k++;
      d.numbering.nums[k] = { abs: absMap[n.abs], ov: L.clone(n.ov || {}) };
      map[nid] = String(k);
    }
    const fix = (bl) => D.walk({ blocks: bl }, (b) => { if (b.t === 'p' && b.pPr.num && map[b.pPr.num.id]) b.pPr.num = Object.assign({}, b.pPr.num, { id: map[b.pPr.num.id] }); });
    fix(blocks);
  };

  /* ---------- split window ---------- */
  A.toggleSplit = function () {
    A.splitOn = !A.splitOn;
    let pane2 = L.$('#split2');
    if (A.splitOn) {
      if (!pane2) {
        pane2 = L.h('div', { id: 'split2', class: 'split2', 'aria-label': 'Second pane' });
        const bar = L.h('div', { class: 'splitter h', id: 'split2-bar', role: 'separator' });
        L.$('#center').insertBefore(bar, L.$('#viewbar'));
        L.$('#center').insertBefore(pane2, bar);
        bar.addEventListener('pointerdown', (e) => {
          e.preventDefault(); bar.setPointerCapture(e.pointerId);
          const y0 = e.clientY, h0 = pane2.offsetHeight;
          const mv = (ev) => { pane2.style.height = L.clamp(h0 - (ev.clientY - y0), 60, window.innerHeight * 0.7) + 'px'; };
          const up = () => { bar.removeEventListener('pointermove', mv); bar.removeEventListener('pointerup', up); };
          bar.addEventListener('pointermove', mv); bar.addEventListener('pointerup', up);
        });
        pane2.addEventListener('click', (e) => {
          const pg = e.target.closest('.pg');
          if (!pg) return;
          const i = +pg.dataset.pg;
          const real = LY.pages[i];
          if (real) real.el.scrollIntoView({ block: 'start' });
        });
      }
      L.$('#split2-bar').hidden = false; pane2.hidden = false;
      A.refreshSplit();
    } else if (pane2) { pane2.hidden = true; L.$('#split2-bar').hidden = true; }
  };
  A.refreshSplit = L.debounce(function () {
    const pane2 = L.$('#split2');
    if (!pane2 || pane2.hidden) return;
    const top = pane2.scrollTop;
    pane2.textContent = '';
    const inner = LY.root.cloneNode(true);
    inner.removeAttribute('id');
    inner.contentEditable = 'false';
    inner.style.position = 'relative'; inner.style.left = '0'; inner.style.top = '8px';
    inner.style.transform = `scale(${LY.zoom})`;
    const sizer = L.h('div', { style: `height:${LY.root.offsetHeight * LY.zoom + 20}px;position:relative;display:flex;justify-content:center` }, inner);
    inner.style.transformOrigin = '50% 0';
    pane2.appendChild(sizer);
    pane2.scrollTop = top;
  }, 300);
  L.bus.on('layout-done', () => { if (A.splitOn) A.refreshSplit(); });

  /* ---------- extend selection (F8) ---------- */
  A.extendStep = function () {
    if (!E.sel) return;
    A.extLevel = E.extend ? (A.extLevel || 0) + 1 : 0;
    E.extend = true;
    if (!A.extLevel) A.extAnchor = E.sel.a;
    else if (A.extLevel === 1) { const w = O.wordExtent(E.sel.f); E.setSel(w[0], w[1]); }
    else if (A.extLevel === 2) { const s = O.sentenceExtent(E.sel.f); E.setSel(s[0], s[1]); }
    else if (A.extLevel === 3) E.selectPara(E.sel.a.p);
    else E.selectAll();
    A.updateStatus();
  };
  L.bus.on('sel', () => {
    if (!E.extend || !A.extAnchor || A.extBusy) return;
    if (A.extLevel) return;
    if (E.sel && !D.eqPos(E.sel.a, A.extAnchor) && D.storyOf(doc(), E.sel.f.p) === D.storyOf(doc(), A.extAnchor.p)) { A.extBusy = true; E.setSel(A.extAnchor, E.sel.f, { noScroll: true }); A.extBusy = false; }
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && E.extend) { E.extend = false; A.extAnchor = null; A.updateStatus(); } }, true);

  /* ---------- go back (Shift+F5): last three edit locations ---------- */
  A.editTrail = [];
  L.bus.on('doc-changed', (info) => {
    if (!E.sel || (info && info.history)) return;
    const p = E.sel.f;
    const last = A.editTrail[0];
    if (last && last.p === p.p && Math.abs(last.o - p.o) < 40) { A.editTrail[0] = { p: p.p, o: p.o }; return; }
    A.editTrail.unshift({ p: p.p, o: p.o });
    A.editTrail.length = Math.min(A.editTrail.length, 4);
  });
  A.goBack = function () {
    if (A.editTrail.length < 2) { if (A.editTrail[0]) A.gotoPos(D.pos(A.editTrail[0].p, A.editTrail[0].o)); return; }
    const t = A.editTrail.shift();
    A.editTrail.push(t);
    const target = A.editTrail[0];
    if (D.info(doc(), target.p)) A.gotoPos(D.pos(target.p, Math.min(target.o, D.plen(target.p))));
  };

  /* ---------- header/footer navigation ---------- */
  A.hfSectionIndex = () => { if (!LY.hfEdit) return 0; const pg = LY.pages[LY.hfEdit.page]; return pg ? pg.si : 0; };
  A.hfLinked = function () {
    const pg = LY.pages[LY.hfEdit.page];
    if (!pg) return false;
    const secs = D.sections(doc());
    const own = secs[pg.si].sect.refs && secs[pg.si].sect.refs[LY.hfEdit.kind] && secs[pg.si].sect.refs[LY.hfEdit.kind][pg.type];
    return !own;
  };
  A.hfToggleLink = function () {
    const d = doc();
    const pg = LY.pages[LY.hfEdit.page];
    const secs = D.sections(d);
    const kind = LY.hfEdit.kind;
    const s = secs[pg.si];
    const holder = s.endPara || d;
    E.edit('Link to Previous', () => {
      D.touchKey(holder, 'sect');
      const sect = holder === d ? d.sect : holder.sect;
      sect.refs = sect.refs || { hdr: {}, ftr: {} };
      sect.refs[kind] = Object.assign({}, sect.refs[kind] || {});
      if (sect.refs[kind][pg.type]) delete sect.refs[kind][pg.type];
      else {
        /* unlink: copy the inherited story into a new one for this section */
        const inh = LY.hfFor(secs, pg.si, kind, pg.type);
        const id = (kind === 'hdr' ? 'h' : 'f') + D.nid();
        D.touchKey(d, 'hf');
        d.hf[id] = { kind, id, blocks: inh ? D.cloneBlocks(inh.blocks) : [D.para([], { style: kind === 'hdr' ? 'Header' : 'Footer' })] };
        sect.refs[kind][pg.type] = id;
      }
      return E.sel;
    });
    A.editHeaderFooter(kind, pg.idx);
  };
  A.hfSwitch = () => { const pg = LY.hfEdit.page; A.editHeaderFooter(LY.hfEdit.kind === 'hdr' ? 'ftr' : 'hdr', pg); const p = LY.pages[pg]; if (p) p.el.scrollIntoView({ block: LY.hfEdit.kind === 'hdr' ? 'start' : 'end' }); };
  A.hfStep = function (dir) {
    /* next page whose header/footer story differs */
    const cur = LY.pages[LY.hfEdit.page];
    const kind = LY.hfEdit.kind;
    const story = (pg) => (kind === 'hdr' ? pg.hdrStory : pg.ftrStory);
    for (let i = LY.hfEdit.page + dir; i >= 0 && i < LY.pages.length; i += dir) {
      const pg = LY.pages[i];
      if (pg.si !== cur.si || pg.type !== cur.type || story(pg) !== story(cur)) { A.editHeaderFooter(kind, i); pg.el.scrollIntoView({ block: 'start' }); return; }
    }
  };
})();
