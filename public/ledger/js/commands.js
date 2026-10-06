/* Ledger — command definitions (menus, toolbars, shortcuts all run these). */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, LY = L.layout, O = L.ops, F = L.formula, NF = L.numfmt, C = L.calc;
  const ui = L.ui;
  const cmd = ui.cmd;
  const G = () => L.grid, A = () => L.app, D = () => L.dlg, E = () => L.editor, CL = () => L.clip;
  const MAXR = M.MAXR, MAXC = M.MAXC;
  const sh = () => G().sheet();
  const wb = () => G().wb;
  const sel = () => G().sel();
  const rgs = () => G().ranges();
  const act = () => { const s = sel(); return LY.styleOf(sh(), s.r, s.c, sh().get(s.r, s.c)); };
  const notEditing = () => !E().active;
  const guard = (key) => G().guardProtect(key, !key);
  const run = (fn) => () => { if (E().active && !E().commit()) return; try { fn(); } catch (e) { A().error(e); } G().paint(); };
  /** apply a style change to the selection */
  const fmt = (delta, label) => run(() => { if (guard('formatCells')) return; O.format(sh(), rgs(), delta, label); });
  const hasMulti = () => rgs().length > 1;
  const wholeCols = () => rgs().every((x) => G().isWholeCols(x));
  const wholeRows = () => rgs().every((x) => G().isWholeRows(x));

  /* ------------------------------------------------------------ File */
  cmd('newBook', { label: '&New...', menuLabel: '&New...', tip: 'New', icon: 'new', key: 'Ctrl+N', run: () => A().newWorkbook() });
  cmd('newBlank', { label: 'New Blank Workbook', icon: 'new', run: () => A().newWorkbook() });
  cmd('open', { label: '&Open...', icon: 'open', key: 'Ctrl+O', run: () => A().openDialog() });
  cmd('close', { label: '&Close', run: () => A().closeWorkbook() });
  cmd('save', { label: '&Save', icon: 'save', key: 'Ctrl+S', run: run(() => A().save()) });
  cmd('saveAs', { label: 'Save &As...', key: 'F12', run: run(() => D().saveAs()) });
  cmd('saveWeb', { label: 'Save as Web Pa&ge...', run: run(() => A().saveAs('html')) });
  cmd('exportPDF', { label: 'Export to &PDF...', icon: 'print', run: run(() => L.print.exportPDF()) });
  cmd('exportCSV', { label: 'Save as &CSV', run: run(() => A().saveAs('csv')) });
  cmd('pageSetup', { label: 'Page Set&up...', icon: 'pageSetup', run: run(() => D().pageSetup()) });
  cmd('setPrintArea', { label: '&Set Print Area', run: run(() => A().setPrintArea(true)) });
  cmd('clearPrintArea', { label: '&Clear Print Area', run: run(() => A().setPrintArea(false)), enabled: () => !!A().printAreaOf(sh()) });
  cmd('printPreview', { label: 'Print Pre&view', icon: 'preview', run: run(() => L.print.preview()) });
  cmd('print', { label: '&Print...', icon: 'print', key: 'Ctrl+P', run: run(() => D().print()) });
  cmd('printQuick', { label: 'Print', tip: 'Print', icon: 'print', run: run(() => L.print.exportPDF({ quick: true })) });
  cmd('properties', { label: 'Propert&ies', icon: 'properties', run: run(() => D().properties()) });
  cmd('exitApp', { label: 'E&xit', run: () => A().closeWorkbook(true) });

  /* ------------------------------------------------------------ Edit */
  cmd('undo', { label: '&Undo', icon: 'undo', key: 'Ctrl+Z', get menuLabel() { return wb() && wb().undo.canUndo() ? '&Undo ' + wb().undo.undoLabel : "Can't &Undo"; }, enabled: () => wb() && wb().undo.canUndo(), run: () => A().undo() });
  cmd('redo', { label: '&Redo', icon: 'redo', key: 'Ctrl+Y', get menuLabel() { return wb() && wb().undo.canRedo() ? '&Redo ' + wb().undo.redoLabel : A().lastCommand ? '&Repeat ' + A().lastCommand.label : "Can't &Repeat"; }, enabled: () => wb() && (wb().undo.canRedo() || !!A().lastCommand), run: () => A().redo() });
  cmd('cut', { label: 'Cu&t', icon: 'cut', key: 'Ctrl+X', run: run(() => CL().copy(null, true)) });
  cmd('copy', { label: '&Copy', icon: 'copy', key: 'Ctrl+C', run: run(() => CL().copy(null, false)) });
  cmd('paste', { label: '&Paste', icon: 'paste', key: 'Ctrl+V', run: run(() => CL().pasteCommand()) });
  cmd('officeClipboard', { label: 'Office Clip&board...', icon: 'clipboard', run: () => L.panes.task.show('clipboard') });
  cmd('pasteSpecial', { label: 'Paste &Special...', enabled: () => CL().canPaste(), run: run(() => D().pasteSpecial()) });
  cmd('pasteValues', { label: 'Paste &Values', enabled: () => CL().canPaste(), run: run(() => CL().pasteInternal({ what: 'values' })) });
  cmd('pasteHyperlink', { label: 'Paste as &Hyperlink', enabled: () => CL().canPaste(), run: run(() => A().pasteAsHyperlink()) });
  cmd('fillDown', { label: '&Down', key: 'Ctrl+D', run: run(() => A().fillDir('down')) });
  cmd('fillRight', { label: '&Right', key: 'Ctrl+R', run: run(() => A().fillDir('right')) });
  cmd('fillUp', { label: '&Up', run: run(() => A().fillDir('up')) });
  cmd('fillLeft', { label: '&Left', run: run(() => A().fillDir('left')) });
  cmd('fillSeries', { label: '&Series...', run: run(() => D().series()) });
  cmd('fillJustify', { label: '&Justify', run: run(() => A().justifyFill()) });
  cmd('fillAcross', { label: '&Across Worksheets...', enabled: () => (A().selectedSheets() || []).length > 1, run: run(() => D().fillAcross()) });
  cmd('clearAll', { label: '&All', run: run(() => { if (guard()) return; O.clear(sh(), rgs(), 'all'); }) });
  cmd('clearFormats', { label: '&Formats', run: run(() => { if (guard('formatCells')) return; O.clear(sh(), rgs(), 'formats'); }) });
  cmd('clearContents', { label: '&Contents', key: 'Del', run: run(() => { if (guard()) return; const d = G().selectedObject(); if (d) { A().deleteObject(d); return; } O.clear(sh(), rgs(), 'contents'); }) });
  cmd('clearComments', { label: 'Co&mments', run: run(() => O.clear(sh(), rgs(), 'comments')) });
  cmd('clearHyperlinks', { label: '&Hyperlinks', run: run(() => O.clear(sh(), rgs(), 'hyperlinks')) });
  cmd('deleteCells', { label: '&Delete...', key: 'Ctrl+-', run: run(() => A().deleteSelection()) });
  cmd('deleteSheet', { label: 'De&lete Sheet', icon: 'deleteSheet', run: run(() => A().deleteSheets()) });
  cmd('moveCopySheet', { label: '&Move or Copy Sheet...', run: run(() => D().moveCopySheet()) });
  cmd('find', { label: '&Find...', icon: 'find', key: 'Ctrl+F', run: run(() => D().find('find')) });
  cmd('replace', { label: 'R&eplace...', icon: 'replace', key: 'Ctrl+H', run: run(() => D().find('replace')) });
  cmd('goTo', { label: '&Go To...', key: 'Ctrl+G', run: run(() => D().goTo()) });
  cmd('goToSpecial', { label: 'Go To &Special...', run: run(() => D().goToSpecial()) });
  cmd('links', { label: 'Lin&ks...', enabled: () => !!(wb() && wb().externalLinks && wb().externalLinks.length), run: run(() => D().links()) });
  cmd('selectAll', { label: 'Select &All', key: 'Ctrl+A', run: run(() => A().selectAllSmart()) });

  /* ------------------------------------------------------------ View */
  cmd('viewNormal', { label: '&Normal', icon: 'normalGrid', radio: true, checked: () => !sh().view.pageBreakPreview, run: run(() => A().setPageBreakPreview(false)) });
  cmd('viewPageBreak', { label: '&Page Break Preview', icon: 'pageBreakView', radio: true, checked: () => !!sh().view.pageBreakPreview, run: run(() => A().setPageBreakPreview(true)) });
  cmd('taskPane', { label: 'Tas&k Pane', key: 'Ctrl+F1', checked: () => !L.$('#taskpane').hidden, run: () => A().toggleTaskPane() });
  cmd('formulaBar', { label: '&Formula Bar', checked: () => A().opts.formulaBar !== false, run: () => { A().opts.formulaBar = A().opts.formulaBar === false; A().applyOpts(); } });
  cmd('statusBar', { label: '&Status Bar', checked: () => A().opts.statusBar !== false, run: () => { A().opts.statusBar = A().opts.statusBar === false; A().applyOpts(); } });
  cmd('headerFooter', { label: '&Header and Footer...', run: run(() => D().pageSetup(2)) });
  cmd('showComments', { label: '&Comments', icon: 'comment', checked: () => !!wb()._showComments, run: run(() => { wb()._showComments = !wb()._showComments; A().toggleToolbar('reviewing', wb()._showComments ? true : undefined); }) });
  cmd('fullScreen', { label: 'F&ull Screen', icon: 'fullScreen', checked: () => L.$('#app').classList.contains('fullscreen'), run: () => A().toggleFullScreen() });
  cmd('zoomDlg', { label: '&Zoom...', icon: 'zoom', run: run(() => D().zoom()) });
  cmd('zoomIn', { label: 'Zoom In', run: () => G().setZoom((sh().view.zoom || 100) + 10) });
  cmd('zoomOut', { label: 'Zoom Out', run: () => G().setZoom((sh().view.zoom || 100) - 10) });
  cmd('showFormulas', { label: 'Formula Auditing &Mode', icon: 'showFormulas', key: 'Ctrl+`', checked: () => !!sh().view.formulas, run: () => { sh().view.formulas = !sh().view.formulas; G().paint(); ui.refresh(); } });
  cmd('gridlinesView', { label: '&Gridlines', checked: () => sh().view.grid !== false, run: () => { sh().view.grid = sh().view.grid === false; G().paint(); } });
  cmd('headingsView', { label: 'Row & Column &Headers', checked: () => sh().view.headings !== false, run: () => { sh().view.headings = sh().view.headings === false; G().paint(); } });
  for (const [id, label] of [['standard', '&Standard'], ['formatting', '&Formatting'], ['drawing', '&Drawing'], ['chart', '&Chart'], ['auditing', 'Formula &Auditing'], ['reviewing', '&Reviewing'], ['picture', '&Picture'], ['borders', '&Borders'], ['pivot', 'PivotTa&ble'], ['forms', 'For&ms'], ['list', '&List']]) {
    cmd('tb_' + id, { label, checked: () => A().toolbarVisible(id), run: () => A().toggleToolbar(id) });
  }
  cmd('customize', { label: '&Customize...', run: run(() => D().customize()) });

  /* ------------------------------------------------------------ Insert */
  cmd('insertCells', { label: 'Ce&lls...', icon: 'insertCells', key: 'Ctrl+Shift++', run: run(() => A().insertSelection()) });
  cmd('insertRows', { label: '&Rows', icon: 'insertRows', run: run(() => A().insertLines('r')) });
  cmd('insertCols', { label: '&Columns', icon: 'insertCols', run: run(() => A().insertLines('c')) });
  cmd('insertSheet', { label: '&Worksheet', icon: 'insertSheet', key: 'Shift+F11', run: run(() => A().insertSheet()) });
  cmd('insertChart', { label: 'C&hart...', icon: 'chartWizard', key: 'F11', run: run(() => L.chartWizard.open()) });
  cmd('chartWizard', { label: 'Chart Wizard', tip: 'Chart Wizard', icon: 'chartWizard', run: run(() => L.chartWizard.open()) });
  cmd('insertSymbol', { label: '&Symbol...', icon: 'symbol', run: run(() => D().symbol()) });
  cmd('insertPageBreak', { get label() { return A().pageBreakAtSel() ? 'Remove Page &Break' : 'Page &Break'; }, icon: 'pageBreak', run: run(() => A().togglePageBreak()) });
  cmd('resetPageBreaks', { label: 'Reset All Page Breaks', run: run(() => A().resetPageBreaks()) });
  cmd('insertFunction', { label: '&Function...', icon: 'fx', key: 'Shift+F3', run: () => D().insertFunction() });
  cmd('functionArgs', { label: 'Function Arguments', run: () => D().functionArgs() });
  cmd('nameDefine', { label: '&Define...', key: 'Ctrl+F3', run: run(() => D().defineName()) });
  cmd('pasteName', { label: '&Paste...', key: 'F3', run: () => D().pasteName() });
  cmd('nameCreate', { label: '&Create...', key: 'Ctrl+Shift+F3', run: run(() => D().createNames()) });
  cmd('nameApply', { label: '&Apply...', run: run(() => D().applyNames()) });
  cmd('insertComment', { get label() { const s = sel(); return sh().comments.get(M.key(s.r, s.c)) ? 'Edit Co&mment' : 'Co&mment'; }, icon: 'comment', key: 'Shift+F2', run: run(() => A().editComment()) });
  cmd('deleteComment', { label: 'Delete Co&mment', enabled: () => { const s = sel(); return !!sh().comments.get(M.key(s.r, s.c)); }, run: run(() => { const s = sel(); O.tx(wb(), 'Delete Comment', () => O.setComment(sh(), s.r, s.c, null)); }) });
  cmd('showHideComment', { label: 'Show/Hide Co&mment', enabled: () => { const s = sel(); return !!sh().comments.get(M.key(s.r, s.c)); }, run: run(() => { const s = sel(); const cm = sh().comments.get(M.key(s.r, s.c)); O.tx(wb(), 'Show Comment', () => O.setComment(sh(), s.r, s.c, Object.assign({}, cm, { visible: !cm.visible }))); }) });
  cmd('insertPicture', { label: '&From File...', icon: 'picture', run: run(() => A().insertPicture()) });
  cmd('insertClipArt', { label: '&Clip Art...', icon: 'clipart', run: run(() => A().insertPicture()) });
  cmd('insertShape', { label: '&AutoShapes', icon: 'autoshapes', run: run(() => L.drawing.startTool('rect')) });
  cmd('insertTextBox', { label: 'Te&xt Box', icon: 'textbox', run: run(() => L.drawing.startTool('textbox')) });
  cmd('hyperlink', { label: 'Hyperl&ink...', icon: 'hyperlink', key: 'Ctrl+K', run: run(() => D().hyperlink()) });
  cmd('removeHyperlink', { label: '&Remove Hyperlink', enabled: () => { const s = sel(); return sh().links.some((l) => M.rangeContains(l.ref, s.r, s.c)); }, run: run(() => O.clear(sh(), rgs(), 'hyperlinks')) });

  /* ------------------------------------------------------------ Format */
  cmd('formatCells', { label: 'C&ells...', icon: 'formatCells', key: 'Ctrl+1', run: run(() => { const d = G().selectedObject(); if (d) { L.drawing.format(d); return; } if (guard('formatCells')) return; D().formatCells(); }) });
  cmd('rowHeight', { label: 'H&eight...', icon: 'rowHeight', run: run(() => { if (guard('formatRows')) return; D().rowHeight(); }) });
  cmd('rowAutofit', { label: '&AutoFit', run: run(() => { if (guard('formatRows')) return; O.autofitRows(sh(), A().selRows()); }) });
  cmd('rowHide', { label: '&Hide', key: 'Ctrl+9', run: run(() => { if (guard('formatRows')) return; O.hideLines(sh(), 'r', A().selRows(), true); }) });
  cmd('rowUnhide', { label: '&Unhide', key: 'Ctrl+Shift+(', run: run(() => { if (guard('formatRows')) return; O.hideLines(sh(), 'r', A().selRows(true), false); }) });
  cmd('colWidth', { label: 'W&idth...', icon: 'colWidth', run: run(() => { if (guard('formatColumns')) return; D().colWidth(); }) });
  cmd('colAutofit', { label: '&AutoFit Selection', run: run(() => { if (guard('formatColumns')) return; const s = rgs()[0]; O.autofitCols(sh(), A().selCols(), G().isWholeCols(s) ? null : { r1: s.r1, r2: s.r2 }); }) });
  cmd('colHide', { label: '&Hide', key: 'Ctrl+0', run: run(() => { if (guard('formatColumns')) return; O.hideLines(sh(), 'c', A().selCols(), true); }) });
  cmd('colUnhide', { label: '&Unhide', key: 'Ctrl+Shift+)', run: run(() => { if (guard('formatColumns')) return; O.hideLines(sh(), 'c', A().selCols(true), false); }) });
  cmd('colStandard', { label: '&Standard Width...', run: run(() => D().standardWidth()) });
  cmd('sheetRename', { label: '&Rename', run: run(() => A().renameSheetInline()) });
  cmd('sheetHide', { label: '&Hide', run: run(() => A().hideSheets()) });
  cmd('sheetUnhide', { label: '&Unhide...', enabled: () => wb().sheets.some((s) => s.state === 'hidden'), run: run(() => D().unhideSheet()) });
  cmd('sheetBackground', { get label() { return sh().background ? 'Delete Bac&kground' : '&Background...'; }, run: run(() => A().sheetBackground()) });
  cmd('tabColor', { label: '&Tab Color...', icon: 'tabColor', run: run(() => D().tabColor()) });
  cmd('autoFormat', { label: '&AutoFormat...', icon: 'tableAutoFormat', run: run(() => { if (guard('formatCells')) return; D().autoFormat(); }) });
  cmd('conditionalFormat', { label: 'Con&ditional Formatting...', run: run(() => { if (guard('formatCells')) return; D().conditionalFormat(); }) });
  cmd('styleDlg', { label: '&Style...', run: run(() => { if (guard('formatCells')) return; D().style(); }) });
  cmd('formatObject', { label: 'Selected &Object...', enabled: () => !!G().selectedObject(), run: run(() => L.drawing.format(G().selectedObject())) });

  /* formatting toolbar */
  cmd('bold', { label: '&Bold', icon: 'bold', key: 'Ctrl+B', checked: () => !!(act().font && act().font.b), run: run(() => { if (guard('formatCells')) return; O.toggleFont(sh(), rgs(), 'b', act()); }) });
  cmd('italic', { label: '&Italic', icon: 'italic', key: 'Ctrl+I', checked: () => !!(act().font && act().font.i), run: run(() => { if (guard('formatCells')) return; O.toggleFont(sh(), rgs(), 'i', act()); }) });
  cmd('underline', { label: '&Underline', icon: 'underline', key: 'Ctrl+U', checked: () => !!(act().font && act().font.u), run: run(() => { if (guard('formatCells')) return; O.toggleFont(sh(), rgs(), 'u', act()); }) });
  cmd('strike', { label: 'Stri&kethrough', icon: 'strike', key: 'Ctrl+5', checked: () => !!(act().font && act().font.strike), run: run(() => { if (guard('formatCells')) return; O.toggleFont(sh(), rgs(), 'strike', act()); }) });
  const halign = (h) => run(() => { if (guard('formatCells')) return; const cur = act().align && act().align.h; O.format(sh(), rgs(), { align: { h: cur === h ? undefined : h } }, 'Align'); });
  cmd('alignLeft', { label: 'Align &Left', icon: 'alignLeft', checked: () => (act().align || {}).h === 'left', run: halign('left') });
  cmd('alignCenter', { label: '&Center', icon: 'alignCenter', checked: () => (act().align || {}).h === 'center', run: halign('center') });
  cmd('alignRight', { label: 'Align &Right', icon: 'alignRight', checked: () => (act().align || {}).h === 'right', run: halign('right') });
  cmd('mergeCenter', { label: '&Merge and Center', tip: 'Merge and Center', icon: 'mergeCenter', checked: () => { const s = sel(); return !!LY.mergeAt(sh(), s.r, s.c); }, run: run(() => {
    if (guard('formatCells')) return;
    const s = sel();
    if (LY.mergeAt(sh(), s.r, s.c) && rgs().every((r) => sh().merges.some((m) => m.r1 === r.r1 && m.c1 === r.c1 && m.r2 === r.r2 && m.c2 === r.c2))) { O.unmerge(sh(), rgs()); return; }
    A().confirmMerge(() => O.merge(sh(), rgs(), { center: true }));
  }) });
  cmd('currencyStyle', { label: '&Currency Style', icon: 'currency', run: fmt({ nf: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' }, 'Currency Style') });
  cmd('percentStyle', { label: '&Percent Style', icon: 'percent', key: 'Ctrl+Shift+%', run: fmt({ nf: '0%' }, 'Percent Style') });
  cmd('commaStyle', { label: 'Co&mma Style', icon: 'comma', run: fmt({ nf: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }, 'Comma Style') });
  cmd('incDecimal', { label: '&Increase Decimal', icon: 'incDecimal', run: run(() => { if (guard('formatCells')) return; const s = sel(); const v = sh().val(s.r, s.c); const code = NF.withDecimals(act().nf || 'General', 1, v); O.format(sh(), rgs(), { nf: code }, 'Increase Decimal'); }) });
  cmd('decDecimal', { label: '&Decrease Decimal', icon: 'decDecimal', run: run(() => { if (guard('formatCells')) return; const s = sel(); const v = sh().val(s.r, s.c); const code = NF.withDecimals(act().nf || 'General', -1, v); O.format(sh(), rgs(), { nf: code }, 'Decrease Decimal'); }) });
  cmd('incIndent', { label: 'Increase I&ndent', icon: 'incIndent', run: run(() => { if (guard('formatCells')) return; O.format(sh(), rgs(), (st) => ({ align: { indent: Math.min(15, ((st.align && st.align.indent) || 0) + 1), h: st.align && (st.align.h === 'right' || st.align.h === 'distributed') ? st.align.h : 'left' } }), 'Increase Indent'); }) });
  cmd('decIndent', { label: 'Decrease In&dent', icon: 'decIndent', run: run(() => { if (guard('formatCells')) return; O.format(sh(), rgs(), (st) => { const n = Math.max(0, ((st.align && st.align.indent) || 0) - 1); return { align: { indent: n || undefined } }; }, 'Decrease Indent'); }) });
  cmd('wrapText', { label: '&Wrap Text', icon: 'wrapText', checked: () => !!(act().align && act().align.wrap), run: run(() => { if (guard('formatCells')) return; const on = !(act().align && act().align.wrap); O.format(sh(), rgs(), { align: { wrap: on || undefined } }, 'Wrap Text'); O.tx(wb(), 'AutoFit', () => { for (const r of A().selRows()) O.autoRow(sh(), r); }); }) });
  cmd('fillColorApply', { label: 'Fill Color', tip: 'Fill Color', icon: 'fillColor', swatch: () => A().lastFill || '#FFFF00', run: run(() => A().applyFill(A().lastFill || '#FFFF00')) });
  cmd('fontColorApply', { label: 'Font Color', tip: 'Font Color', icon: 'fontColor', swatch: () => A().lastFontColor || '#FF0000', run: run(() => A().applyFontColor(A().lastFontColor || '#FF0000')) });
  cmd('bordersApply', { label: 'Borders', tip: 'Borders', icon: 'bordersBottom', run: run(() => A().applyBorderPreset(A().lastBorder || 'bottom')) });
  cmd('formatPainter', { label: 'Format Painter', tip: 'Format Painter', icon: 'painter', checked: () => !!G().painting, run: () => A().startPaint() });
  /* number format shortcuts */
  cmd('fmtGeneral', { label: 'General', key: 'Ctrl+Shift+~', run: fmt({ nf: 'General' }, 'Format') });
  cmd('fmtCurrency', { label: 'Currency', key: 'Ctrl+Shift+$', run: fmt({ nf: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)' }, 'Format') });
  cmd('fmtExp', { label: 'Scientific', key: 'Ctrl+Shift+^', run: fmt({ nf: '0.00E+00' }, 'Format') });
  cmd('fmtDate', { label: 'Date', key: 'Ctrl+Shift+#', run: fmt({ nf: 'd-mmm-yy' }, 'Format') });
  cmd('fmtTime', { label: 'Time', key: 'Ctrl+Shift+@', run: fmt({ nf: 'h:mm AM/PM' }, 'Format') });
  cmd('fmtNumber', { label: 'Number', key: 'Ctrl+Shift+!', run: fmt({ nf: '#,##0.00' }, 'Format') });
  cmd('borderOutline', { label: 'Outline Border', key: 'Ctrl+Shift+&', run: run(() => A().applyBorderPreset('outside')) });
  cmd('borderNone', { label: 'Remove Borders', key: 'Ctrl+Shift+_', run: run(() => A().applyBorderPreset('none')) });

  /* ------------------------------------------------------------ Tools */
  cmd('spelling', { label: '&Spelling...', icon: 'spell', key: 'F7', run: run(() => D().spelling()) });
  cmd('errorChecking', { label: '&Error Checking...', run: run(() => D().errorChecking()) });
  cmd('protectSheet', { get label() { return sh().protection ? 'Un&protect Sheet...' : '&Protect Sheet...'; }, icon: 'protect', run: run(() => D().protectSheet()) });
  cmd('protectWorkbook', { get label() { return wb().protection ? 'Unprotect &Workbook...' : 'Protect &Workbook...'; }, run: run(() => D().protectWorkbook()) });
  cmd('allowEditRanges', { label: 'Allow Users to Edit Ra&nges...', run: run(() => ui.msg('Sheet protection in Ledger follows each cell\'s Locked setting (Format ▸ Cells ▸ Protection). Unlock the cells people may edit, then protect the sheet.', { icon: 'info' })) });
  cmd('goalSeek', { label: '&Goal Seek...', icon: 'goalSeek', run: run(() => D().goalSeek()) });
  cmd('scenarios', { label: 'Sc&enarios...', run: run(() => D().scenarios()) });
  cmd('tracePrecedents', { label: 'Trace &Precedents', icon: 'tracePrec', run: run(() => A().trace('prec')) });
  cmd('traceDependents', { label: 'Trace &Dependents', icon: 'traceDep', run: run(() => A().trace('dep')) });
  cmd('traceError', { label: 'Trace &Error', icon: 'traceError', run: run(() => A().trace('err')) });
  cmd('removeArrows', { label: 'Remove &All Arrows', icon: 'removeArrows', run: () => A().trace(null) });
  cmd('circleInvalid', { label: '&Circle Invalid Data', icon: 'circleInvalid', run: run(() => A().circleInvalid(true)) });
  cmd('clearCircles', { label: 'C&lear Validation Circles', icon: 'clearCircles', run: () => A().circleInvalid(false) });
  cmd('evaluateFormula', { label: 'Evaluate &Formula', icon: 'evaluate', run: run(() => D().evaluateFormula()) });
  cmd('watchWindow', { label: 'Show &Watch Window', icon: 'watch', run: run(() => D().watchWindow()) });
  cmd('addWatch', { label: 'Add &Watch', icon: 'watch', run: run(() => D().watchWindow(true)) });
  cmd('calcNow', { label: 'Calculate &Now', icon: 'calcNow', key: 'F9', run: run(() => A().calculateNow()) });
  cmd('calcSheet', { label: 'Calculate Sheet', key: 'Shift+F9', run: run(() => A().calculateNow()) });
  cmd('autoCorrectDlg', { label: '&AutoCorrect Options...', run: run(() => D().autoCorrect()) });
  cmd('options', { label: '&Options...', run: run(() => D().options()) });
  cmd('macros', { label: '&Macros...', key: 'Alt+F8', run: () => ui.msg(wb().macroEnabled ? 'This workbook contains macros. Ledger keeps them when you save as .xlsm, but cannot run them.' : 'Ledger does not run macros. Workbooks with macros (.xlsm) open and save with their macros intact.', { icon: 'info' }) });
  cmd('addIns', { label: 'Add-&Ins...', run: () => ui.msg('Analysis ToolPak functions (engineering, financial, date) are built in.', { icon: 'info' }) });

  /* ------------------------------------------------------------ Data */
  cmd('sortDlg', { label: '&Sort...', icon: 'sortDlg', run: run(() => { if (guard('sort')) return; D().sort(); }) });
  cmd('sortAsc', { label: 'Sort Ascending', tip: 'Sort Ascending', icon: 'sortAsc', run: run(() => { if (guard('sort')) return; A().quickSort(false); }) });
  cmd('sortDesc', { label: 'Sort Descending', tip: 'Sort Descending', icon: 'sortDesc', run: run(() => { if (guard('sort')) return; A().quickSort(true); }) });
  cmd('autoFilter', { label: 'Auto&Filter', icon: 'autoFilter', checked: () => !!sh().autoFilter || !!A().tableAt(), run: run(() => { if (guard('autoFilter')) return; L.filter.toggle(); }) });
  cmd('showAll', { label: '&Show All', enabled: () => L.filter && L.filter.active(), run: run(() => L.filter.showAll()) });
  cmd('advancedFilter', { label: '&Advanced Filter...', run: run(() => D().advancedFilter()) });
  cmd('dataForm', { label: 'F&orm...', icon: 'dataForm', run: run(() => D().dataForm()) });
  cmd('subtotals', { label: 'Su&btotals...', icon: 'subtotal', run: run(() => { if (guard()) return; D().subtotals(); }) });
  cmd('validation', { label: 'Va&lidation...', icon: 'validation', run: run(() => { if (guard()) return; D().validation(); }) });
  cmd('dataTable', { label: '&Table...', run: run(() => D().dataTable()) });
  cmd('textToColumns', { label: 'T&ext to Columns...', icon: 'textToCols', run: run(() => { if (guard()) return; D().textToColumns(); }) });
  cmd('consolidate', { label: 'Co&nsolidate...', run: run(() => D().consolidate()) });
  cmd('groupRows', { label: '&Group...', icon: 'group', key: 'Alt+Shift+ArrowRight', run: run(() => A().group(1)) });
  cmd('ungroupRows', { label: '&Ungroup...', icon: 'ungroup', key: 'Alt+Shift+ArrowLeft', run: run(() => A().group(-1)) });
  cmd('hideDetail', { label: '&Hide Detail', run: run(() => A().outlineDetail(false)) });
  cmd('showDetail', { label: '&Show Detail', run: run(() => A().outlineDetail(true)) });
  cmd('autoOutline', { label: '&Auto Outline', run: run(() => A().autoOutline()) });
  cmd('clearOutline', { label: '&Clear Outline', run: run(() => A().clearOutline()) });
  cmd('outlineSettings', { label: 'S&ettings...', run: run(() => D().outlineSettings()) });
  cmd('createList', { label: '&Create List...', key: 'Ctrl+L', run: run(() => D().createList()) });
  cmd('totalRow', { label: '&Total Row', checked: () => { const t = A().tableAt(); return !!(t && t.totals); }, enabled: () => !!A().tableAt(), run: run(() => A().toggleTotalRow()) });
  cmd('convertToRange', { label: 'Con&vert to Range', enabled: () => !!A().tableAt(), run: run(() => A().convertTableToRange()) });
  cmd('importText', { label: 'Import &Text File...', icon: 'open', run: async () => { const files = await L.pickFiles('.csv,.txt,.tsv,.prn,.tab'); if (files.length) A().openFile(files[0], null, { wizard: true }).catch((e) => A().error(e)); } });
  cmd('pivotTable', { label: '&PivotTable and PivotChart Report...', icon: 'pivot', run: run(() => D().pivotTable()) });
  cmd('refreshData', { label: '&Refresh Data', run: run(() => A().calculateNow()) });

  /* ------------------------------------------------------------ Window */
  cmd('freezePanes', { get label() { return sh().view.freeze ? 'Un&freeze Panes' : '&Freeze Panes'; }, icon: 'freeze', run: run(() => A().toggleFreeze()) });
  cmd('splitWindow', { get label() { return sh().view.freeze ? 'Remove &Split' : '&Split'; }, icon: 'split', run: run(() => A().toggleFreeze(true)) });
  cmd('newWindow', { label: '&New Window', run: run(() => ui.msg('Ledger shows one window per workbook. Use the Window menu to switch between open workbooks.', { icon: 'info' })) });
  cmd('arrangeWindows', { label: '&Arrange...', run: run(() => ui.msg('Open workbooks are listed at the bottom of the Window menu.', { icon: 'info' })) });

  /* ------------------------------------------------------------ Help */
  cmd('help', { label: 'Ledger &Help', icon: 'help', key: 'F1', run: () => L.panes.task.show('help') });
  cmd('keyboardHelp', { label: '&Keyboard Shortcuts', run: () => L.panes.task.show('keys') });
  cmd('functionsHelp', { label: '&Functions Reference', run: () => D().insertFunction() });
  cmd('about', { label: '&About Ledger 2003', run: () => D().about() });

  /* ------------------------------------------------------------ sheets & misc */
  cmd('nextSheet', { label: 'Next Sheet', key: 'Ctrl+PageDown', run: () => A().nextSheet(1) });
  cmd('prevSheet', { label: 'Previous Sheet', key: 'Ctrl+PageUp', run: () => A().nextSheet(-1) });
  cmd('autoSum', { label: 'AutoSum', tip: 'AutoSum', icon: 'autoSum', key: 'Alt+=', run: run(() => { if (guard()) return; A().autoSum('SUM'); }) });
  cmd('insertDate', { label: 'Insert Date', key: 'Ctrl+;', run: run(() => { if (guard()) return; const s = sel(); O.tx(wb(), 'Typing', () => O.enter(sh(), s.r, s.c, NF.text('m/d/yyyy', Math.floor(NF.jsDateToSerial(new Date(), wb().date1904))))); }) });
  cmd('insertTime', { label: 'Insert Time', key: 'Ctrl+Shift+:', run: run(() => { if (guard()) return; const s = sel(); O.tx(wb(), 'Typing', () => O.enter(sh(), s.r, s.c, NF.text('h:mm AM/PM', NF.jsDateToSerial(new Date(), wb().date1904) % 1))); }) });
  cmd('copyAbove', { label: 'Copy Formula from Above', key: "Ctrl+'", run: run(() => A().copyFromAbove(false)) });
  cmd('copyValueAbove', { label: 'Copy Value from Above', key: 'Ctrl+Shift+"', run: run(() => A().copyFromAbove(true)) });
  cmd('selectRegion', { label: 'Select Current Region', key: 'Ctrl+Shift+*', run: () => { const s = sel(); G().selectRange(O.currentRegion(sh(), s.r, s.c), { r: s.r, c: s.c }); } });
  cmd('selectArray', { label: 'Select Current Array', key: 'Ctrl+/', run: () => { const s = sel(); const a = O.arrayAt(sh(), s.r, s.c); if (a) G().selectRange(a, { r: s.r, c: s.c }); } });
  cmd('pickList', { label: 'Pick From Drop-down &List...', run: () => A().pickList() });
  cmd('dropDownList', { label: 'Pick From Drop-down &List...', key: 'Alt+ArrowDown', run: () => A().pickList() });
  cmd('repeatLast', { label: 'Repeat', key: 'F4', run: () => A().redo() });
  cmd('contextKey', { label: 'Shortcut Menu', key: 'Shift+F10', run: () => A().contextMenuAtSel() });
  cmd('newSheetChart', { label: 'Chart on new sheet', run: run(() => L.chartWizard.open({ sheet: true })) });
  void hasMulti; void wholeCols; void wholeRows; void notEditing; void F; void C; void MAXR; void MAXC;
})(typeof window !== 'undefined' ? window : globalThis);
