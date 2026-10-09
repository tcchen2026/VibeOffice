/* Lectern — command definitions (menus, toolbars, shortcuts). */
(function () {
  'use strict';
  const L = window.L;
  const ui = L.ui;
  const C = ui.cmd;
  const E = () => L.ed;
  const A = () => L.app;
  const T = L.txt;
  const editing = () => L.te.active();
  const hasSel = () => E().selected().length > 0;
  const canRotate = () => hasSel() && E().selected().every(L.model.canRotate);
  const hasSlide = () => !!E().slide();
  const normal = () => A().view === 'normal' || A().view === 'master';
  const textCtx = () => editing() || L.fmt.textShapes().length > 0;
  const fmtState = () => (textCtx() ? L.fmt.state() : null);
  const one = (type) => () => { const s = E().selected(); return s.length > 0 && s.every((x) => (typeof type === 'function' ? type(x) : x.type === type)); };
  const isPic = one('image'), isTable = () => { const s = E().selected(); return s.length === 1 && s[0].type === 'table'; }, isWA = one('wordart');
  const slideCtx = () => A().focusArea === 'slides' || A().view === 'sorter';

  /* ---------- File ---------- */
  /* the toolbar button makes a blank presentation; the menu and Ctrl+N show Blank and the templates */
  C('new', { label: '&New...', menuLabel: '&New...', icon: 'new', key: 'Ctrl+N', tip: 'New', run: (arg, ev) => { if ((ev && ev.currentTarget && ev.currentTarget.closest && ev.currentTarget.closest('.toolbar')) || !window.VO) A().newPresentation(); else VO.newDialog('lectern', (id) => (id ? A().newFromTemplate(id) : A().newPresentation())); } });
  C('open', { label: '&Open...', icon: 'open', key: 'Ctrl+O', tip: 'Open', run: () => A().openDialog() });
  C('close', { label: '&Close', run: () => A().closePresentation() });
  C('save', { label: '&Save', icon: 'save', key: 'Ctrl+S', tip: 'Save', run: () => A().save() });
  C('saveAs', { label: 'Save &As...', run: () => A().saveAs() });
  C('saveWeb', { label: 'Save as Web Pa&ge...', icon: 'webPage', run: () => A().exportAs('html') });
  C('pageSetup', { label: 'Page Set&up...', run: () => L.dlg.pageSetup() });
  C('printPreview', { label: 'Print Pre&view', icon: 'preview', tip: 'Print Preview', run: () => A().setView('preview') });
  C('print', { label: '&Print...', icon: 'print', key: 'Ctrl+P', tip: 'Print', run: () => L.dlg.print() });
  C('properties', { label: 'Propert&ies', icon: 'properties', run: () => L.dlg.properties() });
  C('exitApp', { label: 'E&xit', icon: 'exit', run: () => A().closePresentation() });

  /* ---------- Edit ---------- */
  C('undo', { label: '&Undo', icon: 'undo', key: 'Ctrl+Z', get menuLabel() { return L.hist.undo.length ? '&Undo ' + L.hist.undo[L.hist.undo.length - 1].label : "&Can't Undo"; }, enabled: () => L.hist.undo.length > 0, run: () => A().undo() });
  C('redo', { label: '&Redo', icon: 'redo', key: 'Ctrl+Y', get menuLabel() { return L.hist.redo.length ? '&Redo ' + L.hist.redo[L.hist.redo.length - 1].label : "&Can't Redo"; }, enabled: () => L.hist.redo.length > 0, run: () => A().redo() });
  C('cut', { label: 'Cu&t', icon: 'cut', key: 'Ctrl+X', enabled: () => editing() || hasSel() || slideCtx(), run: () => A().cmdClipboard('cut') });
  C('copy', { label: '&Copy', icon: 'copy', key: 'Ctrl+C', enabled: () => editing() || hasSel() || slideCtx(), run: () => A().cmdClipboard('copy') });
  C('paste', { label: '&Paste', icon: 'paste', key: 'Ctrl+V', enabled: () => !!A().clip, run: () => A().paste() });
  C('pasteSpecial', { label: 'Paste &Special...', enabled: () => !!A().clip, run: () => A().pasteSpecial() });
  C('officeClipboard', { label: 'Office Clip&board...', icon: 'clipboard', run: () => L.panes.task.show('clipboard') });
  C('clear', { label: 'Cle&ar', key: 'Del', enabled: () => hasSel() || slideCtx(), run: () => (slideCtx() ? A().deleteSlides() : A().deleteSel()) });
  C('selectAll', { label: 'Select A&ll', icon: 'selectAll', key: 'Ctrl+A', run: () => A().selectAll() });
  C('duplicate', { label: 'Dupl&icate', icon: 'duplicate', key: 'Ctrl+D', enabled: () => hasSel() || slideCtx(), run: () => (slideCtx() || !hasSel() ? A().duplicateSlides() : A().duplicateSel()) });
  C('deleteSlide', { label: '&Delete Slide', icon: 'delete', enabled: () => L.pres.slides.length > 0 && normal() || A().view === 'sorter', run: () => A().deleteSlides() });
  C('find', { label: '&Find...', icon: 'find', key: 'Ctrl+F', run: () => L.dlg.find(false) });
  C('replace', { label: 'R&eplace...', icon: 'replace', key: 'Ctrl+H', run: () => L.dlg.find(true) });

  /* ---------- View ---------- */
  C('viewNormal', { label: '&Normal', icon: 'normalView', radio: true, checked: () => A().view === 'normal', tip: 'Normal View', run: () => A().setView('normal') });
  C('viewSorter', { label: 'Sli&de Sorter', icon: 'sorterView', radio: true, checked: () => A().view === 'sorter', tip: 'Slide Sorter View', run: () => A().setView('sorter') });
  C('viewNotes', { label: 'Notes &Page', icon: 'notesView', radio: true, checked: () => A().view === 'notes', run: () => A().setView('notes') });
  C('viewMaster', { label: '&Slide Master', icon: 'master', checked: () => A().view === 'master', run: () => A().setView('master') });
  C('closeMaster', { label: '&Close Master View', tbLabel: '&Close Master View', run: () => A().setView('normal') });
  C('insertTitleMaster', { label: 'Insert New &Title Master', icon: 'newSlide', enabled: () => A().view === 'master' && !E().design().titleDeco, run: () => A().insertTitleMaster() });
  C('deleteTitleMaster', { label: '&Delete Title Master', icon: 'delete', enabled: () => A().view === 'master' && E().masterKind === 'title', run: () => A().deleteTitleMaster() });
  C('viewColor', { label: '&Color', radio: true, checked: () => !A().opts.gray && !A().opts.bw, run: () => A().setColorMode('color') });
  C('viewGray', { label: '&Grayscale', icon: 'grayscale', radio: true, checked: () => !!A().opts.gray, tip: 'Color/Grayscale', run: () => A().setColorMode(A().opts.gray ? 'color' : 'gray') });
  C('viewBW', { label: '&Pure Black and White', radio: true, checked: () => !!A().opts.bw, run: () => A().setColorMode('bw') });
  C('taskPane', { label: 'Tas&k Pane', icon: 'taskpane', key: 'Ctrl+F1', checked: () => A().taskOpen(), run: () => A().toggleTask() });
  C('slidesPane', { label: 'Slides and Outline Pane', checked: () => A().leftOpen(), run: () => A().toggleLeft() });
  for (const [id, label] of [['standard', '&Standard'], ['formatting', '&Formatting'], ['drawing', '&Drawing'], ['outlining', '&Outlining'], ['picture', '&Picture'], ['tables', '&Tables and Borders'], ['wordart', '&WordArt']]) {
    C('tb_' + id, { label, checked: () => A().tbShown(id), run: () => A().toggleToolbar(id) });
  }
  C('ruler', { label: '&Ruler', icon: 'ruler', checked: () => !!A().opts.ruler, run: () => { A().opts.ruler = !A().opts.ruler; A().saveOpts(); E().layout(); } });
  C('showGrid', { label: '&Grid and Guides...', icon: 'grid', tip: 'Show/Hide Grid', checked: () => E().grid.show, run: () => { E().grid.show = !E().grid.show; E().layout(); A().saveOpts(); } });
  C('gridGuides', { label: 'Gr&id and Guides...', icon: 'grid', run: () => L.dlg.grid() });
  C('headerFooter', { label: '&Header and Footer...', run: () => L.dlg.headerFooter() });
  C('zoomDlg', { label: '&Zoom...', icon: 'zoom', run: () => L.dlg.zoom() });
  C('zoomFit', { label: '&Fit', run: () => E().setZoom('fit') });

  /* ---------- Insert ---------- */
  C('newSlide', { label: 'New Slide...', menuLabel: '&New Slide...', tbLabel: 'New Slide', icon: 'newSlide', key: 'Ctrl+M', tip: 'New Slide', enabled: () => A().view !== 'master', run: () => A().newSlideCmd() });
  C('duplicateSlide', { label: '&Duplicate Slide', icon: 'duplicate', enabled: () => L.pres.slides.length > 0 && A().view !== 'master', run: () => A().duplicateSlides() });
  C('insertSlideNumber', { label: 'Slide N&umber', icon: 'slideNumber', run: () => A().insertField('slidenum') });
  C('insertDateTime', { label: 'Da&te and Time...', icon: 'date', run: () => L.dlg.dateTime() });
  C('insertSymbol', { label: '&Symbol...', icon: 'symbol', run: () => { A().ensureTextTarget(); L.dlg.symbol(); } });
  C('slidesFromFiles', { label: 'Slides from &Files...', icon: 'insertSlides', run: () => L.dlg.slideFinder() });
  C('slidesFromOutline', { label: 'Slides from Out&line...', icon: 'outlineTxt', run: () => A().slidesFromOutline() });
  C('insertClipArt', { label: '&Clip Art...', icon: 'clipart', tip: 'Insert Clip Art', run: () => L.panes.task.show('clipart') });
  C('insertPicture', { label: '&From File...', icon: 'picture', tip: 'Insert Picture', run: (arg) => A().insertPictureDialog(arg && arg.ph) });
  C('photoAlbum', { label: '&New Photo Album...', icon: 'photoAlbum', run: () => L.dlg.photoAlbum() });
  C('insertWordArt', { label: '&WordArt...', icon: 'wordart', tip: 'Insert WordArt', run: () => A().insertWordArtDialog() });
  C('insertDiagram', { label: 'Dia&gram...', icon: 'diagram', tip: 'Insert Diagram or Organization Chart', run: (arg) => L.dlg.diagramGallery((k) => A().insertDiagram(k, arg && arg.ph)) });
  C('textBox', { label: 'Te&xt Box', icon: 'textbox', tip: 'Text Box', checked: () => !!(E().tool && E().tool.kind === 'text' && !E().tool.vert), run: () => E().setTool(E().tool && E().tool.kind === 'text' ? null : { kind: 'text' }) });
  C('vTextBox', { label: '&Vertical Text Box', icon: 'vtextbox', checked: () => !!(E().tool && E().tool.kind === 'text' && E().tool.vert), run: () => E().setTool({ kind: 'text', vert: true }) });
  C('insertChart', { label: 'C&hart...', icon: 'chart', tip: 'Insert Chart', run: (arg) => A().insertChart(arg && arg.ph) });
  C('insertTable', { label: 'Ta&ble...', icon: 'table', tip: 'Insert Table', run: (arg) => L.dlg.insertTable((r, c) => A().insertTable(r, c, arg && arg.ph)) });
  C('insertMedia', { label: 'Mo&vies and Sounds', icon: 'movie', run: () => ui.msg('Movies and sounds are not supported in this edition. Pictures, charts, tables and diagrams are available from the Insert menu.') });
  C('hyperlink', { label: 'Hyperl&ink...', icon: 'hyperlink', key: 'Ctrl+K', tip: 'Insert Hyperlink', enabled: () => editing() || hasSel(), run: () => L.dlg.hyperlink() });

  /* ---------- Format ---------- */
  C('fontDlg', { label: '&Font...', icon: 'font', key: 'Ctrl+T', enabled: textCtx, run: () => L.dlg.font() });
  C('bulletsDlg', { label: '&Bullets and Numbering...', enabled: textCtx, run: () => L.dlg.bullets() });
  C('lineSpacing', { label: 'Line &Spacing...', icon: 'lineSpacing', enabled: textCtx, run: () => L.dlg.lineSpacing() });
  C('changeCase', { label: 'C&hange Case...', icon: 'changeCase', key: 'Shift+F3', enabled: textCtx, run: () => L.dlg.changeCase() });
  C('replaceFonts', { label: 'Re&place Fonts...', run: () => L.dlg.replaceFonts() });
  C('slideDesign', { label: 'Slide D&esign...', tbLabel: 'Design', icon: 'design', tip: 'Slide Design', run: () => L.panes.task.show('design') });
  C('slideLayout', { label: 'Slide &Layout...', icon: 'layout', run: () => L.panes.task.show('layout') });
  C('background', { label: 'Bac&kground...', icon: 'background', enabled: hasSlide, run: () => L.dlg.background() });
  C('formatObject', { get label() { const s = E().selected()[0]; return s ? (s.type === 'image' ? 'P&icture...' : s.type === 'wordart' ? 'Word&Art...' : s.ph ? 'Pla&ceholder...' : s.type === 'text' ? 'Text B&ox...' : s.type === 'table' ? 'Ta&ble...' : 'AutoSha&pe...') : 'AutoSha&pe...'; }, icon: 'formatPicture', enabled: hasSel, run: (arg) => L.dlg.formatShape(arg) });
  C('alignLeft', { label: 'Align &Left', icon: 'alignLeft', key: 'Ctrl+L', radio: true, enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.algn === 'l'; }, run: () => A().setAlign('l') });
  C('alignCenter', { label: '&Center', icon: 'alignCenter', key: 'Ctrl+E', radio: true, enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.algn === 'ctr'; }, run: () => A().setAlign('ctr') });
  C('alignRight', { label: 'Align &Right', icon: 'alignRight', key: 'Ctrl+R', radio: true, enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.algn === 'r'; }, run: () => A().setAlign('r') });
  C('justify', { label: '&Justify', icon: 'justify', key: 'Ctrl+J', radio: true, enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.algn === 'just'; }, run: () => A().setAlign('just') });
  C('bold', { label: 'Bold', icon: 'bold', key: 'Ctrl+B', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.b; }, run: () => A().toggleRun('b') });
  C('italic', { label: 'Italic', icon: 'italic', key: 'Ctrl+I', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.i; }, run: () => A().toggleRun('i') });
  C('underline', { label: 'Underline', icon: 'underline', key: 'Ctrl+U', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.u; }, run: () => A().toggleRun('u') });
  C('shadowText', { label: 'Shadow', icon: 'shadowText', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.shd; }, run: () => A().toggleRun('shd') });
  C('superscript', { label: 'Superscript', key: 'Ctrl+Shift+=', enabled: textCtx, run: () => A().toggleBaseline(30000) });
  C('subscript', { label: 'Subscript', key: 'Ctrl+=', enabled: textCtx, run: () => A().toggleBaseline(-25000) });
  C('resetChar', { label: 'Reset Character Formatting', key: 'Ctrl+Space', enabled: textCtx, run: () => L.fmt.run((r) => ({ t: r.t, fld: r.fld, link: r.link }), 'Reset Formatting') });
  C('growFont', { label: 'Increase Font Size', icon: 'growFont', key: 'Ctrl+Shift+>', enabled: textCtx, run: () => A().stepFont(1) });
  C('shrinkFont', { label: 'Decrease Font Size', icon: 'shrinkFont', key: 'Ctrl+Shift+<', enabled: textCtx, run: () => A().stepFont(-1) });
  C('numbering', { label: 'Numbering', icon: 'numbering', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.bu === 'num'; }, run: () => A().toggleBullets('num') });
  C('bullets', { label: 'Bullets', icon: 'bullets', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.bu === 'char'; }, run: () => A().toggleBullets('char') });
  C('promote', { label: 'Decrease Indent', icon: 'decIndent', key: 'Alt+Shift+Left', tip: 'Decrease Indent (Promote)', enabled: textCtx, run: () => A().indent(-1) });
  C('demote', { label: 'Increase Indent', icon: 'incIndent', key: 'Alt+Shift+Right', tip: 'Increase Indent (Demote)', enabled: textCtx, run: () => A().indent(1) });
  C('moveParaUp', { label: 'Move Up', icon: 'moveUp', key: 'Alt+Shift+Up', enabled: textCtx, run: () => A().moveParas(-1) });
  C('moveParaDown', { label: 'Move Down', icon: 'moveDown', key: 'Alt+Shift+Down', enabled: textCtx, run: () => A().moveParas(1) });
  C('fontColor', { label: 'Font Color', icon: 'fontColor', tip: 'Font Color', enabled: textCtx, swatch: () => A().lastFontColorHex(), run: () => A().applyFontColor(A().lastFontColor) });
  C('fillColor', { label: 'Fill Color', icon: 'fillColor', tip: 'Fill Color', enabled: hasSel, swatch: () => A().lastFillHex(), run: () => A().applyFill(A().lastFill) });
  C('lineColor', { label: 'Line Color', icon: 'lineColor', tip: 'Line Color', enabled: hasSel, swatch: () => A().lastLineHex(), run: () => A().applyLineColor(A().lastLineColor) });
  C('painter', { label: 'Format Painter', icon: 'painter', tip: 'Format Painter', enabled: () => hasSel() || (E().tool && E().tool.kind === 'painter'), checked: () => !!(E().tool && E().tool.kind === 'painter'), run: (arg, ev) => A().formatPainter(ev && ev.detail === 2) });
  C('showFormatting', { label: 'Show Formatting', icon: 'showFormatting', checked: () => !A().opts.outlinePlain, run: () => { A().opts.outlinePlain = !A().opts.outlinePlain; L.panes.left.render(); } });
  C('expandAll', { label: 'Expand All', icon: 'expandAll', run: () => L.panes.left.setTab('outline') });
  C('summarySlide', { label: 'Summary Slide', icon: 'summary', tip: 'Summary Slide', run: () => A().summarySlide() });

  /* ---------- Tools ---------- */
  C('spelling', { label: '&Spelling...', icon: 'spell', key: 'F7', tip: 'Spelling', run: () => A().spelling() });
  C('autocorrectDlg', { label: '&AutoCorrect Options...', run: () => L.dlg.autocorrect() });
  C('optionsDlg', { label: '&Options...', icon: 'options', run: () => L.dlg.options() });

  /* ---------- Slide Show ---------- */
  C('showFromStart', { label: '&View Show', icon: 'slideshow', key: 'F5', tip: 'Slide Show (from start)', enabled: () => L.pres.slides.length > 0, run: () => L.show.start({ from: 0 }) });
  C('showFromCurrent', { label: 'Slide Show from current slide', icon: 'slideshow', key: 'Shift+F5', tip: 'Slide Show from current slide (Shift+F5)', enabled: () => L.pres.slides.length > 0, run: () => L.show.start({ from: A().view === 'master' ? 0 : E().idx }) });
  C('setupShow', { label: '&Set Up Show...', icon: 'setupShow', run: () => L.dlg.setupShow() });
  C('rehearse', { label: '&Rehearse Timings', icon: 'rehearse', tip: 'Rehearse Timings', enabled: () => L.pres.slides.length > 0, run: () => L.show.start({ from: 0, rehearse: true }) });
  C('actionSettings', { label: 'Act&ion Settings...', icon: 'actionSettings', enabled: hasSel, run: () => L.dlg.actionSettings() });
  C('animSchemes', { label: 'Animatio&n Schemes...', run: () => L.panes.task.show('animSchemes') });
  C('customAnim', { label: 'Custo&m Animation...', icon: 'animation', run: () => L.panes.task.show('customAnim') });
  C('transitionPane', { label: 'Slide &Transition...', icon: 'transition', tbLabel: 'Transition', tip: 'Slide Transition', run: () => L.panes.task.show('transition') });
  C('hideSlide', { label: '&Hide Slide', icon: 'hideSlide', tip: 'Hide Slide', checked: () => { const s = E().slide(); return !!s && !!s.hidden; }, enabled: () => L.pres.slides.length > 0 && A().view !== 'master', run: () => A().toggleHidden() });
  C('speakerNotes', { label: 'Speaker Notes', icon: 'notesView', run: () => A().speakerNotesDlg() });

  /* ---------- Help ---------- */
  C('help', { label: 'Lectern &Help', icon: 'help', key: 'F1', tip: 'Lectern Help', run: () => L.panes.task.show('help') });
  C('about', { label: '&About Lectern', icon: 'about', run: () => L.dlg.about() });
  C('gettingStarted', { label: '&Getting Started', run: () => L.panes.task.show('getting-started') });

  /* ---------- Drawing ---------- */
  C('select', { label: 'Select Objects', icon: 'select', checked: () => !E().tool, run: () => { E().cancelFreeform(false); E().setTool(null); } });
  C('drawLine', { label: 'Line', icon: 'line', checked: () => !!(E().tool && E().tool.kind === 'line' && !E().tool.arrows && E().tool.geom === 'line'), run: (a, ev) => E().setTool({ kind: 'line', geom: 'line', sticky: ev && ev.detail === 2 }) });
  C('drawArrow', { label: 'Arrow', icon: 'arrowTool', checked: () => !!(E().tool && E().tool.kind === 'line' && E().tool.arrows === 'end'), run: (a, ev) => E().setTool({ kind: 'line', geom: 'straightConnector1', arrows: 'end', sticky: ev && ev.detail === 2 }) });
  C('drawDblArrow', { label: 'Double Arrow', icon: 'doubleArrow', run: () => E().setTool({ kind: 'line', geom: 'straightConnector1', arrows: 'both' }) });
  C('drawRect', { label: 'Rectangle', icon: 'rect', checked: () => !!(E().tool && E().tool.geom === 'rect' && E().tool.kind === 'shape'), run: (a, ev) => E().setTool({ kind: 'shape', geom: 'rect', sticky: ev && ev.detail === 2 }) });
  C('drawOval', { label: 'Oval', icon: 'oval', checked: () => !!(E().tool && E().tool.geom === 'ellipse'), run: (a, ev) => E().setTool({ kind: 'shape', geom: 'ellipse', sticky: ev && ev.detail === 2 }) });
  C('drawCurve', { label: 'Curve', icon: 'curve', run: () => E().setTool({ kind: 'curve' }) });
  C('drawFreeform', { label: 'Freeform', icon: 'freeform', run: () => E().setTool({ kind: 'freeform' }) });
  C('drawScribble', { label: 'Scribble', icon: 'scribble', run: () => E().setTool({ kind: 'scribble' }) });
  C('drawElbow', { label: 'Elbow Connector', icon: 'elbow', run: () => E().setTool({ kind: 'line', geom: 'bentConnector3' }) });
  C('drawCurvedConn', { label: 'Curved Connector', icon: 'curve', run: () => E().setTool({ kind: 'line', geom: 'curvedConnector3' }) });
  C('group', { label: '&Group', icon: 'group', key: 'Ctrl+G', enabled: () => E().selected().length > 1, run: () => A().group() });
  C('ungroup', { label: '&Ungroup', icon: 'ungroup', key: 'Ctrl+Shift+G', enabled: () => E().selected().some((s) => s.type === 'group'), run: () => A().ungroup() });
  C('regroup', { label: 'R&egroup', enabled: () => !!A().lastUngroup, run: () => A().regroup() });
  C('bringFront', { label: 'Bring to Fro&nt', icon: 'bringFront', enabled: hasSel, run: () => A().order('front') });
  C('sendBack', { label: 'Send to Bac&k', icon: 'sendBack', enabled: hasSel, run: () => A().order('back') });
  C('bringForward', { label: 'Bring &Forward', icon: 'bringForward', enabled: hasSel, run: () => A().order('forward') });
  C('sendBackward', { label: 'Send &Backward', icon: 'sendBackward', enabled: hasSel, run: () => A().order('backward') });
  for (const [k, l, ic] of [['L', 'Align &Left', 'alignL'], ['C', 'Align &Center', 'alignC'], ['R', 'Align &Right', 'alignR'], ['T', 'Align &Top', 'alignT'], ['M', 'Align &Middle', 'alignM'], ['B', 'Align &Bottom', 'alignB']]) C('align' + k, { label: l, icon: ic, enabled: hasSel, run: () => A().align(k) });
  C('distH', { label: 'Distribute &Horizontally', icon: 'distH', enabled: () => E().selected().length > 1 || A().opts.alignToSlide, run: () => A().distribute('h') });
  C('distV', { label: 'Distribute &Vertically', icon: 'distV', enabled: () => E().selected().length > 1 || A().opts.alignToSlide, run: () => A().distribute('v') });
  C('alignToSlide', { label: 'Relative to &Slide', checked: () => !!A().opts.alignToSlide, run: () => { A().opts.alignToSlide = !A().opts.alignToSlide; } });
  C('freeRotate', { label: 'Free Ro&tate', icon: 'freeRotate', enabled: canRotate, run: () => ui.toast('Drag the green rotation handle to rotate. Hold Shift for 15° steps.') });
  C('rotateLeft', { label: 'Rotate &Left 90°', icon: 'rotateL', tip: 'Rotate Left 90°', enabled: canRotate, run: () => A().rotate(-90) });
  C('rotateRight', { label: 'Rotate &Right 90°', icon: 'rotateR', enabled: canRotate, run: () => A().rotate(90) });
  C('flipH', { label: 'Flip &Horizontal', icon: 'flipH', enabled: canRotate, run: () => A().flip('h') });
  C('flipV', { label: 'Flip &Vertical', icon: 'flipV', enabled: canRotate, run: () => A().flip('v') });
  C('nudgeUp', { label: '&Up', enabled: hasSel, run: () => A().nudge(0, -1) });
  C('nudgeDown', { label: '&Down', enabled: hasSel, run: () => A().nudge(0, 1) });
  C('nudgeLeft', { label: '&Left', enabled: hasSel, run: () => A().nudge(-1, 0) });
  C('nudgeRight', { label: '&Right', enabled: hasSel, run: () => A().nudge(1, 0) });
  C('setDefaults', { label: 'Set AutoShape &Defaults', enabled: () => { const s = E().primary(); return !!s && (s.type === 'shape' || s.type === 'line'); }, run: () => A().setDefaults() });
  C('editPoints', { label: 'Edit Poi&nts', enabled: () => false, run: () => {} });

  /* ---------- Picture ---------- */
  C('picMoreContrast', { label: 'More Contrast', icon: 'moreContrast', enabled: isPic, run: () => A().picAdjust('contrast', 0.15) });
  C('picLessContrast', { label: 'Less Contrast', icon: 'lessContrast', enabled: isPic, run: () => A().picAdjust('contrast', -0.15) });
  C('picMoreBright', { label: 'More Brightness', icon: 'moreBright', enabled: isPic, run: () => A().picAdjust('bright', 0.1) });
  C('picLessBright', { label: 'Less Brightness', icon: 'lessBright', enabled: isPic, run: () => A().picAdjust('bright', -0.1) });
  C('picCrop', { label: 'Crop', icon: 'crop', enabled: isPic, checked: () => !!E().cropMode, run: () => { E().cropMode = !E().cropMode; E().drawOverlay(); } });
  C('picReset', { label: 'Reset Picture', icon: 'resetPicture', enabled: isPic, run: () => A().picReset() });
  C('picTransparent', { label: 'Set Transparent Color', icon: 'transparentColor', enabled: isPic, run: () => A().picTransparentTool() });

  /* ---------- Table ---------- */
  C('rowAbove', { label: 'Insert Rows &Above', icon: 'rowAbove', enabled: isTable, run: () => A().tableCmd('rowAbove') });
  C('rowBelow', { label: 'Insert Rows &Below', icon: 'rowBelow', enabled: isTable, run: () => A().tableCmd('rowBelow') });
  C('colLeft', { label: 'Insert Columns to the &Left', icon: 'colLeft', enabled: isTable, run: () => A().tableCmd('colLeft') });
  C('colRight', { label: 'Insert Columns to the &Right', icon: 'colRight', enabled: isTable, run: () => A().tableCmd('colRight') });
  C('deleteRows', { label: 'Delete &Rows', icon: 'deleteRow', enabled: isTable, run: () => A().tableCmd('deleteRows') });
  C('deleteCols', { label: 'Delete &Columns', icon: 'deleteCol', enabled: isTable, run: () => A().tableCmd('deleteCols') });
  C('mergeCells', { label: '&Merge Cells', icon: 'mergeCells', enabled: () => isTable() && !!E().cellSel, run: () => A().tableCmd('merge') });
  C('splitCell', { label: 'S&plit Cell', icon: 'splitCells', enabled: isTable, run: () => A().tableCmd('split') });
  C('cellTop', { label: 'Align Top', icon: 'alignTop', enabled: isTable, run: () => A().tableCmd('anchor', 't') });
  C('cellMiddle', { label: 'Center Vertically', icon: 'alignMiddle', enabled: isTable, run: () => A().tableCmd('anchor', 'ctr') });
  C('cellBottom', { label: 'Align Bottom', icon: 'alignBottom', enabled: isTable, run: () => A().tableCmd('anchor', 'b') });
  C('distRows', { label: 'Distribute Rows Evenly', icon: 'distRows', enabled: isTable, run: () => A().tableCmd('distRows') });
  C('distCols', { label: 'Distribute Columns Evenly', icon: 'distCols', enabled: isTable, run: () => A().tableCmd('distCols') });
  C('selectTable', { label: 'Select &Table', enabled: isTable, run: () => A().tableCmd('selectAll') });
  C('tableBordersFill', { label: 'Borders and &Fill...', enabled: isTable, run: () => A().tableBordersDlg() });

  /* ---------- WordArt ---------- */
  C('wordartEdit', { label: 'Edit Te&xt...', icon: 'editText', tbLabel: 'Edit Text...', enabled: () => E().selected().length === 1 && isWA(), run: () => A().editWordArt() });
  C('wordartGallery', { label: 'WordArt Gallery', icon: 'waGallery', enabled: isWA, run: () => A().wordartRestyle() });
  C('waSameHeight', { label: 'WordArt Same Letter Heights', icon: 'sameHeight', enabled: isWA, checked: () => { const s = E().primary(); return !!(s && s.wa && s.wa.cap); }, run: () => A().waToggle('cap') });
  C('waVertical', { label: 'WordArt Vertical Text', icon: 'vertText', enabled: isWA, checked: () => { const s = E().primary(); return !!(s && s.wa && s.wa.vert); }, run: () => A().waToggle('vert') });
})();
