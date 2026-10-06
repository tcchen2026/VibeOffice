/* Quire — application controller: window chrome, views, document lifecycle, keyboard, wiring. */
(function () {
  'use strict';
  const L = window.L, D = L.D, LY = L.layout, E = L.ed, O = L.O;
  const { h } = L;
  const ui = L.ui;

  const DEFAULT_OPTS = {
    statusBar: true, ruler: true, vruler: true, taskOpen: true, startupPane: true, docMap: false, thumbs: false,
    toolbars: { standard: true, formatting: true, drawing: true, tables: false, reviewing: 'auto', picture: 'auto', outlining: 'auto', headerFooter: 'auto', mailMerge: false, wordart: 'auto' },
    showMarks: false, marks: { tabs: false, spaces: false, para: false, hidden: false, all: false }, textBoundaries: true, fieldShading: 'selected', bookmarks: false,
    smartQuotes: true, autocorrect: true, autoFormat: true, capSentence: true, autoBullets: true, autoNumbers: true, autoBorders: true, autoHeadings: false, autoUrl: true, ordinals: true, fractions: true, dashes: true,
    spell: true, unit: 'in', recent: [], zoom: 1, view: 'print', balloons: true, pasteOptions: true, dragDrop: true, smartCut: true, typingReplaces: true,
    userName: 'Quire User', userInitials: 'QU', markupMode: 'markup',
  };
  const A = (L.app = Object.assign(L.app || {}, {
    view: 'print', fileName: 'Document1', fileType: 'docx', saved: false, untitled: 1, opts: L.clone(DEFAULT_OPTS),
    lastFontColor: 'FF0000', lastHighlight: 'yellow', lastBorder: 'outside', lastShading: null,
  }));
  const doc = () => D.doc;

  /* ================= options ================= */
  A.loadOpts = () => { A.opts = L.deepMerge(L.clone(DEFAULT_OPTS), L.store.get('opts', {}) || {}); D.unit = A.opts.unit || 'in'; };
  A.saveOpts = () => { L.store.set('opts', A.opts); };
  A.applyOpts = () => {
    D.unit = A.opts.unit || 'in';
    L.store.set('userName', A.opts.userName || 'Quire User');
    L.store.set('userInitials', A.opts.userInitials || '');
    L.$('#statusbar').hidden = !A.opts.statusBar;
    LY.opts.textBoundaries = A.opts.textBoundaries !== false;
    LY.opts.showBookmarks = !!A.opts.bookmarks;
    LY.root.spellcheck = false; /* Quire draws its own spelling and grammar underlines */
    if (L.spell) L.spell.refreshSoon(0);
    LY.root.classList.toggle('fshade', A.opts.fieldShading === 'always');
    A.updateRulers();
    A.updateToolbars();
  };

  /* ================= menus ================= */
  function buildMenus() {
    const recentItems = () => (A.opts.recent || []).slice(0, 4).map((r, i) => ({ label: `&${i + 1} ${r.replace(/&/g, '&&')}`, run: () => ui.msg(`To reopen ${r}, choose File ▸ Open and pick it again. Browsers don't let pages reopen files by themselves.`, { icon: 'info' }) }));
    const menus = [
      { label: '&File', items: () => ['new', 'open', 'close', '-', 'save', 'saveAs', 'saveWeb', 'fileSearch', '-', 'webPreview', '-', 'pageSetup', 'printPreview', 'print', '-', { label: 'Sen&d To', sub: ['sendMail', 'exportPDF'] }, 'properties', ...(A.opts.recent && A.opts.recent.length ? ['-', ...recentItems()] : []), '-', 'exitApp'] },
      {
        label: '&Edit', items: () => ['undo', 'redo', '-', 'cut', 'copy', 'officeClipboard', 'paste', 'pasteSpecial', 'pasteHyperlink', '-',
          { label: 'Cle&ar', sub: ['clearFormats', 'clearContents'] }, 'selectAll', '-', 'find', 'replace', 'goTo', '-', 'links', 'editObject'],
      },
      {
        label: '&View', items: () => ['viewNormal', 'viewWeb', 'viewPrint', 'viewReading', 'viewOutline', '-', 'taskPane',
          { label: '&Toolbars', sub: ['tb_standard', 'tb_formatting', 'tb_drawing', 'tb_tables', 'tb_reviewing', 'tb_picture', 'tb_outlining', 'tb_headerFooter', 'tb_mailMerge', 'tb_wordart', '-', 'customize'] },
          'ruler', '-', 'docMap', 'thumbnails', '-', 'headerFooter', 'footnotes', 'markup', '-', 'fullScreen', 'zoomDlg'],
      },
      {
        label: '&Insert', items: () => ['insertBreak', 'pageNumbers', 'dateTime', { label: '&AutoText', sub: () => L.autocorrect ? L.autocorrect.autoTextMenu() : [] }, 'insertField', 'insertSymbol', 'insertComment', '-',
          { label: 'Refere&nce', sub: ['insertFootnote', 'insertCaption', 'crossReference', 'indexTables'] }, '-',
          { label: '&Picture', icon: 'picture', sub: ['insertClipArt', 'insertPicture', '-', 'newDrawing', { label: '&AutoShapes', icon: 'autoshapes', run: () => A.openAutoShapes() }, 'insertWordArt', 'insertChart'] },
          'insertDiagram', 'textBox', 'insertFile', 'insertObject', 'bookmark', 'hyperlink'],
      },
      {
        label: 'F&ormat', items: () => ['fontDlg', 'paragraphDlg', 'bulletsDlg', 'bordersDlg', 'columnsDlg', 'tabsDlg', 'dropCap', 'textDirection', 'changeCase', '-',
          { label: '&Background', sub: () => backgroundMenu() }, 'themeDlg', 'autoFormat', '-', 'stylesPane', 'revealFormatting', '-', 'formatObject'],
      },
      {
        label: '&Tools', items: () => ['spelling', 'research', { label: '&Language', sub: ['setLanguage', 'thesaurus', 'hyphenation'] }, 'wordCount', 'autoSummarize', '-',
          'trackChanges', 'compareDocs', 'protectDoc', '-', { label: 'L&etters and Mailings', sub: ['mailMerge', 'tb_mailMerge', 'envelopes', 'letterWizard'] }, '-',
          { label: '&Macro', sub: ['macros'] }, 'templates', 'autocorrectDlg', 'customize', 'optionsDlg'],
      },
      {
        label: 'T&able', items: () => ['drawTable', { label: '&Insert', sub: ['insertTable', '-', 'colLeft', 'colRight', 'rowAbove', 'rowBelow', 'insertCells'] },
          { label: '&Delete', sub: ['deleteTable', 'deleteCols', 'deleteRows', 'deleteCells'] }, { label: 'Sele&ct', sub: ['selectTable', 'selectCol', 'selectRow', 'selectCell'] },
          'mergeCells', 'splitCells', 'splitTable', '-', 'tableAutoFormat', { label: '&AutoFit', sub: ['autofitContents', 'autofitWindow', 'fixedWidth', '-', 'distRows', 'distCols'] },
          'headingRows', { label: 'Con&vert', sub: ['textToTable', 'tableToText'] }, 'sortDlg', 'formulaDlg', 'gridlines', '-', 'tableProps'],
      },
      { label: '&Window', items: () => ['newWindow', 'arrangeAll', 'splitWindow', '-', ...A.windowItems()] },
      { label: '&Help', items: ['help', 'showAssistant', '-', 'gettingStarted', 'keyboardHelp', '-', 'about'] },
    ];
    const row = L.$('#menurow');
    A.menuBar = ui.menuBar(row, menus);
    const q = h('input', { type: 'text', id: 'helpq', placeholder: 'Type a question for help', 'aria-label': 'Type a question for help' });
    q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { if (L.panes) { L.panes.helpQuery = q.value; L.panes.task.show('help'); } q.value = ''; E.refocus(); } e.stopPropagation(); });
    row.appendChild(h('div', { class: 'helpbox' }, q, h('span', { class: 'dd-arrow' })));
  }
  function backgroundMenu() {
    return [
      { label: '&No Fill', run: () => A.setBackground(null) },
      { custom: (close) => { const g = h('div', { class: 'color-panel' }); const grid = h('div', { class: 'cp-grid' }); for (const c of L.color.STANDARD) grid.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${c}`, 'aria-label': ui.colorName(c), 'data-tip': ui.colorName(c), onclick: () => { close(); A.setBackground(c.replace('#', '')); } })); g.appendChild(grid); return g; } },
      { label: '&More Colors...', run: () => ui.moreColors(doc().bg ? '#' + doc().bg : null, (c) => { if (c) A.setBackground(c.replace('#', '')); }) },
      'printedWatermark',
    ];
  }
  A.setBackground = (c) => E.edit('Background', () => { D.touchKey(doc(), 'bg'); doc().bg = c; return E.sel; });

  /* ================= toolbars ================= */
  const listPanel = (items, close) => {
    const el = h('div', { class: 'style-list' });
    for (const it of items) {
      const b = h('button', { type: 'button', html: it.html || '' });
      if (it.label) b.appendChild(h('span', { text: it.label }));
      b.addEventListener('click', () => { close(); it.run(); });
      el.appendChild(b);
    }
    return el;
  };
  A.listPanel = listPanel;
  function tableGridPicker(r) {
    ui.openMenu([{
      custom: (close) => {
        const box = h('div', { class: 'tbl-grid-pick' });
        let R = 4, C = 5;
        const cells = h('div', { class: 'tgp-cells' });
        const lbl = h('div', { class: 'tgp-label', text: 'Cancel' });
        let cur = [0, 0];
        const draw = () => {
          cells.style.gridTemplateColumns = `repeat(${C}, 18px)`;
          L.clear(cells);
          for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) {
            const s = h('span', { class: i < cur[0] && j < cur[1] ? 'on' : '' });
            s.addEventListener('pointerenter', () => { cur = [i + 1, j + 1]; if (i + 2 > R && R < 12) R++; if (j + 2 > C && C < 10) C++; draw(); lbl.textContent = `${cur[0]} x ${cur[1]} Table`; });
            s.addEventListener('click', () => { close(); L.tables.insert(cur[0], cur[1]); });
            cells.appendChild(s);
          }
        };
        draw();
        box.append(cells, lbl);
        return box;
      },
    }], r, { cls: 'grid-menu' });
  }
  function columnsPicker(r) {
    ui.openMenu([{
      custom: (close) => {
        const box = h('div', { class: 'tbl-grid-pick' });
        const cells = h('div', { class: 'tgp-cells', style: 'grid-template-columns:repeat(4,22px)' });
        const lbl = h('div', { class: 'tgp-label', text: 'Cancel' });
        for (let j = 0; j < 4; j++) {
          const s = h('span', { style: 'height:30px' });
          s.addEventListener('pointerenter', () => { L.$$('span', cells).forEach((x, k) => x.classList.toggle('on', k <= j)); lbl.textContent = `${j + 1} Column${j ? 's' : ''}`; });
          s.addEventListener('click', () => { close(); A.setColumns(j + 1); });
          cells.appendChild(s);
        }
        box.append(cells, lbl);
        return box;
      },
    }], r, { cls: 'grid-menu' });
  }
  A.setColumns = (n) => E.edit('Columns', () => { const s = A.curSect(); D.touchKey(s.holder, s.key); s.sect.cols = Object.assign({}, s.sect.cols, { n, eq: true, w: [] }); return E.sel; });
  /** the section object governing the caret (holder/key let undo snapshot it) */
  A.curSect = function () {
    const d = doc();
    const p = E.sel ? E.sel.f.p : D.firstPara(d.main.blocks);
    const tb = D.topBlock(d, p);
    if (!tb || tb.cont !== d.main) return { sect: d.sect, holder: d, key: 'sect' };
    for (let i = tb.i; i < d.main.blocks.length; i++) { const b = d.main.blocks[i]; if (b.t === 'p' && b.sect) return { sect: b.sect, holder: b, key: 'sect', para: b }; }
    return { sect: d.sect, holder: d, key: 'sect' };
  };
  function lineSpacingMenu(r) {
    const cur = E.curPara().sp || {};
    const v = cur.rule === 'auto' || !cur.rule ? cur.line || 1 : null;
    ui.openMenu([1, 1.5, 2, 2.5, 3].map((x) => ({ label: String(x.toFixed(1)), checked: () => v === x, run: () => E.formatPara((pPr) => { pPr.sp = Object.assign({}, pPr.sp, { line: x, rule: 'auto' }); }, 'Line Spacing') }))
      .concat(['-', { label: '&More...', run: () => L.dlg.paragraph() }]), r);
  }
  function highlightDD(r, anchor) {
    ui.openMenu([{
      custom: (close) => {
        const el = h('div', { class: 'color-panel' });
        const g = h('div', { class: 'cp-grid hl' });
        for (const n of ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black']) {
          g.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${L.R.HIGHLIGHT[n]}`, 'data-tip': n.replace(/([A-Z])/g, ' $1').replace(/^./, (c) => c.toUpperCase()), 'aria-label': n, onclick: () => { close(); A.lastHighlight = n; A.applyHighlight(n); } }));
        }
        el.appendChild(g);
        el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: 'None', onclick: () => { close(); A.applyHighlight('none'); } }));
        return el;
      },
    }], anchor ? anchor.getBoundingClientRect() : r, { cls: 'color-menu' });
  }
  A.applyHighlight = (n) => { if (E.collapsed()) { A.highlighter = n; L.$('#scroller').classList.add('painting'); A.status('Select text to highlight. Press Esc to stop.'); return; } E.formatRun({ hl: n === 'none' ? undefined : n }, 'Highlight'); };
  function fontColorDD(r, anchor) {
    ui.colorMenu(anchor || r, { mode: 'font' }, (v) => {
      if (v && v.auto) A.applyFontColor('auto');
      else if (typeof v === 'string') { A.lastFontColor = v.replace('#', ''); A.applyFontColor(A.lastFontColor); }
      ui.refresh();
    });
  }
  A.applyFontColor = (c) => E.formatRun({ color: c === 'auto' ? undefined : c }, 'Font Color');
  const BORDER_ITEMS = [
    ['outside', 'bordersOut', 'Outside Border'], ['all', 'bordersAll', 'All Borders'], ['top', 'bordersTop', 'Top Border'], ['bottom', 'bordersBottom', 'Bottom Border'],
    ['left', 'bordersLeft', 'Left Border'], ['right', 'bordersRight', 'Right Border'], ['insideH', 'bordersInH', 'Inside Horizontal Border'], ['insideV', 'bordersInV', 'Inside Vertical Border'],
    ['inside', 'bordersIn', 'Inside Border'], ['diagDown', 'bordersDiagDown', 'Diagonal Down Border'], ['diagUp', 'bordersDiagUp', 'Diagonal Up Border'], ['hline', 'hline', 'Horizontal Line'], ['none', 'bordersNone', 'No Border'],
  ];
  A.BORDER_ITEMS = BORDER_ITEMS;
  function bordersDD(r) {
    ui.openMenu([{
      custom: (close) => {
        const g = h('div', { class: 'border-grid' });
        for (const [k, ic, n] of BORDER_ITEMS) {
          const b = h('button', { type: 'button', 'data-tip': n, 'aria-label': n, html: L.icons.get(ic) });
          b.addEventListener('click', () => { close(); A.lastBorder = k; A.applyBorder(k); ui.refresh(); });
          g.appendChild(b);
        }
        return g;
      },
    }], r, { cls: 'grid-menu' });
  }
  A.borderPen = { val: 'single', sz: 0.5, color: 'auto' };
  /** toolbar border button: paragraphs or table cells */
  A.applyBorder = function (k) {
    if (L.tables && L.tables.inTable() && (E.cellSel || L.tables.wholeCellsSelected())) { L.tables.applyBorder(k); return; }
    if (k === 'hline') { A.insertHorizontalLine(); return; }
    const b = Object.assign({}, A.borderPen);
    E.formatPara((pPr) => {
      const cur = Object.assign({}, pPr.borders || {});
      const set = (side, on) => { if (on) cur[side] = Object.assign({ space: side === 'left' || side === 'right' ? 4 : 1 }, b); else delete cur[side]; };
      const has = (side) => !!(cur[side] && cur[side].val && cur[side].val !== 'nil');
      switch (k) {
        case 'none': pPr.borders = undefined; return;
        case 'outside': case 'all': { const on = !(has('top') && has('bottom') && has('left') && has('right')); ['top', 'bottom', 'left', 'right'].forEach((s) => set(s, on)); if (k === 'all') set('between', on); break; }
        case 'inside': case 'insideH': set('between', !has('between')); break;
        case 'insideV': break;
        default: if (['top', 'bottom', 'left', 'right'].includes(k)) set(k, !has(k));
      }
      pPr.borders = Object.keys(cur).length ? cur : undefined;
    }, 'Borders');
  };
  A.insertHorizontalLine = () => E.edit('Horizontal Line', () => {
    let pos = E.deleteSelection();
    if (pos.o > 0) pos = O.splitPara(pos);
    D.touch(pos.p);
    pos.p.pPr.borders = { bottom: { val: 'single', sz: 1.5, color: 'auto', space: 1 } };
    return O.splitPara(D.pos(pos.p, D.plen(pos.p)));
  });

  A.openAutoShapes = function (at) {
    const r = at || (L.$('[data-as]') ? L.$('[data-as]').getBoundingClientRect() : { left: 100, bottom: 100 });
    const grid = (list) => ({ custom: (close) => { const g = h('div', { class: 'shape-grid' }); for (const k of list) { const p = L.geom.get(k); const b = h('button', { type: 'button', 'data-tip': p.label || k, 'aria-label': p.label || k, html: L.geom.preview(k, 20) }); b.addEventListener('click', () => { close(); L.drawtool && L.drawtool.start({ geom: k, line: !!p.isLine }); }); g.appendChild(b); } return g; } });
    ui.openMenu([
      { label: '&Lines', icon: 'line', sub: ['drawLine', 'drawArrow', 'drawDblArrow'] },
      ...['Basic Shapes', 'Block Arrows', 'Flowchart', 'Stars and Banners', 'Callouts'].map((g) => ({ label: g.replace(/^(\w)/, '&$1'), icon: { 'Basic Shapes': 'autoshapes', 'Block Arrows': 'arrowStyle', Flowchart: 'diagram', 'Stars and Banners': 'animation', Callouts: 'comment' }[g], sub: [grid(L.geom.MENU[g])] })),
      '-', { label: '&More AutoShapes...', icon: 'clipart', run: () => L.panes && L.panes.task.show('clipart') },
    ], { left: r.left, bottom: r.bottom }, { selectFirst: !at });
  };
  function buildToolbars() {
    const host = L.$('#toolbars');
    const stState = () => (E.sel ? styleState() : '');
    const styleCombo = ui.combo({
      id: 'tb-style', width: 118, tip: 'Style', listCls: 'styles',
      options: () => A.styleOptions().map((s) => s.label),
      preview: (v) => { const s = A.styleOptions().find((x) => x.label === v); return s ? s.css : ''; },
      value: stState, onChange: (v) => A.applyStyleByName(v),
    });
    const fontCombo = ui.combo({ id: 'tb-font', width: 136, tip: 'Font', options: () => A.fontOptions(), preview: (v) => `font-family:${L.fontStack(v)};font-size:13px`, listCls: 'fonts', value: () => { const v = E.uniformRun('font'); return v || ''; }, onChange: (v) => A.setFont(v) });
    const sizeCombo = ui.combo({ id: 'tb-size', width: 40, tip: 'Font Size', options: () => L.SIZE_LIST, value: () => { const v = E.uniformRun('sz'); return v ? String(L.round(v, 1)) : ''; }, onChange: (v) => A.setFontSize(parseFloat(v)) });
    const zoomCombo = ui.combo({ id: 'tb-zoom', width: 74, tip: 'Zoom', options: () => ['500%', '200%', '150%', '100%', '75%', '50%', '25%', '10%', 'Page Width', 'Text Width', 'Whole Page', 'Two Pages'], value: () => Math.round(LY.zoom * 100) + '%', onChange: (v) => A.zoomTo(v) });
    A.toolbarsEl = {
      standard: ui.toolbar('standard', 'Standard', ['new', 'open', 'save', 'permission', 'sendMail', 'print', 'printPreview', 'spelling', 'research', '|', 'cut', 'copy', 'paste', 'painter', '|',
        ui.tbSplit('undo', (r) => A.historyMenu(r, 'undo')), ui.tbSplit('redo', (r) => A.historyMenu(r, 'redo')), '|', 'hyperlink', ui.tbButton('tb_tables', { icon: 'tablesBorders' }),
        ui.tbDrop('', 'table', (r) => tableGridPicker(r), 'Insert Table'), 'insertExcel', ui.tbDrop('', 'columns', (r) => columnsPicker(r), 'Columns'), ui.tbButton('tb_drawing', { icon: 'drawing' }), 'docMap', 'showMarks', zoomCombo, 'help', ui.tbButton('viewReading', { label: 'Read', icon: 'read' })]),
      formatting: ui.toolbar('formatting', 'Formatting', ['stylesPane', styleCombo, fontCombo, sizeCombo, 'bold', 'italic', 'underline', '|', 'alignLeft', 'alignCenter', 'alignRight', 'justify', '|',
        ui.tbDrop('', 'lineSpacing', lineSpacingMenu, 'Line Spacing'), '|', 'numbering', 'bullets', 'decIndent', 'incIndent', '|', ui.tbSplit('borderBtn', bordersDD), ui.tbSplit('highlightBtn', highlightDD), ui.tbSplit('fontColorBtn', fontColorDD)]),
      tables: ui.toolbar('tables', 'Tables and Borders', ['drawTable', 'eraser', '|', ui.tbDrop('', 'dashStyle', (r) => A.borderStyleMenu(r), 'Line Style'), ui.tbDrop('', 'lineStyle', (r) => A.borderWeightMenu(r), 'Line Weight'), ui.tbDrop('', 'lineColor', (r, b) => ui.colorMenu(b, { mode: 'plain', automatic: true }, (v) => { A.borderPen.color = v && v.auto ? 'auto' : typeof v === 'string' ? v.replace('#', '') : A.borderPen.color; }), 'Border Color'),
        ui.tbSplit('borderBtn', bordersDD), ui.tbDrop('', 'fillColor', (r, b) => ui.colorMenu(b, { mode: 'fill' }, (v) => { if (L.tables) L.tables.shade(v && v.none ? null : typeof v === 'string' ? v.replace('#', '') : null); }), 'Shading Color'), '|',
        ui.tbDrop('', 'table', (r) => ui.openMenu(['insertTable', '-', 'colLeft', 'colRight', 'rowAbove', 'rowBelow', 'insertCells', '-', 'autofitContents', 'autofitWindow', 'fixedWidth'], r), 'Insert Table'), 'mergeCells', 'splitCells', '|',
        ui.tbDrop('', 'alignTop', (r) => ui.openMenu(['cellTop', 'cellMiddle', 'cellBottom'], r), 'Align'), 'distRows', 'distCols', '|', 'tableAutoFormat', 'textDirection', '|', 'sortAsc', 'sortDesc', 'autoSum']),
      reviewing: ui.toolbar('reviewing', 'Reviewing', [ui.tbDrop('Final Showing Markup', null, (r) => ui.openMenu(['markupFinalMarkup', 'markupFinal', 'markupOriginalMarkup', 'markupOriginal'], r), 'Display for Review'),
        ui.tbDrop('Show', null, (r) => ui.openMenu(['showComments', 'showInsDel', 'showFormatting', '-', 'reviewersMenu', '-', 'balloonsToggle', 'reviewingPane'], r), 'Show'), '|', 'prevChange', 'nextChange', '|',
        ui.tbSplit('acceptChange', (r) => ui.openMenu(['acceptChange', 'acceptShown', 'acceptAll'], r)), ui.tbSplit('rejectChange', (r) => ui.openMenu(['rejectChange', 'rejectShown', 'rejectAll', '-', 'deleteAllComments'], r)), '|',
        'insertComment', 'trackChanges', 'reviewingPane']),
      picture: ui.toolbar('picture', 'Picture', ['insertPicture', ui.tbDrop('', 'colorPic', (r) => ui.openMenu([{ label: '&Automatic', run: () => A.picMode('') }, { label: '&Grayscale', run: () => A.picMode('gray') }, { label: '&Black & White', run: () => A.picMode('bw') }, { label: '&Washout', run: () => A.picMode('washout') }], r), 'Color'),
        'picMoreContrast', 'picLessContrast', 'picMoreBright', 'picLessBright', '|', 'picCrop', 'picRotate', ui.tbDrop('', 'lineStyle', (r) => A.objLineWeightMenu(r), 'Line Style'), '|', ui.tbDrop('', 'textWrap', (r) => A.wrapMenu(r), 'Text Wrapping'), 'formatObject', 'picReset']),
      outlining: ui.toolbar('outlining', 'Outlining', ['promoteH1', 'promote', ui.combo({ id: 'tb-olvl', width: 74, tip: 'Outline Level', options: () => ['Level 1', 'Level 2', 'Level 3', 'Level 4', 'Level 5', 'Level 6', 'Level 7', 'Level 8', 'Level 9', 'Body text'], value: () => { const l = E.sel ? D.headingLevel(doc(), E.sel.f.p) : 0; return l ? 'Level ' + l : 'Body text'; }, onChange: (v) => L.outline && L.outline.setLevel(/body/i.test(v) ? 0 : parseInt(v.replace(/\D/g, ''), 10) || 0) }), 'demote', 'demoteBody', '|', 'moveUp', 'moveDown', 'expand', 'collapse',
        ui.combo({ id: 'tb-olshow', width: 96, tip: 'Show Level', options: () => ['Show All Levels', 'Show Level 1', 'Show Level 2', 'Show Level 3', 'Show Level 4', 'Show Level 5', 'Show Level 6', 'Show Level 7', 'Show Level 8', 'Show Level 9'], value: () => (L.outline && L.outline.showLevel ? 'Show Level ' + L.outline.showLevel : 'Show All Levels'), onChange: (v) => L.outline && L.outline.show(/all/i.test(v) ? 0 : parseInt(v.replace(/\D/g, ''), 10)) }),
        'olFirstLine', 'olFormatting', '|', 'updateTOC', 'gotoTOC']),
      headerFooter: ui.toolbar('headerFooter', 'Header and Footer', [ui.tbDrop('Insert AutoText', null, (r) => ui.openMenu(L.autocorrect ? L.autocorrect.hfAutoText() : [], r), 'Insert AutoText'), '|', 'hfPageNum', 'hfNumPages', 'hfFormatPageNum', 'hfDate', 'hfTime', '|', 'pageSetup', 'hfShowText', 'hfLinkPrev', '|', 'hfSwitch', 'hfPrev', 'hfNext', '|', ui.tbButton('hfClose', { label: 'Close' })]),
      mailMerge: ui.toolbar('mailMerge', 'Mail Merge', ['mmMainDoc', 'mmOpenSource', 'mmRecipients', '|', 'mmAddress', 'mmGreeting', 'mmInsertField', '|', 'mmViewData', 'mmFirst', 'mmPrev', ui.combo({ id: 'tb-mmrec', width: 34, tip: 'Go to Record', options: () => [], value: () => String(L.mailmerge ? L.mailmerge.rec + 1 : 1), onChange: (v) => L.mailmerge && L.mailmerge.go(parseInt(v, 10) - 1) }), 'mmNext', 'mmLast', '|', 'mmMergeNew', 'mmMergePrint']),
      wordart: ui.toolbar('wordart', 'WordArt', ['insertWordArt', ui.tbButton('wordartEdit', { label: 'Edit Text...' }), 'wordartGallery', 'formatObject', ui.tbDrop('', 'waShape', (r) => ui.openMenu(L.drawing.WARPS.map(([k, n]) => ({ label: n, run: () => L.drawing.modifySelected('WordArt Shape', (it) => { it.warp = k; }) })), r), 'WordArt Shape'), ui.tbDrop('', 'textWrap', (r) => A.wrapMenu(r), 'Text Wrapping')]),
    };
    const row1 = h('div', { class: 'tb-row' }, A.toolbarsEl.standard);
    const row2 = h('div', { class: 'tb-row' }, A.toolbarsEl.formatting);
    const row3 = h('div', { class: 'tb-row' }, A.toolbarsEl.tables, A.toolbarsEl.reviewing, A.toolbarsEl.outlining, A.toolbarsEl.headerFooter, A.toolbarsEl.picture, A.toolbarsEl.mailMerge, A.toolbarsEl.wordart);
    host.append(row1, row2, row3);
    A.tbRows = [row1, row2, row3];
    /* drawing toolbar */
    const asBtn = ui.tbDrop('A&utoShapes', null, (r) => A.openAutoShapes(r), 'AutoShapes');
    asBtn.setAttribute('data-as', '1');
    A.toolbarsEl.drawing = ui.toolbar('drawing', 'Drawing', [ui.tbDrop('D&raw', null, (r) => A.drawMenu(r), 'Draw'), 'selectObjects', '|', asBtn, 'drawLine', 'drawArrow', 'drawRect', 'drawOval', 'textBox', 'insertWordArt', 'insertDiagram', 'insertClipArt', 'insertPicture', '|',
      ui.tbSplit('fillColor', (r, b) => A.objColorMenu(b, 'fill')), ui.tbSplit('lineColor', (r, b) => A.objColorMenu(b, 'line')), ui.tbSplit('fontColorBtn', fontColorDD), '|',
      ui.tbDrop('', 'lineStyle', (r) => A.objLineWeightMenu(r), 'Line Style'), ui.tbDrop('', 'dashStyle', (r) => A.objDashMenu(r), 'Dash Style'), ui.tbDrop('', 'arrowStyle', (r) => A.objArrowMenu(r), 'Arrow Style'), ui.tbDrop('', 'shadowStyle', (r) => A.objShadowMenu(r), 'Shadow Style')]);
    L.$('#drawrow').appendChild(A.toolbarsEl.drawing);
    /* view buttons & browse */
    const vb = L.$('#viewbar');
    vb.append(ui.tbButton('viewNormal'), ui.tbButton('viewWeb'), ui.tbButton('viewPrint'), ui.tbButton('viewOutline'), ui.tbButton('viewReading'), h('span', { class: 'vb-sp' }),
      h('span', { class: 'browse' }, ui.tbButton('browsePrev'), ui.tbDrop('', 'browseObj', (r) => A.browseMenu(r), 'Select Browse Object'), ui.tbButton('browseNext')));
  }
  A.fontOptions = () => {
    const recent = (A.opts.recentFonts || []).filter((f) => L.FONT_LIST.includes(f) || true).slice(0, 4);
    const docFonts = new Set();
    for (const k in doc().styles) { const r = doc().styles[k].rPr; if (r && r.font) docFonts.add(r.font); }
    const all = Array.from(new Set(L.FONT_LIST.concat(Array.from(docFonts)))).sort((a, b) => a.localeCompare(b));
    return recent.length ? recent.concat(all.filter((f) => !recent.includes(f))) : all;
  };
  A.setFont = (v) => { if (!v) return; A.opts.recentFonts = [v].concat((A.opts.recentFonts || []).filter((f) => f !== v)).slice(0, 6); A.saveOpts(); E.formatRun({ font: v }, 'Font'); E.refocus(); };
  A.setFontSize = (v) => { if (!(v > 0)) return; v = L.clamp(Math.round(v * 2) / 2, 1, 1638); E.formatRun({ sz: v }, 'Font Size'); E.refocus(); };
  function styleState() {
    const d = doc();
    if (!E.sel) return '';
    const it = D.itemBefore(E.sel.f.p, E.sel.f.o) || D.itemAfter(E.sel.f.p, E.sel.f.o);
    if (it && it.rPr && it.rPr.style && d.styles[it.rPr.style] && it.rPr.style !== 'DefaultParagraphFont') return D.styleDisplayName(d.styles[it.rPr.style]);
    const sid = E.sel.f.p.pPr.style || 'Normal';
    const s = d.styles[sid];
    let name = s ? D.styleDisplayName(s) : sid;
    /* Word shows "Normal + Bold" style names in the box only in the task pane; the box shows the base style */
    return name;
  }
  A.styleState = styleState;
  A.styleOptions = function () {
    const d = doc();
    const used = D.usedStyles(d);
    const ids = Object.keys(d.styles).filter((k) => { const s = d.styles[k]; return (s.type === 'paragraph' || s.type === 'character') && (used.has(k) || s.q || /^Heading[1-3]$/.test(k) || k === 'Normal') && !s.hidden; });
    return ids.map((k) => {
      const s = d.styles[k];
      const r = s.type === 'paragraph' ? D.styleProps(d, k).rPr : D.mergeRPr(d.defaults.rPr, D.charStyleProps(d, k));
      const sz = Math.min(16, r.sz || 12);
      const css = `font-family:${L.fontStack(r.font)};font-size:${Math.round(sz * 1.2)}px;${r.b ? 'font-weight:bold;' : ''}${r.i ? 'font-style:italic;' : ''}${r.color && r.color !== 'auto' ? `color:#${r.color};` : ''}`;
      return { id: k, label: D.styleDisplayName(s), css, type: s.type };
    }).sort((a, b) => a.label.localeCompare(b.label));
  };
  A.applyStyleByName = function (name) {
    const d = doc();
    if (!name || !E.sel) return;
    let s = D.findStyleByName(d, name);
    if (!s) {
      /* typing a new name in the Style box defines a style by example */
      const id = D.uniqueStyleId(d, name);
      const pp = L.clone(E.curPara());
      delete pp.num; delete pp.style;
      const rr = L.clone(E.curRun());
      delete rr.link; delete rr.ins; delete rr.del; delete rr.style;
      E.edit('New Style', () => {
        D.touchKey(d, 'styles');
        d.styles[id] = { id, name, type: 'paragraph', basedOn: E.sel.f.p.pPr.style || 'Normal', next: id, pPr: pp, rPr: rr, custom: true, q: true };
        D.stylesChanged();
        const [a, b] = E.range();
        O.applyStyle(a, b, id);
        return E.sel;
      });
      return;
    }
    const [a, b] = E.range();
    E.edit('Style', () => { O.applyStyle(a, b, s.id); return E.sel; });
    E.refocus();
  };
  A.historyMenu = function (r, kind) {
    const list = kind === 'undo' ? D.undoList() : D.redoList();
    if (!list.length) return;
    ui.openMenu([{ custom: (close) => {
      const el = h('div', { class: 'style-list', style: 'max-height:260px;overflow:auto;min-width:180px' });
      const lbl = h('div', { class: 'tgp-label', text: `${kind === 'undo' ? 'Undo' : 'Redo'} 1 Action` });
      const btns = list.slice(0, 40).map((label, i) => {
        const b = h('button', { type: 'button', text: label });
        b.addEventListener('pointerenter', () => { btns.forEach((x, k) => { x.style.background = k <= i ? '#316ac5' : ''; x.style.color = k <= i ? '#fff' : ''; }); lbl.textContent = `${kind === 'undo' ? 'Undo' : 'Redo'} ${i + 1} Action${i ? 's' : ''}`; });
        b.addEventListener('click', () => { close(); for (let k = 0; k <= i; k++) ui.exec(kind); });
        el.appendChild(b);
        return b;
      });
      const wrap = h('div', { style: 'background:#fff' }, el, lbl);
      return wrap;
    } }], r);
  };

  /* ---------- drawing toolbar helpers ---------- */
  A.drawMenu = function (r) {
    ui.openMenu(['groupObj', 'ungroupObj', '-', { label: '&Order', sub: ['bringFront', 'sendBack', 'bringForward', 'sendBackward', 'bringFrontText', 'sendBehindText'] },
      { label: '&Grid...', run: () => L.dlg.drawingGrid && L.dlg.drawingGrid() }, { label: 'Nud&ge', sub: ['nudgeUp', 'nudgeDown', 'nudgeLeft', 'nudgeRight'] },
      { label: 'Rotate or Fl&ip', sub: ['rotateLeft', 'rotateRight', '-', 'flipH', 'flipV'] }, { label: '&Text Wrapping', sub: () => A.wrapItems() }, '-', 'setDefaults'], r);
  };
  A.wrapItems = () => L.drawing.WRAPS.map(([k, n]) => ({ label: n, checked: () => { const it = L.drawing.selectedItem(); return !!it && (k === 'inline' ? !it.float : it.float && it.float.wrap === k); }, enabled: () => !!L.ed.objSel, run: () => L.drawing.modifySelected('Text Wrapping', (it) => L.drawing.setWrap(it, k)) }))
    .concat(['-', { label: '&Edit Wrap Points', enabled: () => false }]);
  A.wrapMenu = (r) => ui.openMenu(A.wrapItems(), r);
  A.objColorMenu = function (anchor, kind) {
    ui.colorMenu(anchor, { mode: kind === 'fill' ? 'fill' : 'line', effects: kind === 'fill' }, async (v) => {
      if (!E.objSel) { A.status('Select a drawing object first.'); return; }
      if (kind === 'fill') {
        if (v && v.effects) { const it = L.drawing.selectedItem(); const f = L.dlg.fillEffects ? await L.dlg.fillEffects(it && it.fill) : null; if (f) L.drawing.modifySelected('Fill Color', (x) => { x.fill = f; }); return; }
        L.drawing.modifySelected('Fill Color', (it) => { it.fill = v && v.none ? { t: 'none' } : { t: 'solid', c: v, a: 1 }; });
      } else L.drawing.modifySelected('Line Color', (it) => { if (it.t === 'img') it.border = v && v.none ? undefined : { val: 'single', sz: (it.border && it.border.sz) || 0.75, color: String(v).replace('#', '') }; else it.line = v && v.none ? { t: 'none' } : Object.assign({ w: 0.75, dash: 'solid' }, it.line && it.line.t !== 'none' ? it.line : {}, { c: v }); });
    });
  };
  const lineSVG = (w, dash, extra) => `<svg width="90" height="14" viewBox="0 0 90 14"><path d="M2 7h86" stroke="#000" stroke-width="${w}" ${dash ? `stroke-dasharray="${dash}"` : ''} fill="none"/>${extra || ''}</svg>`;
  A.WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
  A.objLineWeightMenu = (r) => ui.openMenu([{ custom: (close) => listPanel(A.WEIGHTS.map((w) => ({ html: lineSVG(Math.max(0.5, w * 1.33)), label: w + ' pt', run: () => L.drawing.modifySelected('Line Style', (it) => { if (it.t === 'img') it.border = Object.assign({ val: 'single', color: '000000' }, it.border || {}, { sz: w }); else it.line = Object.assign({ c: '#000000', dash: 'solid' }, it.line && it.line.t !== 'none' ? it.line : {}, { w }); }) })).concat([{ label: 'More Lines...', run: () => ui.exec('formatObject') }]), close) }], r);
  A.objDashMenu = (r) => ui.openMenu([{ custom: (close) => listPanel([['solid', null], ['sysDot', '1 1'], ['sysDash', '3 1'], ['dash', '4 3'], ['dashDot', '4 3 1 3'], ['lgDash', '8 3'], ['lgDashDot', '8 3 1 3'], ['lgDashDotDot', '8 3 1 3 1 3']].map(([k, da]) => ({ html: lineSVG(2, da), run: () => L.drawing.modifySelected('Dash Style', (it) => { if (it.line && it.line.t !== 'none') it.line = Object.assign({}, it.line, { dash: k }); }) })), close) }], r);
  A.objArrowMenu = (r) => {
    const P = [[null, null], [null, 'triangle'], ['triangle', null], ['triangle', 'triangle'], [null, 'arrow'], [null, 'stealth'], [null, 'oval'], [null, 'diamond'], ['oval', 'oval']];
    const head = (t, x, dir) => !t ? '' : t === 'oval' ? `<circle cx="${x}" cy="7" r="3"/>` : t === 'diamond' ? `<path d="M${x - 4} 7l4-4 4 4-4 4z"/>` : t === 'arrow' ? `<path d="M${x - dir * 6} 2l${dir * 6} 5-${dir * 6} 5" fill="none" stroke="#000" stroke-width="1.5"/>` : `<path d="M${x} 7l${-dir * 8} -4v8z"/>`;
    ui.openMenu([{ custom: (close) => listPanel(P.map(([a, b]) => ({ html: `<svg width="90" height="14" viewBox="0 0 90 14" fill="#000"><path d="M6 7h78" stroke="#000" stroke-width="1.5"/>${head(a, 4, -1)}${head(b, 86, 1)}</svg>`, run: () => L.drawing.modifySelected('Arrow Style', (it) => { if (!it.line || it.line.t === 'none') return; it.line = Object.assign({}, it.line, { head: a ? { type: a, w: 'med', len: 'med' } : undefined, tail: b ? { type: b, w: 'med', len: 'med' } : undefined }); }) })), close) }], r);
  };
  A.objShadowMenu = (r) => {
    const S = [null, [3, 3], [-3, 3], [3, -3], [-3, -3], [0, 4], [4, 0], [6, 6]];
    ui.openMenu([{ custom: (close) => { const g = h('div', { class: 'shape-grid', style: 'grid-template-columns:repeat(4,30px)' }); S.forEach((s) => { const b = h('button', { type: 'button', style: 'width:30px;height:30px', 'aria-label': s ? 'Shadow Style' : 'No Shadow', html: s ? `<svg width="24" height="24"><rect x="${6 + s[0] / 1.5}" y="${6 + s[1] / 1.5}" width="12" height="12" fill="#888"/><rect x="6" y="6" width="12" height="12" fill="#fff" stroke="#000"/></svg>` : '<svg width="24" height="24"><rect x="6" y="6" width="12" height="12" fill="#fff" stroke="#000"/><path d="M3 21L21 3" stroke="#d33"/></svg>' }); b.addEventListener('click', () => { close(); L.drawing.modifySelected('Shadow', (it) => { it.shadow = s ? { c: '#000000', a: 0.5, dx: s[0], dy: s[1], blur: 0 } : undefined; }); }); g.appendChild(b); }); return g; } }], r);
  };
  A.borderStyleMenu = (r) => ui.openMenu([{ custom: (close) => listPanel([['single', null], ['dotted', '1 2'], ['dashed', '4 3'], ['dotDash', '4 3 1 3'], ['double', null], ['thick', null], ['triple', null], ['wave', null], ['threeDEmboss', null]].map(([k, da]) => ({ html: k === 'double' ? '<svg width="90" height="14"><path d="M2 5h86M2 9h86" stroke="#000"/></svg>' : k === 'triple' ? '<svg width="90" height="14"><path d="M2 3h86M2 7h86M2 11h86" stroke="#000"/></svg>' : k === 'wave' ? '<svg width="90" height="14"><path d="M2 7q5-5 10 0t10 0t10 0t10 0t10 0t10 0t10 0t10 0t6 0" fill="none" stroke="#000"/></svg>' : lineSVG(k === 'thick' ? 3 : 1, da), run: () => { A.borderPen.val = k; } })).concat([{ label: 'No Border', run: () => { A.borderPen.val = 'nil'; } }]), close) }], r);
  A.borderWeightMenu = (r) => ui.openMenu([{ custom: (close) => listPanel([0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => ({ html: lineSVG(Math.max(0.5, w * 1.33)), label: w + ' pt', run: () => { A.borderPen.sz = w; } })), close) }], r);
  A.picMode = (m) => L.drawing.modifySelected('Picture Color', (it) => { it.gray = m === 'gray' || undefined; it.bw = m === 'bw' || undefined; it.washout = m === 'washout' || undefined; });
  A.picAdjust = (k, d) => L.drawing.modifySelected(k === 'bright' ? 'Brightness' : 'Contrast', (it) => { it[k] = L.clamp(L.round((it[k] || 0) + d, 2), -1, 1) || undefined; });

  /* ================= status bar ================= */
  let statusTimer = 0;
  function buildStatus() {
    const sb = L.$('#statusbar');
    const cell = (cls, tip) => h('div', { class: 'sb-cell ' + (cls || ''), 'data-tip': tip || null });
    A.sb = {
      page: cell('', 'Page'), sec: cell('', 'Section'), pages: cell('', 'Page / pages'), at: cell('', 'Vertical position'), ln: cell('', 'Line'), col: cell('', 'Column'),
      rec: cell('flag', 'Record macro'), trk: cell('flag btn', 'Track Changes (double-click to toggle)'), ext: cell('flag btn', 'Extend selection'), ovr: cell('flag btn', 'Overtype (double-click to toggle)'),
      lang: cell('btn', 'Language'), spell: cell('btn sb-spell', 'Spelling and grammar status'), msg: cell('grow'),
    };
    A.sb.rec.textContent = 'REC'; A.sb.trk.textContent = 'TRK'; A.sb.ext.textContent = 'EXT'; A.sb.ovr.textContent = 'OVR';
    A.sb.lang.textContent = 'English (U.S.)';
    A.sb.spell.innerHTML = '<svg width="18" height="14" viewBox="0 0 18 14"><rect x="1.5" y="1.5" width="11" height="11" fill="#fff" stroke="#556"/><path d="M3 4h7M3 6.5h7M3 9h5" stroke="#889"/><path d="M10 9l2.5 3 4.5-7" fill="none" stroke="#2f9a3a" stroke-width="1.6"/></svg>';
    A.sb.trk.addEventListener('dblclick', () => ui.exec('trackChanges'));
    A.sb.ovr.addEventListener('dblclick', () => { E.overtype = !E.overtype; A.updateStatus(); });
    A.sb.ext.addEventListener('dblclick', () => { E.extend = !E.extend; A.updateStatus(); });
    A.sb.lang.addEventListener('dblclick', () => ui.exec('setLanguage'));
    A.sb.spell.addEventListener('dblclick', () => ui.exec('spelling'));
    A.sb.page.addEventListener('dblclick', () => ui.exec('goTo'));
    sb.append(A.sb.page, A.sb.sec, A.sb.pages, A.sb.at, A.sb.ln, A.sb.col, A.sb.rec, A.sb.trk, A.sb.ext, A.sb.ovr, A.sb.lang, A.sb.spell, A.sb.msg);
  }
  A.status = (msg, ms) => { if (!A.sb) return; A.sb.msg.textContent = msg; clearTimeout(statusTimer); statusTimer = setTimeout(() => { A.sb.msg.textContent = ''; }, ms || 4000); };
  A.updateStatus = L.debounce(function () {
    if (!A.sb) return;
    const d = doc();
    const n = LY.pages.length;
    const sel = E.sel;
    A.sb.trk.classList.toggle('on', !!d.settings.track);
    A.sb.ovr.classList.toggle('on', !!E.overtype);
    A.sb.ext.classList.toggle('on', !!E.extend);
    if (!sel) return;
    const pi = LY.pageOf(sel.f);
    const pg = pi >= 0 ? LY.pages[pi] : null;
    if (pg) {
      A.sb.page.textContent = 'Page ' + D.fmtNum(pg.numVal, pg.numFmt);
      A.sb.sec.textContent = 'Sec ' + (pg.si + 1);
      A.sb.pages.textContent = `${pi + 1}/${n}`;
      const r = E.caretRect();
      if (r) {
        const pr = pg.el.getBoundingClientRect();
        const atPt = (r.top - pr.top) / LY.zoom / D.PX;
        A.sb.at.textContent = 'At ' + D.fmtLen(atPt);
        /* line number on the page and column within the line */
        const lc = A.lineCol(sel.f, pg);
        A.sb.ln.textContent = 'Ln ' + lc.ln;
        A.sb.col.textContent = 'Col ' + lc.col;
      }
    } else {
      A.sb.page.textContent = LY.view === 'print' ? '' : 'Page 1';
      A.sb.sec.textContent = 'Sec 1';
      A.sb.pages.textContent = LY.view === 'print' ? '' : '';
      A.sb.at.textContent = ''; A.sb.ln.textContent = ''; A.sb.col.textContent = 'Col ' + (sel.f.o + 1);
    }
    const r = E.curRun();
    const lang = r.lang || (d.defaults.rPr && d.defaults.rPr.lang) || 'en-US';
    A.sb.lang.textContent = A.langName(lang);
  }, 60);
  A.LANGS = [['en-US', 'English (U.S.)'], ['en-GB', 'English (U.K.)'], ['en-AU', 'English (Australia)'], ['en-CA', 'English (Canada)'], ['fr-FR', 'French (France)'], ['fr-CA', 'French (Canada)'], ['de-DE', 'German (Germany)'], ['es-ES', 'Spanish (Spain)'], ['es-MX', 'Spanish (Mexico)'], ['it-IT', 'Italian (Italy)'], ['pt-BR', 'Portuguese (Brazil)'], ['pt-PT', 'Portuguese (Portugal)'], ['nl-NL', 'Dutch (Netherlands)'], ['sv-SE', 'Swedish'], ['da-DK', 'Danish'], ['nb-NO', 'Norwegian (Bokmål)'], ['fi-FI', 'Finnish'], ['pl-PL', 'Polish'], ['ru-RU', 'Russian'], ['ja-JP', 'Japanese'], ['zh-CN', 'Chinese (PRC)'], ['ko-KR', 'Korean']];
  A.langName = (code) => { const hit = A.LANGS.find((x) => x[0].toLowerCase() === String(code).toLowerCase()); return hit ? hit[1] : code; };
  A.lineCol = function (pos, pg) {
    const frag = LY.fragFor(pos.p, pos.o);
    if (!frag) return { ln: 1, col: pos.o + 1 };
    const z = LY.zoom;
    const body = frag.closest('.pg-col, .pg-hdr, .pg-ftr, .cf') || pg.body;
    const top = body.getBoundingClientRect().top;
    const lines = LY.lines(frag, top, z);
    let li = 0;
    for (let i = 0; i < lines.length; i++) if (lines[i].o <= pos.o) li = i;
    const col = pos.o - (lines[li] ? lines[li].o : 0) + 1;
    /* lines above this fragment on the page */
    let ln = li + 1;
    if (frag.closest('.pg-col')) {
      const col0 = frag.closest('.pg-col');
      for (const f of col0.querySelectorAll('.p')) {
        if (f === frag) break;
        if (f.compareDocumentPosition(frag) & Node.DOCUMENT_POSITION_FOLLOWING) {
          const lh = parseFloat(getComputedStyle(f.firstChild).lineHeight) || 16;
          ln += Math.max(1, Math.round((f.firstChild.offsetHeight) / lh));
        }
      }
    }
    return { ln, col };
  };
  A.updateTitle = () => { L.$('#titletext').textContent = `${A.fileName} - ${L.APP}`; document.title = `${A.fileName} - Quire 2003`; };

  /* ================= rulers ================= */
  A.updateRulers = function () {
    const w = L.$('#work');
    const show = A.opts.ruler && (A.view === 'print' || A.view === 'normal' || A.view === 'web');
    w.classList.toggle('rulers', !!show);
    w.classList.toggle('vruler', !!(show && A.view === 'print' && A.opts.vruler !== false));
    L.$('#hruler').hidden = !show;
    L.$('#vruler').hidden = !(show && A.view === 'print' && A.opts.vruler !== false);
    L.$('#rcorner').hidden = !show;
    if (L.rulers) L.rulers.draw();
  };

  /* ================= views ================= */
  A.setView = function (v) {
    const app = L.$('#app');
    if (A.view === 'preview' && v !== 'preview' && L.panes && L.panes.preview) L.panes.preview.exit();
    if (A.view === 'reading' && v !== 'reading') { app.classList.remove('reading'); LY.setMultiPage(0); }
    if (v === 'preview') { A.prevView = A.view; }
    A.view = v;
    if (v === 'reading') { A.prevZoom = LY.zoom; }
    LY.view = v === 'reading' || v === 'preview' ? 'print' : v;
    LY.root.classList.toggle('readonly', v === 'reading' || v === 'preview');
    L.$('#scroller').classList.toggle('flowview', v === 'normal' || v === 'outline');
    L.$('#scroller').classList.toggle('webview', v === 'web');
    app.classList.toggle('preview', v === 'preview');
    app.classList.toggle('reading', v === 'reading');
    if (LY.hfEdit && v !== 'print') A.closeHeaderFooter(true);
    LY.render();
    if (v === 'reading') { LY.setMultiPage(2); LY.setZoom(LY.zoomFor('twoPages')); }
    else if (v === 'preview') { if (L.panes && L.panes.preview) L.panes.preview.enter(); }
    else if (A.prevZoom && v !== 'reading') { LY.setZoom(A.prevZoom); A.prevZoom = null; }
    E.restoreDom();
    A.updateRulers();
    A.updateToolbars();
    A.updateStatus();
    ui.refresh();
    L.bus.emit('view', v);
  };
  A.zoomTo = function (v) {
    let z;
    if (/page width/i.test(v)) z = LY.zoomFor('pageWidth');
    else if (/text width/i.test(v)) z = LY.zoomFor('textWidth');
    else if (/whole/i.test(v)) z = LY.zoomFor('wholePage');
    else if (/two/i.test(v)) { LY.setMultiPage(2); z = LY.zoomFor('twoPages'); }
    else { const n = parseFloat(v); if (!(n > 0)) return; z = n / 100; }
    if (!/two/i.test(v) && LY.multi && A.view !== 'reading') LY.setMultiPage(0);
    LY.setZoom(L.clamp(z, 0.1, 5));
    A.opts.zoom = LY.zoom; A.saveOpts();
    if (L.rulers) L.rulers.draw();
    E.focus();
  };

  /* ================= toolbars visibility ================= */
  A.tbShown = (id) => { const v = A.opts.toolbars[id]; return v === true || (v === 'auto' && ctxToolbar(id)); };
  function ctxToolbar(id) {
    const d = doc();
    if (id === 'reviewing') return !!(d.settings.track || Object.keys(d.comments).length || A.hasRevisions);
    if (id === 'picture') return !!(E.objSel && E.objSel.it && E.objSel.it.t === 'img');
    if (id === 'outlining') return A.view === 'outline';
    if (id === 'headerFooter') return !!LY.hfEdit;
    if (id === 'wordart') return !!(E.objSel && E.objSel.it && E.objSel.it.wordart);
    return false;
  }
  A.toggleToolbar = (id) => { const v = A.opts.toolbars[id]; A.opts.toolbars[id] = A.tbShown(id) ? (v === 'auto' || v === true ? false : false) : true; if (A.opts.toolbars[id] === false && ['reviewing', 'picture', 'outlining', 'headerFooter', 'wordart'].includes(id) && v === true) A.opts.toolbars[id] = 'auto'; A.saveOpts(); A.updateToolbars(); };
  A.updateToolbars = function () {
    if (!A.toolbarsEl) return;
    for (const id of Object.keys(A.toolbarsEl)) if (id !== 'drawing') A.toolbarsEl[id].hidden = !A.tbShown(id) || A.view === 'preview';
    if (A.view === 'reading') for (const id of Object.keys(A.toolbarsEl)) if (id !== 'drawing') A.toolbarsEl[id].hidden = id !== 'standard';
    A.tbRows.forEach((r) => { r.hidden = Array.from(r.children).every((c) => c.hidden); });
    L.$('#drawrow').hidden = !A.tbShown('drawing') || A.view === 'preview' || A.view === 'reading' || A.view === 'outline';
    ui.refresh();
  };

  /* ================= task pane / side panes ================= */
  A.toggleTask = (on) => {
    const el = L.$('#taskpane');
    const show = on == null ? el.hidden : on;
    el.hidden = !show;
    A.opts.taskOpen = show;
    if (show && L.panes && !L.panes.task.current) L.panes.task.show('getting-started');
    A.saveOpts();
    LY.fitSizer();
    ui.refresh();
  };
  A.taskOpen = () => !L.$('#taskpane').hidden;
  A.toggleDocMap = (on) => {
    const show = on == null ? L.$('#docmap').hidden || A.opts.thumbs : on;
    A.opts.docMap = show; A.opts.thumbs = false;
    L.$('#docmap').hidden = !show; L.$('#split-map').hidden = !show;
    if (show && L.panes) L.panes.docmap.render();
    LY.fitSizer();
    ui.refresh();
  };
  A.toggleThumbs = (on) => {
    const show = on == null ? !A.opts.thumbs : on;
    A.opts.thumbs = show; A.opts.docMap = false;
    L.$('#docmap').hidden = !show; L.$('#split-map').hidden = !show;
    if (show && L.panes) L.panes.thumbs.render();
    LY.fitSizer();
    ui.refresh();
  };

  /* ================= document lifecycle ================= */
  /* ---------- open windows (File ▸ New and File ▸ Open add a window; Window menu switches) ---------- */
  A.docs = [];
  A.cur = -1;
  const winState = () => ({ doc: D.doc, fileName: A.fileName, saved: A.saved, fileType: A.fileType, hist: D.histState(), sel: D.getSel(), scroll: LY.scroller ? LY.scroller.scrollTop : 0, hasRevisions: A.hasRevisions });
  function stash() { if (A.cur >= 0 && A.docs[A.cur] && D.doc) A.docs[A.cur] = winState(); }
  /** is the current window an untouched blank document that an Open may reuse (as Word does)? */
  const pristine = () => !!D.doc && !D.doc.dirty && !A.saved && !D.canUndo() && /^Document\d+$/.test(A.fileName);
  A.switchTo = function (i) {
    if (i === A.cur || !A.docs[i]) return;
    stash();
    const w = A.docs[i];
    A.cur = i;
    if (LY.hfEdit) { LY.hfEdit = null; LY.root.classList.remove('hfmode'); }
    L.drawing.clearSelection();
    E.objSel = null;
    D.doc = w.doc;
    D.setHistState(w.hist);
    A.fileName = w.fileName; A.saved = w.saved; A.fileType = w.fileType; A.hasRevisions = w.hasRevisions;
    LY.render();
    if (!E.applySerialized(w.sel)) { const first = D.firstPara(w.doc.main.blocks); E.sel = { a: D.pos(first, 0), f: D.pos(first, 0) }; }
    E.restoreDom(true);
    LY.scroller.scrollTop = w.scroll || 0;
    A.updateTitle(); A.updateToolbars(); A.updateStatus();
    L.bus.emit('doc-loaded');
    ui.refresh();
    E.refocus();
  };
  A.windowItems = () => A.docs.map((w, i) => ({ label: `&${i + 1} ${(i === A.cur ? A.fileName : w.fileName).replace(/&/g, '&&')}`, checked: () => i === A.cur, run: () => A.switchTo(i) }));
  A.loadDoc = function (d, name, opts) {
    opts = opts || {};
    if (opts.newWindow || A.cur < 0) { stash(); A.docs.push(null); A.cur = A.docs.length - 1; }
    if (LY.hfEdit) { LY.hfEdit = null; LY.root.classList.remove('hfmode'); }
    L.drawing.clearSelection();
    E.objSel = null;
    D.doc = d;
    D.resetHistory();
    d.dirty = false;
    A.fileName = name || 'Document' + A.untitled;
    A.saved = !!opts.saved;
    A.fileType = opts.type || 'docx';
    A.hasRevisions = false;
    D.walk(d.main, (b) => { if (b.t === 'p' && (b.mark || b.runs.some((it) => it.rPr && (it.rPr.ins || it.rPr.del)))) { A.hasRevisions = true; return false; } return true; });
    E.readOnly = false;
    LY.render();
    const first = D.firstPara(d.main.blocks);
    E.setSel(D.pos(first, 0), null, { noScroll: true });
    LY.scroller.scrollTop = 0;
    A.updateTitle();
    A.updateToolbars();
    A.updateStatus();
    A.docs[A.cur] = winState();
    L.bus.emit('doc-loaded');
    ui.refresh();
    E.refocus();
  };
  async function confirmDiscard() {
    if (!doc() || !doc().dirty) return true;
    const r = await ui.msg(`Do you want to save the changes to ${A.fileName}?`, { icon: 'warn', buttons: ['&Yes', '&No', 'Cancel'] });
    if (r === 0) { await A.save(); return true; }
    return r === 1;
  }
  A.confirmDiscard = confirmDiscard;
  A.newDocument = async function (opts) {
    A.untitled++;
    const d = D.newDoc(opts && opts.modern ? { modern: true } : {});
    if (opts && opts.build) opts.build(d);
    /* a new document opens in its own window, as in Word; an untouched blank one is simply replaced */
    A.loadDoc(d, 'Document' + A.untitled, { newWindow: !pristine() });
  };
  A.closeDocument = async function () {
    if (!(await confirmDiscard())) return;
    if (A.docs.length <= 1) { A.docs = []; A.cur = -1; A.untitled++; A.loadDoc(D.newDoc(), 'Document' + A.untitled); return; }
    const i = A.cur;
    A.docs.splice(i, 1);
    A.cur = -1; /* nothing to stash: the closed window is gone */
    A.switchTo(Math.max(0, i - 1));
  };
  A.openDialog = async function () {
    const files = await L.pickFiles('.docx,.docm,.dotx,.dotm,.doc,.txt,.htm,.html,.rtf,.xml,application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    if (files[0]) A.openFile(files[0]);
  };
  A.openFiles = (files) => { const f = files.find((x) => /\.(docx|docm|dotx|dotm|doc|txt|html?|rtf)$/i.test(x.name)); if (f) A.openFile(f); };
  A.openFile = async function (file) {
    const reuse = pristine();
    const name = file.name.replace(/\.(docx|docm|dotx|dotm|doc|txt|html?|rtf|xml)$/i, '');
    ui.busy(true, `Opening ${file.name}...`);
    try {
      let d, warnings = [];
      if (/\.txt$/i.test(file.name)) d = A.docFromText(await L.readAsText(file));
      else if (/\.html?$/i.test(file.name) && L.htmlio) d = await L.htmlio.docFromHTML(await L.readAsText(file));
      else if (/\.rtf$/i.test(file.name) && L.rtf) d = L.rtf.read(await L.readAsArrayBuffer(file));
      else {
        const buf = await L.readAsArrayBuffer(file);
        const res = await A.readDocx(buf, file.name);
        d = res.doc; warnings = res.warnings || [];
      }
      A.loadDoc(d, name, { saved: /\.docx$/i.test(file.name), type: 'docx', newWindow: !reuse });
      A.opts.recent = [file.name].concat((A.opts.recent || []).filter((x) => x !== file.name)).slice(0, 4);
      A.saveOpts();
      ui.busy(false);
      if (d.isTemplate) A.status('Opened a template: saving creates a new document.');
      if (d.wasEncrypted) A.status('This document is protected with a password to open; saving keeps the protection (Tools ▸ Options ▸ Security).');
      if (warnings.length) ui.msg(warnings.join('\n'), { icon: 'warn' });
      if (d.settings.protect && d.settings.protect.on) A.status('This document is protected: ' + ({ readOnly: 'no changes (read only)', comments: 'comments only', tracked: 'tracked changes', forms: 'filling in forms' }[d.settings.protect.kind] || 'restricted'));
    } catch (e) {
      ui.busy(false);
      console.error(e);
      if (e && e.code === 'cancel') return;
      if (e && e.code === 'ole') ui.msg(e.message, { icon: 'warn' });
      else ui.msg(`${L.APP} can't open ${file.name}. ${e && e.message ? e.message : ''}`, { icon: 'error' });
    }
  };
  /** read .docx bytes; a password-protected file asks for its password (or takes opts.password) */
  A.readDocx = async function (buf, displayName, opts) {
    try { return await L.docx.read(buf); }
    catch (e) {
      if (!(e && e.code === 'ole' && e.encrypted && L.officeCrypto)) throw e;
      let pass = opts && opts.password != null ? opts.password : null;
      if (pass == null) { ui.busy(false); pass = await ui.password(displayName || 'Document'); }
      if (pass == null) { const c = new Error('cancelled'); c.code = 'cancel'; throw c; }
      ui.busy(true, 'Opening...');
      let plain;
      try { plain = await L.officeCrypto.decrypt(e.bytes, pass); }
      catch (err) {
        const b = new Error(err && err.code === 'badpass' ? `The password is incorrect. ${L.APP} cannot open the document.` : err && err.message === 'unsupported' ? `This document is encrypted in a way ${L.APP} cannot read.` : String(err && err.message || err));
        b.code = 'ole';
        b.badPassword = !!(err && err.code === 'badpass');
        throw b;
      }
      const res = await L.docx.read(plain.buffer.slice(plain.byteOffset, plain.byteOffset + plain.byteLength));
      res.doc.wasEncrypted = true;
      res.doc.openPassword = pass; /* kept in memory so saving keeps the protection, as Word does */
      return res;
    }
  };
  /** open raw bytes without prompting (used by drag-and-drop of blobs and by automated checks) */
  A.openBytes = async function (bytes, name, opts) {
    const buf = bytes instanceof ArrayBuffer ? bytes : bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    const res = await A.readDocx(buf, name, opts);
    A.loadDoc(res.doc, String(name || 'Document').replace(/\.(docx|docm|dotx|dotm)$/i, ''), { saved: true, type: 'docx', newWindow: !pristine() });
    return res;
  };
  A.docFromText = function (text) {
    const d = D.newDoc();
    d.main.blocks = String(text).replace(/\r\n?/g, '\n').split('\n').map((line) => {
      const p = D.para([]);
      const parts = line.split('\t');
      parts.forEach((t, i) => { if (i) p.runs.push(D.item('tab')); if (t) p.runs.push(D.text(t)); });
      return p;
    });
    if (!d.main.blocks.length) d.main.blocks.push(D.para());
    return d;
  };
  A.save = async function () {
    if (!A.saved) return A.saveAs();
    return A.exportAs(A.fileType || 'docx', A.fileName);
  };
  A.saveAs = function () { return new Promise((res) => (L.dlg && L.dlg.saveAs ? L.dlg.saveAs(async (name, type) => { if (type === 'docx' || type === 'dotx') { A.fileName = name; A.fileType = type; A.saved = true; A.updateTitle(); } await A.exportAs(type, name); res(); }) : A.exportAs('docx', A.fileName).then(res))); };
  A.docStats = function () {
    const d = doc();
    const st = D.stats(D.allParas(d));
    return { pages: LY.pages.length || 1, words: st.words, chars: st.chars, charsNoSp: st.charsNoSp, paras: st.paras, lines: Math.max(st.paras, Math.round(st.chars / 70)) };
  };
  A.exportAs = async function (type, name) {
    name = name || A.fileName;
    const d = doc();
    ui.busy(true, 'Saving...');
    try {
      if (L.fields && L.fields.beforeSave) L.fields.beforeSave();
      let blob, ext = type;
      d.props.modified = new Date().toISOString();
      d.props.lastModifiedBy = A.opts.userName || 'Quire User';
      if (type === 'docx' || type === 'dotx') {
        d.props.revision = (d.props.revision || 0) + 1;
        blob = await L.docx.write(d, { stats: A.docStats(), template: type === 'dotx' });
        /* a password to open: encrypt the package (Agile encryption, readable by Word 2010+ and LibreOffice) */
        if (d.openPassword && L.officeCrypto) blob = new Blob([await L.officeCrypto.encrypt(new Uint8Array(await blob.arrayBuffer()), d.openPassword)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      }
      else if (type === 'html') blob = await L.htmlio.webPage(d, name);
      else if (type === 'txt') blob = new Blob([L.htmlio ? L.htmlio.plainText(d) : ''], { type: 'text/plain;charset=utf-8' });
      else if (type === 'pdf') blob = await A.buildPDF();
      else if (type === 'rtf' && L.rtf) { blob = new Blob([await L.rtf.write(d)], { type: 'application/rtf' }); }
      ui.busy(false);
      if (!blob) return;
      const r = await L.saveFile(`${name}.${ext}`, blob);
      if (r === 'saved' && (type === 'docx' || type === 'dotx')) { d.dirty = false; A.status(`Saved ${name}.${ext}`); ui.refresh(); }
      return r;
    } catch (e) {
      ui.busy(false);
      console.error(e);
      ui.msg(`The file could not be saved: ${e.message || e}`, { icon: 'error' });
    }
  };

  /* ---------- PDF (print) ---------- */
  function docCSS() {
    let css = '';
    for (const sheet of document.styleSheets) {
      let rules;
      try { rules = sheet.cssRules; } catch (e) { continue; }
      for (const r of rules) if (r.selectorText && /(^|[\s,])\.(pg|p|pc|lbl|tb|img|shp|tbl|cf|nref|fdyn|ffcb|crop|pg-|frm|flt|anc|sym|raw|zw|fnsep|fnote|hbrk|sectmark|cmk|cmr|ins|del|hid|lnk|sgrp|tbx|geo|wa|chart)/.test(r.selectorText)) css += r.cssText + '\n';
    }
    return css + '.crop{display:none}.pg-hdr,.pg-ftr{opacity:1!important}.cmk,.cmr{background:none!important}td.csel{background-image:none!important}.p.chg::before{display:none}';
  }
  async function mediaDataMap() {
    const map = new Map();
    for (const [, m] of L.media.all()) { try { map.set(m.url, await L.blobToDataURL(m.view || m.blob)); } catch (e) { /* skip */ } }
    return map;
  }
  A.rasterizePage = async function (pgEl, scale, css, map) {
    const w = pgEl.offsetWidth, hh = pgEl.offsetHeight;
    const clone = pgEl.cloneNode(true);
    clone.style.margin = '0'; clone.style.boxShadow = 'none';
    for (const e of clone.querySelectorAll('[contenteditable]')) e.removeAttribute('contenteditable');
    let xhtml = new XMLSerializer().serializeToString(clone);
    for (const [u, dd] of map) xhtml = xhtml.split(u).join(dd);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${hh}"><foreignObject x="0" y="0" width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;width:${w}px;height:${hh}px;color:#000;font-size:16px;line-height:normal"><style>${css.replace(/</g, '\\3c ')}</style>${xhtml}</div></foreignObject></svg>`;
    const img = await L.loadImage('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg));
    const c = document.createElement('canvas');
    c.width = Math.round(w * scale); c.height = Math.round(hh * scale);
    const g = c.getContext('2d');
    g.fillStyle = '#fff';
    g.fillRect(0, 0, c.width, c.height);
    g.scale(scale, scale);
    g.drawImage(img, 0, 0, w, hh);
    return c;
  };
  A.buildPDF = async function (range) {
    const prevView = A.view;
    if (LY.view !== 'print') { LY.view = 'print'; LY.render(); }
    const map = await mediaDataMap();
    const css = docCSS();
    const pages = [];
    const list = LY.pages.filter((p, i) => !range || range.includes(i + 1));
    const sel = window.getSelection(); sel.removeAllRanges();
    let n = 0;
    for (const p of list) {
      ui.busy(true, `Creating PDF... (page ${++n} of ${list.length})`);
      const c = await A.rasterizePage(p.el, 2.2, css, map);
      const jpeg = new Uint8Array(await (await new Promise((res, rej) => c.toBlob((b) => (b ? res(b) : rej(new Error('This browser blocked the page image.'))), 'image/jpeg', 0.92))).arrayBuffer());
      pages.push({ jpeg, iw: c.width, ih: c.height, pw: p.sect.pgW, ph: p.sect.pgH });
    }
    if (prevView !== A.view || LY.view !== (prevView === 'reading' || prevView === 'preview' ? 'print' : prevView)) { LY.view = prevView === 'reading' || prevView === 'preview' ? 'print' : prevView; LY.render(); }
    E.restoreDom(true);
    return L.buildPDF(pages, { title: doc().props.title || A.fileName, author: doc().props.creator || '' });
  };
  A.printPDF = async function () {
    try { const blob = await A.buildPDF(); ui.busy(false); await L.saveFile(`${A.fileName}.pdf`, blob); }
    catch (e) { ui.busy(false); ui.msg('The PDF could not be created: ' + (e.message || e), { icon: 'error' }); }
  };

  /* ================= headers & footers ================= */
  A.editHeaderFooter = function (kind, pageIdx) {
    if (A.view !== 'print') A.setView('print');
    const pg = LY.pages[pageIdx == null ? Math.max(0, LY.pageOf(E.sel.f)) : pageIdx] || LY.pages[0];
    const d = doc();
    const secs = D.sections(d);
    kind = kind || 'hdr';
    const ty = pg.type;
    /* create the story when the section has none */
    let story = LY.hfFor(secs, pg.si, kind, ty);
    if (!story) {
      E.edit(kind === 'hdr' ? 'Header' : 'Footer', () => {
        const id = (kind === 'hdr' ? 'h' : 'f') + D.nid();
        D.touchKey(d, 'hf');
        d.hf[id] = { kind, id, blocks: [D.para([], { style: kind === 'hdr' ? 'Header' : 'Footer' })] };
        const s = secs[pg.si].sect;
        const holder = secs[pg.si].endPara || d;
        D.touchKey(holder, 'sect');
        const ss = secs[pg.si].endPara ? secs[pg.si].endPara.sect : d.sect;
        ss.refs = ss.refs || { hdr: {}, ftr: {} };
        ss.refs[kind] = Object.assign({}, ss.refs[kind], { [ty]: id });
        void s;
        return E.sel;
      });
      story = LY.hfFor(D.sections(d), pg.si, kind, ty);
    }
    LY.hfEdit = { kind, page: pg.idx };
    LY.root.classList.add('hfmode');
    LY.render();
    labelHF();
    D.reindex(d);
    const p = story && story.paras && story.paras[0] ? story.paras[0] : story && D.firstPara(story.blocks);
    if (p) E.setSel(D.pos(p, 0));
    A.updateToolbars();
    ui.refresh();
  };
  function labelHF() {
    const d = doc();
    const secs = D.sections(d);
    for (const pg of LY.pages) {
      const multi = secs.length > 1;
      const tyName = { first: 'First Page ', even: 'Even Page ', default: doc().settings.evenOdd ? 'Odd Page ' : '' }[pg.type] || '';
      pg.hdr.dataset.label = `${tyName}Header${multi ? ' -Section ' + (pg.si + 1) + '-' : ''}`;
      pg.ftr.dataset.label = `${tyName}Footer${multi ? ' -Section ' + (pg.si + 1) + '-' : ''}`;
      const refH = LY.hfRefFor(secs, pg.si, 'hdr', pg.type), refF = LY.hfRefFor(secs, pg.si, 'ftr', pg.type);
      if (refH && refH.si < pg.si) pg.hdr.dataset.same = '1'; else delete pg.hdr.dataset.same;
      if (refF && refF.si < pg.si) pg.ftr.dataset.same = '1'; else delete pg.ftr.dataset.same;
    }
  }
  L.bus.on('layout-done', () => { if (LY.hfEdit) labelHF(); });
  A.closeHeaderFooter = function (silent) {
    LY.hfEdit = null;
    LY.root.classList.remove('hfmode');
    if (!silent) {
      LY.render();
      const p = D.firstPara(doc().main.blocks);
      const pg = LY.pages[0];
      void pg;
      E.setSel(D.pos(p, 0));
    }
    A.updateToolbars();
  };

  /* ================= misc helpers used by commands ================= */
  A.gotoPos = function (pos, sel2) {
    E.setSel(pos, sel2 || pos);
    E.focus();
  };
  A.gotoBookmark = function (name) {
    const d = doc();
    for (const st of D.stories(d)) {
      D.reindex(d);
      for (const p of st.paras || []) {
        let o = 0;
        for (const it of p.runs) { if (it.t === 'bs' && it.name === name) { A.gotoPos(D.pos(p, o)); return true; } o += D.ilen(it); }
      }
    }
    A.status(`Bookmark "${name}" not found.`);
    return false;
  };
  A.gotoNote = function (key) {
    const [kind, id] = key.split(':');
    const d = doc();
    const st = (kind === 'fn' ? d.fn : d.en)[id];
    if (!st) return;
    D.reindex(d);
    const p = D.firstPara(st.blocks);
    if (p) A.gotoPos(D.pos(p, D.plen(p)));
  };
  A.contextMenu = function (e) {
    e.preventDefault();
    const inTable = L.tables && L.tables.inTable();
    const items = ['cut', 'copy', 'paste', '-'];
    if (E.objSel) items.push({ label: 'Text W&rapping', sub: () => A.wrapItems() }, 'formatObject', '-');
    const ln = e.target.closest && e.target.closest('[data-link]');
    if (ln) items.push('hyperlink', { label: '&Open Hyperlink', run: () => { const a = document.createElement('a'); a.href = ln.dataset.link; a.target = '_blank'; a.rel = 'noopener'; a.click(); } }, { label: '&Remove Hyperlink', run: () => ui.exec('removeHyperlink') }, '-');
    if (inTable) items.push({ label: '&Insert', sub: ['colLeft', 'colRight', 'rowAbove', 'rowBelow', 'insertCells'] }, { label: '&Delete', sub: ['deleteTable', 'deleteCols', 'deleteRows', 'deleteCells'] }, 'mergeCells', 'splitCells', '-', 'distRows', 'distCols', '-', 'bordersDlg', { label: 'Cell Alignment', sub: ['cellTop', 'cellMiddle', 'cellBottom'] }, 'tableAutoFormat', 'tableProps', '-');
    const f = L.fields && L.fields.at(E.sel && E.sel.f);
    if (f) items.push('updateField', 'editField', 'toggleCodes', '-');
    if (L.review && L.review.revisionAt && L.review.revisionAt()) items.push('acceptChange', 'rejectChange', '-');
    const cm = L.review && L.review.commentAt && L.review.commentAt();
    if (cm) items.push({ label: '&Edit Comment', run: () => L.review.editComment(cm) }, { label: 'Delete Co&mment', run: () => L.review.deleteComment(cm) }, '-');
    items.push('fontDlg', 'paragraphDlg', 'bulletsDlg', '-', 'hyperlink', 'research');
    /* on a misspelled word or a grammar problem Word shows only the proofing menu */
    if (!E.objSel && L.spell && L.spell.suggestionsAt) { const s = L.spell.suggestionsAt(e); if (s && s.length) { ui.openMenu(s.concat(['-', 'cut', 'copy', 'paste']), { x: e.clientX, y: e.clientY }); return; } }
    ui.openMenu(items, { x: e.clientX, y: e.clientY });
  };
  A.browseKind = 'page';
  A.browseMenu = function (r) {
    const kinds = [['field', 'Browse by Field'], ['endnote', 'Browse by Endnote'], ['footnote', 'Browse by Footnote'], ['comment', 'Browse by Comment'], ['section', 'Browse by Section'], ['page', 'Browse by Page'], ['goto', 'Go To...'], ['find', 'Find...'], ['edit', 'Browse by Edits'], ['heading', 'Browse by Heading'], ['graphic', 'Browse by Graphic'], ['table', 'Browse by Table']];
    ui.openMenu(kinds.map(([k, n]) => ({ label: n, checked: () => A.browseKind === k, run: () => { if (k === 'goto') ui.exec('goTo'); else if (k === 'find') ui.exec('find'); else { A.browseKind = k; A.status(n); } } })), r);
  };
  A.browse = function (dir) {
    const d = doc();
    const k = A.browseKind;
    if (k === 'page') {
      const i = Math.max(0, LY.pageOf(E.sel.f)) + dir;
      const pg = LY.pages[L.clamp(i, 0, LY.pages.length - 1)];
      if (pg && pg.start && pg.start.block) { const p = pg.start.block.t === 'p' ? pg.start.block : D.firstPara([pg.start.block]); if (p) A.gotoPos(D.pos(p, pg.start.block.t === 'p' ? pg.start.sub || 0 : 0)); pg.el.scrollIntoView({ block: 'start' }); }
      return;
    }
    if (L.panes && L.panes.browseTo) L.panes.browseTo(k, dir);
    void d;
  };

  /* ================= insert pictures ================= */
  A.insertPictureDialog = async function () {
    const files = await L.pickFiles('image/*,.emf,.wmf', true);
    if (files.length) A.insertPictureFiles(files);
  };
  A.insertPictureFiles = async function (files) {
    const items = [];
    for (const f of files) {
      const buf = new Uint8Array(await L.readAsArrayBuffer(f));
      const ext = (f.name.split('.').pop() || '').toLowerCase();
      let view = null;
      if (ext === 'emf' || ext === 'wmf') view = await L.metafile.toPNG(buf, ext, 1600);
      const blob = new Blob([buf], { type: f.type || L.extToMime(ext) });
      const id = L.media.add(blob, f.name, view);
      let w = 200, hh = 150;
      try {
        const img = await L.loadImage(L.media.url(id));
        w = img.naturalWidth * 0.75; hh = img.naturalHeight * 0.75;
      } catch (e) { /* keep default */ }
      /* fit to the text column */
      const s = A.curSect().sect;
      const maxW = s.pgW - s.ml - s.mr - (s.gutter || 0);
      if (w > maxW) { hh *= maxW / w; w = maxW; }
      items.push(D.item('img', { media: id, w: L.round(w, 2), h: L.round(hh, 2), name: f.name, alt: '' }));
    }
    if (!items.length) return;
    E.edit('Insert Picture', () => { let pos = E.deleteSelection(); for (const it of items) pos = O.insertItem(pos, it); return pos; });
  };

  /* ================= keyboard ================= */
  const keyName = (e) => {
    let k = e.key;
    if (k === ' ') k = 'Space';
    if (k.length === 1) k = k.toUpperCase();
    if (e.code && /^Digit\d$/.test(e.code) && (e.ctrlKey || e.altKey)) k = e.code.slice(5);
    if (e.code === 'Equal' && (e.ctrlKey || e.metaKey)) k = '=';
    if (e.code === 'Minus' && (e.ctrlKey || e.metaKey)) k = '-';
    if (e.code === 'BracketLeft' && (e.ctrlKey || e.metaKey)) k = '[';
    if (e.code === 'BracketRight' && (e.ctrlKey || e.metaKey)) k = ']';
    if (e.code === 'Comma' && (e.ctrlKey || e.metaKey) && e.shiftKey) k = '<';
    if (e.code === 'Period' && (e.ctrlKey || e.metaKey) && e.shiftKey) k = '>';
    if (e.code === 'Period' && (e.ctrlKey || e.metaKey) && e.altKey) k = '.';
    if (e.code === 'Digit8' && (e.ctrlKey || e.metaKey) && e.shiftKey) k = '8';
    return (e.ctrlKey || e.metaKey ? 'Ctrl+' : '') + (e.altKey ? 'Alt+' : '') + (e.shiftKey ? 'Shift+' : '') + k;
  };
  A.SHORTCUTS = {
    'Ctrl+N': 'new', 'Ctrl+O': 'open', 'Ctrl+S': 'save', 'Shift+F12': 'save', 'F12': 'saveAs', 'Ctrl+P': 'print', 'Ctrl+F2': 'printPreview', 'Ctrl+Alt+I': 'printPreview', 'Ctrl+W': 'close', 'Ctrl+F4': 'close',
    'Ctrl+Z': 'undo', 'Alt+Backspace': 'undo', 'Ctrl+Y': 'redo', 'F4': 'redo', 'Ctrl+A': 'selectAll', 'Ctrl+F': 'find', 'Ctrl+H': 'replace', 'Ctrl+G': 'goTo', 'F5': 'goTo',
    'Ctrl+B': 'bold', 'Ctrl+I': 'italic', 'Ctrl+U': 'underline', 'Ctrl+Shift+D': 'dblUnderline', 'Ctrl+Shift+W': 'wordUnderline', 'Ctrl+=': 'subscript', 'Ctrl+Shift+=': 'superscript', 'Ctrl+Shift++': 'superscript',
    'Ctrl+Shift+A': 'allCaps', 'Ctrl+Shift+K': 'smallCaps', 'Ctrl+Shift+H': 'hiddenText', 'Ctrl+Space': 'resetChar', 'Ctrl+Shift+Z': 'resetChar', 'Ctrl+Q': 'resetPara',
    'Ctrl+L': 'alignLeft', 'Ctrl+E': 'alignCenter', 'Ctrl+R': 'alignRight', 'Ctrl+J': 'justify', 'Ctrl+1': 'single', 'Ctrl+2': 'double', 'Ctrl+5': 'onePointFive', 'Ctrl+0': 'spaceBefore',
    'Ctrl+M': 'incIndent', 'Ctrl+Shift+M': 'decIndent', 'Ctrl+T': 'hangIndent', 'Ctrl+Shift+T': 'unhangIndent',
    'Ctrl+Shift+N': 'styleNormal', 'Ctrl+Alt+1': 'styleH1', 'Ctrl+Alt+2': 'styleH2', 'Ctrl+Alt+3': 'styleH3', 'Ctrl+Shift+L': 'styleListBullet',
    'Ctrl+]': 'growFont1', 'Ctrl+[': 'shrinkFont1', 'Ctrl+Shift+>': 'growFont', 'Ctrl+Shift+<': 'shrinkFont', 'Ctrl+Shift+.': 'growFont', 'Ctrl+Shift+,': 'shrinkFont',
    'Ctrl+D': 'fontDlg', 'Ctrl+Shift+F': 'focusFont', 'Ctrl+Shift+P': 'focusSize', 'Ctrl+Shift+S': 'focusStyle', 'Ctrl+Shift+C': 'copyFormat', 'Ctrl+Shift+V': 'pasteFormat', 'Shift+F3': 'changeCaseCycle',
    'Ctrl+K': 'hyperlink', 'Ctrl+Alt+F': 'insertFootnoteNow', 'Ctrl+Alt+D': 'insertEndnoteNow', 'Ctrl+Alt+M': 'insertComment', 'Ctrl+Shift+E': 'trackChanges', 'Ctrl+Shift+F5': 'bookmark',
    'Alt+Shift+D': 'insertDateField', 'Alt+Shift+T': 'insertTimeField', 'Alt+Shift+P': 'insertPageField', 'Ctrl+Alt+C': 'symCopyright', 'Ctrl+Alt+R': 'symRegistered', 'Ctrl+Alt+T': 'symTrademark', 'Ctrl+Alt+.': 'symEllipsis',
    'Ctrl+-': 'optHyphen', 'Ctrl+Shift+-': 'nbHyphen', 'Ctrl+Alt+-': 'emDash', 'Ctrl+Alt+Shift+-': 'emDash',
    'F7': 'spelling', 'Shift+F7': 'thesaurus', 'F9': 'updateField', 'Alt+F9': 'toggleCodes', 'Shift+F9': 'toggleCodes', 'Ctrl+F9': 'insertFieldBraces', 'Ctrl+Shift+F9': 'unlinkField', 'Ctrl+F11': 'lockField', 'Ctrl+Shift+F11': 'unlockField',
    'F1': 'help', 'Ctrl+F1': 'taskPane', 'Shift+F1': 'revealFormatting', 'Ctrl+Shift+8': 'showMarks', 'Ctrl+Shift+*': 'showMarks', 'Ctrl+*': 'showMarks',
    'Ctrl+Alt+P': 'viewPrint', 'Ctrl+Alt+N': 'viewNormal', 'Ctrl+Alt+O': 'viewOutline', 'Alt+Ctrl+S': 'splitWindow', 'Ctrl+Alt+Shift+S': 'stylesPane',
    'Ctrl+Shift+F12': 'print', 'Alt+F8': 'macros', 'F8': 'extendSel', 'Shift+F5': 'goBack', 'Alt+Ctrl+Z': 'goBack',
  };
  A.shortcut = function (e) {
    const k = keyName(e);
    const cmd = A.SHORTCUTS[k];
    if (!cmd) return false;
    /* let the browser handle clipboard keys inside the editor */
    ui.exec(cmd);
    return true;
  };
  document.addEventListener('keydown', (e) => {
    if (ui.dialogOpen()) { ui.dialogKey(e); return; }
    if (ui.menuOpen()) { if (ui.menuKey(e)) { e.preventDefault(); e.stopPropagation(); } return; }
    if (e.key === 'F10' && !e.shiftKey) { e.preventDefault(); A.menuBar.openIndex(0); return; }
    if (e.altKey && !e.ctrlKey && !e.shiftKey && e.key.length === 1 && /[a-z]/i.test(e.key)) { if (A.menuBar.openByKey(e.key)) { e.preventDefault(); return; } }
    const tgt = e.target;
    const inEditor = tgt && tgt.closest && (tgt.closest('#pages') || tgt.closest('.ed-host'));
    const inInput = tgt && (tgt.tagName === 'INPUT' || tgt.tagName === 'TEXTAREA' || tgt.tagName === 'SELECT');
    if (e.key === 'Escape') {
      if (A.highlighter) { A.highlighter = null; L.$('#scroller').classList.remove('painting'); A.status(''); return; }
      if (A.painter) { A.painter = null; L.$('#scroller').classList.remove('painting'); return; }
      if (A.view === 'preview' || A.view === 'reading') { A.setView(A.prevView || 'print'); return; }
      if (L.$('#app').classList.contains('fullscreen')) { ui.exec('fullScreen'); return; }
      if (LY.hfEdit) { A.closeHeaderFooter(); return; }
    }
    if (inEditor) return; /* the editor's own keydown handler calls A.shortcut */
    if (inInput) { if (e.key === 'Escape') { tgt.blur(); E.refocus(); } return; }
    const k = keyName(e);
    const cmd = A.SHORTCUTS[k];
    if (cmd && !/^(cut|copy|paste)$/.test(cmd)) { e.preventDefault(); ui.exec(cmd); }
  });

  /* ================= drag & drop of files ================= */
  function setupDrop() {
    let hint = null, depth = 0;
    const app = L.$('#app');
    app.addEventListener('dragenter', (e) => { if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return; depth++; if (!hint) { hint = h('div', { class: 'drop-hint', text: 'Drop a document to open it, or pictures to insert them' }); document.body.appendChild(hint); } });
    app.addEventListener('dragleave', () => { depth--; if (depth <= 0 && hint) { hint.remove(); hint = null; depth = 0; } });
    app.addEventListener('dragover', (e) => { if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) e.preventDefault(); });
    app.addEventListener('drop', (e) => {
      if (hint) { hint.remove(); hint = null; } depth = 0;
      if (e.defaultPrevented) return;
      const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []);
      if (!files.length) return;
      e.preventDefault();
      if (files.some((f) => /\.(docx|docm|dotx|dotm|doc|txt|html?|rtf)$/i.test(f.name))) { A.openFiles(files); return; }
      const imgs = files.filter((f) => /^image\//.test(f.type) || /\.(emf|wmf)$/i.test(f.name));
      if (imgs.length) A.insertPictureFiles(imgs);
    });
  }

  /* ================= splitter (document map width) ================= */
  function setupSplitters() {
    const sp = L.$('#split-map');
    sp.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      sp.setPointerCapture(e.pointerId);
      const x0 = e.clientX, w0 = L.$('#docmap').offsetWidth;
      const mv = (ev) => { L.$('#docmap').style.width = L.clamp(w0 + ev.clientX - x0, 100, 480) + 'px'; };
      const up = () => { sp.removeEventListener('pointermove', mv); sp.removeEventListener('pointerup', up); LY.fitSizer(); };
      sp.addEventListener('pointermove', mv); sp.addEventListener('pointerup', up);
    });
  }

  /* ================= recovery (per-browser convenience) ================= */
  const autosave = L.debounce(async () => {
    const d = doc();
    if (!d || !d.dirty) return;
    try {
      const blob = await L.docx.write(d, { stats: A.docStats() });
      if (blob.size > 4e6) { L.store.del('recover'); return; }
      const data = await L.blobToDataURL(blob);
      L.store.set('recover', { at: Date.now(), name: A.fileName, data });
    } catch (e) { /* storage is optional */ }
  }, 15000);
  async function offerRecovery() {
    const snap = L.store.get('recover', null);
    if (!snap || !snap.data) return false;
    const r = await ui.msg(`Quire found an unsaved version of "${snap.name}" from ${new Date(snap.at).toLocaleString()}.\nDo you want to restore it?`, { title: 'Document Recovery', icon: 'question', buttons: ['&Restore', '&Discard'] });
    if (r !== 0) { L.store.del('recover'); return false; }
    try {
      const bytes = await (await fetch(snap.data)).arrayBuffer();
      const res = await L.docx.read(bytes);
      A.loadDoc(res.doc, snap.name + ' (Recovered)');
      doc().dirty = true;
      return true;
    } catch (e) { L.store.del('recover'); return false; }
  }

  /* ================= wiring ================= */
  function wire() {
    L.bus.on('sel', () => { A.updateStatus(); ui.refresh(); if (L.rulers) L.rulers.drawSoon(); if (A.highlighter && !E.collapsed() && !A.hlBusy) { A.hlBusy = true; setTimeout(() => { if (!E.collapsed()) E.formatRun({ hl: A.highlighter === 'none' ? undefined : A.highlighter }, 'Highlight'); A.hlBusy = false; }, 0); } if (A.painter && !E.collapsed()) A.applyPainter(); });
    L.bus.on('objsel', () => { A.updateToolbars(); ui.refresh(); });
    L.bus.on('doc-changed', () => { autosave(); A.updateStatus(); ui.refresh(); if (doc().settings.track || Object.keys(doc().comments).length) A.updateToolbars(); });
    L.bus.on('layout-done', () => { A.updateStatus(); if (L.rulers) L.rulers.drawSoon(); if (L.drawing) L.drawing.placeHandles(); E.placeHandles(); });
    L.bus.on('zoom', () => { ui.refresh(); if (L.rulers) L.rulers.draw(); if (L.drawing) L.drawing.placeHandles(); E.placeHandles(); });
    LY.scroller.addEventListener('scroll', () => { if (L.rulers) L.rulers.drawSoon(); });
    window.addEventListener('beforeunload', (e) => { const any = (doc() && doc().dirty) || A.docs.some((w, i) => i !== A.cur && w && w.doc && w.doc.dirty); if (any) { e.preventDefault(); e.returnValue = ''; } });
    window.addEventListener('resize', L.debounce(() => { if (A.view === 'web') LY.render(); else LY.fitSizer(); if (L.rulers) L.rulers.draw(); }, 200));
    LY.root.addEventListener('mouseup', () => { if (A.painter && !E.collapsed()) A.applyPainter(); });
    document.addEventListener('keydown', (e) => { LY.root.classList.toggle('ctrl', e.ctrlKey || e.metaKey); });
    document.addEventListener('keyup', (e) => { LY.root.classList.toggle('ctrl', e.ctrlKey || e.metaKey); });
  }

  /* ================= format painter ================= */
  A.startPainter = function (sticky) {
    if (!E.sel) return;
    A.painter = { rPr: L.clone(E.curRun()), pPr: L.clone(E.sel.f.p.pPr), sticky: !!sticky, whole: !E.collapsed() && E.range()[0].o === 0 };
    delete A.painter.rPr.link; delete A.painter.rPr.ins; delete A.painter.rPr.del;
    L.$('#scroller').classList.add('painting');
    A.status('Select the text to apply the formatting to.');
    ui.refresh();
  };
  A.applyPainter = function () {
    const pt = A.painter;
    if (!pt || E.collapsed()) return;
    const [a, b] = E.range();
    const keep = E.sel;
    const src = pt.rPr;
    E.edit('Format Painter', () => {
      O.formatRange(a, b, () => { const r = {}; for (const k of ['font', 'sz', 'b', 'i', 'u', 'uColor', 'strike', 'dstrike', 'color', 'hl', 'caps', 'smallCaps', 'vert', 'spacing', 'shd', 'style']) if (src[k] !== undefined) r[k] = L.clone(src[k]); return r; });
      if (a.p !== b.p || (a.o === 0 && b.o >= D.plen(b.p))) O.setParaProps(E.selectedParas(), () => L.clone(pt.pPr));
      return keep;
    });
    if (!pt.sticky) { A.painter = null; L.$('#scroller').classList.remove('painting'); }
    ui.refresh();
  };

  /* ================= sample document ================= */
  A.buildSample = function () {
    const d = D.newDoc();
    const P = (text, pPr, rPr) => D.para(text ? [D.text(text, rPr || {})] : [], pPr || {});
    const bl = d.main.blocks;
    bl.length = 0;
    d.props.title = 'Welcome to Quire 2003';
    d.props.creator = 'Quire';
    bl.push(P('Welcome to Quire 2003', { style: 'Title' }));
    bl.push(P('A word processor in the spirit of Word 2003 — it opens and saves real .docx files.', { style: 'Subtitle' }, { i: true }));
    bl.push(P('Getting around', { style: 'Heading1' }));
    bl.push(D.para([D.text('Type anywhere on this page. Everything works the way you remember: '), D.text('bold', { b: true }), D.text(', '), D.text('italic', { i: true }), D.text(', '), D.text('underline', { u: 'single' }), D.text(', '), D.text('highlighting', { hl: 'yellow' }), D.text(', '), D.text('colors', { color: 'C00000' }), D.text(', styles, bullets, numbering, tables, headers and footers, footnotes, comments, tracked changes, fields and more.')], { jc: 'both' }));
    bl.push(P('Open a document with File ▸ Open, or drop a .docx onto the window. Save with File ▸ Save (Ctrl+S); File ▸ Save As offers Word Document, Word Template, Web Page, Plain Text and PDF.', { num: { id: '1', lvl: 0 } }));
    bl.push(P('Switch views with the buttons in the lower-left corner: Normal, Web Layout, Print Layout, Outline and Reading Layout.', { num: { id: '1', lvl: 0 } }));
    bl.push(P('Right-click for the shortcut menu; the task pane (Ctrl+F1) holds Styles and Formatting, Reveal Formatting, Clip Art and the Mail Merge wizard.', { num: { id: '1', lvl: 0 } }));
    bl.push(P('A small table', { style: 'Heading2' }));
    const t = D.simpleTable(4, 3, 432);
    const rows = [['Feature', 'Menu', 'Shortcut'], ['Find and Replace', 'Edit', 'Ctrl+H'], ['Insert footnote', 'Insert ▸ Reference', 'Ctrl+Alt+F'], ['Track changes', 'Tools', 'Ctrl+Shift+E']];
    rows.forEach((r, ri) => r.forEach((v, ci) => { t.rows[ri].cells[ci].blocks[0].runs = [D.text(v, ri === 0 ? { b: true } : {})]; if (ri === 0) t.rows[ri].cells[ci].tcPr.shd = { val: 'clear', fill: 'D9E2F3' }; }));
    t.rows[0].trPr.header = true;
    bl.push(t);
    const fnId = '1';
    d.fn[fnId] = { kind: 'fn', id: fnId, blocks: [D.para([D.item('fn', { id: fnId, self: true }, { style: 'FootnoteReference' }), D.text(' Footnotes are numbered automatically and placed at the bottom of the page.')], { style: 'FootnoteText' })] };
    bl.push(D.para([D.text('Tabs, leaders and fields'), D.item('fn', { id: fnId }, { style: 'FootnoteReference' })], { style: 'Heading2' }));
    bl.push(D.para([D.text('Chapter One'), D.item('tab'), D.text('7')], { tabs: [{ pos: 432, al: 'right', leader: 'dot' }] }));
    bl.push(D.para([D.text('Chapter Two'), D.item('tab'), D.text('19')], { tabs: [{ pos: 432, al: 'right', leader: 'dot' }] }));
    const fp = D.para([D.text('This document was last printed on ')]);
    fp.runs.push(D.item('fb', { fid: 'f1', instr: ' DATE \\@ "MMMM d, yyyy" ' }), D.item('fs', { fid: 'f1' }), D.text(L.fmtDate(new Date(), 'datetime4')), D.item('fe', { fid: 'f1' }), D.text('.'));
    bl.push(fp);
    bl.push(P(''));
    /* header & footer with page numbers */
    d.hf.h1 = { kind: 'hdr', id: 'h1', blocks: [D.para([D.text('Quire 2003'), D.item('tab'), D.item('tab'), D.text('Sample document')], { style: 'Header' })] };
    const ftr = D.para([D.item('tab'), D.text('Page ')], { style: 'Footer' });
    ftr.runs.push(D.item('fb', { fid: 'f2', instr: ' PAGE ' }), D.item('fs', { fid: 'f2' }), D.text('1'), D.item('fe', { fid: 'f2' }), D.text(' of '), D.item('fb', { fid: 'f3', instr: ' NUMPAGES ' }), D.item('fs', { fid: 'f3' }), D.text('1'), D.item('fe', { fid: 'f3' }));
    d.hf.f1 = { kind: 'ftr', id: 'f1', blocks: [ftr] };
    d.sect.refs = { hdr: { default: 'h1' }, ftr: { default: 'f1' } };
    return d;
  };

  /* ================= init ================= */
  A.init = async function () {
    A.loadOpts();
    L.$('#appicon').innerHTML = L.icons.app ? L.icons.app(16) : '';
    buildMenus();
    buildToolbars();
    buildStatus();
    LY.mount(L.$('#scroller'));
    E.init();
    if (L.rulers) L.rulers.mount(L.$('#hruler'), L.$('#vruler'), L.$('#rcorner'));
    if (L.panes) { L.panes.task.mount(L.$('#taskpane')); L.panes.docmap.mount(L.$('#docmap')); if (L.panes.reviewing) L.panes.reviewing.mount(L.$('#revpane')); }
    setupDrop();
    setupSplitters();
    wire();
    L.bus.emit('app-ready');
    const narrow = window.innerWidth < 820;
    if (narrow) A.opts.taskOpen = false;
    L.$('#taskpane').hidden = !A.opts.taskOpen;
    LY.zoom = narrow ? Math.max(0.4, (window.innerWidth - 40) / 816) : A.opts.zoom || 1;
    A.loadDoc(A.buildSample(), 'Document1');
    A.view = 'x';
    A.setView(A.opts.view && ['print', 'normal', 'web', 'outline'].includes(A.opts.view) ? A.opts.view : 'print');
    if (L.panes && A.opts.taskOpen) L.panes.task.show('getting-started');
    A.applyOpts();
    if (A.opts.docMap) A.toggleDocMap(true);
    setTimeout(() => offerRecovery(), 700);
    ui.refresh();
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => A.init());
  else A.init();
})();
