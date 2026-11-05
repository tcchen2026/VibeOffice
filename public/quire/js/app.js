/* Quire — application controller: window chrome, views, document lifecycle, keyboard, wiring. */
(function () {
  'use strict';
  const L = window.L, D = L.D, LY = L.layout, E = L.ed, O = L.O;
  const { h } = L;
  const ui = L.ui;

  /* original application icon: a folded quire of pages with a ribbon marker */
  L.icons.app = function (size) {
    const sz = size || 16;
    return `<svg class="appico" width="${sz}" height="${sz}" viewBox="0 0 16 16" aria-hidden="true"><path d="M3.5 1.5h7l3 3v10h-10z" fill="#fff" stroke="#1c4a9c"/><path d="M10.5 1.5v3h3" fill="#d6e2f7" stroke="#1c4a9c"/><path d="M1.5 4.5h7v9h-7z" fill="#2b5fbf" stroke="#173f86"/><circle cx="5" cy="8.8" r="2.2" fill="none" stroke="#fff" stroke-width="1.1"/><path d="M5.7 9.8l1.6 1.9" stroke="#fff" stroke-width="1.2"/><path d="M10 7.5h2.5M10 9.5h2.5M10 11.5h2.5" stroke="#7a8aa6"/></svg>`;
  };
  /* the editor around commands, toolbar boxes and dialogs (common/ui.js) */
  ui.hooks = {
    beforeExec() { if (L.ed && L.ed.syncSel && !L.ed.objSel) L.ed.syncSel(); },
    refocus() { if (L.ed) L.ed.refocus(); },
    commitEdit() { if (L.ed && L.ed.textSave) L.ed.textSave(); },
    ownsText: (el) => !!(L.layout && L.layout.root && L.layout.root.contains(el)),
  };

  const DEFAULT_OPTS = {
    statusBar: true, ruler: true, vruler: true, taskOpen: true, startupPane: true, paneFollows: true, docMap: false, thumbs: false,
    toolbars: { standard: true, formatting: true, drawing: true, tables: false, reviewing: 'auto', picture: 'auto', outlining: 'auto', headerFooter: 'auto', mailMerge: false, wordart: 'auto' },
    showMarks: false, marks: { tabs: false, spaces: false, para: false, hidden: false, all: false }, textBoundaries: true, fieldShading: 'selected', bookmarks: false,
    smartQuotes: true, autocorrect: true, autoFormat: true, capSentence: true, autoBullets: true, autoNumbers: true, autoBorders: true, autoHeadings: false, autoUrl: true, ordinals: true, fractions: true, dashes: true,
    spell: true, unit: (L.locale && L.locale.get().unit) || 'in', recent: [], zoom: 1, view: 'print', balloons: true, pasteOptions: true, dragDrop: true, smartCut: true, typingReplaces: true,
    userName: 'Quire User', userInitials: 'QU', markupMode: 'markup',
  };
  const A = (L.app = Object.assign(L.app || {}, {
    view: 'print', fileName: 'Document1', fileType: 'docx', saved: false, untitled: 1, opts: L.clone(DEFAULT_OPTS),
    lastFontColor: 'FF0000', lastHighlight: 'yellow', lastBorder: 'outside', lastShading: null,
  }));
  const doc = () => D.doc;

  /* ================= options ================= */
  /* the options that follow the UI's language, with their English values (common/locales.js) */
  const EN_OPTS = { unit: 'in' };
  A.loadOpts = () => { A.opts = L.deepMerge(L.clone(DEFAULT_OPTS), L.locale.savedOpts(L.store.get('opts', {}), EN_OPTS)); D.unit = A.opts.unit || 'in'; };
  A.saveOpts = () => { L.store.set('opts', L.locale.optsToSave(A.opts, DEFAULT_OPTS, EN_OPTS)); };
  A.applyOpts = () => {
    document.body.classList.toggle('bluebg', !!A.opts.blueBg);
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
    const recentItems = () => (window.VO ? VO.recentMenuItems(4) : []);
    const menus = [
      { label: '&File', items: () => ['new', 'open', 'close', '-', 'save', 'saveAs', 'saveWeb', 'fileSearch', '-', 'webPreview', '-', 'pageSetup', 'printPreview', 'print', 'exportPDF', '-', { label: 'Sen&d To', sub: ['sendMail'] }, 'properties', ...(L.file.native() ? ['versions'] : []), ...((recentItemsList) => (recentItemsList.length ? ['-', ...recentItemsList] : []))(recentItems()), '-', 'exitApp'] },
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
      { label: '&Help|menu bar', items: ['help', 'showAssistant', '-', 'gettingStarted', 'keyboardHelp', '-', 'about'] },
    ];
    const row = L.$('#menurow');
    A.menuBar = ui.menuBar(row, menus);
    row.appendChild(ui.searchBox());
  }
  function backgroundMenu() {
    return [
      { label: '&No Fill', run: () => A.setBackground(null) },
      { custom: (close) => { const g = h('div', { class: 'color-panel' }); const grid = h('div', { class: 'cp-grid' }); for (const c of L.color.STANDARD) grid.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${c}`, 'aria-label': ui.colorName(c), 'data-tip': ui.colorName(c), onclick: () => { close(); A.setBackground(c.replace('#', '')); } })); g.appendChild(grid); return g; } },
      { label: '&More Colors...', run: () => ui.moreColors(doc().bg ? '#' + doc().bg : null, (c) => { if (c) A.setBackground(c.replace('#', '')); }) },
      'printedWatermark',
    ];
  }
  A.setBackground = (c) => E.edit(L.N('Background'), () => { D.touchKey(doc(), 'bg'); doc().bg = c; return E.sel; });

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
  function tableGridPicker(r, more) {
    ui.openMenu([{
      custom: (close) => {
        const box = h('div', { class: 'tbl-grid-pick' });
        let R = 4, C = 5;
        const cells = h('div', { class: 'tgp-cells' });
        const lbl = h('div', { class: 'tgp-label', text: 'Cancel' });
        let cur = [0, 0];
        // the cells stay put while the grid grows (a cell replaced under the pointer between press and release
        // gets no click, so the table was never inserted)
        const draw = () => {
          cells.style.gridTemplateColumns = `repeat(${C}, 18px)`;
          const have = new Map([...cells.children].map((s) => [s.dataset.ij, s]));
          let next = cells.firstChild;
          for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) {
            let s = have.get(i + ',' + j);
            if (s) next = s.nextSibling; else cells.insertBefore(s = h('span', { 'data-ij': i + ',' + j }), next);
            s.classList.toggle('on', i < cur[0] && j < cur[1]);
          }
        };
        cells.addEventListener('pointerover', (e) => {
          const s = e.target.closest('span[data-ij]');
          if (!s) return;
          const [i, j] = s.dataset.ij.split(',').map(Number);
          cur = [i + 1, j + 1]; if (i + 2 > R && R < 12) R++; if (j + 2 > C && C < 10) C++;
          draw(); lbl.textContent = L.t('{0} x {1} Table', cur[0], cur[1]);
        });
        cells.addEventListener('click', () => { if (cur[0]) { close(); L.tables.insert(cur[0], cur[1]); } });
        draw();
        box.append(cells, lbl);
        return box;
      },
    }, ...(more || [])], r, { cls: 'grid-menu' });
  }
  function columnsPicker(r, more) {
    ui.openMenu([{
      custom: (close) => {
        const box = h('div', { class: 'tbl-grid-pick' });
        const cells = h('div', { class: 'tgp-cells', style: 'grid-template-columns:repeat(4,22px)' });
        const lbl = h('div', { class: 'tgp-label', text: 'Cancel' });
        for (let j = 0; j < 4; j++) {
          const s = h('span', { style: 'height:30px' });
          s.addEventListener('pointerenter', () => { L.$$('span', cells).forEach((x, k) => x.classList.toggle('on', k <= j)); lbl.textContent = j ? L.t('{0} Columns', j + 1) : L.t('{0} Column', j + 1); });
          s.addEventListener('click', () => { close(); A.setColumns(j + 1); });
          cells.appendChild(s);
        }
        box.append(cells, lbl);
        return box;
      },
    }, ...(more || [])], r, { cls: 'grid-menu' });
  }
  A.setColumns = (n) => E.edit(L.N('Columns'), () => { const s = A.curSect(); D.touchKey(s.holder, s.key); s.sect.cols = Object.assign({}, s.sect.cols, { n, eq: true, w: [] }); return E.sel; });
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
  const HL_NAMES = [['yellow', L.N('Yellow')], ['green', L.N('Bright Green')], ['cyan', L.N('Turquoise')], ['magenta', L.N('Pink')], ['blue', L.N('Blue')], ['red', L.N('Red')], ['darkBlue', L.N('Dark Blue')], ['darkCyan', L.N('Teal')],
    ['darkGreen', L.N('Green')], ['darkMagenta', L.N('Violet')], ['darkRed', L.N('Dark Red')], ['darkYellow', L.N('Dark Yellow')], ['darkGray', L.N('Gray-50%')], ['lightGray', L.N('Gray-25%')], ['black', L.N('Black')]];
  function highlightDD(r, anchor) {
    ui.openMenu([{
      custom: (close) => {
        const el = h('div', { class: 'color-panel' });
        const g = h('div', { class: 'cp-grid hl' });
        for (const [n, name] of HL_NAMES) {
          g.appendChild(h('button', { class: 'cp-sw', type: 'button', style: `background:${L.R.HIGHLIGHT[n]}`, 'data-tip': name, 'aria-label': name, onclick: () => { close(); A.lastHighlight = n; A.applyHighlight(n); } }));
        }
        el.appendChild(g);
        el.appendChild(h('button', { class: 'cp-wide', type: 'button', text: 'None', onclick: () => { close(); A.applyHighlight('none'); } }));
        return el;
      },
    }], anchor ? anchor.getBoundingClientRect() : r, { cls: 'color-menu' });
  }
  A.applyHighlight = (n) => { if (E.collapsed()) { A.highlighter.start(true, n); return; } E.formatRun({ hl: n === 'none' ? undefined : n }, L.N('Highlight')); };
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
    }, L.N('Borders'));
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
      ...[['Basic Shapes', '&Basic Shapes', 'autoshapes'], ['Block Arrows', '&Block Arrows', 'arrowStyle'], ['Flowchart', '&Flowchart', 'diagram'], ['Stars and Banners', '&Stars and Banners', 'animation'], ['Callouts', '&Callouts', 'comment']].map(([g, label, icon]) => ({ label, icon, sub: [grid(L.geom.MENU[g])] })),
      '-', { label: '&More AutoShapes...', icon: 'clipart', run: () => L.panes && L.panes.task.show('clipart') },
    ], { left: r.left, bottom: r.bottom }, { selectFirst: !at });
  };
  /* the toolbars' boxes and menus, also used by the ribbon (js/ribbon.js): each call makes a new box */
  A.menus = { tableGrid: (r, more) => tableGridPicker(r, more), columns: (r, more) => columnsPicker(r, more), lineSpacing: (r) => lineSpacingMenu(r), borders: (r, b) => bordersDD(r, b), highlight: (r, b) => highlightDD(r, b), fontColor: (r, b) => fontColorDD(r, b) };
  A.parts = {
    style: (o) => ui.combo(Object.assign({
      id: 'tb-style', width: 118, tip: 'Style', listCls: 'styles',
      options: () => A.styleOptions().map((s) => s.label),
      preview: (v) => { const s = A.styleOptions().find((x) => x.label === v); return s ? s.css : ''; },
      value: () => (E.sel ? styleState() : ''), onChange: (v) => A.applyStyleByName(v),
    }, o)),
    font: (o) => ui.combo(Object.assign({ id: 'tb-font', width: 136, tip: 'Font', options: () => A.fontOptions(), preview: (v) => `font-family:${L.fontStack(v)};font-size:13px`, note: L.fontNote, listCls: 'fonts', value: () => (A.typingEA() && E.uniformRun('fontEA')) || E.uniformRun('font') || '', onChange: (v) => A.setFont(v) }, o)),
    size: (o) => ui.combo(Object.assign({ id: 'tb-size', width: 44, tip: 'Font Size', options: () => L.SIZE_LIST, value: () => { const v = E.uniformRun('sz'); return v ? String(L.round(v, 1)) : ''; }, onChange: (v) => A.setFontSize(parseFloat(v)) }, o)),
    zoom: (o) => ui.combo(Object.assign({ id: 'tb-zoom', width: 64, tip: 'Zoom', options: () => ['500%', '200%', '150%', '100%', '75%', '50%', '25%', '10%'].concat([L.N('Page Width'), L.N('Text Width'), L.N('Whole Page'), L.N('Two Pages')].map((v) => ({ value: v, label: L.t(v) }))), value: () => Math.round(LY.zoom * 100) + '%', onChange: (v) => A.zoomTo(v) }, o)),
  };
  function buildToolbars() {
    const host = L.$('#toolbars');
    const styleCombo = A.parts.style(), fontCombo = A.parts.font(), sizeCombo = A.parts.size(), zoomCombo = A.parts.zoom();
    A.toolbarsEl = {
      standard: ui.toolbar('standard', 'Standard', ['new', 'open', 'save', 'permission', '|', 'sendMail', 'print', 'printPreview', 'spelling', 'research', '|', 'cut', 'copy', 'paste', 'painter', '|',
        ui.tbSplit('undo', (r) => A.historyMenu(r, 'undo')), ui.tbSplit('redo', (r) => A.historyMenu(r, 'redo')), '|', 'hyperlink', ui.tbButton('tb_tables', { icon: 'tablesBorders' }),
        ui.tbDrop('', 'table', (r) => tableGridPicker(r), 'Insert Table'), 'insertExcel', ui.tbDrop('', 'columns', (r) => columnsPicker(r), 'Columns'), ui.tbButton('tb_drawing', { icon: 'drawing' }), '|', 'docMap', 'showMarks', '|', zoomCombo, '|', 'help', ui.tbButton('viewReading', { label: 'Read', icon: 'read' })]),
      formatting: ui.toolbar('formatting', 'Formatting', ['stylesPane', styleCombo, fontCombo, sizeCombo, '|', 'bold', 'italic', 'underline', '|', 'alignLeft', 'alignCenter', 'alignRight', 'justify', '|',
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
      outlining: ui.toolbar('outlining', 'Outlining', ['promoteH1', 'promote', ui.combo({ id: 'tb-olvl', width: 74, tip: 'Outline Level', options: () => [1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => L.t('Level {0}', n)).concat(L.t('Body text')), value: () => { const l = E.sel ? D.headingLevel(doc(), E.sel.f.p) : 0; return l ? L.t('Level {0}', l) : L.t('Body text'); }, onChange: (v) => L.outline && L.outline.setLevel(parseInt(v.replace(/\D/g, ''), 10) || 0) }), 'demote', 'demoteBody', '|', 'moveUp', 'moveDown', 'expand', 'collapse',
        ui.combo({ id: 'tb-olshow', width: 96, tip: 'Show Level', options: () => [L.t('Show All Levels')].concat([1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => L.t('Show Level {0}', n))), value: () => (L.outline && L.outline.showLevel ? L.t('Show Level {0}', L.outline.showLevel) : L.t('Show All Levels')), onChange: (v) => L.outline && L.outline.show(parseInt(v.replace(/\D/g, ''), 10) || 0) }),
        'olFirstLine', 'olFormatting', '|', 'updateTOC', 'gotoTOC']),
      headerFooter: ui.toolbar('headerFooter', 'Header and Footer', [ui.tbDrop('Insert AutoText', null, (r) => ui.openMenu(L.autocorrect ? L.autocorrect.hfAutoText() : [], r), 'Insert AutoText'), '|', 'hfPageNum', 'hfNumPages', 'hfFormatPageNum', 'hfDate', 'hfTime', '|', 'pageSetup', 'hfShowText', 'hfLinkPrev', '|', 'hfSwitch', 'hfPrev', 'hfNext', '|', ui.tbButton('hfClose', { label: 'Close' })]),
      mailMerge: ui.toolbar('mailMerge', 'Mail Merge', ['mmMainDoc', 'mmOpenSource', 'mmRecipients', '|', 'mmAddress', 'mmGreeting', 'mmInsertField', '|', 'mmViewData', 'mmFirst', 'mmPrev', ui.combo({ id: 'tb-mmrec', width: 34, tip: 'Go to Record', options: () => [], value: () => String(L.mailmerge ? L.mailmerge.rec + 1 : 1), onChange: (v) => L.mailmerge && L.mailmerge.go(parseInt(v, 10) - 1) }), 'mmNext', 'mmLast', '|', 'mmMergeNew', 'mmMergePrint']),
      wordart: ui.toolbar('wordart', 'LettersArt', ['insertWordArt', ui.tbButton('wordartEdit', { label: 'Edit Text...' }), 'wordartGallery', 'formatObject', ui.tbDrop('', 'waShape', (r) => ui.openMenu(L.drawing.WARPS.map(([k, n]) => ({ label: n, run: () => L.drawing.modifySelected('LettersArt Shape', (it) => { it.warp = k; }) })), r), 'LettersArt Shape'), ui.tbDrop('', 'textWrap', (r) => A.wrapMenu(r), 'Text Wrapping')]),
    };
    const row1 = h('div', { class: 'tb-row' }, A.toolbarsEl.standard);
    const row2 = h('div', { class: 'tb-row' }, A.toolbarsEl.formatting);
    const row3 = h('div', { class: 'tb-row' }, A.toolbarsEl.tables, A.toolbarsEl.reviewing, A.toolbarsEl.outlining, A.toolbarsEl.headerFooter, A.toolbarsEl.picture, A.toolbarsEl.mailMerge, A.toolbarsEl.wordart);
    host.append(h('div', { class: 'tb-pair' }, row1, row2), row3);
    A.tbRows = [row1, row2, row3];
    ui.toolbarPrefs(A.opts);
    /* drawing toolbar */
    const asBtn = ui.tbDrop('A&utoShapes', null, (r) => A.openAutoShapes(r), 'AutoShapes');
    asBtn.setAttribute('data-as', '1');
    A.toolbarsEl.drawing = ui.toolbar('drawing', 'Drawing', [ui.tbDrop('D&raw', null, (r) => A.drawMenu(r), 'Draw'), 'selectObjects', '|', asBtn, 'drawLine', 'drawArrow', 'drawRect', 'drawOval', 'textBox', 'insertWordArt', 'insertDiagram', 'insertClipArt', 'insertPicture', '|',
      ui.tbSplit('fillColor', (r, b) => A.objColorMenu(b, 'fill')), ui.tbSplit('lineColor', (r, b) => A.objColorMenu(b, 'line')), ui.tbSplit('fontColorBtn', fontColorDD), '|',
      ui.tbDrop('', 'lineStyle', (r) => A.objLineWeightMenu(r), 'Line Style'), ui.tbDrop('', 'dashStyle', (r) => A.objDashMenu(r), 'Dash Style'), ui.tbDrop('', 'arrowStyle', (r) => A.objArrowMenu(r), 'Arrow Style'), ui.tbDrop('', 'shadowStyle', (r) => A.objShadowMenu(r), 'Shadow Style')]);
    L.$('#drawrow').appendChild(A.toolbarsEl.drawing);
  }
  A.fontOptions = () => {
    const docFonts = [doc().defaults.rPr.font];
    for (const k in doc().styles) { const r = doc().styles[k].rPr; if (r && r.font) docFonts.push(r.font); }
    return L.fontSections({ major: doc().theme?.major, minor: doc().theme?.minor, extra: docFonts });
  };
  /* as in Word, an East Asian font is for all the text (its letters too), a Western font for the Latin text only */
  A.setFont = (v) => { if (!v) return; L.addRecentFont(v); E.formatRun(L.fontIsEA(v) ? { font: v, fontEA: v } : { font: v }, L.N('Font')); E.refocus(); };
  /** whether East Asian text is typed here, as Word tells by an East Asian keyboard: the selection has East Asian
      characters, or the character before the insertion point is one (after it, at a paragraph's start); with no text
      there, the UI's language stands in for the keyboard, which the page can't see (Japanese, Korean, Chinese: the
      languages whose new documents have an East Asian language, common/locales.js). The Font box and the status bar
      then show the East Asian font and language. */
  A.typingEA = function () {
    if (!E.sel) return false;
    let seen = false;
    const look = (t) => { t = t.replace(/[\s\ufffc]/g, ''); if (t) seen = true; return L.hasEAText(t); };
    if (E.collapsed()) {
      const { p, o } = E.sel.f;
      if (look(D.ptext(p, Math.max(0, o - 1), o) || D.ptext(p, o, o + 1))) return true;
    } else {
      const [a, b] = E.range();
      for (const p of D.parasBetween(doc(), a.p, b.p)) if (look(D.ptext(p, p === a.p ? a.o : 0, p === b.p ? b.o : undefined))) return true;
    }
    const loc = !seen && L.locale && L.locale.get();
    return !!(loc && loc.rPr && loc.rPr.langEA);
  };
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
    const ids = Object.keys(d.styles).filter((k) => { const s = d.styles[k]; return (s.type === 'paragraph' || s.type === 'character') && (used.has(k) || s.q || /^Heading[1-3]$/.test(k) || /^heading [1-3]$/i.test(s.name || '') || k === 'Normal') && !s.hidden && !s.sameAs; });
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
    E.edit(L.N('Style'), () => { O.applyStyle(a, b, s.id); return E.sel; });
    E.refocus();
  };
  A.historyMenu = (r, kind) => ui.historyMenu(r, kind, kind === 'undo' ? D.undoList() : D.redoList());

  /* ---------- drawing toolbar helpers ---------- */
  A.drawMenu = function (r) {
    ui.openMenu(['groupObj', 'ungroupObj', 'regroupObj', '-', { label: '&Order', sub: ['bringFront', 'sendBack', 'bringForward', 'sendBackward', 'bringFrontText', 'sendBehindText'] },
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
  A.objLineWeightMenu = (r) => ui.openMenu([{ custom: (close) => listPanel(A.WEIGHTS.map((w) => ({ html: lineSVG(Math.max(0.5, w * 1.33)), label: L.t('{0} pt', w), run: () => L.drawing.modifySelected('Line Style', (it) => { if (it.t === 'img') it.border = Object.assign({ val: 'single', color: '000000' }, it.border || {}, { sz: w }); else it.line = Object.assign({ c: '#000000', dash: 'solid' }, it.line && it.line.t !== 'none' ? it.line : {}, { w }); }) })).concat([{ label: 'More Lines...', run: () => ui.exec('formatObject') }]), close) }], r);
  A.objDashMenu = (r) => ui.openMenu([{ custom: (close) => listPanel([['solid', null], ['sysDot', '1 1'], ['sysDash', '3 1'], ['dash', '4 3'], ['dashDot', '4 3 1 3'], ['lgDash', '8 3'], ['lgDashDot', '8 3 1 3'], ['lgDashDotDot', '8 3 1 3 1 3']].map(([k, da]) => ({ html: lineSVG(2, da), run: () => L.drawing.modifySelected('Dash Style', (it) => { if (it.line && it.line.t !== 'none') it.line = Object.assign({}, it.line, { dash: k }); }) })), close) }], r);
  A.objArrowMenu = (r) => {
    const P = [[null, null], [null, 'triangle'], ['triangle', null], ['triangle', 'triangle'], [null, 'arrow'], [null, 'stealth'], [null, 'oval'], [null, 'diamond'], ['oval', 'oval']];
    const head = (t, x, dir) => !t ? '' : t === 'oval' ? `<circle cx="${x}" cy="7" r="3"/>` : t === 'diamond' ? `<path d="M${x - 4} 7l4-4 4 4-4 4z"/>` : t === 'arrow' ? `<path d="M${x - dir * 6} 2l${dir * 6} 5-${dir * 6} 5" fill="none" stroke="#000" stroke-width="1.5"/>` : `<path d="M${x} 7l${-dir * 8} -4v8z"/>`;
    ui.openMenu([{ custom: (close) => listPanel(P.map(([a, b]) => ({ html: `<svg width="90" height="14" viewBox="0 0 90 14" fill="#000"><path d="M6 7h78" stroke="#000" stroke-width="1.5"/>${head(a, 4, -1)}${head(b, 86, 1)}</svg>`, run: () => L.drawing.modifySelected('Arrow Style', (it) => { if (!it.line || it.line.t === 'none') return; it.line = Object.assign({}, it.line, { head: a ? { type: a, w: 'med', len: 'med' } : undefined, tail: b ? { type: b, w: 'med', len: 'med' } : undefined }); }) })), close) }], r);
  };
  A.objShadowMenu = (r) => {
    const S = [null, [3, 3], [-3, 3], [3, -3], [-3, -3], [0, 4], [4, 0], [6, 6]];
    ui.openMenu([{ custom: (close) => { const g = h('div', { class: 'shape-grid', style: 'grid-template-columns:repeat(4,30px)' }); S.forEach((s) => { const b = h('button', { type: 'button', style: 'width:30px;height:30px', 'aria-label': s ? 'Shadow Style' : 'No Shadow', html: s ? `<svg width="24" height="24"><rect x="${6 + s[0] / 1.5}" y="${6 + s[1] / 1.5}" width="12" height="12" fill="#888"/><rect x="6" y="6" width="12" height="12" fill="#fff" stroke="#000"/></svg>` : '<svg width="24" height="24"><rect x="6" y="6" width="12" height="12" fill="#fff" stroke="#000"/><path d="M3 21L21 3" stroke="#d33"/></svg>' }); b.addEventListener('click', () => { close(); L.drawing.modifySelected(L.N('Shadow'), (it) => { it.shadow = s ? { c: '#000000', a: 0.5, dx: s[0], dy: s[1], blur: 0 } : undefined; }); }); g.appendChild(b); }); return g; } }], r);
  };
  A.borderStyleMenu = (r) => ui.openMenu([{ custom: (close) => listPanel([['single', null], ['dotted', '1 2'], ['dashed', '4 3'], ['dotDash', '4 3 1 3'], ['double', null], ['thick', null], ['triple', null], ['wave', null], ['threeDEmboss', null]].map(([k, da]) => ({ html: k === 'double' ? '<svg width="90" height="14"><path d="M2 5h86M2 9h86" stroke="#000"/></svg>' : k === 'triple' ? '<svg width="90" height="14"><path d="M2 3h86M2 7h86M2 11h86" stroke="#000"/></svg>' : k === 'wave' ? '<svg width="90" height="14"><path d="M2 7q5-5 10 0t10 0t10 0t10 0t10 0t10 0t10 0t10 0t6 0" fill="none" stroke="#000"/></svg>' : lineSVG(k === 'thick' ? 3 : 1, da), run: () => { A.borderPen.val = k; } })).concat([{ label: 'No Border', run: () => { A.borderPen.val = 'nil'; } }]), close) }], r);
  A.borderWeightMenu = (r) => ui.openMenu([{ custom: (close) => listPanel([0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => ({ html: lineSVG(Math.max(0.5, w * 1.33)), label: L.t('{0} pt', w), run: () => { A.borderPen.sz = w; } })), close) }], r);
  A.picMode = (m) => L.drawing.modifySelected('Picture Color', (it) => { it.gray = m === 'gray' || undefined; it.bw = m === 'bw' || undefined; it.washout = m === 'washout' || undefined; });
  A.picAdjust = (k, d) => L.drawing.modifySelected(k === 'bright' ? L.N('Brightness') : L.N('Contrast'), (it) => { it[k] = L.clamp(L.round((it[k] || 0) + d, 2), -1, 1) || undefined; });

  /* ================= status bar ================= */
  let statusTimer = 0;
  function buildStatus() {
    const sb = L.$('#statusbar');
    const cell = (cls, tip) => h('div', { class: 'sb-cell ' + (cls || ''), 'data-tip': tip || null });
    A.sb = {
      page: cell('', L.N('Page')), sec: cell('opt', L.N('Section')), pages: cell('', 'Page / pages'), at: cell('opt', 'Vertical position'), ln: cell('', L.N('Line|position')), col: cell('', L.N('Column|position')),
      rec: cell('flag opt', 'Record macro'), trk: cell('flag act opt', 'Track Changes (double-click to toggle)'), ext: cell('flag act opt', 'Extend selection'), ovr: cell('flag act opt', 'Overtype (double-click to toggle)'),
      lang: cell('act', L.N('Language')), spell: cell('act sb-spell', 'Spelling and grammar status'), msg: cell('grow'),
    };
    A.sb.rec.textContent = L.t('REC'); A.sb.trk.textContent = L.t('TRK'); A.sb.ext.textContent = L.t('EXT'); A.sb.ovr.textContent = L.t('OVR');
    A.sb.lang.textContent = A.langName('en-US');
    A.sb.spell.innerHTML = L.icons.get('spellStatus');
    A.sb.trk.addEventListener('dblclick', () => ui.exec('trackChanges'));
    A.sb.ovr.addEventListener('dblclick', () => { E.overtype = !E.overtype; A.updateStatus(); });
    A.sb.ext.addEventListener('dblclick', () => { E.extend = !E.extend; A.updateStatus(); });
    A.sb.lang.addEventListener('dblclick', () => ui.exec('setLanguage'));
    /* click: check this document's spelling and grammar as you type, or stop (off when a document opens) */
    A.sb.spell.addEventListener('click', () => {
      if (!L.spell) return;
      if (!A.opts.spell) { A.opts.spell = true; A.saveOpts && A.saveOpts(); }
      L.spell.hidden = !L.spell.hidden;
      A.status(L.spell.hidden ? 'Spelling and grammar are no longer marked in this document.' : 'Checking spelling and grammar as you type in this document.');
    });
    const spellIcon = () => {
      const on = L.spell && !L.spell.hidden && A.opts.spell !== false;
      A.sb.spell.classList.toggle('off', !on);
      A.sb.spell.dataset.tip = on ? L.t('Spelling and grammar are checked as you type. Click to stop.') : L.t('Spelling and grammar are not checked as you type. Click to check this document.');
    };
    L.bus.on('proof-state', spellIcon);
    L.bus.on('doc-loaded', spellIcon);
    spellIcon();
    A.sb.page.addEventListener('dblclick', () => ui.exec('goTo'));
    // the right end, as in the other apps (ui.js): browse buttons, view buttons, the zoom slider, the task pane's toggle
    const zoom = ui.zoomSlider({ min: 10, max: 500, value: () => Math.round(LY.zoom * 100), dialog: 'zoomDlg',
      set: (p, final) => (final ? A.zoomTo(p + '%', { focus: false }) : A.zoomLive(p)) });
    const browse = ui.statusGroup(['browsePrev', ui.tbDrop('', 'browseObj', (r) => A.browseMenu(r), 'Select Browse Object'), 'browseNext']);
    browse.classList.add('sb-browse');
    const views = ui.statusGroup(['viewNormal', 'viewWeb', 'viewPrint', 'viewOutline', 'viewReading']);
    views.classList.add('sb-views');
    sb.append(A.sb.page, A.sb.sec, A.sb.pages, A.sb.at, A.sb.ln, A.sb.col, A.sb.rec, A.sb.trk, A.sb.ext, A.sb.ovr, A.sb.lang, A.sb.spell, A.sb.msg, h('div', { class: 'sb-right' }, browse, views, zoom, ui.paneToggle('taskPane')));
  }
  A.status = (msg, ms) => { if (!A.sb) return; A.sb.msg.textContent = L.t(msg); clearTimeout(statusTimer); statusTimer = setTimeout(() => { A.sb.msg.textContent = ''; }, ms || 4000); };
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
      A.sb.page.textContent = L.t('Page {0}', D.fmtNum(pg.numVal, pg.numFmt));
      A.sb.sec.textContent = L.t('Sec {0}', pg.si + 1);
      A.sb.pages.textContent = `${pi + 1}/${n}`;
      const r = E.caretRect();
      if (r) {
        const pr = pg.el.getBoundingClientRect();
        const atPt = (r.top - pr.top) / LY.zoom / D.PX;
        A.sb.at.textContent = L.t('At {0}', D.fmtLen(atPt));
        /* line number on the page and column within the line */
        const lc = A.lineCol(sel.f, pg);
        A.sb.ln.textContent = L.t('Ln {0}', lc.ln);
        A.sb.col.textContent = L.t('Col {0}', lc.col);
      }
    } else {
      A.sb.page.textContent = LY.view === 'print' ? '' : L.t('Page {0}', 1);
      A.sb.sec.textContent = L.t('Sec {0}', 1);
      A.sb.pages.textContent = LY.view === 'print' ? '' : '';
      A.sb.at.textContent = ''; A.sb.ln.textContent = ''; A.sb.col.textContent = L.t('Col {0}', sel.f.o + 1);
    }
    const r = E.curRun(), def = d.defaults.rPr || {};
    const lang = (A.typingEA() && (r.langEA || def.langEA)) || r.lang || def.lang || 'en-US';
    A.sb.lang.textContent = A.langName(lang);
  }, 60);
  /* [code, name]: the names are UI text, shown through L.t; A.langName gives the name as shown */
  A.LANGS = [['en-US', L.N('English (U.S.)')], ['en-GB', L.N('English (U.K.)')], ['en-AU', L.N('English (Australia)')], ['en-CA', L.N('English (Canada)')], ['fr-FR', L.N('French (France)')], ['fr-CA', L.N('French (Canada)')], ['de-DE', L.N('German (Germany)')], ['es-ES', L.N('Spanish (Spain)')], ['es-MX', L.N('Spanish (Mexico)')], ['it-IT', L.N('Italian (Italy)')], ['pt-BR', L.N('Portuguese (Brazil)')], ['pt-PT', L.N('Portuguese (Portugal)')], ['nl-NL', L.N('Dutch (Netherlands)')], ['sv-SE', L.N('Swedish')], ['da-DK', L.N('Danish')], ['nb-NO', L.N('Norwegian (Bokmål)')], ['fi-FI', L.N('Finnish')], ['pl-PL', L.N('Polish')], ['ru-RU', L.N('Russian')], ['ja-JP', L.N('Japanese')], ['zh-CN', L.N('Chinese (Simp.)')], ['zh-HK', L.N('Chinese (Trad.)')], ['ko-KR', L.N('Korean')]];
  A.langName = (code) => { const hit = A.LANGS.find((x) => x[0].toLowerCase() === String(code).toLowerCase()); return hit ? L.t(hit[1]) : code; };
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
          ln += LY.lineCount(f);
        }
      }
    }
    return { ln, col };
  };
  A.updateTitle = () => { document.title = `${L.file.titled(A.fileName)} - Quire`; };

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
    ui.dragCancel();   /* a view switch ends a drag in progress */
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
  /** zoom to v ('NN%', 'Page Width', 'Text Width', 'Whole Page', 'Two Pages'); the Zoom box and dialog give the
      document the keyboard afterwards, the status bar's slider and the ribbon's buttons pass { focus: false } */
  A.zoomTo = function (v, o) {
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
    if (!o || o.focus !== false) E.focus();
  };
  /** while the slider is dragged: the same zoom without saving it; Web Layout, which lays the whole document out
      again for a zoom, waits for the release */
  A.zoomLive = function (pct) {
    if (LY.view === 'web') return;
    if (LY.multi && A.view !== 'reading') LY.setMultiPage(0);
    LY.setZoom(L.clamp(pct / 100, 0.1, 5));
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
  /** what is selected, for the ribbon's contextual tabs: 'table', 'picture', 'shape', 'headerFooter', 'outlining' */
  A.inContext = (k) => {
    const it = E.objSel && E.objSel.it;
    if (k === 'table') return !!(L.tables && L.tables.inTable());
    if (k === 'shape') return !!(it && (it.t === 'shape' || it.t === 'group' || it.t === 'text' || it.tb || it.wordart));
    return ctxToolbar(k);
  };
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
    ui.dragCancel();
    stash();
    const w = A.docs[i];
    A.cur = i;
    if (LY.hfEdit) { LY.hfEdit = null; LY.root.classList.remove('hfmode'); }
    E.clearObjSel();
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
  A.windowItems = () => (window.VO && VO.oneDocument ? VO.windowItems() : A.docs.map((w, i) => ({ label: `&${i + 1} ${(i === A.cur ? A.fileName : w.fileName).replace(/&/g, '&&')}`, raw: true, checked: () => i === A.cur, run: () => A.switchTo(i) })));
  A.loadDoc = function (d, name, opts) {
    opts = opts || {};
    /* the desktop app: one document per window, so a new one goes to a window of its own */
    if (opts.newWindow && A.cur >= 0 && window.VO && VO.oneDocument) { A.toWindow(d, name); return; }
    ui.dragCancel();
    if (opts.newWindow || A.cur < 0) { stash(); A.docs.push(null); A.cur = A.docs.length - 1; }
    if (LY.hfEdit) { LY.hfEdit = null; LY.root.classList.remove('hfmode'); }
    E.clearObjSel();
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
    const r = await ui.msg(L.t('Do you want to save the changes to {0}?', A.fileName), { icon: 'warn', buttons: ['&Yes', '&No', 'Cancel'] });
    if (r === 0) return await A.save() === 'saved';
    if (r === 1 && window.VO) VO.discard(doc());   // its unsaved version goes too
    return r === 1;
  }
  A.confirmDiscard = confirmDiscard;
  /** a document made here (New with content, mail merge's letters) in a window of its own, as a .docx (the desktop
   *  app's one document per window) */
  A.toWindow = async function (d, name) {
    const keep = D.doc;
    ui.busy(true, 'Saving...');
    try {
      D.doc = d;   // the writer looks up list numbering in D.doc
      const blob = await L.docx.write(d, { format: 'docx' });
      D.doc = keep;
      ui.busy(false);
      await VO.made(blob, name + '.docx');
    } catch (e) {
      D.doc = keep;
      ui.busy(false);
      console.error(e);
      ui.msg(String(e && e.message || e), { icon: 'error' });
    }
  };
  /** a blank document as the Office of the UI's language makes it (common/locales.js) */
  const blank = () => D.applyLocale(D.newDoc(), L.locale && L.locale.get());
  A.newDocument = async function (opts) {
    /* the desktop app: a blank document in a window of its own (one with content goes through loadDoc) */
    if (!opts && !pristine() && window.VO && VO.oneDocument) return VO.go(VO.url('quire', { new: true })).catch((e) => ui.msg(String(e), { icon: 'error' }));
    A.untitled++;
    const d = opts && opts.modern ? D.newDoc({ modern: true }) : blank();
    if (opts && opts.build) opts.build(d);
    /* a new document opens in its own window, as in Word; an untouched blank one is simply replaced */
    A.loadDoc(d, 'Document' + A.untitled, { newWindow: !pristine() });
  };
  /** close the current document (asking to save); false when the user cancels */
  A.closeDocument = async function () {
    if (!(await confirmDiscard())) return false;
    if (window.VO && VO.oneDocument) { VO.exit(); return true; }   // its window closes
    if (A.docs.length <= 1) { A.docs = []; A.cur = -1; A.untitled++; A.loadDoc(blank(), 'Document' + A.untitled); return true; }
    const i = A.cur;
    A.docs.splice(i, 1);
    A.cur = -1; /* nothing to stash: the closed window is gone */
    A.switchTo(Math.max(0, i - 1));
    return true;
  };
  A.openDialog = async function () {
    if (window.VO && VO.oneDocument) return VO.openDialog();
    const files = await L.pickFiles('.docx,.docm,.dotx,.dotm,.doc,.txt,.htm,.html,.rtf,.xml,.md,.markdown,.zip,application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    if (files[0]) A.openFile(files[0]);
  };
  A.openFiles = (files) => { const f = files.find((x) => /\.(docx|docm|dotx|dotm|doc|txt|html?|rtf|md|markdown|zip)$/i.test(x.name)); if (f) A.openFile(f); };
  A.openFile = async function (file) {
    if (window.VO && await VO.elsewhere(file)) return;   // the desktop app: a window of its own
    const reuse = pristine();
    const name = file.name.replace(/\.(docx|docm|dotx|dotm|doc|txt|html?|rtf|xml|md|markdown|mdown|mkd|zip)$/i, '');
    ui.busy(true, L.t('Opening {0}...', file.name));
    try {
      let d, warnings = [], type = 'docx';
      if (/\.txt$/i.test(file.name)) d = A.docFromText(await L.readAsText(file));
      else if (L.mdio && L.mdio.isMarkdown(file.name)) { d = await L.mdio.read(await L.readAsText(file)); type = 'md'; }
      else if (L.mdio && L.mdio.isPackage(file.name)) { d = await L.mdio.readZip(await L.readAsArrayBuffer(file)); type = 'mdzip'; }
      else if (/\.html?$/i.test(file.name) && L.htmlio) d = await L.htmlio.docFromHTML(await L.readAsText(file));
      else if (/\.rtf$/i.test(file.name) && L.rtf) d = L.rtf.read(await L.readAsArrayBuffer(file));
      else {
        const buf = await L.readAsArrayBuffer(file);
        const res = await A.readDocx(buf, file.name);
        d = res.doc; warnings = res.warnings || []; type = d.ooxmlFormat || type;
      }
      A.loadDoc(d, name, { saved: /\.(docx|docm|dotx|dotm)$/i.test(file.name) || type !== 'docx', type, newWindow: !reuse });
      if (window.VO) VO.opened(file, doc());
      ui.busy(false);
      if (d.isTemplate) A.status('Opened a template. Save keeps its template format.');
      if (d.wasEncrypted) A.status('This document is protected with a password to open; saving keeps the protection (Tools ▸ Options ▸ Security).');
      if (warnings.length) ui.msg(warnings.join('\n'), { icon: 'warn' });
      if (type === 'md') { const n = L.mdio.missingPictures(d); if (n) A.status(n === 1 ? L.t('A picture is not in the .md file and shows as a box. Open a .zip holding the file and its pictures to see them.') : L.t('{0} pictures are not in the .md file and show as a box. Open a .zip holding the file and its pictures to see them.', n)); }
      if (d.settings.protect && d.settings.protect.on) A.status({ readOnly: L.N('This document is protected: no changes (read only)'), comments: L.N('This document is protected: comments only'), tracked: L.N('This document is protected: tracked changes'), forms: L.N('This document is protected: filling in forms') }[d.settings.protect.kind] || L.N('This document is protected: restricted'));
    } catch (e) {
      ui.busy(false);
      console.error(e);
      if (e && e.code === 'cancel') return;
      if (e && e.code === 'ole') ui.msg(e.message, { icon: 'warn' });
      else ui.msg(L.t("{0} can't open {1}. {2}", L.APP, file.name, e && e.message ? e.message : ''), { icon: 'error' });
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
        const b = new Error(err && err.code === 'badpass' ? L.t('The password is incorrect. {0} cannot open the document.', L.APP) : err && err.message === 'unsupported' ? L.t('This document is encrypted in a way {0} cannot read.', L.APP) : String(err && err.message || err));
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
  /* what Quire saves, for the desktop app's Save dialog (L.file); the document types are saves, the rest exports */
  L.file.types = [{ id: 'docx', label: L.N('Word Document (*.docx)'), ext: ['docx'] }, { id: 'docm', label: L.N('Word Macro-Enabled Document (*.docm)'), ext: ['docm'] },
    { id: 'dotx', label: L.N('Word Template (*.dotx)'), ext: ['dotx'] }, { id: 'dotm', label: L.N('Word Macro-Enabled Template (*.dotm)'), ext: ['dotm'] },
    { id: 'html', label: L.N('Single File Web Page (*.htm)'), ext: ['htm', 'html'] }, { id: 'rtf', label: L.N('Rich Text Format (*.rtf)'), ext: ['rtf'] },
    { id: 'txt', label: L.N('Plain Text (*.txt)'), ext: ['txt'] }, { id: 'md', label: L.N('Markdown (*.md)'), ext: ['md', 'markdown'] },
    { id: 'mdzip', label: L.N('Markdown with pictures (*.zip)'), ext: ['zip'] }, { id: 'pdf', label: L.N('PDF (*.pdf)'), ext: ['pdf'] }];
  const DOC_TYPES = ['docx', 'docm', 'dotx', 'dotm', 'md', 'mdzip'];
  /** Save: the document's own file (the desktop app: in place, when it has one it may change), else Save As */
  A.save = async function () {
    if (!A.saved || (L.file.native() && !L.file.inPlace())) return A.saveAs();
    return A.exportAs(A.fileType || 'docx', A.fileName, L.file.native() ? 'save' : undefined);
  };
  A.saveAs = function () {
    /* the desktop app: the system's Save dialog picks the name and type, then the document is written there */
    if (L.file.native()) return (async () => {
      const c = await L.file.choose('saveAs', DOC_TYPES.includes(A.fileType) ? A.fileType : 'docx', A.fileName);
      return c ? A.exportAs(c.type, c.name, 'chosen') : 'declined';
    })();
    if (!L.dlg?.saveAs) return A.exportAs('docx', A.fileName);
    return new Promise(resolve => {
      let writing = false;
      const dialog = L.dlg.saveAs(async (name, type) => { writing = true; resolve(await A.exportAs(type, name)); });
      dialog.done.then(() => { if (!writing) resolve('declined'); });
    });
  };
  A.docStats = function () {
    const d = doc();
    const st = D.stats(D.allParas(d));
    return { pages: LY.pages.length || 1, words: st.words, chars: st.chars, charsNoSp: st.charsNoSp, paras: st.paras, lines: Math.max(st.paras, Math.round(st.chars / 70)) };
  };
  /** writes the document as `type`: a download (web), or in the desktop app (`how`) its own file ('save') or what the
   *  Save dialog picked ('chosen'); without `how` there, an export that asks where. A document type saved where it
   *  becomes the document's file renames the document and clears its changes */
  A.exportAs = async function (type, name, how) {
    name = name || A.fileName;
    const d = doc();
    const before = window.VO ? VO.changeCount(d) : 0;
    ui.busy(true, 'Saving...');
    try {
      if (L.fields && L.fields.beforeSave) L.fields.beforeSave();
      let blob, ext = type;
      d.props.modified = new Date().toISOString();
      d.props.lastModifiedBy = A.opts.userName || 'Quire User';
      if (/^(docx|docm|dotx|dotm)$/.test(type)) {
        d.props.revision = (d.props.revision || 0) + 1;
        blob = await L.docx.write(d, { stats: A.docStats(), format: type });
        /* a password to open: encrypt the package (Agile encryption, readable by Word 2010+ and LibreOffice) */
        if (d.openPassword && L.officeCrypto) {
          const original = blob;
          blob = new Blob([await L.officeCrypto.encrypt(new Uint8Array(await original.arrayBuffer()), d.openPassword)], { type: original.type });
          blob.dropped = original.dropped;
        }
      }
      else if (type === 'md') {
        /* pictures need the .zip form: a .md file only refers to them */
        if (L.mdio.newPictures(d)) {
          ui.busy(false);
          const r = await ui.msg('This document has pictures that a .md file cannot hold. Save it as Markdown with pictures (.zip), which holds the .md file and its pictures?', { icon: 'warn', buttons: ['Save as .&zip', 'Save .md &without them', 'Cancel'] });
          if (r !== 0 && r !== 1) { if (how === 'chosen') L.file.cancel(); return 'declined'; }
          if (r === 0) {
            type = ext = 'mdzip';
            /* the desktop app: a .zip is another file, so the Save dialog asks again */
            if (how) { const c = await L.file.choose('saveAs', 'mdzip', name, ['mdzip']); if (!c) return 'declined'; name = c.name; how = 'chosen'; }
            else if (A.fileType === 'md') A.fileType = 'mdzip';
          }
          ui.busy(true, 'Saving...');
        }
        if (type === 'md') blob = new Blob([L.mdio.write(d).text], { type: 'text/markdown;charset=utf-8' });
      }
      if (type === 'mdzip') { blob = await L.mdio.writeZip(d, name); ext = 'zip'; }
      else if (type === 'html') blob = await L.htmlio.webPage(d, name);
      else if (type === 'txt') blob = new Blob([L.htmlio ? L.htmlio.plainText(d) : ''], { type: 'text/plain;charset=utf-8' });
      else if (type === 'pdf') blob = await A.buildPDF();
      else if (type === 'rtf' && L.rtf) { blob = new Blob([await L.rtf.write(d)], { type: 'application/rtf' }); }
      ui.busy(false);
      if (!blob) return;
      const losses = L.opc.formats[type] ? await ui.compatibility(d, blob.dropped) : [];
      if (losses === null) { if (how === 'chosen') L.file.cancel(); return 'declined'; }
      const w = how ? await L.file.write(blob, how, { type, name: `${name}.${ext}` }) : { result: await L.saveFile(`${name}.${ext}`, blob) };
      const r = w.result;
      if (w.name) name = w.name.replace(/\.[^.]+$/, '');   // as the Save dialog left it
      if (r === 'saved') L.opc.acknowledge(d, losses);
      /* the document's own file now: on the web any save of a document type; the desktop app's Save, or a Save As it bound */
      const own = !L.file.native() || how === 'save' || w.bound;
      if (r === 'saved' && own && DOC_TYPES.includes(type)) { A.fileName = name; A.fileType = type; A.saved = true; A.updateTitle(); d.dirty = !!window.VO && VO.changeCount(d) !== before; A.status(L.t('Saved {0}', `${name}.${ext}`)); ui.refresh(); if (window.VO) VO.saved(`${name}.${ext}`, blob, d); }
      return r;
    } catch (e) {
      ui.busy(false);
      console.error(e);
      ui.msg(L.t('The file could not be saved: {0}', e.message || e), { icon: 'error' });
    }
  };

  /* ---------- PDF (print) ---------- */
  /** the document's own rules (`.p …`, `table.tbl …`), for the page pictures of PDF export */
  function docCSS() {
    let css = '';
    for (const sheet of document.styleSheets) {
      let rules;
      try { rules = sheet.cssRules; } catch (e) { continue; }
      for (const r of rules) if (r.selectorText && /(^|[\s,>]|\b(?:table|td))\.(pg|p|pc|lbl|tb|img|shp|tbl|cf|nref|fdyn|ffcb|crop|pg-|frm|flt|anc|sym|raw|zw|fnsep|fnote|hbrk|sectmark|cmk|cmr|ins|del|hid|lnk|sgrp|tbx|geo|wa|chart)/.test(r.selectorText)) css += r.cssText + '\n';
    }
    return css + '.crop{display:none}.pg-hdr,.pg-ftr{opacity:1!important}.cmk,.cmr{background:none!important}td.csel{background-image:none!important}.p.chg::before{display:none}';
  }
  /** the pictures as data: URLs, by their blob: URLs; within: only those an element shows (a page's thumbnail) */
  async function mediaDataMap(within) {
    const html = within ? within.outerHTML : null;
    const map = new Map();
    for (const [, m] of L.media.all()) { if (html && !html.includes(m.url)) continue; try { map.set(m.url, await L.blobToDataURL(m.view || m.blob)); } catch (e) { /* skip */ } }
    return map;
  }
  A.rasterizePage = async function (pgEl, scale, css, map) {
    const w = pgEl.offsetWidth, hh = pgEl.offsetHeight;
    const clone = pgEl.cloneNode(true);
    clone.style.margin = '0'; clone.style.boxShadow = 'none';
    for (const e of clone.querySelectorAll('.format-marks, .sectmark, .hbrk')) e.remove();
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
      ui.busy(true, L.t('Creating PDF... (page {0} of {1})', ++n, list.length));
      const c = await A.rasterizePage(p.el, 2.2, css, map);
      const jpeg = new Uint8Array(await (await new Promise((res, rej) => c.toBlob((b) => (b ? res(b) : rej(new Error(L.t('This browser blocked the page image.')))), 'image/jpeg', 0.92))).arrayBuffer());
      pages.push({ jpeg, iw: c.width, ih: c.height, pw: p.sect.pgW, ph: p.sect.pgH });
    }
    if (prevView !== A.view || LY.view !== (prevView === 'reading' || prevView === 'preview' ? 'print' : prevView)) { LY.view = prevView === 'reading' || prevView === 'preview' ? 'print' : prevView; LY.render(); }
    E.restoreDom(true);
    return L.buildPDF(pages, { title: doc().props.title || A.fileName, author: doc().props.creator || '' });
  };
  A.printPDF = async function () {
    try { const blob = await A.buildPDF(); ui.busy(false); await L.saveFile(`${A.fileName}.pdf`, blob); }
    catch (e) { ui.busy(false); ui.msg(L.t('The PDF could not be created: {0}', e.message || e), { icon: 'error' }); }
  };

  /* ================= headers & footers ================= */
  /** the caret at a point of the page, when that point is in story (or, without a story, in the body) */
  const caretAt = (ev, story) => {
    const pos = ev && E.posAtPoint(ev.clientX, ev.clientY);
    return pos && D.storyOf(doc(), pos.p) === (story || doc().main) ? pos : null;
  };
  /** open the header or footer of kind on page pageIdx (the caret's page); ev, the double-click that opened it:
      the view stays, and the caret goes where it was double-clicked. The body's selection is kept for the way back. */
  A.editHeaderFooter = function (kind, pageIdx, ev) {
    if (A.view !== 'print') A.setView('print');
    if (!LY.hfEdit && E.sel && D.storyOf(doc(), E.sel.f.p) === doc().main) hfReturn = { a: E.sel.a, f: E.sel.f };
    const top = LY.scroller.scrollTop;
    const pg = LY.pages[pageIdx == null ? Math.max(0, LY.pageOf(E.sel.f)) : pageIdx] || LY.pages[0];
    const d = doc();
    const secs = D.sections(d);
    kind = kind || 'hdr';
    const ty = pg.type;
    /* create the story when the section has none */
    let story = LY.hfFor(secs, pg.si, kind, ty);
    if (!story) {
      E.edit(kind === 'hdr' ? L.N('Header') : L.N('Footer'), () => {
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
    if (ev) LY.scroller.scrollTop = top;
    const p = story && story.paras && story.paras[0] ? story.paras[0] : story && D.firstPara(story.blocks);
    const at = caretAt(ev, story);
    if (at) E.setSel(at); else if (p) E.setSel(D.pos(p, 0));
    A.updateToolbars();
    ui.refresh();
  };
  const HF_NAMES = { first: [L.N('First Page Header'), L.N('First Page Footer')], even: [L.N('Even Page Header'), L.N('Even Page Footer')], odd: [L.N('Odd Page Header'), L.N('Odd Page Footer')], default: [L.N('Header'), L.N('Footer')] };
  function labelHF() {
    const d = doc();
    const secs = D.sections(d);
    for (const pg of LY.pages) {
      const multi = secs.length > 1;
      const ty = pg.type === 'first' || pg.type === 'even' ? pg.type : doc().settings.evenOdd ? 'odd' : 'default';
      const label = (name) => (multi ? L.t('{0} -Section {1}-', L.t(name), pg.si + 1) : L.t(name));
      pg.hdr.dataset.label = label(HF_NAMES[ty][0]);
      pg.ftr.dataset.label = label(HF_NAMES[ty][1]);
      const refH = LY.hfRefFor(secs, pg.si, 'hdr', pg.type), refF = LY.hfRefFor(secs, pg.si, 'ftr', pg.type);
      if (refH && refH.si < pg.si) pg.hdr.dataset.same = '1'; else delete pg.hdr.dataset.same;
      if (refF && refF.si < pg.si) pg.ftr.dataset.same = '1'; else delete pg.ftr.dataset.same;
    }
  }
  L.bus.on('layout-done', () => { if (LY.hfEdit) labelHF(); });
  /** back to the body: the caret where ev (a double-click on the body) was, else where it was before the header */
  let hfReturn = null;
  A.closeHeaderFooter = function (silent, ev) {
    LY.hfEdit = null;
    LY.root.classList.remove('hfmode');
    if (!silent) {
      const top = LY.scroller.scrollTop;
      LY.render();
      LY.scroller.scrollTop = top;
      const back = hfReturn && D.byId(doc(), hfReturn.a.p.id) === hfReturn.a.p && D.byId(doc(), hfReturn.f.p.id) === hfReturn.f.p ? hfReturn : null;
      const at = caretAt(ev);
      if (at) E.setSel(at);
      else if (back) E.setSel(back.a, back.f);
      else E.setSel(D.pos(D.firstPara(doc().main.blocks), 0));
    }
    hfReturn = null;
    A.updateToolbars();
  };

  /* ================= misc helpers used by commands ================= */
  /** select pos (to sel2) and show it; opts.focus false: a modeless dialog (Find and Replace) keeps the keyboard */
  A.gotoPos = function (pos, sel2, opts) {
    E.clearObjSel();   /* the text found or gone to is the selection now, not a shape selected before */
    if (opts && opts.focus === false) {
      E.setSel(pos, sel2 || pos, { noDom: true });
      /* an edit just before (Replace) left the page's own selection behind: it would show, and typing would go there */
      const s = window.getSelection();
      if (s.rangeCount && LY.root && LY.root.contains(s.anchorNode)) s.removeAllRanges();
      E.scrollToCaret(true);
      E.drawSel(true);
      return;
    }
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
    /* a Markdown link to a heading uses GitHub's anchor for it ("#getting-started") */
    if (L.md) {
      const used = new Set();
      for (const p of D.allParas(d)) if (D.headingLevel(d, p) && L.md.slug(D.plainText(p), used) === name.toLowerCase()) { A.gotoPos(D.pos(p, 0)); return true; }
    }
    A.status(L.t('Bookmark "{0}" not found.', name));
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
  /** Word's menu for a selected object: its text, grouping and order, wrapping, caption, Format…, hyperlink */
  A.objectMenu = function (it) {
    const fl = !!(E.objSel && E.objSel.float);
    const format = it.t === 'img' ? 'Format &Picture...' : it.wordart ? 'Format Letters&Art...' : it.tb && it.txBox !== false ? 'Format Te&xt Box...' : it.t === 'shape' ? 'Format AutoSha&pe...' : 'Format &Object...';
    const items = ['cut', 'copy', 'paste', '-'];
    if (L.drawing.takesText(it)) items.push('shapeText', '-');
    else if (it.wordart) items.push('wordartEdit', '-');
    if (fl) items.push({ label: 'Gro&uping', sub: ['groupObj', 'ungroupObj', 'regroupObj'] }, { label: '&Order', sub: ['bringFront', 'sendBack', 'bringForward', 'sendBackward', '-', 'bringFrontText', 'sendBehindText'] }, '-');
    items.push({ label: 'Text W&rapping', sub: () => A.wrapItems() });
    if (it.t === 'shape' && !it.wordart) items.push('setDefaults');
    items.push('insertCaption', { cmd: 'formatObject', label: format }, '-', 'hyperlink');
    return items;
  };
  /** Word's Synonyms submenu: the thesaurus' words for the one at the caret (a click puts it there), and Thesaurus… */
  A.synonymItems = function () {
    const r = E.sel && (E.collapsed() ? O.wordAt(E.sel.f) : E.range());
    const word = r ? O.textRange(r[0], r[1]).trim() : '';
    const TH = L.thesaurus, e = TH && TH.data && /^[\p{L}'’-]+$/u.test(word) ? TH.lookup(word) : null;
    const cap = (s) => (/^\p{Lu}/u.test(word) ? s.charAt(0).toUpperCase() + s.slice(1) : s);
    const syns = e ? [...new Set(e.senses.flatMap((s) => s.syns))].slice(0, 10).map(cap) : [];
    return (syns.length ? syns.map((s) => ({ label: s, raw: true, run: () => { E.setSel(r[0], r[1]); E.insertTextAt(s, L.N('Synonym')); } })) : [{ label: '(No Suggestions)', disabled: true }]).concat(['-', 'thesaurus']);
  };
  A.contextMenu = function (e) {
    e.preventDefault();
    /* a right-click away from the selected object (or in its text) selects there first, as in Word */
    const onObj = E.objSel && e.target.closest && !e.target.closest('.tbx') && ((E.objSel.el && E.objSel.el.contains(e.target)) || e.target.closest('.obj-handles'));
    if (E.objSel && !onObj) E.clearObjSel();
    if (!E.objSel) E.syncSel();   /* the press moved the caret, but 'selectionchange' comes after this */
    /* a selected object: its own menu, without the text's */
    const obj = onObj && L.drawing && L.drawing.selectedItem();
    if (obj) { ui.openMenu(A.objectMenu(obj), { x: e.clientX, y: e.clientY }); return; }
    if (L.thesaurus) L.thesaurus.load().catch(() => {});
    const inTable = L.tables && L.tables.inTable();
    const items = ['cut', 'copy', 'paste', '-'];
    const ln = e.target.closest && e.target.closest('[data-link]');
    if (ln) items.push({ label: '&Edit Hyperlink...', cmd: 'hyperlink' }, { label: '&Open Hyperlink', run: () => L.ed.followLink(ln.dataset.link) }, { label: '&Remove Hyperlink', run: () => ui.exec('removeHyperlink') }, '-');
    if (inTable) items.push({ label: '&Insert', sub: ['colLeft', 'colRight', 'rowAbove', 'rowBelow', 'insertCells'] }, { label: '&Delete', sub: ['deleteTable', 'deleteCols', 'deleteRows', 'deleteCells'] }, 'mergeCells', 'splitCells', '-', 'distRows', 'distCols', '-', 'bordersDlg', 'textDirection', { label: 'Cell Alignment', sub: ['cellTop', 'cellMiddle', 'cellBottom'] }, { label: '&AutoFit', sub: ['autofitContents', 'autofitWindow', 'fixedWidth', '-', 'distRows', 'distCols'] }, '-', 'insertCaption', 'tableAutoFormat', 'tableProps', '-');
    const f = L.fields && L.fields.at(E.sel && E.sel.f);
    if (f) items.push('updateField', 'editField', 'toggleCodes', '-');
    const control = L.preserve.controlAt(E.sel && E.sel.f);
    if (control && !control.control.converted && ['date', 'dropDownList', 'comboBox', 'checkbox'].includes(control.control.type)) items.push({ label: 'Content Control &Value...', run: () => L.preserve.controlDialog(control) }, '-');
    if (L.review && L.review.revisionAt && L.review.revisionAt()) items.push('acceptChange', 'rejectChange', '-');
    const cm = L.review && L.review.commentAt && L.review.commentAt();
    if (cm) items.push({ label: '&Edit Comment', run: () => L.review.editComment(cm) }, { label: 'Delete Co&mment', run: () => L.review.deleteComment(cm) }, '-');
    items.push('fontDlg', 'paragraphDlg', 'bulletsDlg', '-', ...(ln ? [] : ['hyperlink']), 'research', { label: 'S&ynonyms', sub: () => A.synonymItems() }, '-', { cmd: 'insertComment', label: 'New Co&mment' });
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
      if (e.button !== 0) return;
      e.preventDefault();
      const map = L.$('#docmap'), w0 = map.offsetWidth, style0 = map.style.width;
      ui.drag(e, {
        move: (ev, d) => { map.style.width = L.clamp(w0 + d.dx, 100, 480) + 'px'; },
        up: (ev, d) => { if (d.moved) LY.fitSizer(); },
        cancel: () => { map.style.width = style0; },
      });
    });
  }

  /* ================= wiring ================= */
  function wire() {
    L.bus.on('sel', () => { A.updateStatus(); ui.refresh(); if (L.rulers) L.rulers.drawSoon(); });
    L.bus.on('sel-done', () => { A.painter.release(); A.highlighter.release(); });
    L.bus.on('objsel', () => { A.updateToolbars(); ui.refresh(); });
    L.bus.on('doc-changed', () => { if (window.VO) VO.changed(doc()); A.updateStatus(); ui.refresh(); if (doc().settings.track || Object.keys(doc().comments).length) A.updateToolbars(); });
    L.bus.on('layout-done', () => { A.updateStatus(); if (L.rulers) L.rulers.drawSoon(); if (L.drawing) L.drawing.placeHandles(); E.placeHandles(); });
    L.bus.on('zoom', () => { ui.refresh(); if (L.rulers) L.rulers.draw(); if (L.drawing) L.drawing.placeHandles(); E.placeHandles(); });
    LY.scroller.addEventListener('scroll', () => { if (L.rulers) L.rulers.drawSoon(); });
    // Ctrl+wheel (a trackpad's pinch too) zooms in 10 % steps, as in Word and the other apps; Web Layout lays the
    // document out again for each step, so it doesn't follow the wheel
    LY.scroller.addEventListener('wheel', (e) => {
      if (!(e.ctrlKey || e.metaKey) || !e.deltaY) return;
      e.preventDefault();
      if (LY.view === 'web') return;
      const p = Math.round(LY.zoom * 100), n = e.deltaY < 0 ? Math.floor(p / 10) * 10 + 10 : Math.ceil(p / 10) * 10 - 10;
      A.zoomTo(L.clamp(n, 10, 500) + '%', { focus: false });
    }, { passive: false });
    window.addEventListener('beforeunload', (e) => { const any = (doc() && doc().dirty) || A.docs.some((w, i) => i !== A.cur && w && w.doc && w.doc.dirty); if (any) { e.preventDefault(); e.returnValue = ''; } });
    window.addEventListener('resize', L.debounce(() => { if (A.view === 'web') LY.render(); else LY.fitSizer(); if (L.rulers) L.rulers.draw(); }, 200));
    document.addEventListener('keydown', (e) => { LY.root.classList.toggle('ctrl', e.ctrlKey || e.metaKey); });
    document.addEventListener('keyup', (e) => { LY.root.classList.toggle('ctrl', e.ctrlKey || e.metaKey); });
  }

  /* ================= format painter ================= */
  const PAINTED = ['font', 'sz', 'b', 'i', 'u', 'uColor', 'strike', 'dstrike', 'color', 'hl', 'caps', 'smallCaps', 'vert', 'spacing', 'shd', 'style'];
  /** the formatting at the selection, as the Format Painter and Copy Formatting carry it */
  A.pickFormat = () => (E.sel ? { rPr: L.clone(E.curRun()), pPr: L.clone(E.sel.f.p.pPr) } : null);
  /** put fmt on the selection, or on the word a click landed in (Word behaviour); false: nothing to paint.
      The character formatting replaces the target's; its links and revisions stay. */
  A.paintFormat = function (fmt) {
    if (!E.sel) return false;
    let [a, b] = E.range();
    if (E.collapsed()) {
      const w = O.wordAt(E.sel.f);
      if (!w || !(w[0].o < a.o && a.o < w[1].o)) return false;
      [a, b] = w;
    }
    const keep = E.sel, src = fmt.rPr;
    E.edit('Format Painter', () => {
      O.formatRange(a, b, (r) => { for (const k of PAINTED) { if (src[k] !== undefined) r[k] = L.clone(src[k]); else delete r[k]; } return r; });
      if (a.p !== b.p || (a.o === 0 && b.o >= D.plen(b.p))) O.setParaProps(E.selectedParas(), () => L.clone(fmt.pPr));
      return keep;
    });
    ui.refresh();
    return true;
  };
  /* the brushes paint when the mouse is let go ('sel-done'), not while a selection is being dragged out */
  A.painter = ui.brush({ el: () => L.$('#scroller'), pick: A.pickFormat, apply: A.paintFormat, hint: 'Select the text to apply the formatting to.', status: (t) => A.status(t) });
  A.highlighter = ui.brush({
    el: () => L.$('#scroller'),
    pick: () => A.lastHighlight,
    apply: (n) => { if (E.collapsed()) return false; E.formatRun({ hl: n === 'none' ? undefined : n }, L.N('Highlight')); return true; },
    hint: 'Select text to highlight. Press Esc to stop.',
    status: (t) => A.status(t),
  });

  /* ================= sample document ================= */
  A.buildSample = function () {
    const d = D.newDoc();
    const P = (text, pPr, rPr) => D.para(text ? [D.text(text, rPr || {})] : [], pPr || {});
    const bl = d.main.blocks;
    bl.length = 0;
    d.props.title = 'Welcome to Quire';
    d.props.creator = 'Quire';
    bl.push(P('Welcome to Quire', { style: 'Title' }));
    bl.push(P('A word processor in the spirit of Word — it opens and saves real .docx files.', { style: 'Subtitle' }, { i: true }));
    bl.push(P('Getting around', { style: 'Heading1' }));
    bl.push(D.para([D.text('Type anywhere on this page. Everything works the way you remember: '), D.text('bold', { b: true }), D.text(', '), D.text('italic', { i: true }), D.text(', '), D.text('underline', { u: 'single' }), D.text(', '), D.text('highlighting', { hl: 'yellow' }), D.text(', '), D.text('colors', { color: 'C00000' }), D.text(', styles, bullets, numbering, tables, headers and footers, footnotes, comments, tracked changes, fields and more.')], { jc: 'both' }));
    bl.push(P('Open a document with File ▸ Open, or drop a .docx onto the window. Save with File ▸ Save (Ctrl+S); File ▸ Save As offers Word Document, Word Template, Web Page, Plain Text, Markdown and PDF.', { num: { id: '1', lvl: 0 } }));
    bl.push(P('Switch views with the buttons in the lower-right corner: Normal, Web Layout, Print Layout, Outline and Reading Layout; the slider beside them zooms.', { num: { id: '1', lvl: 0 } }));
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
    d.hf.h1 = { kind: 'hdr', id: 'h1', blocks: [D.para([D.text('Quire'), D.item('tab'), D.item('tab'), D.text('Sample document')], { style: 'Header' })] };
    const ftr = D.para([D.item('tab'), D.text('Page ')], { style: 'Footer' });
    ftr.runs.push(D.item('fb', { fid: 'f2', instr: ' PAGE ' }), D.item('fs', { fid: 'f2' }), D.text('1'), D.item('fe', { fid: 'f2' }), D.text(' of '), D.item('fb', { fid: 'f3', instr: ' NUMPAGES ' }), D.item('fs', { fid: 'f3' }), D.text('1'), D.item('fe', { fid: 'f3' }));
    d.hf.f1 = { kind: 'ftr', id: 'f1', blocks: [ftr] };
    d.sect.refs = { hdr: { default: 'h1' }, ftr: { default: 'f1' } };
    return d;
  };

  /* ================= init ================= */
  A.init = async function () {
    A.loadOpts();
    if (window.VO && L.icons.app) VO.setIcon(L.icons.app(32));
    /* the labels the page's CSS draws on pages and marks, in the UI's language */
    for (const [v, text] of [['--t-blank', L.t('(This page intentionally left blank by a section break)')], ['--t-same', L.t('Same as Previous')],
      ['--t-object', L.t('[object]')], ['--t-sect', L.t('Section Break (Next Page)')]]) document.documentElement.style.setProperty(v, JSON.stringify(text));
    buildMenus();
    buildToolbars();
    L.ribbonUI.init();
    buildStatus();
    LY.mount(L.$('#scroller'));
    E.init();
    if (L.rulers) L.rulers.mount(L.$('#hruler'), L.$('#vruler'), L.$('#rcorner'));
    if (L.panes) { L.panes.task.mount(L.$('#taskpane')); L.panes.docmap.mount(L.$('#docmap')); if (L.panes.reviewing) L.panes.reviewing.mount(L.$('#revpane')); }
    /* the task pane follows the selected object to its tool set, opening and closing itself (common/ribbon.js) */
    const taskPref = { get: () => A.opts.taskOpen, set: (v) => { A.opts.taskOpen = v; A.saveOpts(); } };
    if (L.panes) L.ribbonUI.follow({ task: L.panes.task, sets: { picture: true, drawing: true, table: { when: () => L.tables.wholeSelected() } }, on: () => A.opts.paneFollows !== false, open: () => A.toggleTask(true), close: () => A.toggleTask(false), pref: taskPref,
      room: () => { const s = LY.scroller; return s.getBoundingClientRect().left + s.clientLeft + s.clientWidth; },
      target: () => { const t = !E.objSel && E.cellSel && E.cellSel.tbl, el = E.objSel ? E.objSel.el : t && LY.root.querySelector(`table.tbl[data-tid="${t.id}"]`); return el ? el.getBoundingClientRect() : null; } });
    setupDrop();
    setupSplitters();
    wire();
    L.bus.emit('app-ready');
    const narrow = window.innerWidth < 820;
    if (narrow) A.opts.taskOpen = false;
    L.$('#taskpane').hidden = !A.opts.taskOpen;
    LY.zoom = narrow ? Math.max(0.4, (window.innerWidth - 40) / 816) : A.opts.zoom || 1;
    /* opened from the Start Center: a template, or blank (a new document, or to open a file into); else the sample */
    const launch = window.VO && VO.launch;
    if (launch && launch.n) A.untitled = launch.n;   // the desktop app's number for it
    A.loadDoc(launch && launch.template && L.templates ? L.templates.build(launch.template) : launch ? blank() : A.buildSample(), 'Document' + A.untitled);
    A.view = 'x';
    A.setView(A.opts.view && ['print', 'normal', 'web', 'outline'].includes(A.opts.view) ? A.opts.view : 'print');
    if (L.panes && A.opts.taskOpen) L.panes.task.show('getting-started');
    A.applyOpts();
    if (A.opts.docMap) A.toggleDocMap(true);
    ui.refresh();
    if (window.VO) VO.attach('quire', {
      async open(f, o) {
        await A.openFile(f);
        /* a document a page made (desktop app): never saved */
        if (o && o.made && doc()) { A.saved = false; A.fileName = o.name.replace(/\.(docx|docm|dotx|dotm)$/i, ''); A.updateTitle(); }
        if (o?.draft && doc()) L.opc.recoverLosses(doc(), o.lossState);
        /* an unsaved version: the changes are not saved yet, and a never-saved document still needs Save As */
        if (o && o.draft && doc()) { doc().dirty = true; A.saved = !!o.saved; A.fileName = o.name.replace(/\.(docx|docm|dotx|dotm|md|markdown|zip)$/i, ''); A.updateTitle(); A.status('Recovered the unsaved changes. Save the document to keep them.'); ui.refresh(); }
      },
      /* a Markdown document's unsaved version is Markdown too, so recovering it loses nothing */
      async snapshot(d) {
        const t = winOf(d).fileType;
        if (t === 'md' && L.mdio) return new Blob([L.mdio.write(d).text], { type: 'text/markdown' });
        if (t === 'mdzip' && L.mdio) return L.mdio.writeZip(d, winOf(d).fileName);
        const format = L.opc.formats[t] ? t : d.ooxmlFormat || 'docx';
        const blob = await L.docx.write(d, { stats: A.docStats(), format });
        return d.openPassword && L.officeCrypto ? new Blob([await L.officeCrypto.encrypt(new Uint8Array(await blob.arrayBuffer()), d.openPassword)], { type: blob.type }) : blob;
      },
      docInfo(d) { const w = winOf(d); const ext = { md: '.md', mdzip: '.zip' }[w.fileType] || '.' + (L.opc.formats[w.fileType] ? w.fileType : d.ooxmlFormat || 'docx'); return { name: (w.fileName || 'Document') + ext, lossState: L.opc.lossState(d) }; },
      isDirty: (d) => !!(d && d.dirty),
      current: () => doc(),
      autosave: () => A.opts.autoRecover !== false,
      showRecovery() { if (!L.panes) return; A.opts.taskOpen = true; L.$('#taskpane').hidden = false; L.panes.task.show('recovery'); },
      pristine,
      close: confirmDiscard,
      thumb: pageThumb,
      message: (t) => ui.msg(t, { icon: 'warn' }),
      templates: () => (L.templates ? L.templates.list().map(({ id, name }) => ({ id, name })) : []),
    });
  };
  const winOf = (d) => (d === doc() ? A : A.docs.find((x) => x && x.doc === d) || A);
  /** Recent Files picture: the first page, drawn as for PDF (print layout only) */
  async function pageThumb() {
    const p = LY.view === 'print' && LY.pages && LY.pages[0];
    return p ? A.rasterizePage(p.el, 400 / p.el.offsetWidth, docCSS(), await mediaDataMap(p.el)) : null;
  }
  /* once the page and the UI's dictionary are here (common/i18n.js) */
  const boot = () => (window.VO && VO.i18n ? VO.i18n.whenReady(() => A.init()) : A.init());
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
