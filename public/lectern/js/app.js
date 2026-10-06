/* Lectern — application controller. */
(function () {
  'use strict';
  const L = window.L;
  const { h } = L;
  const ui = L.ui;
  const M = L.model;
  const T = L.txt;
  const E = L.ed;
  const TE = L.te;

  const DEFAULT_OPTS = {
    startupPane: true, layoutPaneOnNew: true, statusBar: true, ruler: false, showPopup: true, showPopbar: true, endBlack: true,
    smartQuotes: true, autocorrect: true, capSentence: true, spell: true, autoPreview: true, gray: false, bw: false, outlinePlain: false, alignToSlide: false,
    toolbars: { standard: true, formatting: true, drawing: true, outlining: false, picture: 'auto', tables: 'auto', wordart: 'auto' },
    grid: { show: false, snap: true, size: 6 }, leftW: 196, notesH: 92, leftOpen: true, taskOpen: true,
  };
  const A = (L.app = Object.assign(L.app || {}, {
    view: 'normal', prevView: null, fileName: 'Presentation1', fileType: 'pptx', saved: false,
    slideSel: new Set([0]), clipHistory: [], clip: null, focusArea: null,
    printOpts: { what: 'slides', per: 6, color: 'color', frame: false, hidden: false, range: null },
    opts: L.clone(DEFAULT_OPTS),
    lastFill: 'accent1', lastLineColor: 'tx1', lastFontColor: '#FF0000',
  }));

  /* ================= options ================= */
  A.loadOpts = () => {
    const o = L.store.get('opts', null);
    A.opts = L.deepMerge(L.clone(DEFAULT_OPTS), o || {});
    E.grid = Object.assign(E.grid, A.opts.grid);
  };
  A.saveOpts = () => { A.opts.grid = { show: E.grid.show, snap: E.grid.snap, size: E.grid.size }; L.store.set('opts', A.opts); };
  A.applyOpts = () => {
    L.$('#statusbar').hidden = !A.opts.statusBar;
    A.updateToolbars();
    E.layout();
    E.render();
  };

  /* ================= UI construction ================= */
  function buildMenus() {
    const autoShapeSub = (group) => ({ custom: (close) => shapeGrid(L.geom.MENU[group], close) });
    const menus = [
      { label: '&File', items: ['new', 'open', 'close', '-', 'save', 'saveAs', 'saveWeb', '-', 'pageSetup', 'printPreview', 'print', '-', 'properties', '-', 'exitApp'] },
      { label: '&Edit', items: ['undo', 'redo', '-', 'cut', 'copy', 'officeClipboard', 'paste', 'pasteSpecial', '-', 'clear', 'selectAll', 'duplicate', 'deleteSlide', '-', 'find', 'replace'] },
      {
        label: '&View', items: () => ['viewNormal', 'viewSorter', 'showFromStart', 'viewNotes', '-',
          { label: '&Master', sub: ['viewMaster'] },
          { label: 'Co&lor/Grayscale', sub: ['viewColor', 'viewGray', 'viewBW'] }, '-',
          'taskPane', 'slidesPane', { label: '&Toolbars', sub: ['tb_standard', 'tb_formatting', 'tb_drawing', 'tb_outlining', 'tb_picture', 'tb_tables', 'tb_wordart'] }, 'ruler', 'gridGuides', '-', 'headerFooter', '-', 'zoomDlg'],
      },
      {
        label: '&Insert', items: ['newSlide', 'duplicateSlide', '-', 'insertSlideNumber', 'insertDateTime', 'insertSymbol', '-', 'slidesFromFiles', 'slidesFromOutline', '-',
          { label: '&Picture', icon: 'picture', sub: ['insertClipArt', 'insertPicture', 'photoAlbum', '-', { label: '&AutoShapes', icon: 'autoshapes', run: () => openAutoShapes() }, 'insertWordArt'] },
          'insertDiagram', 'textBox', 'insertMedia', 'insertChart', 'insertTable', '-', 'hyperlink'],
      },
      { label: 'F&ormat', items: ['fontDlg', 'bulletsDlg', { label: '&Alignment', sub: ['alignLeft', 'alignCenter', 'alignRight', 'justify'] }, 'lineSpacing', 'changeCase', 'replaceFonts', '-', 'slideDesign', 'slideLayout', 'background', '-', 'formatObject'] },
      { label: '&Tools', items: ['spelling', '-', 'autocorrectDlg', '-', 'optionsDlg'] },
      {
        label: 'Sli&de Show', items: () => ['showFromStart', 'setupShow', 'rehearse', '-',
          { label: 'Act&ion Buttons', sub: [{ custom: (close) => shapeGrid(L.geom.MENU['Action Buttons'], close) }] }, 'actionSettings', '-',
          'animSchemes', 'customAnim', 'transitionPane', '-', 'hideSlide'],
      },
      { label: '&Window', items: () => [{ label: '&1 ' + A.fileName, checked: () => true, run: () => {} }] },
      { label: '&Help', items: ['help', 'gettingStarted', '-', 'about'] },
    ];
    void autoShapeSub;
    const row = L.$('#menurow');
    A.menuBar = ui.menuBar(row, menus);
    const q = h('input', { type: 'text', id: 'helpq', placeholder: 'Type a question for help', 'aria-label': 'Type a question for help' });
    q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { L.panes.helpQuery = q.value; L.panes.task.show('help'); q.value = ''; E.refocus(); } e.stopPropagation(); });
    row.appendChild(h('div', { class: 'helpbox' }, q, h('span', { class: 'dd-arrow' })));
  }

  function shapeGrid(list, close) {
    const g = h('div', { class: 'shape-grid' });
    for (const k of list) {
      const p = L.geom.get(k);
      const b = h('button', { type: 'button', 'data-tip': p.label, 'aria-label': p.label, html: L.geom.preview(k, 20) });
      b.addEventListener('click', () => { close(); E.setTool(p.isLine ? { kind: 'line', geom: k } : { kind: 'shape', geom: k }); });
      g.appendChild(b);
    }
    return g;
  }
  function openAutoShapes(at) {
    const r = at || (L.$('[data-as]') ? L.$('[data-as]').getBoundingClientRect() : { left: 100, bottom: 100 });
    ui.openMenu([
      { label: '&Lines', icon: 'line', sub: ['drawLine', 'drawArrow', 'drawDblArrow', 'drawCurve', 'drawFreeform', 'drawScribble'] },
      { label: 'Co&nnectors', icon: 'elbow', sub: [{ label: 'Straight Connector', icon: 'line', run: () => E.setTool({ kind: 'line', geom: 'straightConnector1' }) }, 'drawElbow', 'drawCurvedConn', { label: 'Elbow Arrow Connector', icon: 'elbow', run: () => E.setTool({ kind: 'line', geom: 'bentConnector3', arrows: 'end' }) }] },
      ...['Basic Shapes', 'Block Arrows', 'Flowchart', 'Stars and Banners', 'Callouts', 'Action Buttons'].map((g) => ({ label: g.replace(/^(\w)/, '&$1'), icon: { 'Basic Shapes': 'autoshapes', 'Block Arrows': 'arrowStyle', Flowchart: 'diagram', 'Stars and Banners': 'animation', Callouts: 'comment', 'Action Buttons': 'actionSettings' }[g], sub: [{ custom: (close) => shapeGrid(L.geom.MENU[g], close) }] })),
      '-', { label: '&More AutoShapes...', icon: 'clipart', run: () => L.panes.task.show('clipart') },
    ], { left: r.left, bottom: r.bottom }, { selectFirst: !at });
  }
  A.openAutoShapes = openAutoShapes;

  const WEIGHTS = L.dlg.LINE_WEIGHTS;
  function listPanel(items, close) {
    const el = h('div', { class: 'style-list' });
    for (const it of items) {
      const b = h('button', { type: 'button', html: it.html || '' });
      if (it.label) b.appendChild(h('span', { text: it.label }));
      b.addEventListener('click', () => { close(); it.run(); });
      el.appendChild(b);
    }
    return el;
  }
  const lineSVG = (w, dash, extra) => `<svg width="90" height="14" viewBox="0 0 90 14"><path d="M2 7h86" stroke="#000" stroke-width="${w}" ${dash ? `stroke-dasharray="${dash}"` : ''} fill="none"/>${extra || ''}</svg>`;
  function lineStyleMenu(r) {
    ui.openMenu([{ custom: (close) => listPanel(WEIGHTS.map((w) => ({ html: lineSVG(Math.max(0.5, w * 1.33)), label: w + ' pt', run: () => A.setLineProp({ w }) })).concat([{ label: 'More Lines...', run: () => ui.exec('formatObject', 'Colors and Lines') }]), close) }], r);
  }
  function dashMenu(r) {
    ui.openMenu([{ custom: (close) => listPanel(L.dlg.DASHES.map(([k, n]) => ({ html: lineSVG(2, (L.render.dashArray(k, 2) || '').replace(/,/g, ' ')), label: n, run: () => A.setLineProp({ dash: k }) })), close) }], r);
  }
  const ARROW_PRESETS = [
    [null, null], [null, 'triangle'], ['triangle', null], ['triangle', 'triangle'], [null, 'arrow'], [null, 'stealth'], [null, 'oval'], [null, 'diamond'], ['oval', 'triangle'], ['oval', 'oval'], ['diamond', 'diamond'],
  ];
  function arrowMenu(r) {
    const head = (t, x, dir) => !t ? '' : t === 'oval' ? `<circle cx="${x}" cy="7" r="3"/>` : t === 'diamond' ? `<path d="M${x - 4} 7l4-4 4 4-4 4z"/>` : t === 'arrow' ? `<path d="M${x - dir * 6} 2l${dir * 6} 5-${dir * 6} 5" fill="none" stroke="#000" stroke-width="1.5"/>` : `<path d="M${x} 7l${-dir * 8} -4v8z"/>`;
    ui.openMenu([{ custom: (close) => listPanel(ARROW_PRESETS.map(([a, b]) => ({ html: `<svg width="90" height="14" viewBox="0 0 90 14" fill="#000"><path d="M6 7h78" stroke="#000" stroke-width="1.5"/>${head(a, 4, -1)}${head(b, 86, 1)}</svg>`, run: () => A.setLineProp({ head: a, tail: b }) })).concat([{ label: 'More Arrows...', run: () => ui.exec('formatObject', 'Colors and Lines') }]), close) }], r);
  }
  const SHADOWS = [null, [3, 3], [-3, 3], [3, -3], [-3, -3], [0, 4], [4, 0], [6, 6], [-6, 6], [2, 2, 4], [0, 6, 6], [5, 5, 3]];
  function shadowMenu(r) {
    ui.openMenu([{
      custom: (close) => {
        const g = h('div', { class: 'shape-grid', style: 'grid-template-columns:repeat(4,30px)' });
        SHADOWS.forEach((s) => {
          const b = h('button', { type: 'button', style: 'width:30px;height:30px', 'aria-label': s ? 'Shadow style' : 'No Shadow', 'data-tip': s ? 'Shadow Style' : 'No Shadow', html: s ? `<svg width="24" height="24"><rect x="${6 + s[0] / 1.5}" y="${6 + s[1] / 1.5}" width="12" height="12" fill="#888" ${s[2] ? 'filter="blur(1px)"' : ''}/><rect x="6" y="6" width="12" height="12" fill="#fff" stroke="#000"/></svg>` : '<svg width="24" height="24"><rect x="6" y="6" width="12" height="12" fill="#fff" stroke="#000"/><path d="M3 21L21 3" stroke="#d33" /></svg>' });
          b.addEventListener('click', () => { close(); A.setShadow(s ? { c: '#000000', a: 0.5, dx: s[0], dy: s[1], blur: s[2] || 0 } : null); });
          g.appendChild(b);
        });
        return g;
      },
    }], r);
  }
  function drawMenu(r) {
    ui.openMenu([
      'group', 'ungroup', 'regroup', '-',
      { label: '&Order', sub: ['bringFront', 'sendBack', 'bringForward', 'sendBackward'] }, 'gridGuides',
      { label: 'Nud&ge', sub: ['nudgeUp', 'nudgeDown', 'nudgeLeft', 'nudgeRight'] },
      { label: '&Align or Distribute', sub: ['alignL', 'alignC', 'alignR', 'alignT', 'alignM', 'alignB', '-', 'distH', 'distV', '-', 'alignToSlide'] },
      { label: 'Rotate or Fl&ip', sub: ['freeRotate', 'rotateLeft', 'rotateRight', '-', 'flipH', 'flipV'] }, '-',
      { label: 'C&hange AutoShape', enabled: () => E.selected().some((s) => s.type === 'shape' || s.type === 'text'), sub: ['Basic Shapes', 'Block Arrows', 'Flowchart', 'Stars and Banners', 'Callouts', 'Action Buttons'].map((g) => ({ label: g, sub: [{ custom: (close) => { const el = shapeGrid(L.geom.MENU[g], close); L.$$('button', el).forEach((b, i) => { const nb = b.cloneNode(true); nb.addEventListener('click', () => { close(); A.changeAutoShape(L.geom.MENU[g][i]); }); b.replaceWith(nb); }); return el; } }] })) },
      'setDefaults',
    ], r);
  }
  function colorDD(kind) {
    return (r, anchor) => ui.colorMenu(anchor || r, { mode: kind === 'font' ? 'font' : kind, effects: kind === 'fill' }, async (v) => {
      if (kind === 'fill') {
        if (v && v.none) A.applyFill({ t: 'none' });
        else if (v && v.effects) { const s = E.primary(); const f = await L.dlg.fillEffects(s && s.fill && s.fill.t !== 'none' ? s.fill : { t: 'solid', c: A.lastFill }); if (f) A.applyFill(f); }
        else if (typeof v === 'string') { A.lastFill = v; A.applyFill(v); }
      } else if (kind === 'line') {
        if (v && v.none) A.applyLineColor(null);
        else if (typeof v === 'string') { A.lastLineColor = v; A.applyLineColor(v); }
      } else {
        if (v && v.auto) A.applyFontColor(null);
        else if (typeof v === 'string') { A.lastFontColor = v; A.applyFontColor(v); }
      }
      ui.refresh();
    });
  }
  function buildToolbars() {
    const host = L.$('#toolbars');
    const fontCombo = ui.combo({ id: 'tb-font', width: 136, tip: 'Font', options: () => L.FONT_LIST, preview: (v) => `font-family:${L.fontStack(v)};font-size:13px`, listCls: 'fonts', value: () => { const s = textState(); return s ? s.font : ''; }, enabled: () => !!textState(), onChange: (v) => A.setFontName(v) });
    const sizeCombo = ui.combo({ id: 'tb-size', width: 44, tip: 'Font Size', options: () => L.SIZE_LIST, value: () => { const s = textState(); return s && s.sz ? String(L.round(s.sz, 1)) : ''; }, enabled: () => !!textState(), onChange: (v) => A.setFontSize(parseFloat(v)) });
    const zoomCombo = ui.combo({ id: 'tb-zoom', width: 52, tip: 'Zoom', options: () => ['Fit', '400%', '200%', '150%', '100%', '75%', '66%', '50%', '33%', '25%'], value: () => (E.zoom === 'fit' ? 'Fit' : E.zoomPct() + '%'), onChange: (v) => { if (/fit/i.test(v)) E.setZoom('fit'); else { const n = parseFloat(v); if (n > 0) E.setZoom(n / 100); } } });
    A.toolbarsEl = {
      standard: ui.toolbar('standard', 'Standard', ['new', 'open', 'save', '|', 'print', 'printPreview', 'spelling', '|', 'cut', 'copy', 'paste', 'painter', '|', 'undo', 'redo', '|', 'insertChart', 'insertTable', ui.tbButton('tb_tables', { icon: 'tablesBorders' }), 'hyperlink', '|', 'expandAll', 'showFormatting', 'showGrid', ui.tbSplit('viewGray', (r) => ui.openMenu(['viewColor', 'viewGray', 'viewBW'], r)), zoomCombo, 'help']),
      formatting: ui.toolbar('formatting', 'Formatting', [fontCombo, sizeCombo, 'bold', 'italic', 'underline', 'shadowText', '|', 'alignLeft', 'alignCenter', 'alignRight', '|', 'numbering', 'bullets', '|', 'growFont', 'shrinkFont', '|', 'promote', 'demote', '|', ui.tbSplit('fontColor', colorDD('font')), '|', ui.tbButton('slideDesign', { label: 'Design' }), ui.tbButton('newSlide', { label: 'New Slide' })]),
      outlining: ui.toolbar('outlining', 'Outlining', ['promote', 'demote', 'moveParaUp', 'moveParaDown', '|', 'summarySlide', 'showFormatting']),
      picture: ui.toolbar('picture', 'Picture', ['insertPicture', ui.tbDrop('', 'colorPic', (r) => ui.openMenu([{ label: '&Automatic', run: () => A.picMode('') }, { label: '&Grayscale', run: () => A.picMode('gray') }, { label: '&Black & White', run: () => A.picMode('bw') }, { label: '&Washout', run: () => A.picMode('wash') }], r), 'Color'), 'picMoreContrast', 'picLessContrast', 'picMoreBright', 'picLessBright', '|', 'picCrop', 'rotateLeft', ui.tbDrop('', 'lineStyle', lineStyleMenu, 'Line Style'), '|', ui.tbButton('formatObject', { icon: 'formatPicture' }), 'picTransparent', 'picReset']),
      tables: ui.toolbar('tables', 'Tables and Borders', [ui.tbDrop('', 'dashStyle', (r) => ui.openMenu([{ custom: (close) => listPanel(L.dlg.DASHES.map(([k, n]) => ({ html: lineSVG(2, (L.render.dashArray(k, 2) || '').replace(/,/g, ' ')), label: n, run: () => { A.tblBorder.dash = k; } })), close) }], r), 'Border Style'),
        ui.tbDrop('', 'lineStyle', (r) => ui.openMenu([{ custom: (close) => listPanel(WEIGHTS.map((w) => ({ html: lineSVG(Math.max(0.5, w * 1.33)), label: w + ' pt', run: () => { A.tblBorder.w = w; } })), close) }], r), 'Border Width'),
        ui.tbDrop('', 'lineColor', (r, b) => ui.colorMenu(b, { mode: 'plain', grid: true }, (v) => { if (typeof v === 'string') A.tblBorder.c = v; }), 'Border Color'),
        ui.tbDrop('', 'bordersAll', (r) => ui.openMenu([{ custom: (close) => { const g = h('div', { class: 'shape-grid', style: 'grid-template-columns:repeat(4,24px)' }); for (const [k, ic, n] of [['all', 'bordersAll', 'All Borders'], ['out', 'bordersOut', 'Outside Borders'], ['in', 'bordersIn', 'Inside Borders'], ['none', 'bordersNone', 'No Border'], ['t', 'bordersTop', 'Top Border'], ['b', 'bordersBottom', 'Bottom Border'], ['l', 'bordersLeft', 'Left Border'], ['r', 'bordersRight', 'Right Border']]) { const bt = h('button', { type: 'button', 'data-tip': n, 'aria-label': n, html: L.icons.get(ic) }); bt.addEventListener('click', () => { close(); A.tableBorders(k); }); g.appendChild(bt); } return g; } }], r), 'Borders'),
        ui.tbDrop('', 'fillColor', (r, b) => ui.colorMenu(b, { mode: 'fill', effects: true }, async (v) => { if (v && v.none) A.tableFill({ t: 'none' }); else if (v && v.effects) { const f = await L.dlg.fillEffects({ t: 'solid', c: 'accent1' }); if (f) A.tableFill(f); } else if (typeof v === 'string') A.tableFill({ t: 'solid', c: v, a: 1 }); }), 'Fill Color'),
        '|', ui.tbDrop('Table', null, (r) => ui.openMenu(['insertTable', '-', 'colLeft', 'colRight', 'rowAbove', 'rowBelow', '-', 'deleteCols', 'deleteRows', '-', 'mergeCells', 'splitCell', '-', 'selectTable', '-', 'tableBordersFill'], r), 'Table'),
        'mergeCells', 'splitCell', '|', 'cellTop', 'cellMiddle', 'cellBottom', '|', 'distRows', 'distCols']),
      wordart: ui.toolbar('wordart', 'WordArt', ['insertWordArt', ui.tbButton('wordartEdit', { label: 'Edit Text...' }), 'wordartGallery', ui.tbButton('formatObject', { icon: 'formatPicture' }),
        ui.tbDrop('', 'waShape', (r) => ui.openMenu(L.render.WARPS.map(([k, n]) => ({ label: n, run: () => A.waSet({ warp: k }), checked: () => { const s = E.primary(); return !!(s && s.wa && (s.wa.warp || 'textPlain') === k); } })), r), 'WordArt Shape'),
        'waSameHeight', 'waVertical',
        ui.tbDrop('', 'alignCenter', (r) => ui.openMenu([['l', '&Left Align'], ['ctr', '&Center'], ['r', '&Right Align']].map(([k, n]) => ({ label: n, run: () => A.waSet({ algn: k }) })), r), 'WordArt Alignment'),
        ui.tbDrop('', 'charSpacing', (r) => ui.openMenu([[-3, '&Very Tight'], [-1.5, '&Tight'], [0, '&Normal'], [2, '&Loose'], [4, 'V&ery Loose']].map(([k, n]) => ({ label: n, run: () => A.waSet({ spc: k }) })), r), 'WordArt Character Spacing')]),
      master: ui.toolbar('master', 'Slide Master View', ['insertTitleMaster', 'deleteTitleMaster', '|', ui.tbButton('closeMaster', { label: 'Close Master View' })]),
      sorter: ui.toolbar('sorter', 'Slide Sorter', ['hideSlide', 'rehearse', '|', 'summarySlide', 'speakerNotes', '|', ui.tbButton('transitionPane', { label: 'Transition' }), ui.tbButton('slideDesign', { label: 'Design' }), ui.tbButton('newSlide', { label: 'New Slide' })]),
    };
    const row1 = h('div', { class: 'tb-row' }, A.toolbarsEl.standard);
    const row2 = h('div', { class: 'tb-row' }, A.toolbarsEl.formatting, A.toolbarsEl.outlining);
    const row3 = h('div', { class: 'tb-row' }, A.toolbarsEl.picture, A.toolbarsEl.tables, A.toolbarsEl.wordart, A.toolbarsEl.master, A.toolbarsEl.sorter);
    host.append(row1, row2, row3);
    A.tbRows = [row1, row2, row3];
    /* drawing toolbar */
    const asBtn = ui.tbDrop('A&utoShapes', null, (r) => openAutoShapes(r), 'AutoShapes');
    asBtn.setAttribute('data-as', '1');
    A.toolbarsEl.drawing = ui.toolbar('drawing', 'Drawing', [ui.tbDrop('D&raw', null, drawMenu, 'Draw'), 'select', '|', asBtn, 'drawLine', 'drawArrow', 'drawRect', 'drawOval', 'textBox', 'insertWordArt', 'insertDiagram', 'insertClipArt', 'insertPicture', '|',
      ui.tbSplit('fillColor', colorDD('fill')), ui.tbSplit('lineColor', colorDD('line')), ui.tbSplit('fontColor', colorDD('font')), '|',
      ui.tbDrop('', 'lineStyle', lineStyleMenu, 'Line Style'), ui.tbDrop('', 'dashStyle', dashMenu, 'Dash Style'), ui.tbDrop('', 'arrowStyle', arrowMenu, 'Arrow Style'), ui.tbDrop('', 'shadowStyle', shadowMenu, 'Shadow Style')]);
    L.$('#drawrow').appendChild(A.toolbarsEl.drawing);
    /* view buttons */
    const vb = L.$('#viewbar');
    vb.append(ui.tbButton('viewNormal'), ui.tbButton('viewSorter'), ui.tbButton('showFromCurrent'));
  }
  A.tblBorder = { c: 'tx1', w: 1, dash: 'solid' };
  let tsCache = null, tsTick = 0;
  function textState() {
    const now = performance.now();
    if (now - tsTick > 30) { tsCache = (TE.active() || L.fmt.textShapes().length) ? L.fmt.state() : null; tsTick = now; }
    return tsCache;
  }
  L.bus.on('ui-refresh', () => { tsTick = 0; });

  A.tbShown = (id) => { const v = A.opts.toolbars[id]; return v === true || (v === 'auto' && ctxToolbar(id)); };
  function ctxToolbar(id) {
    const sel = E.selected();
    if (id === 'picture') return sel.some((s) => s.type === 'image');
    if (id === 'tables') return sel.some((s) => s.type === 'table');
    if (id === 'wordart') return sel.some((s) => s.type === 'wordart');
    return false;
  }
  A.toggleToolbar = (id) => { const v = A.opts.toolbars[id]; A.opts.toolbars[id] = v === true ? (['picture', 'tables', 'wordart'].includes(id) ? false : false) : true; A.saveOpts(); A.updateToolbars(); };
  A.updateToolbars = function () {
    if (!A.toolbarsEl) return;
    for (const id of ['standard', 'formatting', 'drawing', 'outlining', 'picture', 'tables', 'wordart']) A.toolbarsEl[id].hidden = !A.tbShown(id) || (A.view === 'preview');
    A.toolbarsEl.outlining.hidden = !(A.opts.toolbars.outlining === true || (L.panes.left.tab === 'outline' && A.view === 'normal' && A.opts.toolbars.outlining !== false && false));
    A.toolbarsEl.master.hidden = A.view !== 'master';
    A.toolbarsEl.sorter.hidden = A.view !== 'sorter';
    A.tbRows.forEach((r) => { r.hidden = Array.from(r.children).every((c) => c.hidden); });
    L.$('#drawrow').hidden = !A.tbShown('drawing') || A.view === 'sorter' || A.view === 'preview';
    E.layout();
  };

  /* ================= status bar ================= */
  let statusTimer = 0;
  function buildStatus() {
    const sb = L.$('#statusbar');
    A.sbSlide = h('div', { class: 'sb-cell fit', 'aria-live': 'polite' });
    A.sbMsg = h('div', { class: 'sb-cell grow' });
    A.sbDesign = h('div', { class: 'sb-cell fit' });
    A.sbLang = h('div', { class: 'sb-cell fit', text: 'English (U.S.)' });
    const zr = h('input', { type: 'range', min: 10, max: 400, value: 100, id: 'sb-zoom', 'aria-label': 'Zoom' });
    zr.addEventListener('input', () => E.setZoom(zr.value / 100));
    A.sbZoom = h('div', { class: 'sb-cell fit sb-zoom' }, h('span', { text: '100%' }), zr);
    sb.append(A.sbSlide, A.sbDesign, A.sbMsg, A.sbLang, A.sbZoom);
    L.bus.on('zoom', () => { A.sbZoom.firstChild.textContent = E.zoomPct() + '%'; zr.value = E.zoomPct(); ui.refresh(); });
  }
  A.status = (msg) => { if (!A.sbMsg) return; A.sbMsg.textContent = msg; clearTimeout(statusTimer); statusTimer = setTimeout(() => { A.sbMsg.textContent = ''; }, 3000); };
  A.updateStatus = function () {
    if (!A.sbSlide) return;
    const n = L.pres.slides.length;
    if (A.view === 'master') A.sbSlide.textContent = E.masterKind === 'title' ? 'Title Master' : 'Slide Master';
    else if (A.view === 'sorter') A.sbSlide.textContent = 'Slide Sorter';
    else A.sbSlide.textContent = n ? `Slide ${E.idx + 1} of ${n}` : 'No slides';
    const d = E.design();
    A.sbDesign.textContent = d ? d.name : '';
  };
  A.updateTitle = () => { L.$('#titletext').textContent = `${L.APP} - [${A.fileName}]`; document.title = `${A.fileName} - Lectern 2003`; };

  /* ================= views ================= */
  A.setView = function (v) {
    if (v === A.view && v !== 'master') return;
    if (TE.active()) TE.end();
    if (L.panes.left.tab === 'outline') L.panes.left.commitOutline();
    if (v === 'preview') A.prevView = A.view;
    if (A.view === 'master' && v !== 'master') { E.syncMaster(); E.masterSlide = null; }
    A.view = v;
    const left = L.$('#leftpane'), center = L.$('#center'), split = L.$('#split-left');
    const alt = A.altView;
    L.panes.sorter.el.classList.toggle('on', v === 'sorter');
    L.panes.notesView.el.classList.toggle('on', v === 'notes');
    L.panes.preview.el.classList.toggle('on', v === 'preview');
    const showLeft = (v === 'normal' || v === 'master') && A.opts.leftOpen;
    left.hidden = !showLeft; split.hidden = !showLeft;
    center.hidden = !(v === 'normal' || v === 'master');
    alt.hidden = !(v === 'sorter' || v === 'notes');
    L.$('#notes').hidden = v === 'master';
    L.$('#split-notes').hidden = v === 'master';
    if (v === 'master') { E.view = 'master'; E.masterKind = 'slide'; E.buildMaster(); E.sel = []; }
    else E.view = 'normal';
    if (v === 'sorter') { A.slideSel = new Set([E.idx]); requestAnimationFrame(() => { L.panes.sorter.render(); L.panes.sorter.el.focus({ preventScroll: true }); }); }
    if (v === 'notes') requestAnimationFrame(() => L.panes.notesView.render());
    if (v === 'preview') requestAnimationFrame(() => L.panes.preview.render());
    if (v === 'normal' || v === 'master') { L.panes.left.render(); E.render(); L.panes.notes.render(); }
    A.updateToolbars();
    A.updateStatus();
    L.panes.task.refresh('view');
    ui.refresh();
  };
  A.setColorMode = (m) => { A.opts.gray = m === 'gray'; A.opts.bw = m === 'bw'; A.saveOpts(); E.render(); };
  A.toggleTask = (on) => {
    const el = L.$('#taskpane');
    const show = on == null ? el.hidden : on;
    const changed = el.hidden === show;
    el.hidden = !show;
    A.opts.taskOpen = show;
    if (show && on == null && !L.panes.task.current) L.panes.task.show('getting-started');
    else if (show && changed && on == null) L.panes.task.render();
    if (changed) { E.layout(); if (A.view === 'sorter') L.panes.sorter.render(); }
    ui.refresh();
  };
  A.taskOpen = () => !L.$('#taskpane').hidden;
  A.toggleLeft = (on) => { const show = on == null ? !A.opts.leftOpen : on; A.opts.leftOpen = show; const v = A.view === 'normal' || A.view === 'master'; L.$('#leftpane').hidden = !(show && v); L.$('#split-left').hidden = !(show && v); if (show) L.panes.left.render(); E.layout(); ui.refresh(); };
  A.leftOpen = () => !!A.opts.leftOpen;

  /* ================= presentation lifecycle ================= */
  A.loadPres = function (pres, name, opts) {
    if (TE.active()) TE.detach(true);
    L.pres = pres;
    L.hist.clear();
    E.idx = 0; E.sel = []; E.tool = null; E.view = 'normal'; E.masterSlide = null;
    A.slideSel = new Set([0]);
    A.fileName = name || 'Presentation1';
    A.saved = !!(opts && opts.saved);
    A.view = 'x';
    A.setView('normal');
    A.updateTitle();
    L.bus.emit('slides-changed');
    E.refocus();
  };
  async function confirmDiscard() {
    if (!L.hist.dirty) return true;
    const r = await ui.msg(`Do you want to save the changes you made to ${A.fileName}?`, { icon: 'warn', buttons: ['&Yes', '&No', 'Cancel'] });
    if (r === 0) { await A.save(); return true; }
    return r === 1;
  }
  A.newPresentation = async function () {
    if (!(await confirmDiscard())) return;
    A.untitled = (A.untitled || 1) + 1;
    A.loadPres(M.newPresentation(Object.assign({ design: 'default' }, A.newSize())), 'Presentation' + A.untitled);
    if (A.opts.startupPane) L.panes.task.show('layout');
  };
  /** slide size for new presentations (Page Setup ▸ "Use for new presentations") */
  A.newSize = () => { const z = A.opts.slideSize; return z && z.w > 0 && z.h > 0 ? { w: z.w, h: z.h } : {}; };
  A.closePresentation = async function () {
    if (!(await confirmDiscard())) return;
    A.loadPres(M.newPresentation(Object.assign({ design: 'default' }, A.newSize())), 'Presentation1');
  };
  A.openDialog = async function () {
    const files = await L.pickFiles('.pptx,.ppsx,.potx,.pptm,.potm,.ppsm,.txt,application/vnd.openxmlformats-officedocument.presentationml.presentation');
    if (files[0]) A.openFile(files[0]);
  };
  A.openFile = async function (file) {
    if (/\.txt$/i.test(file.name)) { if (!(await confirmDiscard())) return; const text = await L.readAsText(file); const pres = M.newPresentation(Object.assign({ design: 'default', empty: true }, A.newSize())); A.loadPres(pres, file.name.replace(/\.txt$/i, '')); A.outlineToSlides(text, 0); return; }
    if (/\.ppt$/i.test(file.name)) { ui.msg(`${file.name} is in the older binary PowerPoint 97–2003 format. Open it in PowerPoint, Keynote, Google Slides or LibreOffice and save it as .pptx, then open the .pptx here.`, { icon: 'warn' }); return; }
    if (!(await confirmDiscard())) return;
    ui.busy(true, `Opening ${file.name}...`);
    try {
      const buf = await L.readAsArrayBuffer(file);
      const pres = await L.pptx.read(buf, { progress: (i, n) => ui.busy(true, `Opening ${file.name}... (slide ${i} of ${n})`) });
      if (!pres.slides.length) pres.slides.push(M.newSlide(pres, 'title', Object.keys(pres.designs)[0]));
      A.loadPres(pres, file.name.replace(/\.(pptx|ppsx|potx|pptm|ppsm|potm)$/i, ''), { saved: true });
      A.fileType = /\.ppsx$/i.test(file.name) ? 'ppsx' : /\.potx$/i.test(file.name) ? 'potx' : 'pptx';
      ui.busy(false);
      if (pres.repaired) {
        const lost = pres.repaired.parts.filter((p) => /slides\/slide\d+\.xml$/.test(p)).length;
        ui.msg(`${L.APP} found a problem with content in ${file.name} and repaired the presentation.` + (lost ? `\n${lost} slide${lost > 1 ? 's' : ''} could not be recovered.` : '\nCheck the slides before you save over the original.'), { icon: 'warn' });
      } else if (/\.ppsx$/i.test(file.name)) L.show.start({ from: 0 });
    } catch (e) {
      ui.busy(false);
      console.error(e);
      if (e && e.code === 'ole') ui.msg(e.message, { icon: 'warn' });
      else ui.msg(`${L.APP} can't read ${file.name}. ${e && e.message ? e.message : ''}\nMake sure the file is a PowerPoint 2007 or later presentation (.pptx).`, { icon: 'error' });
    }
  };
  A.save = async function () {
    if (!A.saved) return A.saveAs();
    return A.exportAs(A.fileType || 'pptx', A.fileName);
  };
  A.saveAs = function () { return new Promise((res) => L.dlg.saveAs(async (name, type) => { if (['pptx', 'ppsx', 'potx'].includes(type)) { A.fileName = name; A.fileType = type; A.saved = true; A.updateTitle(); } await A.exportAs(type, name); res(); })); };
  A.exportAs = async function (type, name) {
    name = name || A.fileName;
    if (TE.active()) TE.sync();
    if (L.panes.left.tab === 'outline') L.panes.left.commitOutline();
    ui.busy(true, 'Saving...');
    try {
      let blob, ext = type;
      if (type === 'pptx' || type === 'ppsx' || type === 'potx') blob = await L.pptx.write(L.pres, { format: type });
      else if (type === 'html') blob = await A.buildWebPage();
      else if (type === 'png') blob = await A.slideToPNG(E.slide() || L.pres.slides[0], Math.round(L.pres.W * 2.6667));
      else if (type === 'pdf') blob = await A.buildPDF('slides');
      else if (type === 'txt') blob = new Blob([A.outlineText()], { type: 'text/plain' });
      ui.busy(false);
      if (!blob) return;
      const r = await L.saveFile(`${name}.${ext}`, blob);
      if (r === 'saved' && ['pptx', 'ppsx', 'potx'].includes(type)) { L.hist.dirty = false; A.status(`Saved ${name}.${ext}`); }
    } catch (e) {
      ui.busy(false);
      console.error(e);
      ui.msg(`The file could not be saved: ${e.message || e}`, { icon: 'error' });
    }
  };
  A.outlineText = function () {
    const lines = [];
    L.pres.slides.forEach((s) => {
      lines.push(M.slideTitle(s) || '');
      const b = s.shapes.find((x) => x.ph && x.type === 'text' && ['body', 'obj', 'subTitle'].includes(x.ph.type));
      if (b) for (const p of b.tx.ps) { const t = T.paraText(p); if (t) lines.push('\t'.repeat((p.lvl || 0) + 1) + t.replace(/\n/g, ' ')); }
    });
    return lines.join('\r\n');
  };
  /* --- render slides for export --- */
  function slideCSS(extra) {
    let css = '';
    const re = extra ? /(^|[\s,])\.sl[\s.,:]|^\.sl$|\.thumb-box|\.pa|\.tx|\.txi|\.shp|\.pic|\.pv-|\.np-/ : /(^|[\s,])\.sl[\s.,:]|^\.sl$|\.thumb-box|\.pa|\.tx|\.txi|\.shp|\.pic/;
    for (const sheet of document.styleSheets) {
      let rules;
      try { rules = sheet.cssRules; } catch (e) { continue; }
      for (const r of rules) if (r.selectorText && re.test(r.selectorText)) css += r.cssText + '\n';
    }
    return css;
  }
  async function mediaDataMap() {
    const map = new Map();
    for (const [, m] of L.media.all()) map.set(m.url, await L.blobToDataURL(m.view || m.blob));
    return map;
  }
  function inlineMedia(html, map) { for (const [u, d] of map) html = html.split(u).join(d); return html; }
  /** Rasterize a detached DOM element (laid out at w×h CSS px) into a canvas `scale` times larger. */
  A.rasterize = async function (el, w, hh, scale, opts) {
    const map = (opts && opts.map) || await mediaDataMap();
    const xhtml = inlineMedia(new XMLSerializer().serializeToString(el), map);
    const css = (opts && opts.css) || slideCSS(opts && opts.extraCSS);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${hh}"><foreignObject x="0" y="0" width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;width:${w}px;height:${hh}px;font:11px Arial,Arimo,sans-serif"><style>${css.replace(/</g, '\\3c ')}</style>${xhtml}</div></foreignObject></svg>`;
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
  const canvasBlob = (c, type, q) => new Promise((res, rej) => { try { c.toBlob((b) => (b ? res(b) : rej(new Error('This browser blocked the image export.'))), type, q); } catch (e) { rej(new Error('This browser blocked the image export.')); } });
  A.slideCanvas = function (slide, width, map) {
    const pres = L.pres;
    const el = L.render.slide(pres, slide, { mode: 'thumb', index: pres.slides.indexOf(slide) });
    return A.rasterize(el, pres.W, pres.H, width / pres.W, { map });
  };
  A.slideToPNG = async function (slide, width) { return canvasBlob(await A.slideCanvas(slide, width), 'image/png'); };
  /** PDF of the slides (one slide per page, page = slide size) or of the print-preview pages. */
  A.buildPDF = async function (source) {
    const pages = [];
    const map = await mediaDataMap();
    const jpeg = async (c) => new Uint8Array(await (await canvasBlob(c, 'image/jpeg', 0.92)).arrayBuffer());
    if (source === 'preview') {
      const els = L.$$('.pv-page', L.panes.preview.scroll);
      const css = slideCSS(true);
      let n = 0;
      for (const p of els) {
        ui.busy(true, `Creating PDF... (page ${++n} of ${els.length})`);
        const w = p.offsetWidth, hh = p.offsetHeight;
        const clone = p.cloneNode(true);
        clone.style.margin = '0'; clone.style.boxShadow = 'none'; clone.style.position = 'relative';
        const portrait = hh > w;
        const pw = portrait ? 612 : 792, ph = portrait ? 792 : 612;
        const c = await A.rasterize(clone, w, hh, (pw * 2.5) / w, { map, css });
        pages.push({ jpeg: await jpeg(c), iw: c.width, ih: c.height, pw, ph });
      }
    } else {
      const slides = L.pres.slides.filter((s) => !s.hidden);
      let n = 0;
      for (const s of slides) {
        ui.busy(true, `Creating PDF... (slide ${++n} of ${slides.length})`);
        const c = await A.slideCanvas(s, Math.round(L.pres.W * 2.5), map);
        pages.push({ jpeg: await jpeg(c), iw: c.width, ih: c.height, pw: L.pres.W, ph: L.pres.H });
      }
    }
    return L.buildPDF(pages, { title: L.pres.props.title || A.fileName, author: L.pres.props.author || '' });
  };
  A.previewPDF = async function () {
    ui.busy(true, 'Creating PDF...');
    try { const blob = await A.buildPDF('preview'); ui.busy(false); await L.saveFile(`${A.fileName} (${{ slides: 'Slides', handouts: 'Handouts', notes: 'Notes', outline: 'Outline' }[A.printOpts.what] || 'Print'}).pdf`, blob); }
    catch (e) { ui.busy(false); console.error(e); ui.msg('The PDF could not be created: ' + (e.message || e), { icon: 'error' }); }
  };
  A.buildWebPage = async function () {
    const pres = L.pres;
    const map = await mediaDataMap();
    const slides = pres.slides.filter((s) => !s.hidden).map((s) => inlineMedia(new XMLSerializer().serializeToString(L.render.slide(pres, s, { mode: 'thumb', index: pres.slides.indexOf(s) })), map));
    const notes = pres.slides.filter((s) => !s.hidden).map((s) => s.notes || '');
    const title = L.esc(pres.props.title || A.fileName);
    return new Blob([`<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${title}</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Arimo:ital,wght@0,400;0,700;1,400;1,700&family=Tinos:ital,wght@0,400;0,700;1,400;1,700&family=Cousine:wght@400;700&family=Carlito:ital,wght@0,400;0,700;1,400;1,700&family=Caladea:wght@400;700&display=swap">
<style>html,body{height:100%;margin:0;background:#1e1e1e;color:#eee;font:13px Tahoma,Verdana,sans-serif}#stage{position:absolute;left:0;top:0;transform-origin:0 0}#bar{position:fixed;left:0;right:0;bottom:0;display:flex;gap:8px;align-items:center;justify-content:center;padding:8px;background:rgba(0,0,0,.6)}button{font:inherit;padding:4px 12px}#notes{max-width:70ch;white-space:pre-wrap;color:#ccc}${slideCSS()}</style></head><body>
<div id="stage"></div><div id="bar"><button id="prev" aria-label="Previous slide">◀</button><span id="num"></span><button id="next" aria-label="Next slide">▶</button></div>
<script>const S=${JSON.stringify(slides)};const N=${JSON.stringify(notes)};const W=${pres.W},H=${pres.H};let i=0;const st=document.getElementById('stage');
function fit(){const k=Math.min(innerWidth/W,(innerHeight-44)/H);st.style.transform='translate('+((innerWidth-W*k)/2)+'px,'+((innerHeight-44-H*k)/2)+'px) scale('+k+')';}
function show(n){i=Math.max(0,Math.min(S.length-1,n));st.innerHTML=S[i];document.getElementById('num').textContent=(i+1)+' / '+S.length;}
document.getElementById('prev').onclick=()=>show(i-1);document.getElementById('next').onclick=()=>show(i+1);
addEventListener('keydown',e=>{if(['ArrowRight','PageDown',' ','Enter'].includes(e.key))show(i+1);if(['ArrowLeft','PageUp','Backspace'].includes(e.key))show(i-1);if(e.key==='Home')show(0);if(e.key==='End')show(S.length-1);});
addEventListener('resize',fit);fit();show(0);<\/script></body></html>`], { type: 'text/html' });
  };

  /* ================= history ================= */
  A.undo = () => { if (TE.active()) TE.sync(); if (L.hist.doUndo()) afterHistory(); };
  A.redo = () => { if (TE.active()) TE.sync(); if (L.hist.doRedo()) afterHistory(); };
  function afterHistory() {
    /* restoreSel runs from the history event */
  }
  L.bus.on('history', (entry) => {
    if (E.view === 'master') E.buildMaster();
    E.restoreSel(entry.sel);
    A.slideSel = new Set([E.idx]);
    L.bus.emit('slides-changed');
  });

  /* ================= slides ================= */
  A.selectedSlides = function () {
    const n = L.pres.slides.length;
    if ((A.view === 'sorter' || A.focusArea === 'slides') && A.slideSel.size) return Array.from(A.slideSel).filter((i) => i < n).sort((a, b) => a - b).map((i) => L.pres.slides[i]);
    return E.slide() && !E.slide().isMaster ? [E.slide()] : [];
  };
  const selIdx = () => A.selectedSlides().map((s) => L.pres.slides.indexOf(s));
  A.insertSlide = function (layout, at) {
    if (TE.active()) TE.end();
    const pres = L.pres;
    const cur = pres.slides[E.idx];
    const design = cur ? cur.design : Object.keys(pres.designs)[0];
    L.hist.push('New Slide');
    const s = M.newSlide(pres, layout || 'text', design);
    const i = at != null ? at : pres.slides.length ? Math.max(...selIdx(), E.idx) + 1 : 0;
    pres.slides.splice(i, 0, s);
    A.slideSel = new Set([i]);
    E.goto(i, { force: true });
    L.bus.emit('slides-changed');
    return s;
  };
  A.newSlideCmd = function () {
    const cur = E.slide();
    const layout = !cur ? 'title' : cur.layout === 'title' ? 'text' : cur.layout;
    A.insertSlide(layout);
    if (A.opts.layoutPaneOnNew) L.panes.task.show('layout');
  };
  A.duplicateSlides = function () {
    const idxs = selIdx();
    if (!idxs.length) return;
    L.hist.push('Duplicate Slide');
    const copies = idxs.map((i) => M.dupSlide(L.pres.slides[i]));
    const at = Math.max(...idxs) + 1;
    L.pres.slides.splice(at, 0, ...copies);
    A.slideSel = new Set(copies.map((_, k) => at + k));
    E.goto(at, { force: true });
    L.bus.emit('slides-changed');
  };
  A.deleteSlides = function () {
    const idxs = selIdx();
    if (!idxs.length) return;
    if (TE.active()) TE.detach(true);
    L.hist.push('Delete Slide');
    for (const i of idxs.slice().sort((a, b) => b - a)) L.pres.slides.splice(i, 1);
    const n = L.pres.slides.length;
    const ni = Math.min(idxs[0], n - 1);
    E.idx = Math.max(0, ni);
    A.slideSel = new Set(n ? [E.idx] : []);
    M.gcDesigns(L.pres);
    E.sel = [];
    E.render();
    L.bus.emit('slides-changed');
  };
  A.moveSlides = function (idxs, target) {
    if (!idxs.length) return;
    const pres = L.pres;
    const moving = idxs.map((i) => pres.slides[i]);
    const before = pres.slides.slice();
    const rest = pres.slides.filter((s) => !moving.includes(s));
    let t = target - idxs.filter((i) => i < target).length;
    t = L.clamp(t, 0, rest.length);
    const next = rest.slice(0, t).concat(moving, rest.slice(t));
    if (next.every((s, i) => s === before[i])) return;
    L.hist.push('Move Slide');
    pres.slides = next;
    A.slideSel = new Set(moving.map((s) => next.indexOf(s)));
    E.idx = next.indexOf(moving[0]);
    E.render();
    L.bus.emit('slides-changed');
  };
  A.toggleHidden = function () {
    const ss = A.selectedSlides();
    if (!ss.length) return;
    L.hist.push('Hide Slide');
    const v = !ss[0].hidden;
    for (const s of ss) s.hidden = v;
    L.bus.emit('slides-changed');
  };
  A.applyLayoutToSel = function (key, reapply) {
    if (A.view === 'master') return;
    const ss = A.selectedSlides();
    if (!ss.length) { A.insertSlide(key); return; }
    if (TE.active()) TE.end();
    L.hist.push('Slide Layout');
    for (const s of ss) M.applyLayout(L.pres, s, key);
    void reapply;
    E.sel = [];
    E.render();
    L.bus.emit('slides-changed');
  };
  A.applyDesign = function (key, all) {
    if (TE.active()) TE.end();
    L.hist.push('Apply Design Template');
    M.applyDesign(L.pres, key, all ? null : A.selectedSlides());
    if (A.view === 'master') E.buildMaster();
    E.render();
    L.bus.emit('slides-changed', { design: true });
    L.panes.task.refresh('design');
  };
  A.applyColorScheme = function (cs, all) {
    L.hist.push('Color Scheme');
    const targets = all ? L.pres.slides : A.selectedSlides();
    const ids = new Set(targets.map((s) => s.design));
    if (all) for (const id of ids) L.pres.designs[id].colors = Object.assign({}, L.pres.designs[id].colors, L.clone(cs));
    else for (const id of ids) {
      const users = L.pres.slides.filter((s) => s.design === id);
      if (users.every((s) => targets.includes(s))) { L.pres.designs[id].colors = Object.assign({}, L.pres.designs[id].colors, L.clone(cs)); continue; }
      const d = L.clone(L.pres.designs[id]); d.id = L.uid('dsn'); d.colors = Object.assign({}, d.colors, L.clone(cs)); L.pres.designs[d.id] = d;
      for (const s of targets) if (s.design === id) s.design = d.id;
    }
    if (A.view === 'master') E.buildMaster();
    E.render();
    L.bus.emit('slides-changed', { design: true });
    L.panes.task.refresh('design');
  };
  A.editColorScheme = function () {
    const d = E.design();
    const cs = L.clone(d.colors);
    const rows = M.SCHEME_SLOTS.map(([k, label]) => {
      const p = L.dlg.colorPick(cs[k], (c) => { cs[k] = c[0] === '#' ? c : M.resolveColor(c, d); });
      return ui.field(label + ':', p, { cls: 'wide' });
    });
    ui.dialog({ title: 'Edit Color Scheme', body: h('div', { class: 'col' }, ...rows), width: 360, buttons: [{ label: '&Apply', primary: true, onClick: () => A.applyColorScheme(cs, true) }, { label: 'Cancel' }] });
  };
  A.browseDesign = async function () {
    const files = await L.pickFiles('.pptx,.potx,.ppsx');
    if (!files[0]) return;
    ui.busy(true, 'Reading design...');
    try {
      const src = await L.pptx.read(await L.readAsArrayBuffer(files[0]));
      ui.busy(false);
      const d = Object.values(src.designs)[0];
      if (!d) return;
      d.id = L.uid('dsn');
      d.name = files[0].name.replace(/\.\w+$/, '');
      L.hist.push('Apply Design Template');
      L.pres.designs[d.id] = d;
      for (const s of L.pres.slides) A.moveToDesign(s, d);
      M.gcDesigns(L.pres);
      E.render();
      L.bus.emit('slides-changed', { design: true });
    } catch (e) { ui.busy(false); ui.msg('That design could not be read: ' + e.message, { icon: 'error' }); }
  };
  /** move a slide to a design, relocating untouched placeholders */
  A.moveToDesign = function (s, d) {
    const old = L.pres.designs[s.design];
    const oldF = old ? M.layoutFrames(s.layout, old) : [];
    const newF = M.layoutFrames(s.layout, d);
    for (const sh of s.shapes) {
      if (!sh.ph) continue;
      const of = oldF.find((f) => f.type === sh.ph.type && (f.idx || 0) === (sh.ph.idx || 0));
      const nf = newF.find((f) => f.type === sh.ph.type && (f.idx || 0) === (sh.ph.idx || 0)) || newF.find((f) => f.type === sh.ph.type);
      if (nf && (!of || (Math.abs(of.x - sh.x) < 2 && Math.abs(of.y - sh.y) < 2))) { sh.x = nf.x; sh.y = nf.y; sh.w = nf.w; if (sh.type !== 'table') sh.h = nf.h; }
    }
    s.design = d.id;
    delete s.lkey;
  };
  A.applyAnimScheme = function (def, all) {
    const targets = all ? L.pres.slides : A.selectedSlides();
    L.hist.push('Animation Scheme');
    for (const s of targets) {
      const ph = (t) => s.shapes.find((x) => x.ph && t.includes(x.ph.type));
      const title = ph(['title', 'ctrTitle']), body = ph(['body', 'obj', 'subTitle']);
      s.anims = (s.anims || []).filter((a) => !(title && a.sid === title.id) && !(body && a.sid === body.id));
      if (!def) { s.trans = Object.assign({}, s.trans, { type: 'none' }); continue; }
      s.trans = Object.assign({ spd: 'med', click: true, after: null }, s.trans, def.tr, { dur: undefined });
      delete s.trans.dur;
      const mk = (sh, spec, start, by) => {
        if (!sh || !spec || T.isEmpty(sh.tx)) return;
        s.anims.push({ id: L.uid('a'), sid: sh.id, cls: 'entr', eff: spec[0], dir: spec[1] || L.anim.defaultDir('entr', spec[0]), start, dur: spec[0] === 'appear' ? 1 : 500, delay: 0, by: by || 'all' });
      };
      mk(title, def.title, 'after');
      mk(body, def.body, 'click', def.body && def.body[2] === 'all' ? 'all' : 'para');
    }
    L.bus.emit('slides-changed');
    L.panes.task.refresh('anim');
  };
  A.setTransition = function (fn) {
    const ss = A.view === 'sorter' ? A.selectedSlides() : A.selectedSlides().length ? A.selectedSlides() : [E.slide()];
    L.hist.push('Slide Transition', 'trans');
    for (const s of ss) { if (!s) continue; s.trans = Object.assign({ type: 'none', spd: 'fast', click: true, after: null }, s.trans); fn(s); }
    L.bus.emit('slides-changed', { trans: true });
  };
  A.transitionToAll = function () {
    const src = E.slide();
    if (!src) return;
    L.hist.push('Apply Transition to All');
    for (const s of L.pres.slides) s.trans = L.clone(src.trans);
    L.bus.emit('slides-changed', { trans: true });
    A.status('Transition applied to all slides');
  };
  A.addAnimation = function (cls, eff) {
    const slide = E.slide();
    const shapes = E.selected();
    if (!slide || !shapes.length) return;
    L.hist.push('Add Effect');
    let last = null;
    for (const sh of shapes) {
      const top = E.topOf(sh.id) || sh;
      last = { id: L.uid('a'), sid: top.id, cls, eff, dir: L.anim.defaultDir(cls, eff), start: 'click', dur: cls === 'emph' || cls === 'path' ? 2000 : eff === 'appear' || eff === 'disappear' ? 1 : 500, delay: 0, by: 'all' };
      (slide.anims = slide.anims || []).push(last);
    }
    L.panes.selectAnim(last.id);
    L.panes.task.refresh('anim');
    L.bus.emit('slide-modified', E.idx);
    if (A.opts.autoPreview) L.panes.playSlide(last.id);
  };
  /**
   * Change the slide size the way current PowerPoint does ("Ensure Fit"):
   * backgrounds, master graphics and placeholders reflow to the new frame; everything else
   * (pictures, AutoShapes, charts, tables, WordArt, groups) keeps its proportions and is
   * scaled uniformly and centred, and text shrinks only if the slide got smaller.
   */
  const scaleSizes = (o, k) => {
    if (!o || typeof o !== 'object' || k === 1) return;
    if (Array.isArray(o)) { o.forEach((x) => scaleSizes(x, k)); return; }
    for (const key of Object.keys(o)) {
      const v = o[key];
      if (typeof v === 'number' && (key === 'sz' || key === 'pts' || key === 'marL' || key === 'indent' || key === 'fsz')) o[key] = L.round(v * k, key === 'sz' || key === 'fsz' ? 1 : 2);
      else if (v && typeof v === 'object' && key !== 'cmds' && key !== 'path' && key !== 'stops') scaleSizes(v, k);
    }
  };
  const scaleText = (tx, k) => { if (tx && k !== 1) { scaleSizes(tx.ps, k); scaleSizes(tx.lst, k); } };
  function uniform(sh, k, ox, oy) {
    sh.x = L.round(sh.x * k + ox, 2); sh.y = L.round(sh.y * k + oy, 2); sh.w = L.round(sh.w * k, 2); sh.h = L.round(sh.h * k, 2);
    if (k !== 1) {
      scaleText(sh.tx, k);
      if (sh.line && sh.line.w && k < 1) sh.line.w = L.round(Math.max(0.25, sh.line.w * k), 2);
      if (sh.tbl) { sh.tbl.cols = sh.tbl.cols.map((c) => L.round(c * k, 2)); sh.tbl.rows.forEach((r) => { r.h = L.round(r.h * k, 2); r.cells.forEach((c) => scaleText(c.tx, k)); }); }
      if (sh.chart && sh.chart.fsz) sh.chart.fsz = L.round(sh.chart.fsz * k, 1);
      if (sh.shadow) { sh.shadow.dx = (sh.shadow.dx || 0) * k; sh.shadow.dy = (sh.shadow.dy || 0) * k; }
    }
    if (sh.kids) sh.kids.forEach((c) => uniform(c, k, ox, oy));
  }
  function stretch(sh, sx, sy) {
    sh.x *= sx; sh.y *= sy; sh.w *= sx; sh.h *= sy;
    if (sh.kids) sh.kids.forEach((c) => stretch(c, sx, sy));
    if (sh.tbl) { sh.tbl.cols = sh.tbl.cols.map((c) => c * sx); sh.tbl.rows.forEach((r) => { r.h *= sy; }); }
  }
  const sameBox = (a, b) => a && b && Math.abs(a.x - b.x) < 1.5 && Math.abs(a.y - b.y) < 1.5 && Math.abs(a.w - b.w) < 1.5;
  A.resizeSlides = function (W, H) {
    const pres = L.pres;
    if (Math.abs(W - pres.W) < 0.5 && Math.abs(H - pres.H) < 0.5) return;
    if (TE.active()) TE.end();
    L.hist.push('Page Setup');
    const oW = pres.W, oH = pres.H;
    const sx = W / oW, sy = H / oH, k = Math.min(sx, sy);
    const ox = (W - oW * k) / 2, oy = (H - oH * k) / 2;
    const frameOf = (s, d) => M.layoutFrames(s.layout, d);
    const phFrame = (frames, d, ph) => frames.find((f) => f.type === ph.type && (f.idx || 0) === (ph.idx || 0)) || (d.ph && d.ph[ph.type] && ['dt', 'ftr', 'sldNum'].includes(ph.type) ? Object.assign({ type: ph.type }, d.ph[ph.type]) : null);
    const before = new Map(pres.slides.map((s) => [s, frameOf(s, M.design(pres, s))]));
    const beforeDesign = new Map(pres.slides.map((s) => [s, M.design(pres, s)]));
    /* 1. masters: built-in designs are regenerated at the new size, imported ones stretch */
    const oldPh = {};
    for (const id in pres.designs) {
      const d = pres.designs[id];
      oldPh[id] = L.clone(d.ph);
      const builtin = d.key && d.key !== 'imported' && M.DESIGNS.some((x) => x.key === d.key);
      if (builtin) {
        const nd = M.buildDesign(d.key, W, H);
        d.deco = nd.deco; d.ph = nd.ph;
        if (d.titleDeco) d.titleDeco = nd.titleDeco || d.titleDeco.map((x) => (stretch(x, sx, sy), x));
      } else {
        (d.deco || []).forEach((x) => stretch(x, sx, sy));
        (d.titleDeco || []).forEach((x) => stretch(x, sx, sy));
        if (d.layoutDecos) Object.values(d.layoutDecos).forEach((a) => a.forEach((x) => stretch(x, sx, sy)));
        for (const key in d.ph) { const r = d.ph[key]; r.x *= sx; r.y *= sy; r.w *= sx; r.h *= sy; }
      }
      if (k !== 1) { scaleSizes(d.tx, k); scaleSizes(d.footer, k); }
    }
    /* 2. slides */
    for (const s of pres.slides) {
      const d = M.design(pres, s);
      const oldD = Object.assign({}, beforeDesign.get(s), { ph: oldPh[s.design] || d.ph });
      const of = before.get(s), nf = frameOf(s, d);
      for (const sh of s.shapes) {
        if (sh.ph) {
          const o = phFrame(of, oldD, sh.ph), n = phFrame(nf, d, sh.ph);
          /* a placeholder the user moved or resized is treated like any other object */
          if (!(o && n && sameBox(sh, o))) { uniform(sh, k, ox, oy); continue; }
          const box = { x: n.x, y: n.y, w: n.w, h: n.h };
          if (sh.type === 'text') { Object.assign(sh, box); scaleText(sh.tx, k); continue; }
          /* picture / chart / table in a placeholder: fit inside the new box without distortion */
          const kk = Math.min(box.w / sh.w, box.h / sh.h);
          const cx = box.x + box.w / 2, cy = box.y + box.h / 2;
          uniform(sh, kk, 0, 0);
          M.translate(sh, cx - (sh.x + sh.w / 2), cy - (sh.y + sh.h / 2));
          continue;
        }
        uniform(sh, k, ox, oy);
      }
    }
    pres.W = W; pres.H = H;
    E.sel = [];
    E.layout();
    E.render();
    L.bus.emit('slides-changed', { design: true });
  };
  A.insertTitleMaster = function () {
    const d = E.design();
    L.hist.push('Insert Title Master');
    d.titleDeco = L.clone(d.deco).map((x) => Object.assign(x, { id: L.uid('d') }));
    E.masterKind = 'title'; E.buildMaster(); E.render(); L.panes.left.render();
  };
  A.deleteTitleMaster = function () {
    const d = E.design();
    L.hist.push('Delete Title Master');
    d.titleDeco = null; d.titleBg = null;
    E.masterKind = 'slide'; E.buildMaster(); E.render(); L.panes.left.render();
  };
  A.summarySlide = function () {
    const ss = A.selectedSlides().length > 1 ? A.selectedSlides() : L.pres.slides;
    const titles = ss.map((s) => M.slideTitle(s)).filter(Boolean);
    if (!titles.length) { ui.msg('The selected slides have no titles to summarize.'); return; }
    const at = L.pres.slides.indexOf(ss[0]);
    const s = A.insertSlide('text', Math.max(0, at));
    const t = s.shapes.find((x) => x.ph.type === 'title'), b = s.shapes.find((x) => x.ph.type === 'body');
    T.setPlain(t.tx, 'Summary Slide');
    b.tx.ps = titles.map((x) => T.para(x));
    E.render(); L.bus.emit('slides-changed');
  };
  A.speakerNotesDlg = function () {
    const s = E.slide();
    if (!s) return;
    const ta = h('textarea', { id: 'sn-text', rows: 10, style: 'width:100%' });
    ta.value = s.notes || '';
    ui.dialog({ title: 'Speaker Notes', body: h('div', { class: 'col' }, h('label', { for: 'sn-text', text: `Slide ${E.idx + 1}:` }), ta), width: 420, buttons: [{ label: 'Close', primary: true, onClick: () => { if (ta.value !== (s.notes || '')) { L.hist.push('Notes'); s.notes = ta.value; L.panes.notes.render(); } } }] });
  };
  A.outlineToSlides = function (text, at) {
    const lines = String(text).replace(/\r/g, '').split('\n');
    L.hist.push('Slides from Outline');
    let cur = null, i = at != null ? at : L.pres.slides.length;
    const did = (L.pres.slides[E.idx] || {}).design || Object.keys(L.pres.designs)[0];
    for (const raw of lines) {
      if (!raw.trim()) continue;
      const m = /^(\t*| *)(.*)$/.exec(raw);
      const depth = m[1].includes('\t') ? m[1].length : Math.floor(m[1].length / 4);
      if (depth === 0) { cur = M.newSlide(L.pres, 'text', did); T.setPlain(cur.shapes[0].tx, m[2].trim()); cur.shapes[1].tx.ps = []; L.pres.slides.splice(i++, 0, cur); }
      else if (cur) cur.shapes[1].tx.ps.push(T.para(m[2].trim(), null, null, Math.min(8, depth - 1)));
    }
    for (const s of L.pres.slides) { const b = s.shapes.find((x) => x.ph && x.ph.type === 'body'); if (b && !b.tx.ps.length) b.tx.ps = [T.para('')]; }
    E.render(); L.bus.emit('slides-changed');
  };
  A.slidesFromOutline = async function () {
    const f = await L.pickFiles('.txt,text/plain');
    if (!f[0]) return;
    A.outlineToSlides(await L.readAsText(f[0]), E.idx + 1);
  };
  A.insertSlidesFrom = function (src, idxs, keep) {
    L.hist.push('Insert Slides');
    const cur = E.slide();
    const target = cur ? L.pres.designs[cur.design] : Object.values(L.pres.designs)[0];
    const at = E.idx + 1;
    const added = [];
    const designMap = {};
    for (const i of idxs) {
      const s = M.dupSlide(src.slides[i]);
      if (keep) {
        if (!designMap[s.design]) { const d = L.clone(src.designs[s.design]); d.id = L.uid('dsn'); L.pres.designs[d.id] = d; designMap[s.design] = d.id; }
        s.design = designMap[s.design];
      } else {
        L.pres.designs[s.design] = src.designs[s.design];
        A.moveToDesign(s, target);
        delete L.pres.designs[src.slides[i].design];
      }
      added.push(s);
    }
    L.pres.slides.splice(at, 0, ...added);
    M.gcDesigns(L.pres);
    E.goto(at, { force: true });
    L.bus.emit('slides-changed');
  };
  A.loadSample = async function () {
    if (L.hist.dirty && !(await confirmDiscard())) return;
    A.loadPres(buildSample(), 'Harbor & Pine — 2027 Expansion (sample)');
  };

  /* ================= shape operations ================= */
  A.deleteSel = function () {
    const shapes = E.selected();
    if (!shapes.length) return;
    E.commit('Clear', () => {
      const slide = E.slide();
      for (const sh of shapes) {
        const f = E.find(sh.id);
        if (!f) continue;
        /* deleting a filled placeholder empties it first (PowerPoint behaviour) */
        if (sh.ph && sh.type === 'text' && !T.isEmpty(sh.tx) && !f.parent) { sh.tx.ps = [{ lvl: 0, pp: L.clone(sh.tx.ps[0].pp || {}), rs: [], end: sh.tx.ps[0].rs[0] ? T.runProps(sh.tx.ps[0].rs[0]) : sh.tx.ps[0].end }]; continue; }
        f.list.splice(f.list.indexOf(sh), 1);
        if (f.parent) { if (f.parent.kids.length === 1) { const pf = E.find(f.parent.id); const k = f.parent.kids[0]; pf.list.splice(pf.list.indexOf(f.parent), 1, k); } else E.fitGroup(f.parent); }
        slide.anims = (slide.anims || []).filter((a) => a.sid !== sh.id);
      }
      E.sel = [];
    });
  };
  A.duplicateSel = function () {
    const shapes = E.selected().filter((s) => !E.find(s.id).parent);
    if (!shapes.length) return;
    E.commit('Duplicate', () => {
      const slide = E.slide();
      E.sel = shapes.map((s) => { const c = M.dup(s); M.translate(c, 12, 12); slide.shapes.push(c); return c.id; });
    });
  };
  A.selectAll = function () {
    if (TE.active()) { const t = TE.target(); TE.setSel(T.allRange(t.tx)); return; }
    if (A.focusArea === 'slides' || A.view === 'sorter') { A.slideSel = new Set(L.pres.slides.map((_, i) => i)); L.panes.left.markCurrent(); if (A.view === 'sorter') L.panes.sorter.mark(); return; }
    E.selectAll();
  };
  A.group = function () {
    const shapes = E.selected().filter((s) => !E.find(s.id).parent);
    if (shapes.length < 2) return;
    E.commit('Group', () => {
      const slide = E.slide();
      const order = slide.shapes.filter((s) => shapes.includes(s));
      const at = slide.shapes.indexOf(order[order.length - 1]);
      const g = { id: L.uid('s'), type: 'group', name: 'Group ' + (slide.shapes.length + 1), rot: 0, kids: order };
      Object.assign(g, M.groupBounds(g));
      slide.shapes.splice(at + 1, 0, g);
      slide.shapes = slide.shapes.filter((s) => !order.includes(s));
      for (const s of order) delete s.ph;
      slide.anims = (slide.anims || []).filter((a) => !order.some((s) => s.id === a.sid));
      E.sel = [g.id];
    });
  };
  A.ungroup = function () {
    const groups = E.selected().filter((s) => s.type === 'group');
    if (!groups.length) return;
    E.commit('Ungroup', () => {
      const slide = E.slide();
      const out = [];
      for (const g of groups) {
        const f = E.find(g.id);
        const kids = g.kids;
        if (g.rot) {
          const cx = g.x + g.w / 2, cy = g.y + g.h / 2;
          for (const k of kids) { const [nx, ny] = L.rotPt(k.x + k.w / 2, k.y + k.h / 2, cx, cy, g.rot); M.translate(k, nx - (k.x + k.w / 2), ny - (k.y + k.h / 2)); k.rot = ((k.rot || 0) + g.rot) % 360; }
        }
        f.list.splice(f.list.indexOf(g), 1, ...kids);
        out.push(...kids.map((k) => k.id));
        A.lastUngroup = { ids: kids.map((k) => k.id), name: g.name, slide: slide.id };
      }
      E.sel = out;
    });
  };
  A.regroup = function () {
    const lu = A.lastUngroup;
    if (!lu || E.slide().id !== lu.slide) return;
    E.select(lu.ids.filter((id) => E.shape(id)));
    A.group();
    A.lastUngroup = null;
  };
  A.order = function (op) {
    const shapes = E.selected();
    if (!shapes.length) return;
    E.commit('Order', () => {
      for (const sh of op === 'front' || op === 'forward' ? shapes.slice().reverse() : shapes) {
        const f = E.find(sh.id);
        const list = f.list;
        const i = list.indexOf(sh);
        list.splice(i, 1);
        const j = op === 'front' ? list.length : op === 'back' ? 0 : op === 'forward' ? Math.min(list.length, i + 1) : Math.max(0, i - 1);
        list.splice(j, 0, sh);
      }
    });
  };
  A.align = function (k) {
    const shapes = E.selected();
    if (!shapes.length) return;
    const toSlide = A.opts.alignToSlide || shapes.length === 1;
    const bb = (s) => (s.type === 'group' ? M.groupBounds(s) : L.rotBounds(s));
    const ref = toSlide ? { x: 0, y: 0, w: L.pres.W, h: L.pres.H } : L.unionBounds(shapes.map(bb));
    E.commit('Align', () => {
      for (const s of shapes) {
        const b = bb(s);
        let dx = 0, dy = 0;
        if (k === 'L') dx = ref.x - b.x; if (k === 'C') dx = ref.x + ref.w / 2 - (b.x + b.w / 2); if (k === 'R') dx = ref.x + ref.w - (b.x + b.w);
        if (k === 'T') dy = ref.y - b.y; if (k === 'M') dy = ref.y + ref.h / 2 - (b.y + b.h / 2); if (k === 'B') dy = ref.y + ref.h - (b.y + b.h);
        M.translate(s, dx, dy);
        const f = E.find(s.id); if (f.parent) E.fitGroup(f.parent);
      }
    });
  };
  A.distribute = function (axis) {
    const shapes = E.selected();
    const bb = (s) => (s.type === 'group' ? M.groupBounds(s) : L.rotBounds(s));
    if (shapes.length < 2 && !A.opts.alignToSlide) return;
    const items = shapes.map((s) => ({ s, b: bb(s) })).sort((a, b) => (axis === 'h' ? a.b.x - b.b.x : a.b.y - b.b.y));
    const ref = A.opts.alignToSlide ? { x: 0, y: 0, w: L.pres.W, h: L.pres.H } : L.unionBounds(items.map((i) => i.b));
    const total = items.reduce((a, i) => a + (axis === 'h' ? i.b.w : i.b.h), 0);
    const gap = ((axis === 'h' ? ref.w : ref.h) - total) / Math.max(1, items.length - 1);
    E.commit('Distribute', () => {
      let pos = axis === 'h' ? ref.x : ref.y;
      for (const it of items) {
        if (axis === 'h') M.translate(it.s, pos - it.b.x, 0); else M.translate(it.s, 0, pos - it.b.y);
        pos += (axis === 'h' ? it.b.w : it.b.h) + gap;
      }
    });
  };
  A.rotate = (deg) => E.commit('Rotate', () => { for (const s of E.selected()) if (s.type !== 'table') { if (s.type === 'line') { const [a, b] = L.geom.lineEnds(s); const cx = s.x + s.w / 2, cy = s.y + s.h / 2; L.geom.setLineEnds(s, L.rotPt(a[0], a[1], cx, cy, deg), L.rotPt(b[0], b[1], cx, cy, deg)); } else s.rot = (((s.rot || 0) + deg) % 360 + 360) % 360; } });
  A.flip = (axis) => E.commit('Flip', () => {
    for (const s of E.selected()) {
      if (s.type === 'group') { const g = s; for (const k of g.kids) { if (axis === 'h') { k.x = g.x + g.w - (k.x - g.x) - k.w; k.flipH = !k.flipH; } else { k.y = g.y + g.h - (k.y - g.y) - k.h; k.flipV = !k.flipV; } k.rot = k.rot ? 360 - k.rot : 0; } continue; }
      if (axis === 'h') s.flipH = !s.flipH; else s.flipV = !s.flipV;
      if (s.rot) s.rot = (360 - s.rot) % 360;
    }
  });
  A.nudge = function (dx, dy, fine) {
    const step = fine ? 1 : E.grid.snap ? E.grid.size : 1;
    const shapes = E.selected();
    if (!shapes.length) return;
    L.hist.push('Nudge', 'nudge');
    for (const s of shapes) { M.translate(s, dx * step, dy * step); const f = E.find(s.id); if (f.parent) E.fitGroup(f.parent); }
    if (E.view === 'master') E.syncMaster();
    if (shapes.length === 1 && !E.find(shapes[0].id).parent) E.renderShape(shapes[0].id); else E.render();
    E.touched();
  };
  A.setDefaults = function () {
    const s = E.primary();
    if (!s) return;
    if (s.type === 'line') L.pres.lineDefaults = { line: L.clone(s.line) };
    else L.pres.shapeDefaults = { fill: L.clone(s.fill), line: L.clone(s.line), shadow: L.clone(s.shadow) };
    A.status('AutoShape defaults set');
  };
  A.changeAutoShape = function (geom) {
    E.commit('Change AutoShape', () => { for (const s of E.selected()) if (s.type === 'shape' || s.type === 'text') { s.geom = geom; delete s.path; s.adj = L.clone(L.geom.get(geom).adj); if (!s.adj) delete s.adj; if (s.type === 'text') s.type = 'shape'; } });
  };
  /* fills & lines */
  A.lastFillHex = () => M.resolveColor(A.lastFill, E.design());
  A.lastLineHex = () => M.resolveColor(A.lastLineColor, E.design());
  A.lastFontColorHex = () => M.resolveColor(A.lastFontColor, E.design());
  A.applyFill = function (v) {
    const f = typeof v === 'string' ? { t: 'solid', c: v, a: 1 } : v;
    const shapes = E.selected();
    if (!shapes.length) return;
    if (shapes.length === 1 && shapes[0].type === 'table') { A.tableFill(f); return; }
    E.commit('Fill Color', () => {
      for (const s of shapes) {
        if (s.type === 'wordart') { s.wa.fill = L.clone(f); continue; }
        if (s.type === 'line') continue;
        const keepA = s.fill && s.fill.t === 'solid' && f.t === 'solid' && s.fill.a < 1 ? s.fill.a : null;
        s.fill = L.clone(f);
        if (keepA != null) s.fill.a = keepA;
      }
    });
  };
  A.applyLineColor = function (c) {
    const shapes = E.selected();
    if (!shapes.length) return;
    E.commit('Line Color', () => {
      for (const s of shapes) {
        const tgt = s.type === 'wordart' ? s.wa : s;
        if (!c) { tgt.line = { t: 'none', w: (tgt.line && tgt.line.w) || 0.75 }; continue; }
        tgt.line = Object.assign({ w: 0.75, dash: 'solid' }, tgt.line && tgt.line.t !== 'none' ? tgt.line : {}, { c });
        delete tgt.line.t;
      }
    });
  };
  A.setLineProp = function (p) {
    const shapes = E.selected();
    if (!shapes.length) return;
    E.commit('Line Style', () => {
      for (const s of shapes) {
        const tgt = s.type === 'wordart' ? s.wa : s;
        if (!tgt.line || tgt.line.t === 'none') tgt.line = { c: 'tx1', w: 0.75, dash: 'solid' };
        if (p.w != null) tgt.line.w = p.w;
        if (p.dash) tgt.line.dash = p.dash;
        if ('head' in p) { tgt.line.head = p.head ? { type: p.head, w: 'med', len: 'med' } : undefined; tgt.line.tail = p.tail ? { type: p.tail, w: 'med', len: 'med' } : undefined; if (!tgt.line.head) delete tgt.line.head; if (!tgt.line.tail) delete tgt.line.tail; }
      }
    });
  };
  A.setShadow = (sh) => E.commit('Shadow', () => { for (const s of E.selected()) { if (s.type === 'wordart') s.wa.shadow = L.clone(sh); else s.shadow = L.clone(sh); } });
  /* text formatting */
  A.applyFontColor = function (c) {
    L.fmt.run((r) => { const n = Object.assign({}, r); if (c) n.color = c; else delete n.color; delete n.fill; return n; }, 'Font Color', (wa) => { wa.fill = c ? { t: 'solid', c, a: 1 } : { t: 'solid', c: 'tx1', a: 1 }; });
  };
  A.setFontName = (v) => { if (!v) return; L.fmt.run((r) => Object.assign({}, r, { font: v }), 'Font', (wa) => { wa.font = v; }); };
  A.setFontSize = (v) => { if (!(v > 0)) return; v = L.clamp(v, 1, 4000); L.fmt.run((r) => Object.assign({}, r, { sz: v }), 'Font Size'); };
  A.toggleRun = function (key) {
    const st = L.fmt.state();
    if (!st) return;
    const on = !st[key];
    const label = { b: 'Bold', i: 'Italic', u: 'Underline', shd: 'Shadow' }[key];
    L.fmt.run((r) => { const n = Object.assign({}, r); if (key === 'u') { if (on) n.u = 'sng'; else n.u = 'none'; } else n[key] = on ? true : false; return n; }, label, (wa) => { if (key === 'b' || key === 'i') wa[key] = on; if (key === 'shd') wa.shadow = on ? { c: '#000000', a: 0.4, dx: 3, dy: 3, blur: 0 } : null; });
  };
  A.toggleBaseline = function (v) {
    const st = L.fmt.state();
    const cur = st && st.r && st.r.base;
    L.fmt.run((r) => { const n = Object.assign({}, r); if (cur === v) delete n.base; else n.base = v; return n; }, v > 0 ? 'Superscript' : 'Subscript');
  };
  A.setAlign = (a) => {
    if (!TE.active() && E.selected().some((s) => s.type === 'wordart')) { A.waSet({ algn: a }); return; }
    L.fmt.para((p) => { p.pp = Object.assign({}, p.pp, { algn: a }); }, 'Alignment');
  };
  A.stepFont = function (d) {
    const st = L.fmt.state();
    if (!st) return;
    const list = L.SIZE_LIST;
    const step = (sz) => { if (d > 0) return list.find((x) => x > sz + 0.01) || Math.round(sz * 1.1); const s = list.slice().reverse().find((x) => x < sz - 0.01); return s || Math.max(1, Math.round(sz * 0.9)); };
    L.fmt.run((r) => Object.assign({}, r, { sz: step(r.sz || st.sz || 18) }), d > 0 ? 'Increase Font Size' : 'Decrease Font Size');
  };
  A.toggleBullets = function (kind) {
    const st = L.fmt.state();
    if (!st) return;
    const off = st.bu === kind;
    L.fmt.para((p, i, t) => {
      const bu = off ? { t: 'none' } : kind === 'num' ? { t: 'num', scheme: 'arabicPeriod' } : { t: 'char', ch: L.style.level(t ? t.sh : null, p.lvl, E.design()).bu.ch || '•' };
      if (kind === 'char' && !off && (!bu.ch || bu.ch === undefined)) bu.ch = '•';
      p.pp = Object.assign({}, p.pp, { bu });
      const cls = L.style.cls(t ? t.sh : null);
      if (!off && cls === 'other' && !(p.pp.marL > 0)) { p.pp.marL = 27 + (p.lvl || 0) * 36; p.pp.indent = -27; }
      if (off && cls === 'other') { delete p.pp.marL; delete p.pp.indent; }
      if (bu.t === 'char' && (!bu.ch || bu.ch === 'undefined')) bu.ch = '•';
    }, kind === 'num' ? 'Numbering' : 'Bullets');
  };
  A.indent = function (d) {
    L.fmt.para((p) => { p.lvl = L.clamp((p.lvl || 0) + d, 0, 8); }, d > 0 ? 'Demote' : 'Promote');
    if (L.panes.left.tab === 'outline') L.panes.left.render();
  };
  A.moveParas = function (d) {
    if (!TE.active()) return;
    const t = TE.target();
    TE.sync();
    const s = T.norm(TE.currentSel());
    const ps = t.tx.ps;
    if ((d < 0 && s.p0 === 0) || (d > 0 && s.p1 >= ps.length - 1)) return;
    L.hist.push('Move Paragraph');
    const block = ps.splice(s.p0, s.p1 - s.p0 + 1);
    ps.splice(s.p0 + d, 0, ...block);
    TE.rerender({ p0: s.p0 + d, o0: s.o0, p1: s.p1 + d, o1: s.o1 });
  };
  A.insertField = function (f) {
    A.ensureTextTarget();
    if (!TE.active()) return;
    L.hist.push('Insert ' + (f === 'slidenum' ? 'Slide Number' : 'Date'));
    TE.insertField(f);
  };
  /** make sure there's an active text editing target; returns true if editing */
  A.ensureTextTarget = function () {
    if (TE.active()) return true;
    const s = E.primary();
    if (s && (s.tx || s.type === 'shape')) { TE.begin(s.id, { atEnd: true }); return TE.active(); }
    const slide = E.slide();
    if (!slide) return false;
    L.hist.push('Insert Text Box');
    const tb = M.newTextBox(L.pres, slide, L.pres.W / 2 - 72, L.pres.H / 2 - 18, 144, 36, false);
    tb.tx.wrap = false;
    slide.shapes.push(tb);
    E.sel = [tb.id];
    E.render();
    TE.begin(tb.id, { atEnd: true });
    return true;
  };
  A.formatPainter = function (sticky) {
    if (E.tool && E.tool.kind === 'painter') { E.setTool(null); return; }
    const fmt = E.pickFormat();
    if (!fmt) return;
    E.setTool({ kind: 'painter', fmt, sticky: !!sticky });
    A.status(sticky ? 'Format Painter is locked on — press Esc to stop.' : 'Click an object or select text to apply the formatting.');
  };
  A.spelling = function () {
    A.opts.spell = true;
    A.saveOpts();
    L.$$('[contenteditable="true"]').forEach((el) => { el.spellcheck = true; });
    ui.msg('Spelling is checked as you type: misspelled words get a wavy red underline while you edit text.\nHold Shift and right-click a flagged word to see your browser\'s suggestions.', { title: 'Spelling', icon: 'info' });
  };

  /* ================= pictures ================= */
  A.insertPictureDialog = async function (phId) {
    const files = await L.pickFiles('image/*,.wmf,.emf', !phId);
    if (files.length) A.insertPictureFiles(files, phId);
  };
  async function imgSize(url) { try { const im = await L.loadImage(url); return [im.naturalWidth || 300, im.naturalHeight || 200]; } catch (e) { return [300, 200]; } }
  A.insertPictureFiles = async function (files, phId) {
    const slide = E.slide();
    if (!slide) return;
    const ph = phId ? E.shape(phId) : null;
    const ids = [];
    L.hist.push('Insert Picture');
    let k = 0;
    for (const f of files) {
      let view = null;
      if (/\.(wmf|emf)$/i.test(f.name) && L.metafile) view = await L.metafile.toPNG(new Uint8Array(await f.arrayBuffer()), /emf$/i.test(f.name) ? 'emf' : 'wmf');
      const id = L.media.add(/\.(wmf|emf)$/i.test(f.name) ? new Blob([f], { type: /emf$/i.test(f.name) ? 'image/x-emf' : 'image/x-wmf' }) : f, f.name, view);
      let [w, hh] = await imgSize(L.media.url(id));
      w *= 0.75; hh *= 0.75;
      const box = ph ? { x: ph.x, y: ph.y, w: ph.w, h: ph.h } : { x: L.pres.W * 0.05, y: L.pres.H * 0.05, w: L.pres.W * 0.9, h: L.pres.H * 0.9 };
      const sc = Math.min(1, box.w / w, box.h / hh);
      if (ph) { const s2 = Math.min(box.w / w, box.h / hh); w *= s2; hh *= s2; } else { w *= sc; hh *= sc; }
      const sh = M.newImage(slide, id, box.x + (box.w - w) / 2 + k * 12, box.y + (box.h - hh) / 2 + k * 12, w, hh);
      sh.name = 'Picture ' + (slide.shapes.length + 1);
      sh.alt = f.name.replace(/\.\w+$/, '');
      if (ph) { sh.ph = L.clone(ph.ph); const i = slide.shapes.indexOf(ph); slide.shapes.splice(i, 1, sh); }
      else slide.shapes.push(sh);
      ids.push(sh.id);
      k++;
    }
    E.sel = ids;
    E.render();
    E.touched();
    A.updateToolbars();
  };
  A.insertClipArt = async function (item) {
    const blob = await L.clipart.toPNG(item, 512);
    const file = new File([blob], item.name + '.png', { type: 'image/png' });
    const ph = E.primary() && E.primary().ph && E.primary().type === 'text' && T.isEmpty(E.primary().tx) && ['obj'].includes(E.primary().ph.type) ? E.primary().id : null;
    await A.insertPictureFiles([file], ph);
    const s = E.primary();
    if (s && !ph) { const size = 180; s.x = (L.pres.W - size) / 2; s.y = (L.pres.H - size) / 2; s.w = size; s.h = size; E.render(); }
  };
  A.picMode = (m) => E.commit('Picture Color', () => { for (const s of E.selected()) if (s.type === 'image') { s.img = Object.assign({}, s.img); if (m) s.img.mode = m; else delete s.img.mode; } });
  A.picAdjust = (k, d) => E.commit(k === 'bright' ? 'Brightness' : 'Contrast', () => { for (const s of E.selected()) if (s.type === 'image') { s.img = Object.assign({}, s.img); s.img[k] = L.clamp(L.round((s.img[k] || 0) + d, 2), -1, 1); } });
  A.picReset = () => E.commit('Reset Picture', () => { for (const s of E.selected()) if (s.type === 'image') { s.img = {}; s.crop = { l: 0, t: 0, r: 0, b: 0 }; } });
  A.picTransparentTool = function () {
    const s = E.primary();
    if (!s || s.type !== 'image') return;
    A.status('Click the color in the picture you want to make transparent.');
    const el = E.elOf(s.id);
    const once = async (e) => {
      E.scroller.removeEventListener('pointerdown', once, true);
      e.preventDefault(); e.stopPropagation();
      const img = el && el.querySelector('img');
      if (!img) return;
      const r = img.getBoundingClientRect();
      const nx = (e.clientX - r.left) / r.width, ny = (e.clientY - r.top) / r.height;
      if (nx < 0 || ny < 0 || nx > 1 || ny > 1) return;
      const im = await L.loadImage(L.media.url(s.media));
      const c = document.createElement('canvas');
      c.width = im.naturalWidth; c.height = im.naturalHeight;
      const g = c.getContext('2d');
      g.drawImage(im, 0, 0);
      let data;
      try { data = g.getImageData(0, 0, c.width, c.height); } catch (er) { ui.msg('This picture cannot be edited.'); return; }
      const px = (Math.floor(ny * c.height) * c.width + Math.floor(nx * c.width)) * 4;
      const [tr, tg, tb] = [data.data[px], data.data[px + 1], data.data[px + 2]];
      for (let i = 0; i < data.data.length; i += 4) if (Math.abs(data.data[i] - tr) < 12 && Math.abs(data.data[i + 1] - tg) < 12 && Math.abs(data.data[i + 2] - tb) < 12) data.data[i + 3] = 0;
      g.putImageData(data, 0, 0);
      const blob = await new Promise((res) => c.toBlob(res, 'image/png'));
      E.commit('Set Transparent Color', () => { s.media = L.media.add(blob, 'picture.png'); });
    };
    E.scroller.addEventListener('pointerdown', once, true);
  };

  /* ================= tables ================= */
  A.insertTable = function (rows, cols, phId) {
    const slide = E.slide();
    if (!slide) return;
    const ph = phId ? E.shape(phId) : null;
    const d = E.design();
    const box = ph || d.ph.body;
    L.hist.push('Insert Table');
    const t = M.newTable(L.pres, slide, rows, cols, box.x, box.y, box.w);
    t.y = ph ? ph.y : box.y;
    if (ph) { t.ph = L.clone(ph.ph); slide.shapes.splice(slide.shapes.indexOf(ph), 1, t); }
    else { t.y = (L.pres.H - t.h) / 2; slide.shapes.push(t); }
    E.sel = [t.id];
    E.render();
    TE.begin(t.id, { cell: [0, 0] });
    E.touched();
    A.updateToolbars();
  };
  function tableCtx() {
    const sh = TE.active() && TE.target() && TE.target().sh.type === 'table' ? TE.target().sh : E.selected().find((s) => s.type === 'table');
    if (!sh) return null;
    const rows = sh.tbl.rows.length, cols = sh.tbl.cols.length;
    let r0 = 0, c0 = 0, r1 = rows - 1, c1 = cols - 1;
    if (E.cellSel) { r0 = Math.min(E.cellSel.r0, E.cellSel.r1); r1 = Math.max(E.cellSel.r0, E.cellSel.r1); c0 = Math.min(E.cellSel.c0, E.cellSel.c1); c1 = Math.max(E.cellSel.c0, E.cellSel.c1); }
    else if (TE.active() && TE.state.cell) { [r0, c0] = TE.state.cell; r1 = r0; c1 = c0; }
    return { sh, r0, c0, r1, c1, editing: TE.active() };
  }
  const cloneCell = (c) => { const n = L.clone(c); n.tx.ps = [{ lvl: 0, pp: L.clone((c.tx.ps[0] && c.tx.ps[0].pp) || {}), rs: [], end: c.tx.ps[0] && c.tx.ps[0].rs[0] ? T.runProps(c.tx.ps[0].rs[0]) : (c.tx.ps[0] && c.tx.ps[0].end) || { sz: 18 } }]; delete n.gs; delete n.rs; delete n.hm; delete n.vm; return n; };
  A.tableAddRow = function (sh, at, below) {
    const src = sh.tbl.rows[at];
    const row = { h: src.h, cells: src.cells.map(cloneCell) };
    sh.tbl.rows.splice(below ? at + 1 : at, 0, row);
    sh.h += row.h;
  };
  A.tableCmd = function (op, arg) {
    const c = tableCtx();
    if (!c) return;
    const { sh } = c;
    const snap = TE.active() ? TE.snapshot() : null;
    if (snap) TE.detach();
    L.hist.push('Table');
    const t = sh.tbl;
    switch (op) {
      case 'rowAbove': for (let i = c.r0; i <= c.r1; i++) A.tableAddRow(sh, c.r0, false); break;
      case 'rowBelow': for (let i = c.r0; i <= c.r1; i++) A.tableAddRow(sh, c.r1, true); break;
      case 'colLeft': case 'colRight': {
        const n = c.c1 - c.c0 + 1;
        for (let k = 0; k < n; k++) {
          const at = op === 'colLeft' ? c.c0 : c.c1 + 1;
          const src = op === 'colLeft' ? c.c0 : c.c1;
          t.cols.splice(at, 0, t.cols[src]);
          sh.w += t.cols[src];
          for (const r of t.rows) r.cells.splice(at, 0, cloneCell(r.cells[src]));
        }
        break;
      }
      case 'deleteRows': t.rows.splice(c.r0, c.r1 - c.r0 + 1); break;
      case 'deleteCols': { const n = c.c1 - c.c0 + 1; sh.w -= t.cols.slice(c.c0, c.c1 + 1).reduce((a, b) => a + b, 0); t.cols.splice(c.c0, n); for (const r of t.rows) r.cells.splice(c.c0, n); break; }
      case 'merge': {
        const tl = t.rows[c.r0].cells[c.c0];
        for (let r = c.r0; r <= c.r1; r++) for (let q = c.c0; q <= c.c1; q++) {
          if (r === c.r0 && q === c.c0) continue;
          const cell = t.rows[r].cells[q];
          if (!T.isEmpty(cell.tx)) tl.tx.ps = tl.tx.ps.concat(cell.tx.ps);
          cell.tx.ps = [T.para('')];
          delete cell.gs; delete cell.rs;
          if (r === c.r0) cell.hm = true; else cell.vm = true;
          if (r > c.r0 && q > c.c0) cell.hm = true;
        }
        tl.gs = c.c1 - c.c0 + 1; tl.rs = c.r1 - c.r0 + 1;
        if (tl.gs === 1) delete tl.gs; if (tl.rs === 1) delete tl.rs;
        E.cellSel = null;
        break;
      }
      case 'split': {
        const cell = t.rows[c.r0].cells[c.c0];
        const gs = cell.gs || 1, rs = cell.rs || 1;
        if (gs === 1 && rs === 1) { L.hist.undo.pop(); if (snap) TE.resume(snap); ui.msg('Split Cell divides a merged cell back into its original cells. Select a merged cell first.'); return; }
        for (let r = c.r0; r < c.r0 + rs; r++) for (let q = c.c0; q < c.c0 + gs; q++) { const x = t.rows[r] && t.rows[r].cells[q]; if (x) { delete x.hm; delete x.vm; } }
        delete cell.gs; delete cell.rs;
        break;
      }
      case 'anchor': for (let r = c.r0; r <= c.r1; r++) for (let q = c.c0; q <= c.c1; q++) t.rows[r].cells[q].tx.anchor = arg; break;
      case 'distRows': { const total = t.rows.slice(c.r0, c.r1 + 1).reduce((a, r) => a + r.h, 0); const n = c.r1 - c.r0 + 1; for (let r = c.r0; r <= c.r1; r++) t.rows[r].h = L.round(total / n, 2); break; }
      case 'distCols': { const total = t.cols.slice(c.c0, c.c1 + 1).reduce((a, b) => a + b, 0); const n = c.c1 - c.c0 + 1; for (let q = c.c0; q <= c.c1; q++) t.cols[q] = L.round(total / n, 2); break; }
      case 'selectAll': L.hist.undo.pop(); E.cellSel = { r0: 0, c0: 0, r1: t.rows.length - 1, c1: t.cols.length - 1 }; if (snap) TE.resume(snap); E.cellSel = { r0: 0, c0: 0, r1: t.rows.length - 1, c1: t.cols.length - 1 }; E.drawOverlay(); return;
      default: break;
    }
    if (!t.rows.length || !t.cols.length) { const f = E.find(sh.id); f.list.splice(f.list.indexOf(sh), 1); E.sel = []; E.render(); E.touched(); return; }
    sh.h = t.rows.reduce((a, r) => a + r.h, 0);
    E.render({ keepEdit: false });
    if (snap && E.shape(sh.id)) { snap.cell = [Math.min(snap.cell ? snap.cell[0] : 0, t.rows.length - 1), Math.min(snap.cell ? snap.cell[1] : 0, t.cols.length - 1)]; snap.sel = null; TE.resume(snap); }
    E.touched();
  };
  A.tableBorders = function (kind) {
    const c = tableCtx();
    if (!c) return;
    const b = L.clone(A.tblBorder);
    const none = { t: 'none', w: 1 };
    const snap = TE.active() ? TE.snapshot() : null;
    if (snap) TE.detach();
    L.hist.push('Borders');
    const t = c.sh.tbl;
    for (let r = c.r0; r <= c.r1; r++) for (let q = c.c0; q <= c.c1; q++) {
      const cell = t.rows[r].cells[q];
      cell.bd = cell.bd || {};
      const edge = { t: r === c.r0, b: r === c.r1, l: q === c.c0, r: q === c.c1 };
      for (const side of ['l', 'r', 't', 'b']) {
        let on;
        if (kind === 'all') on = true; else if (kind === 'none') on = false;
        else if (kind === 'out') on = edge[side] ? true : undefined; else if (kind === 'in') on = !edge[side] ? true : undefined;
        else on = kind === side && edge[side] ? true : undefined;
        if (on === true) cell.bd[side] = L.clone(b); else if (on === false) cell.bd[side] = L.clone(none);
      }
    }
    E.render({ keepEdit: false });
    if (snap) TE.resume(snap);
    E.touched();
  };
  A.tableFill = function (f) {
    const c = tableCtx();
    if (!c) return;
    const snap = TE.active() ? TE.snapshot() : null;
    if (snap) TE.detach();
    L.hist.push('Fill Color');
    for (let r = c.r0; r <= c.r1; r++) for (let q = c.c0; q <= c.c1; q++) c.sh.tbl.rows[r].cells[q].fill = L.clone(f);
    E.render({ keepEdit: false });
    if (snap) TE.resume(snap);
    E.touched();
  };
  A.tableBordersDlg = function () {
    const c = tableCtx();
    if (!c) return;
    const bc = L.dlg.colorPick(A.tblBorder.c, (v) => { A.tblBorder.c = v; });
    const bw = ui.select(L.dlg.LINE_WEIGHTS.map((w) => [w, w + ' pt']), A.tblBorder.w, (v) => { A.tblBorder.w = +v; });
    const bs = ui.select(L.dlg.DASHES, A.tblBorder.dash, (v) => { A.tblBorder.dash = v; });
    const btns = h('div', { class: 'row' }, ...[['all', 'bordersAll'], ['out', 'bordersOut'], ['in', 'bordersIn'], ['none', 'bordersNone'], ['t', 'bordersTop'], ['b', 'bordersBottom'], ['l', 'bordersLeft'], ['r', 'bordersRight']].map(([k, ic]) => { const b = h('button', { type: 'button', class: 'btn small', html: L.icons.get(ic), 'aria-label': k }); b.addEventListener('click', () => A.tableBorders(k)); return b; }));
    const cf = L.dlg.fillPicker({ t: 'none' }, 'fill', (f) => A.tableFill(f));
    const anchor = ui.select([['t', 'Top'], ['ctr', 'Middle'], ['b', 'Bottom']], 't', (v) => A.tableCmd('anchor', v));
    const tabs = ui.tabs([{ label: 'Borders', body: h('div', { class: 'col' }, ui.field('&Style:', bs), ui.field('&Color:', bc), ui.field('&Width:', bw), h('label', { text: 'Click a button to apply borders:' }), btns) }, { label: 'Fill', body: h('div', { class: 'col' }, ui.field('Fill &color:', cf)) }, { label: 'Text Box', body: h('div', { class: 'col' }, ui.field('Text &alignment:', anchor)) }]);
    ui.dialog({ title: 'Format Table', body: tabs, width: 380, buttons: [{ label: 'OK', primary: true }] });
  };

  /* ================= charts, WordArt, diagrams ================= */
  A.insertChart = function (phId) {
    const slide = E.slide();
    if (!slide) return;
    const ph = phId ? E.shape(phId) : null;
    const box = ph || { x: L.pres.W * 0.2, y: L.pres.H * 0.22, w: L.pres.W * 0.6, h: L.pres.H * 0.6 };
    L.hist.push('Insert Chart');
    const sh = { id: L.uid('s'), type: 'chart', name: 'Chart ' + (slide.shapes.length + 1), x: box.x, y: box.y, w: box.w, h: box.h, rot: 0, chart: L.chart.sample() };
    sh.chart.series.forEach((s, i) => { s.color = L.chart.PALETTE[i]; });
    if (ph) { sh.ph = L.clone(ph.ph); slide.shapes.splice(slide.shapes.indexOf(ph), 1, sh); } else slide.shapes.push(sh);
    E.sel = [sh.id];
    E.render(); E.touched();
    L.chart.edit(sh, (c) => E.commit('Chart', () => { sh.chart = c; }));
  };
  ui.cmd('editChart', { label: 'Edit Chart Data...', icon: 'chart', enabled: () => { const s = E.primary(); return !!s && s.type === 'chart'; }, run: () => { const s = E.primary(); L.chart.edit(s, (c) => E.commit('Chart', () => { s.chart = c; })); } });
  A.insertWordArtDialog = function () {
    L.dlg.wordartGallery((idx) => {
      const st = L.dlg.WA_STYLES[idx];
      L.dlg.wordartText({ text: 'Your Text Here', font: st.font, b: st.b }, (o) => {
        const slide = E.slide();
        if (!slide) return;
        const lines = o.text.split('\n');
        const maxLen = Math.max(...lines.map((x) => x.length), 1);
        const hh = Math.min(L.pres.H * 0.6, o.size * 1.5 * lines.length);
        const w = Math.min(L.pres.W * 0.85, maxLen * o.size * 0.75);
        L.hist.push('Insert WordArt');
        const sh = { id: L.uid('s'), type: 'wordart', name: 'WordArt ' + (slide.shapes.length + 1), x: (L.pres.W - w) / 2, y: (L.pres.H - hh) / 2, w, h: hh, rot: 0, wa: Object.assign(L.clone(st), { text: o.text, font: o.font, b: o.b, i: o.i, style: idx }) };
        slide.shapes.push(sh);
        E.sel = [sh.id];
        E.render(); E.touched(); A.updateToolbars();
      });
    });
  };
  A.editWordArt = function () {
    const s = E.primary();
    if (!s || s.type !== 'wordart') return;
    L.dlg.wordartText(s.wa, (o) => E.commit('Edit WordArt Text', () => { Object.assign(s.wa, { text: o.text, font: o.font, b: o.b, i: o.i }); }));
  };
  A.wordartRestyle = function () {
    const s = E.primary();
    L.dlg.wordartGallery((idx) => E.commit('WordArt Style', () => { for (const x of E.selected()) if (x.type === 'wordart') { const st = L.clone(L.dlg.WA_STYLES[idx]); x.wa = Object.assign(x.wa, { fill: st.fill, line: st.line, shadow: st.shadow, warp: st.warp, style: idx }); } }), s && s.wa ? s.wa.style : 0);
  };
  A.waSet = (o) => E.commit('WordArt', () => { for (const s of E.selected()) if (s.type === 'wordart') Object.assign(s.wa, o); });
  A.waToggle = (k) => E.commit('WordArt', () => { for (const s of E.selected()) if (s.type === 'wordart') s.wa[k] = !s.wa[k]; });
  A.insertDiagram = function (kind, phId) {
    const slide = E.slide();
    if (!slide) return;
    const ph = phId ? E.shape(phId) : null;
    const box = ph ? { x: ph.x, y: ph.y, w: ph.w, h: ph.h } : { x: L.pres.W * 0.15, y: L.pres.H * 0.22, w: L.pres.W * 0.7, h: L.pres.H * 0.66 };
    const kids = buildDiagram(kind, box);
    L.hist.push('Insert Diagram');
    const g = { id: L.uid('s'), type: 'group', name: L.dlg.DIAGRAMS.find((d) => d[0] === kind)[1] + ' ' + (slide.shapes.length + 1), rot: 0, kids };
    Object.assign(g, M.groupBounds(g));
    if (ph) slide.shapes.splice(slide.shapes.indexOf(ph), 1, g); else slide.shapes.push(g);
    E.sel = [g.id];
    E.render(); E.touched();
  };
  function buildDiagram(kind, b) {
    const shp = (geom, x, y, w, hh, fill, text, extra) => Object.assign({ id: L.uid('s'), type: 'shape', name: 'Diagram shape', geom, x, y, w, h: hh, rot: 0, fill, line: { c: 'tx1', w: 0.75, dash: 'solid' }, tx: T.body([T.para(text || '', { algn: 'ctr' }, { sz: Math.max(10, Math.round(Math.min(b.w, b.h) / 22)), color: 'tx1' })], { anchor: 'ctr', wrap: true }) }, extra || {});
    const ln = (x1, y1, x2, y2) => M.newLine(L.pres, { shapes: [] }, 'line', [x1, y1], [x2, y2]);
    const acc = (i) => ({ t: 'solid', c: ['accent1', 'accent2', 'accent3', 'accent6', 'accent5', 'accent4'][i % 6], a: 1 });
    const cx = b.x + b.w / 2, cy = b.y + b.h / 2;
    const out = [];
    switch (kind) {
      case 'org': {
        const bw = b.w * 0.26, bh = b.h * 0.2;
        out.push(shp('rect', cx - bw / 2, b.y, bw, bh, acc(0), 'Director'));
        const ys = b.y + b.h * 0.6;
        out.push(ln(cx, b.y + bh, cx, b.y + bh + (ys - b.y - bh) / 2));
        const xs = [b.x + b.w * 0.02, cx - bw / 2, b.x + b.w * 0.98 - bw];
        out.push(ln(xs[0] + bw / 2, b.y + bh + (ys - b.y - bh) / 2, xs[2] + bw / 2, b.y + bh + (ys - b.y - bh) / 2));
        xs.forEach((x, i) => { out.push(ln(x + bw / 2, b.y + bh + (ys - b.y - bh) / 2, x + bw / 2, ys)); out.push(shp('rect', x, ys, bw, bh, acc(0), ['Operations', 'Marketing', 'Finance'][i])); });
        break;
      }
      case 'cycle': {
        const r = Math.min(b.w, b.h) * 0.42;
        for (let i = 0; i < 3; i++) {
          const a0 = -90 + i * 120 + 12, a1 = a0 + 96;
          out.push(shp('circularArrow', cx - r, cy - r, r * 2, r * 2, acc(i), '', { adj: { adj1: 9000, adj2: 1142319, adj3: 20457681, adj4: Math.round((((a0 % 360) + 360) % 360) * 60000), adj5: Math.round((((a1 % 360) + 360) % 360) * 60000) }, line: { t: 'none' } }));
          const am = ((a0 + 48) * Math.PI) / 180;
          out.push(shp('rect', cx + Math.cos(am) * r * 1.02 - b.w * 0.11, cy + Math.sin(am) * r * 1.02 - b.h * 0.06, b.w * 0.22, b.h * 0.12, { t: 'none' }, ['Plan', 'Do', 'Review'][i], { line: { t: 'none' }, type: 'text' }));
        }
        break;
      }
      case 'radial': {
        const r = Math.min(b.w, b.h) * 0.38, cr = r * 0.42, sr = r * 0.3;
        for (let i = 0; i < 4; i++) { const a = (-90 + i * 90) * Math.PI / 180; out.push(ln(cx + Math.cos(a) * cr, cy + Math.sin(a) * cr, cx + Math.cos(a) * (r - sr), cy + Math.sin(a) * (r - sr))); }
        out.push(shp('ellipse', cx - cr, cy - cr, cr * 2, cr * 2, acc(0), 'Core'));
        for (let i = 0; i < 4; i++) { const a = (-90 + i * 90) * Math.PI / 180; out.push(shp('ellipse', cx + Math.cos(a) * r - sr, cy + Math.sin(a) * r - sr, sr * 2, sr * 2, acc(i + 1), ['North', 'East', 'South', 'West'][i])); }
        break;
      }
      case 'pyramid': {
        const n = 4, top = b.y, base = b.y + b.h, pw = Math.min(b.w, b.h * 1.3);
        for (let i = 0; i < n; i++) {
          const y0 = top + (b.h * i) / n, y1 = top + (b.h * (i + 1)) / n;
          const w0 = (pw * (y0 - top)) / b.h, w1 = (pw * (y1 - top)) / b.h;
          const x = cx - w1 / 2, w = w1, hh = y1 - y0;
          const cmds = i === 0 ? [['M', w / 2, 0], ['L', w, hh], ['L', 0, hh], ['Z']] : [['M', (w - w0) / 2, 0], ['L', (w + w0) / 2, 0], ['L', w, hh], ['L', 0, hh], ['Z']];
          out.push(shp('custom', x, y0, w, hh, acc(i), ['Vision', 'Strategy', 'Programs', 'Operations'][i], { path: { paths: [{ w, h: hh, cmds, fill: 'norm', stroke: true }] } }));
        }
        void base;
        break;
      }
      case 'venn': {
        const r = Math.min(b.w, b.h) * 0.3;
        [[-0.55, -0.35], [0.55, -0.35], [0, 0.55]].forEach(([dx, dy], i) => out.push(shp('ellipse', cx + dx * r - r, cy + dy * r - r, r * 2, r * 2, { t: 'solid', c: ['accent1', 'accent2', 'accent3'][i], a: 0.55 }, ['People', 'Process', 'Tools'][i])));
        break;
      }
      case 'target': {
        const r = Math.min(b.w, b.h) * 0.46;
        [1, 0.68, 0.36].forEach((k, i) => out.push(shp('ellipse', cx - r * k, cy - r * k, r * k * 2, r * k * 2, acc(i), '', { tx: undefined })));
        ['Awareness', 'Interest', 'Decision'].forEach((t, i) => out.push(shp('rect', cx + r * 1.05, cy - r + i * r * 0.5 + r * 0.25, b.w * 0.2, r * 0.3, { t: 'none' }, t, { type: 'text', line: { t: 'none' } })));
        out.forEach((s) => { if (s.tx === undefined) delete s.tx; });
        break;
      }
      default: break;
    }
    return out;
  }
  A.createPhotoAlbum = async function (pics, layout, frame, caps) {
    const pres = L.pres;
    const did = (pres.slides[E.idx] || pres.slides[0] || {}).design || Object.keys(pres.designs)[0];
    L.hist.push('Photo Album');
    const t = M.newSlide(pres, 'title', did);
    T.setPlain(t.shapes[0].tx, 'Photo Album');
    T.setPlain(t.shapes[1].tx, 'by ' + (pres.props.author || 'Lectern'));
    pres.slides.push(t);
    const per = layout === 'fit' || /^1/.test(layout) ? 1 : /^2/.test(layout) ? 2 : 4;
    const withTitle = /t$/.test(layout);
    for (let i = 0; i < pics.length; i += per) {
      const s = M.newSlide(pres, withTitle ? 'titleOnly' : 'blank', did);
      if (withTitle) T.setPlain(s.shapes[0].tx, 'Photos ' + (i / per + 1));
      const top = withTitle ? pres.H * 0.24 : pres.H * 0.05, area = { x: pres.W * 0.05, y: top, w: pres.W * 0.9, h: pres.H * 0.95 - top };
      if (layout === 'fit') { area.x = 0; area.y = 0; area.w = pres.W; area.h = pres.H; }
      const cells = per === 1 ? [area] : per === 2 ? [{ x: area.x, y: area.y, w: area.w / 2 - 6, h: area.h }, { x: area.x + area.w / 2 + 6, y: area.y, w: area.w / 2 - 6, h: area.h }] : [0, 1, 2, 3].map((k) => ({ x: area.x + (k % 2) * (area.w / 2 + 6), y: area.y + Math.floor(k / 2) * (area.h / 2 + 6), w: area.w / 2 - 6, h: area.h / 2 - 6 }));
      for (let k = 0; k < per && i + k < pics.length; k++) {
        const p = pics[i + k];
        const id = L.media.add(p.file, p.name);
        let [w, hh] = await imgSize(p.url);
        const cell = Object.assign({}, cells[k]);
        if (caps) cell.h -= 24;
        const sc = Math.min(cell.w / w, cell.h / hh);
        w *= sc; hh *= sc;
        const sh = M.newImage(s, id, cell.x + (cell.w - w) / 2, cell.y + (cell.h - hh) / 2, w, hh);
        sh.alt = p.name.replace(/\.\w+$/, '');
        if (frame === 'roundRect') sh.geom = 'roundRect';
        if (frame === 'simple') sh.line = { c: '#FFFFFF', w: 6, dash: 'solid' };
        if (frame === 'black') sh.line = { c: '#000000', w: 6, dash: 'solid' };
        if (frame === 'shadow') sh.shadow = { c: '#000000', a: 0.45, dx: 4, dy: 4, blur: 6 };
        s.shapes.push(sh);
        if (caps) { const tb = M.newTextBox(pres, s, cell.x, cell.y + cell.h + 2, cell.w, 22, true); tb.tx.ps = [T.para(sh.alt, { algn: 'ctr' }, { sz: 12 })]; s.shapes.push(tb); }
      }
      pres.slides.push(s);
    }
    E.goto(pres.slides.indexOf(t), { force: true });
    L.bus.emit('slides-changed');
  };

  /* ================= AutoContent wizard ================= */
  const AUTOCONTENT = {
    'Recommending a Strategy': [['Recommending a Strategy', ['Your name', 'Date']], ['Vision Statement', ['Where we want to be in three years', 'What success looks like for customers']], ['Goals', ['Goal and its measure', 'Goal and its measure', 'Goal and its measure']], ['Today\'s Situation', ['Strengths we build on', 'Gaps we need to close', 'What competitors are doing']], ['Options', ['Option A — cost, benefit, risk', 'Option B — cost, benefit, risk', 'Option C — cost, benefit, risk']], ['Recommendation', ['The option we propose and why', 'What it will cost', 'How we will know it is working']], ['Next Steps', ['Decision needed today', 'Owners and dates', 'When we report back']]],
    'Project Overview': [['Project Overview', ['Project lead', 'Date']], ['Objectives', ['What the project will deliver', 'Who benefits']], ['Scope', ['In scope', 'Out of scope']], ['Timeline', ['Milestone and date', 'Milestone and date', 'Milestone and date']], ['Team and Roles', ['Sponsor', 'Lead', 'Contributors']], ['Risks', ['Risk — likelihood — mitigation', 'Risk — likelihood — mitigation']], ['Status Reporting', ['How often we report', 'Where to find updates']]],
    'Status Report': [['Status Report', ['Team or project', 'Reporting period']], ['Summary', ['Overall status: on track / at risk / off track', 'Highlights since last report']], ['Progress', ['Completed', 'In progress', 'Not started']], ['Issues', ['Issue — impact — help needed']], ['Next Period', ['Planned work', 'Decisions needed']]],
    'Brainstorming Session': [['Brainstorming Session', ['Facilitator', 'Date']], ['Ground Rules', ['Every idea counts', 'Build on others\' ideas', 'Quantity first, judgment later']], ['The Question', ['State the problem in one sentence']], ['Ideas', ['', '', '']], ['Grouping the Ideas', ['Theme', 'Theme', 'Theme']], ['Next Steps', ['Ideas to explore', 'Owners and dates']]],
    'Training': [['Training', ['Instructor', 'Date']], ['Learning Objectives', ['After this session you will be able to...', '...', '...']], ['Agenda', ['Topic one', 'Topic two', 'Practice', 'Review']], ['Topic One', ['Key concept', 'Example', 'Try it']], ['Topic Two', ['Key concept', 'Example', 'Try it']], ['Review', ['What we covered', 'Where to learn more', 'Questions']]],
    'Selling a Product or Service': [['Product Name', ['Presenter', 'Date']], ['The Problem', ['What customers struggle with today']], ['Our Solution', ['What it does', 'How it is different']], ['Benefits', ['Saves time', 'Saves money', 'Reduces risk']], ['Pricing', ['Plans and options']], ['Getting Started', ['How to buy', 'Who to contact']]],
  };
  A.autoContent = function () {
    const types = Object.keys(AUTOCONTENT);
    const list = h('select', { id: 'ac-type', size: types.length, style: 'width:240px' }, ...types.map((t) => h('option', { value: t, text: t })));
    list.selectedIndex = 0;
    const title = h('input', { type: 'text', id: 'ac-title', style: 'width:240px', placeholder: 'Presentation title' });
    const footer = h('input', { type: 'text', id: 'ac-footer', style: 'width:240px', placeholder: 'Footer (optional)' });
    const dsn = ui.select(M.DESIGNS.map((d) => [d.key, d.name]), 'azure');
    const num = ui.check('Slide &number', true), date = ui.check('&Date last updated', false);
    ui.dialog({
      title: 'AutoContent Wizard', body: h('div', { class: 'row' }, h('div', { class: 'col' }, h('label', { for: 'ac-type', text: 'Presentation type:' }), list), h('div', { class: 'col' }, ui.field('&Title:', title), ui.field('&Footer:', footer), ui.field('D&esign:', dsn), num, date)), width: 560, buttons: [{
        label: '&Finish', primary: true, onClick: async () => {
          if (!(await confirmDiscard())) return;
          const spec = AUTOCONTENT[list.value];
          const pres = M.newPresentation(Object.assign({ design: dsn.value, empty: true }, A.newSize()));
          const did = Object.keys(pres.designs)[0];
          spec.forEach(([t, bullets], i) => {
            const s = M.newSlide(pres, i === 0 ? 'title' : 'text', did);
            T.setPlain(s.shapes[0].tx, i === 0 && title.value ? title.value : t);
            if (i === 0) T.setPlain(s.shapes[1].tx, bullets.join('\n'));
            else s.shapes[1].tx.ps = bullets.map((b) => T.para(b));
            pres.slides.push(s);
          });
          pres.hf = Object.assign(pres.hf, { num: num.input.checked, dt: date.input.checked, ftr: !!footer.value, ftrText: footer.value, notOnTitle: true });
          A.loadPres(pres, title.value || list.value);
        },
      }, { label: 'Cancel' }],
    });
  };

  /* ================= clipboard ================= */
  const clipText = (shapes) => shapes.map((s) => (s.tx ? T.plain(s.tx) : s.wa ? s.wa.text : s.tbl ? s.tbl.rows.map((r) => r.cells.map((c) => T.plain(c.tx)).join('\t')).join('\n') : '')).filter(Boolean).join('\n');
  function pushClip(c) {
    A.clip = c;
    c.id = L.uid('clip');
    A.clipHistory.unshift(c);
    if (A.clipHistory.length > 24) A.clipHistory.pop();
    L.panes.task.refresh('clip');
    ui.refresh();
  }
  A.copyShapes = function (cut) {
    const shapes = E.selected().filter((s) => !(E.find(s.id) && E.find(s.id).parent && false));
    if (!shapes.length) return null;
    const c = { kind: 'shapes', shapes: L.clone(shapes), design: E.design().id, label: (shapes.length > 1 ? shapes.length + ' objects: ' : '') + (clipText(shapes) || shapes[0].name), text: clipText(shapes) };
    pushClip(c);
    if (cut) A.deleteSel();
    return c;
  };
  A.copySlides = function (cut) {
    const ss = A.selectedSlides();
    if (!ss.length) return null;
    const designs = {};
    for (const s of ss) designs[s.design] = L.clone(L.pres.designs[s.design]);
    const c = { kind: 'slides', slides: L.clone(ss), designs, label: ss.length + ' slide' + (ss.length > 1 ? 's' : '') + ': ' + (M.slideTitle(ss[0]) || 'Untitled'), text: ss.map((s) => M.allText(s)).join('\n\n') };
    pushClip(c);
    if (cut) A.deleteSlides();
    return c;
  };
  A.cmdClipboard = function (op) {
    if (TE.active()) {
      const txt = TE.selectedText();
      if (txt) pushClip({ kind: 'text', text: txt, label: txt.slice(0, 80) });
      try { document.execCommand(op); } catch (e) { /* ignore */ }
      return;
    }
    const c = A.focusArea === 'slides' || A.view === 'sorter' ? A.copySlides(op === 'cut') : A.copyShapes(op === 'cut');
    if (c && c.text && navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(c.text).catch(() => {});
  };
  A.paste = function () { if (A.clip) A.pasteItem(A.clip); };
  A.pasteItem = function (c) {
    if (!c) return;
    if (c.kind === 'text') {
      if (TE.active()) { L.hist.push('Paste'); TE.insertText(c.text); return; }
      return A.pasteText(c.text);
    }
    if (c.kind === 'slides') {
      const at = (A.view === 'sorter' || A.focusArea === 'slides') && A.slideSel.size ? Math.max(...A.slideSel) + 1 : E.idx + 1;
      L.hist.push('Paste Slides');
      const added = c.slides.map((s) => { const d = M.dupSlide(s); if (!L.pres.designs[d.design]) L.pres.designs[d.design] = L.clone(c.designs[s.design]); return d; });
      L.pres.slides.splice(at, 0, ...added);
      A.slideSel = new Set(added.map((_, k) => at + k));
      E.goto(at, { force: true });
      L.bus.emit('slides-changed');
      return;
    }
    if (c.kind === 'shapes') {
      if (TE.active()) { L.hist.push('Paste'); TE.insertText(c.text || ''); return; }
      const slide = E.slide();
      if (!slide) return;
      E.commit('Paste', () => {
        const ids = [];
        for (const s of c.shapes) {
          const d = M.dup(s);
          const clash = slide.shapes.some((x) => Math.abs(x.x - d.x) < 1 && Math.abs(x.y - d.y) < 1);
          if (clash) M.translate(d, 12, 12);
          slide.shapes.push(d);
          ids.push(d.id);
        }
        E.sel = ids;
      });
      A.updateToolbars();
    }
  };
  A.pasteText = function (text) {
    const slide = E.slide();
    if (!slide || !text) return;
    const lines = text.replace(/\r/g, '').split('\n');
    L.hist.push('Paste');
    const tb = M.newTextBox(L.pres, slide, L.pres.W * 0.1, L.pres.H * 0.3, L.pres.W * 0.8, 40, true);
    tb.tx.ps = lines.map((l) => T.para(l));
    slide.shapes.push(tb);
    E.sel = [tb.id];
    E.render();
    E.touched();
  };
  A.pasteSpecial = function () {
    const c = A.clip;
    if (!c) return;
    let mode = 'native';
    const body = h('div', { class: 'col' }, h('div', { text: 'Source: ' + c.label.slice(0, 60) }), ui.radio('ps', '&Paste as ' + (c.kind === 'slides' ? 'slides' : 'Lectern objects'), true, () => { mode = 'native'; }), ui.radio('ps', 'Paste as &Unformatted Text', false, () => { mode = 'text'; }));
    ui.dialog({ title: 'Paste Special', body, width: 360, buttons: [{ label: 'OK', primary: true, onClick: () => (mode === 'text' ? (TE.active() ? (L.hist.push('Paste'), TE.insertText(c.text || '')) : A.pasteText(c.text || '')) : A.pasteItem(c)) }, { label: 'Cancel' }] });
  };
  let pendingClip = null;
  document.addEventListener('copy', (e) => onClipEvent(e, 'copy'));
  document.addEventListener('cut', (e) => onClipEvent(e, 'cut'));
  function inNativeText(t) { return t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable); }
  function onClipEvent(e, op) {
    pendingClip = null;
    if (L.show.active || ui.dialogOpen()) return;
    const t = document.activeElement;
    if (TE.active()) { const txt = TE.selectedText(); if (txt) pushClip({ kind: 'text', text: txt, label: txt.slice(0, 80) }); return; }
    if (inNativeText(t)) return;
    const c = A.focusArea === 'slides' || A.view === 'sorter' ? A.copySlides(op === 'cut') : A.copyShapes(op === 'cut');
    if (c) { e.preventDefault(); try { e.clipboardData.setData('text/plain', c.text || ''); e.clipboardData.setData('application/x-lectern', c.id); } catch (er) { /* ignore */ } }
  }
  document.addEventListener('paste', (e) => {
    pendingClip = null;
    if (L.show.active || ui.dialogOpen()) return;
    if (TE.active()) return; /* text editor handles its own paste */
    if (inNativeText(document.activeElement)) return;
    const dt = e.clipboardData;
    if (!dt) return;
    e.preventDefault();
    const marker = dt.getData('application/x-lectern');
    const files = Array.from(dt.files || []).filter((f) => /^image\//.test(f.type));
    const text = dt.getData('text/plain');
    if (A.clip && (marker === A.clip.id || (text && text === A.clip.text))) { A.pasteItem(A.clip); return; }
    if (files.length) { A.insertPictureFiles(files); return; }
    if (text) { const sel = E.primary(); if (sel && sel.tx && E.sel.length === 1) { TE.begin(sel.id, { atEnd: true }); L.hist.push('Paste'); TE.insertText(text); } else A.pasteText(text); }
  });

  /* ================= context menus ================= */
  A.contextShape = function (e) {
    const s = E.primary();
    const items = ['cut', 'copy', 'paste', '-'];
    if (s && s.type === 'wordart') items.push('wordartEdit', 'wordartGallery');
    else if (s && s.type === 'chart') items.push('editChart');
    else if (s && (s.tx || s.type === 'shape')) items.push({ label: '&Edit Text', run: () => TE.begin(s.id, { atEnd: true }) });
    if (s && s.type === 'table') items.push({ label: '&Table', sub: ['colLeft', 'colRight', 'rowAbove', 'rowBelow', '-', 'deleteCols', 'deleteRows', '-', 'mergeCells', 'splitCell', '-', 'tableBordersFill'] });
    items.push('-', { label: '&Grouping', sub: ['group', 'ungroup', 'regroup'] }, { label: '&Order', sub: ['bringFront', 'sendBack', 'bringForward', 'sendBackward'] }, '-', 'setDefaults', 'customAnim', 'actionSettings', 'formatObject', 'hyperlink');
    ui.contextMenu(e, items);
  };
  A.contextSlide = (e) => ui.contextMenu(e, ['cut', 'copy', 'paste', '-', 'ruler', 'gridGuides', '-', 'slideDesign', 'slideLayout', 'background', '-', 'transitionPane', 'newSlide']);
  A.contextText = (e) => ui.contextMenu(e, ['cut', 'copy', 'paste', '-', 'fontDlg', 'bulletsDlg', 'lineSpacing', '-', 'hyperlink', '-', { label: 'Exit Edit Text', run: () => { const id = TE.target().sh.id; TE.end(); E.select([id]); } }]);
  A.contextThumb = (e) => ui.contextMenu(e, ['cut', 'copy', 'paste', '-', 'newSlide', 'deleteSlide', 'duplicateSlide', '-', 'slideDesign', 'slideLayout', 'background', 'transitionPane', '-', 'hideSlide']);

  /* ================= keyboard ================= */
  function keyName(e) {
    const parts = [];
    if (e.ctrlKey || e.metaKey) parts.push('Ctrl');
    if (e.altKey) parts.push('Alt');
    if (e.shiftKey) parts.push('Shift');
    let k = e.key;
    if (k === ' ') k = 'Space';
    if (k.length === 1) k = k.toUpperCase();
    if (k === '>' || (k === '.' && e.shiftKey)) k = '>';
    if (k === '<' || (k === ',' && e.shiftKey)) k = '<';
    parts.push(k);
    return parts.join('+');
  }
  const SHORTCUTS = {
    'Ctrl+N': 'new', 'Ctrl+O': 'open', 'Ctrl+S': 'save', 'Ctrl+P': 'print', 'Ctrl+Z': 'undo', 'Ctrl+Y': 'redo', 'Ctrl+Shift+Z': 'redo',
    'Ctrl+M': 'newSlide', 'Ctrl+D': 'duplicate', 'Ctrl+F': 'find', 'Ctrl+H': 'replace', 'Ctrl+K': 'hyperlink', 'Ctrl+T': 'fontDlg',
    'Ctrl+B': 'bold', 'Ctrl+I': 'italic', 'Ctrl+U': 'underline', 'Ctrl+E': 'alignCenter', 'Ctrl+L': 'alignLeft', 'Ctrl+R': 'alignRight', 'Ctrl+J': 'justify',
    'Ctrl+Shift+>': 'growFont', 'Ctrl+Shift+<': 'shrinkFont', 'Ctrl+]': 'growFont', 'Ctrl+[': 'shrinkFont', 'Ctrl+=': 'subscript', 'Ctrl+Shift+=': 'superscript', 'Ctrl+Shift++': 'superscript', 'Ctrl+Space': 'resetChar',
    'Shift+F3': 'changeCase', 'Ctrl+G': 'group', 'Ctrl+Shift+G': 'ungroup', 'F5': 'showFromStart', 'Shift+F5': 'showFromCurrent', 'F7': 'spelling', 'F1': 'help', 'Ctrl+F1': 'taskPane',
    'Alt+Shift+ArrowLeft': 'promote', 'Alt+Shift+ArrowRight': 'demote', 'Alt+Shift+ArrowUp': 'moveParaUp', 'Alt+Shift+ArrowDown': 'moveParaDown', 'F12': 'saveAs', 'Ctrl+Shift+S': 'saveAs',
  };
  const TEXT_SAFE = new Set(['bold', 'italic', 'underline', 'alignCenter', 'alignLeft', 'alignRight', 'justify', 'growFont', 'shrinkFont', 'subscript', 'superscript', 'resetChar', 'fontDlg', 'hyperlink', 'changeCase', 'promote', 'demote', 'moveParaUp', 'moveParaDown', 'undo', 'redo', 'save', 'open', 'new', 'print', 'newSlide', 'find', 'replace', 'showFromStart', 'showFromCurrent', 'help', 'spelling', 'taskPane', 'saveAs']);
  document.addEventListener('keydown', (e) => {
    if (L.show.active) return;
    if (ui.dialogOpen()) { ui.dialogKey(e); return; }
    if (ui.menuOpen()) { if (ui.menuKey(e)) { e.preventDefault(); e.stopPropagation(); } return; }
    if (e.key === 'F10' || (e.key === 'Alt' && false)) { e.preventDefault(); A.menuBar.openIndex(0); return; }
    if (e.altKey && !e.ctrlKey && !e.shiftKey && e.key.length === 1 && /[a-z]/i.test(e.key)) { if (A.menuBar.openByKey(e.key)) { e.preventDefault(); return; } }
    const k = keyName(e);
    const tgt = e.target;
    const inInput = tgt && (tgt.tagName === 'INPUT' || tgt.tagName === 'TEXTAREA' || tgt.tagName === 'SELECT');
    if (inInput) { if (k === 'Escape') { tgt.blur(); E.refocus(); } return; }
    if (L.panes.preview.el.classList.contains('on') && k === 'Escape') { A.setView(A.prevView || 'normal'); return; }
    const cmd = SHORTCUTS[k];
    /* text editing */
    if (TE.active()) {
      if (TE.onKey(e)) return;
      if (cmd && TEXT_SAFE.has(cmd)) { e.preventDefault(); ui.exec(cmd); return; }
      if (k === 'Ctrl+A') return;
      return;
    }
    const otherEditable = tgt && tgt.isContentEditable;
    if (otherEditable) {
      if (cmd && ['save', 'open', 'new', 'print', 'undo', 'redo', 'showFromStart', 'showFromCurrent', 'help', 'newSlide', 'taskPane', 'saveAs'].includes(cmd) && !(L.panes.left.tab === 'outline' && (cmd === 'undo' || cmd === 'redo'))) { e.preventDefault(); ui.exec(cmd); }
      return;
    }
    if (k === 'Ctrl+C' || k === 'Ctrl+X' || k === 'Ctrl+V' || k === 'Ctrl+Insert' || k === 'Shift+Delete' || k === 'Shift+Insert') {
      const op = /C$|Insert$/.test(k) && !/Shift\+Insert/.test(k) ? 'copy' : /X$|Shift\+Delete/.test(k) ? 'cut' : 'paste';
      pendingClip = op;
      setTimeout(() => { if (pendingClip === op) { pendingClip = null; if (op === 'paste') A.paste(); else A.cmdClipboard(op); } }, 60);
      return;
    }
    if (cmd) { e.preventDefault(); ui.exec(cmd); return; }
    if (k === 'Ctrl+A') { e.preventDefault(); A.selectAll(); return; }
    if (k === 'Escape') {
      if (E.tool) { E.cancelFreeform(true); E.setTool(null); return; }
      if (E.cropMode) { E.cropMode = false; E.drawOverlay(); return; }
      if (A.view === 'master') { return; }
      const f = E.sel.length === 1 && E.find(E.sel[0]);
      if (f && f.parent) { E.select([E.topOf(E.sel[0]).id]); return; }
      E.clearSel(); return;
    }
    /* slide list / sorter */
    if (A.focusArea === 'slides' || A.view === 'sorter') {
      const n = L.pres.slides.length;
      if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); A.deleteSlides(); return; }
      if (k === 'Enter') { e.preventDefault(); A.newSlideCmd(); return; }
      const step = { ArrowUp: -1, ArrowLeft: -1, ArrowDown: 1, ArrowRight: 1, PageUp: -1, PageDown: 1 }[e.key];
      if (step && !e.ctrlKey) {
        e.preventDefault();
        const cols = A.view === 'sorter' ? Math.max(1, Math.round(L.panes.sorter.el.clientWidth / ((L.$('.so-it') && L.$('.so-it').offsetWidth + 24) || 200))) : 1;
        const d = A.view === 'sorter' && (e.key === 'ArrowUp' || e.key === 'ArrowDown') ? step * cols : step;
        const i = L.clamp(E.idx + d, 0, n - 1);
        if (e.shiftKey) A.slideSel.add(i); else A.slideSel = new Set([i]);
        if (A.view === 'sorter') { E.idx = i; L.panes.sorter.mark(); const el = L.$(`.so-it[data-i="${i}"]`); if (el) el.scrollIntoView({ block: 'nearest' }); }
        else E.goto(i);
        return;
      }
      if (k === 'Home' || k === 'End') { e.preventDefault(); const i = k === 'Home' ? 0 : n - 1; A.slideSel = new Set([i]); if (A.view === 'sorter') { E.idx = i; L.panes.sorter.mark(); } else E.goto(i); return; }
      return;
    }
    if (A.view !== 'normal' && A.view !== 'master') return;
    /* slide editing area */
    const sel = E.selected();
    if (k === 'Delete' || k === 'Backspace') { if (sel.length) { e.preventDefault(); A.deleteSel(); } return; }
    if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); if (A.view === 'normal') E.goto(E.idx + (k === 'PageDown' ? 1 : -1)); return; }
    if (k === 'Home' || k === 'End') { if (!sel.length && A.view === 'normal') { e.preventDefault(); E.goto(k === 'Home' ? 0 : L.pres.slides.length - 1); } return; }
    const arrows = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] };
    if (arrows[e.key] && !e.altKey) {
      e.preventDefault();
      if (sel.length) A.nudge(arrows[e.key][0], arrows[e.key][1], e.ctrlKey);
      else if (A.view === 'normal') E.goto(E.idx + (e.key === 'ArrowDown' || e.key === 'ArrowRight' ? 1 : -1));
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      const shapes = E.slide() ? E.slide().shapes : [];
      if (!shapes.length) return;
      const i = shapes.findIndex((s) => s.id === E.sel[0]);
      const j = (i + (e.shiftKey ? -1 : 1) + shapes.length) % shapes.length;
      E.select([shapes[i < 0 ? 0 : j].id]);
      return;
    }
    if ((k === 'Enter' || k === 'F2') && sel.length === 1) {
      const s = sel[0];
      e.preventDefault();
      if (s.type === 'chart') ui.exec('editChart');
      else if (s.type === 'wordart') A.editWordArt();
      else if (s.tx || s.type === 'shape' || s.type === 'table') TE.begin(s.id, { selectAll: k === 'F2', atEnd: k === 'Enter' });
      return;
    }
    if (k === 'Ctrl+Enter') {
      e.preventDefault();
      const phs = (E.slide() ? E.slide().shapes : []).filter((s) => s.ph && s.type === 'text');
      const i = phs.findIndex((s) => s.id === E.sel[0]);
      if (i + 1 < phs.length) TE.begin(phs[i + 1].id, { selectAll: true }); else A.insertSlide('text');
      return;
    }
    /* typing on a selected text object replaces its text */
    if (sel.length === 1 && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      const s = sel[0];
      if (s.tx || s.type === 'shape') {
        e.preventDefault();
        TE.begin(s.id, { selectAll: true });
        if (TE.active()) { L.hist.push('Typing'); TE.insertText(e.key); }
      }
    }
  });

  /* ================= drag & drop ================= */
  function setupDrop() {
    let hint = null, depth = 0;
    const app = L.$('#app');
    app.addEventListener('dragenter', (e) => { if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return; depth++; if (!hint) { hint = h('div', { class: 'drop-hint', text: 'Drop a .pptx to open it, or pictures to insert them' }); document.body.appendChild(hint); } });
    app.addEventListener('dragleave', () => { depth--; if (depth <= 0 && hint) { hint.remove(); hint = null; depth = 0; } });
    app.addEventListener('dragover', (e) => { if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) e.preventDefault(); });
    app.addEventListener('drop', (e) => {
      if (hint) { hint.remove(); hint = null; } depth = 0;
      const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []);
      if (!files.length) return;
      e.preventDefault();
      const deck = files.find((f) => /\.(pptx|ppsx|potx|pptm|ppt)$/i.test(f.name));
      if (deck) { A.openFile(deck); return; }
      const imgs = files.filter((f) => /^image\//.test(f.type));
      if (imgs.length) A.insertPictureFiles(imgs);
      const txt = files.find((f) => /\.txt$/i.test(f.name));
      if (txt) L.readAsText(txt).then((t) => A.outlineToSlides(t, E.idx + 1));
    });
  }

  /* ================= splitters ================= */
  function setupSplitters() {
    const sl = L.$('#split-left'), sn = L.$('#split-notes');
    sl.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      sl.setPointerCapture(e.pointerId);
      const x0 = e.clientX, w0 = L.$('#leftpane').offsetWidth;
      const mv = (ev) => { const w = L.clamp(w0 + ev.clientX - x0, 110, 420); L.$('#leftpane').style.width = w + 'px'; A.opts.leftW = w; };
      const up = () => { sl.removeEventListener('pointermove', mv); sl.removeEventListener('pointerup', up); L.panes.left.render(); E.layout(); A.saveOpts(); };
      sl.addEventListener('pointermove', mv); sl.addEventListener('pointerup', up);
    });
    sn.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      sn.setPointerCapture(e.pointerId);
      const y0 = e.clientY, h0 = L.$('#notes').offsetHeight;
      const mv = (ev) => { const hh = L.clamp(h0 - (ev.clientY - y0), 24, 420); L.$('#notes').style.height = hh + 'px'; A.opts.notesH = hh; };
      const up = () => { sn.removeEventListener('pointermove', mv); sn.removeEventListener('pointerup', up); E.layout(); A.saveOpts(); };
      sn.addEventListener('pointermove', mv); sn.addEventListener('pointerup', up);
    });
  }

  /* ================= recovery (per-browser convenience) ================= */
  const autosave = L.debounce(async () => {
    if (!L.hist.dirty) return;
    try {
      const media = {};
      let total = 0;
      for (const [id, m] of L.media.all()) { if (total > 3e6) break; const d = await L.blobToDataURL(m.blob); total += d.length; media[id] = d; }
      const snap = { at: Date.now(), name: A.fileName, pres: JSON.parse(L.hist.snapshot()), media };
      const s = JSON.stringify(snap);
      if (s.length < 4.5e6) L.store.set('recover', snap); else L.store.del('recover');
    } catch (e) { /* storage is optional */ }
  }, 4000);
  async function offerRecovery() {
    const snap = L.store.get('recover', null);
    if (!snap || !snap.pres || !snap.pres.slides || !snap.pres.slides.length) return false;
    const r = await ui.msg(`Lectern found an unsaved version of "${snap.name}" from ${new Date(snap.at).toLocaleString()}.\nDo you want to restore it?`, { title: 'Document Recovery', icon: 'question', buttons: ['&Restore', '&Discard'] });
    if (r !== 0) { L.store.del('recover'); return false; }
    for (const id in snap.media) {
      try { const blob = await (await fetch(snap.media[id])).blob(); const m = L.media.add(blob, id); const re = new RegExp(id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'); snap.pres = JSON.parse(JSON.stringify(snap.pres).replace(re, m)); } catch (e) { /* skip */ }
    }
    const pres = Object.assign(M.newPresentation({ empty: true }), snap.pres);
    A.loadPres(pres, snap.name + ' (Recovered)');
    L.hist.dirty = true;
    return true;
  }

  /* ================= sample presentation ================= */
  function buildSample() {
    const pres = M.newPresentation({ design: 'azure', empty: true });
    const did = Object.keys(pres.designs)[0];
    pres.props.title = 'Harbor & Pine Coffee — 2027 Expansion Plan (sample)';
    pres.props.author = 'Sample';
    const add = (layout) => { const s = M.newSlide(pres, layout, did); pres.slides.push(s); return s; };
    const ph = (s, t) => s.shapes.find((x) => x.ph && x.ph.type === t);
    /* 1 */
    let s = add('title');
    T.setPlain(ph(s, 'ctrTitle').tx, 'Harbor & Pine Coffee');
    ph(s, 'subTitle').tx.ps = [T.para('2027 Expansion Plan'), T.para('Sample presentation · Board review, October 2026', null, { sz: 18, i: true })];
    s.trans = { type: 'fade', spd: 'med', click: true, after: null };
    s.notes = 'This is a sample deck. Open your own .pptx from File ▸ Open, or start fresh with File ▸ New.';
    /* 2 */
    s = add('text');
    T.setPlain(ph(s, 'title').tx, 'Where we stand today');
    const b = ph(s, 'body');
    b.tx.ps = T.fromLines(['14 cafés across three coastal cities', 'Same-store sales up 11% year over year', '\tMorning traffic is the strongest daypart', 'Wholesale beans now 22% of revenue', 'Roastery running at 64% of capacity']);
    s.anims = [{ id: L.uid('a'), sid: b.id, cls: 'entr', eff: 'fade', start: 'click', dur: 500, delay: 0, by: 'para' }];
    s.trans = { type: 'push', dir: 'u', spd: 'med', click: true, after: null };
    s.notes = 'Lead with momentum: comparable-store growth and a roastery with headroom.';
    /* 3 */
    s = add('chart');
    T.setPlain(ph(s, 'title').tx, 'Revenue by region ($K)');
    const cp = ph(s, 'chart');
    const chart = { id: L.uid('s'), type: 'chart', name: 'Chart 2', ph: L.clone(cp.ph), x: cp.x, y: cp.y, w: cp.w, h: cp.h, rot: 0, chart: { kind: 'col', title: '', legend: 'b', gridY: true, labels: false, fsz: 12, cats: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Portland', vals: [410, 455, 520, 560], color: '#4A7BC8' }, { name: 'Seattle', vals: [380, 402, 446, 490], color: '#E0A526' }, { name: 'Monterey', vals: [120, 138, 171, 205], color: '#2D6BB5' }] } };
    s.shapes.splice(s.shapes.indexOf(cp), 1, chart);
    s.anims = [{ id: L.uid('a'), sid: chart.id, cls: 'entr', eff: 'wipe', dir: 'b', start: 'after', dur: 1000, delay: 0, by: 'all' }];
    s.trans = { type: 'wipe', dir: 'r', spd: 'med', click: true, after: null };
    /* 4 */
    s = add('table');
    T.setPlain(ph(s, 'title').tx, 'Rollout schedule');
    const tp = ph(s, 'tbl');
    const t = M.newTable(pres, s, 5, 4, tp.x, tp.y, tp.w);
    t.ph = L.clone(tp.ph);
    const data = [['Phase', 'Market', 'New stores', 'Opens'], ['1', 'Tacoma, WA', '2', 'March 2027'], ['2', 'Santa Cruz, CA', '3', 'June 2027'], ['3', 'Bend, OR', '2', 'September 2027'], ['4', 'Boise, ID', '3', 'January 2028']];
    data.forEach((row, ri) => row.forEach((v, ci) => {
      const cell = t.tbl.rows[ri].cells[ci];
      cell.tx.ps = [T.para(v, { algn: ci === 0 || ci === 2 ? 'ctr' : 'l' }, ri === 0 ? { sz: 20, b: true, color: 'lt1' } : { sz: 20 })];
      cell.fill = ri === 0 ? { t: 'solid', c: 'accent1', a: 1 } : ri % 2 ? { t: 'solid', c: '#E4ECF8', a: 1 } : { t: 'solid', c: '#FFFFFF', a: 1 };
      for (const k of ['l', 'r', 't', 'b']) cell.bd[k] = { c: '#9BB7E3', w: 1, dash: 'solid' };
      cell.tx.anchor = 'ctr';
    }));
    t.tbl.rows.forEach((r) => { r.h = 46; });
    t.h = 46 * 5;
    t.tbl.cols = [tp.w * 0.14, tp.w * 0.38, tp.w * 0.2, tp.w * 0.28];
    s.shapes.splice(s.shapes.indexOf(tp), 1, t);
    s.trans = { type: 'cover', dir: 'l', spd: 'med', click: true, after: null };
    s.notes = 'Ten new stores over five quarters; each phase starts only after the prior one is cash-positive.';
    /* 5 */
    s = add('titleOnly');
    T.setPlain(ph(s, 'title').tx, 'From lease to first pour');
    const steps = [['Site & lease', '8 weeks'], ['Build-out', '12 weeks'], ['Hire & train', '4 weeks'], ['Soft open', '2 weeks']];
    const W = pres.W, bw = W * 0.22, gap = W * 0.01, x0 = (W - (bw * 4 + gap * 3)) / 2, y0 = pres.H * 0.42;
    steps.forEach(([label, dur], i) => {
      const sh = M.newShape(pres, s, i === 0 ? 'homePlate' : 'chevron', x0 + i * (bw + gap), y0, bw, pres.H * 0.16);
      sh.fill = { t: 'grad', stops: [{ p: 0, c: L.color.lighten('#4A7BC8', 0.25), a: 1 }, { p: 1, c: '#2A4E8F', a: 1 }], ang: 90, path: 'lin' };
      sh.line = { t: 'none', w: 0.75 };
      sh.tx.ps = [T.para(label, { algn: 'ctr' }, { sz: 15, b: true, color: 'lt1' })];
      sh.tx.ins = [4, 3.6, 4, 3.6];
      s.shapes.push(sh);
      const tb = M.newTextBox(pres, s, x0 + i * (bw + gap), y0 + pres.H * 0.18, bw, 30, true);
      tb.tx.ps = [T.para(dur, { algn: 'ctr' }, { sz: 16, color: 'accent1', i: true })];
      s.shapes.push(tb);
      s.anims.push({ id: L.uid('a'), sid: sh.id, cls: 'entr', eff: 'wipe', dir: 'l', start: i === 0 ? 'click' : 'after', dur: 500, delay: 0, by: 'all' });
      s.anims.push({ id: L.uid('a'), sid: tb.id, cls: 'entr', eff: 'fade', start: 'with', dur: 500, delay: 200, by: 'all' });
    });
    const total = M.newTextBox(pres, s, x0, y0 + pres.H * 0.3, bw * 4 + gap * 3, 30, true);
    total.tx.ps = [T.para('About 26 weeks from signed lease to opening day', { algn: 'ctr' }, { sz: 18 })];
    s.shapes.push(total);
    s.trans = { type: 'fade', spd: 'med', click: true, after: null };
    /* 6 */
    s = add('text');
    T.setPlain(ph(s, 'title').tx, 'What we need from the board');
    const b6 = ph(s, 'body');
    b6.tx.ps = T.fromLines(['Approve the $4.2M capital plan', 'Green-light a second roaster in Q2', 'Name an executive sponsor for Boise']);
    b6.w = pres.W * 0.62;
    const call = M.newShape(pres, s, 'wedgeRoundRectCallout', pres.W * 0.7, pres.H * 0.3, pres.W * 0.25, pres.H * 0.2);
    call.adj = { adj1: -62000, adj2: 48000 };
    call.fill = { t: 'solid', c: 'accent2', a: 1 };
    call.line = { c: '#9A6B00', w: 1, dash: 'solid' };
    call.shadow = { c: '#000000', a: 0.3, dx: 3, dy: 3, blur: 4 };
    call.tx.ps = [T.para('Decision needed today', { algn: 'ctr' }, { sz: 20, b: true, color: 'dk1' })];
    s.shapes.push(call);
    s.anims = [{ id: L.uid('a'), sid: b6.id, cls: 'entr', eff: 'flyIn', dir: 'l', start: 'click', dur: 500, delay: 0, by: 'para' }, { id: L.uid('a'), sid: call.id, cls: 'entr', eff: 'zoom', start: 'after', dur: 500, delay: 250, by: 'all' }, { id: L.uid('a'), sid: call.id, cls: 'emph', eff: 'teeter', start: 'after', dur: 1000, delay: 0, by: 'all' }];
    s.trans = { type: 'push', dir: 'l', spd: 'med', click: true, after: null };
    /* 7 */
    s = add('blank');
    const wa = { id: L.uid('s'), type: 'wordart', name: 'WordArt 1', x: pres.W * 0.15, y: pres.H * 0.3, w: pres.W * 0.7, h: pres.H * 0.32, rot: 0, wa: Object.assign(L.clone(L.dlg.WA_STYLES[4]), { text: 'Thank you', font: 'Arial Black', b: true, warp: 'textArchUp', style: 4 }) };
    s.shapes.push(wa);
    const q = M.newTextBox(pres, s, pres.W * 0.15, pres.H * 0.66, pres.W * 0.7, 36, true);
    q.tx.ps = [T.para('Questions? expansion@harborandpine.example', { algn: 'ctr' }, { sz: 20, color: 'tx1' })];
    s.shapes.push(q);
    s.anims = [{ id: L.uid('a'), sid: wa.id, cls: 'entr', eff: 'zoom', start: 'after', dur: 750, delay: 0, by: 'all' }];
    s.trans = { type: 'circle', spd: 'med', click: true, after: null };
    pres.hf = Object.assign(pres.hf, { num: true, notOnTitle: true });
    return pres;
  }
  A.buildSample = buildSample;

  /* ================= wiring ================= */
  function wire() {
    L.bus.on('slide-changed', () => {
      A.slideSel = A.view === 'sorter' ? A.slideSel : new Set([E.idx]);
      L.panes.left.markCurrent();
      L.panes.notes.render();
      if (A.view === 'notes') L.panes.notesView.render();
      L.panes.task.refresh('slide');
      A.updateStatus();
      A.updateToolbars();
      ui.refresh();
    });
    L.bus.on('slide-changed-silent', () => { L.panes.notes.render(); A.updateStatus(); L.panes.task.refresh('slide'); });
    const thumbSoon = L.debounce((i) => { L.panes.left.refreshThumb(i); if (A.view === 'sorter') L.panes.sorter.render(); }, 350);
    L.bus.on('slide-modified', (i) => { if (E.view === 'master') { L.panes.left.render(); return; } thumbSoon(i); autosave(); if (L.panes.left.tab === 'outline' && !TE.active()) L.panes.left.render(); });
    L.bus.on('slides-changed', (o) => {
      if (E.idx >= L.pres.slides.length) E.idx = Math.max(0, L.pres.slides.length - 1);
      if (!(o && o.keepOutline)) L.panes.left.render();
      if (A.view === 'sorter') L.panes.sorter.render();
      if (A.view === 'notes') L.panes.notesView.render();
      L.panes.notes.render();
      L.panes.task.refresh(o && o.design ? 'design' : o && o.trans ? 'trans' : 'slides');
      A.updateStatus();
      autosave();
      ui.refresh();
    });
    L.bus.on('selection', () => { A.updateToolbars(); L.panes.task.refresh('selection'); });
    L.bus.on('edit-start', () => { A.updateToolbars(); ui.refresh(); });
    L.bus.on('edit-end', () => { ui.refresh(); });
    window.addEventListener('beforeunload', (e) => { if (L.hist.dirty) { e.preventDefault(); e.returnValue = ''; } });
    window.addEventListener('resize', L.debounce(() => { if (A.view === 'sorter') L.panes.sorter.render(); if (A.view === 'notes') L.panes.notesView.render(); if (A.view === 'preview') L.panes.preview.render(); L.panes.left.render(); }, 250));
    A.outlineDirty = () => {};
  }

  /* ================= init ================= */
  A.init = async function () {
    A.loadOpts();
    L.$('#appicon').innerHTML = L.icons.app(16);
    buildMenus();
    buildToolbars();
    buildStatus();
    L.pres = buildSample();
    L.panes.left.mount(L.$('#leftpane'));
    L.panes.notes.mount(L.$('#notes'));
    L.panes.task.mount(L.$('#taskpane'));
    E.mount(L.$('#editorwrap'));
    A.altView = h('div', { id: 'altview', style: 'flex:1;min-width:0;position:relative', hidden: true });
    L.$('#main').insertBefore(A.altView, L.$('#taskpane'));
    L.panes.sorter.mount(A.altView);
    L.panes.notesView.mount(A.altView);
    L.panes.preview.mount(L.$('#main'));
    L.$('#leftpane').style.width = (A.opts.leftW || 196) + 'px';
    L.$('#notes').style.height = (A.opts.notesH || 92) + 'px';
    const narrow = window.innerWidth < 820;
    if (narrow) { A.opts.taskOpen = false; if (window.innerWidth < 640) A.opts.leftOpen = false; }
    L.$('#taskpane').hidden = !A.opts.taskOpen;
    setupSplitters();
    setupDrop();
    wire();
    A.loadPres(L.pres, 'Harbor & Pine — 2027 Expansion (sample)');
    const wantTask = A.opts.taskOpen;
    L.panes.task.show('getting-started');
    if (!wantTask) { L.$('#taskpane').hidden = true; A.opts.taskOpen = false; }
    A.applyOpts();
    E.layout();
    const prompt = window.self !== window.top ? null : null;
    void prompt;
    setTimeout(() => { offerRecovery(); }, 600);
    ui.refresh();
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => A.init());
  else A.init();
})();
