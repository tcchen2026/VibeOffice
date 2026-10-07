/* Ledger — the task pane (Getting Started, New Workbook, Help, Clipboard, Clip Art) and the spreadsheet templates. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, F = L.formula, O = L.ops;
  const { h } = L;
  const ui = L.ui;
  const P = (L.panes = {});
  const A = () => L.app;
  const G = () => L.grid;

  /* ================= shell ================= */
  const TP = (P.task = { stack: [], pos: -1, registry: {} });
  TP.mount = function (el) {
    TP.el = el;
    const nav = (icon, label) => h('button', { class: 'tp-nav', type: 'button', 'aria-label': label, 'data-tip': label, html: L.icons.get(icon) || label[0] });
    TP.back = nav('back', 'Back'); TP.fwd = nav('forward', 'Forward'); TP.home = nav('home', 'Home');
    TP.title = h('div', { class: 'tp-title', role: 'button', tabindex: '0', 'aria-haspopup': 'menu' }, h('span', { text: '' }), h('span', { class: 'dd-arrow' }));
    const close = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Close task pane', 'data-tip': 'Close', html: '<svg width="10" height="10" viewBox="0 0 10 10"><path d="M1 1l8 8M9 1L1 9" stroke="#000" stroke-width="1.6"/></svg>' });
    TP.body = h('div', { class: 'tp-body' });
    el.append(h('div', { class: 'tp-head' }, TP.back, TP.fwd, TP.home, TP.title, close), TP.body);
    TP.back.addEventListener('click', () => TP.go(TP.pos - 1));
    TP.fwd.addEventListener('click', () => TP.go(TP.pos + 1));
    TP.home.addEventListener('click', () => TP.show('getting-started'));
    close.addEventListener('click', () => A().toggleTask(false));
    const openList = () => { const r = TP.title.getBoundingClientRect(); ui.openMenu(Object.values(TP.registry).filter((p) => !p.hidden).map((p) => ({ label: p.title, run: () => TP.show(p.id), checked: () => TP.current === p.id })), { left: r.left, bottom: r.bottom }); };
    TP.title.addEventListener('click', openList);
    TP.title.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openList(); } });
    L.bus.on('clipboard', () => TP.refresh('clip'));
  };
  TP.register = (p) => { TP.registry[p.id] = p; };
  TP.show = function (id) {
    if (!TP.registry[id] || !TP.el) return;
    if (TP.el.hidden) A().toggleTask(true);
    if (TP.stack[TP.pos] !== id) { TP.stack = TP.stack.slice(0, TP.pos + 1); TP.stack.push(id); TP.pos = TP.stack.length - 1; }
    TP.render();
  };
  TP.go = function (pos) { if (pos < 0 || pos >= TP.stack.length) return; TP.pos = pos; TP.render(); };
  TP.render = function () {
    const id = TP.stack[TP.pos];
    TP.current = id;
    const p = TP.registry[id];
    if (!p) return;
    TP.title.firstChild.textContent = p.title;
    TP.back.disabled = TP.pos <= 0; TP.fwd.disabled = TP.pos >= TP.stack.length - 1;
    TP.body.textContent = '';
    p.render(TP.body);
  };
  TP.refresh = function (reason) {
    const p = TP.registry[TP.current];
    if (!p || !TP.el || TP.el.hidden) return;
    if (p.live) { TP.body.textContent = ''; p.render(TP.body); }
    void reason;
  };
  const sec = (title, ...kids) => h('div', { class: 'tp-sec' }, title ? h('div', { class: 'tp-h' }, title) : null, ...kids);
  const link = (icon, label, fn, o) => h('button', Object.assign({ class: 'tp-link', type: 'button', onclick: fn }, o || {}), icon ? h('span', { html: L.icons.get(icon) || '' }) : null, h('span', { text: label }));
  P.sec = sec; P.link = link;

  /* ---------- Getting Started ---------- */
  /* ---------- Document Recovery: unsaved versions kept by common/suite.js ---------- */
  TP.register({ id: 'recovery', title: 'Document Recovery', render(b) { if (window.VO) VO.recoveryPane(b); } });
  TP.register({
    id: 'getting-started', title: 'Getting Started',
    render(b) {
      const q = h('input', { type: 'text', id: 'gs-search', placeholder: 'Search help', style: 'flex:1;min-width:0' });
      q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { P.helpQuery = q.value; TP.show('help'); } e.stopPropagation(); });
      b.append(
        h('div', { class: 'tp-row tp-brand' }, h('span', { html: L.icons.app ? L.icons.app(32) : '' }), h('b', { text: 'Ledger 2003 Web Edition' })),
        sec('Search for:', h('div', { class: 'tp-row' }, q, ui.button('Go', () => { P.helpQuery = q.value; TP.show('help'); }, { class: 'btn small' }))),
        sec('Open', window.VO ? VO.recentLinks(4) : null, link('open', 'More...', () => ui.exec('open')), link('new', 'Create a new workbook...', () => TP.show('new'))),
        sec('Learn', link('help', 'What Ledger can do', () => { P.helpQuery = ''; TP.show('help'); }), link('fx', 'Functions reference', () => L.dlg.insertFunction()), link('help', 'Keyboard shortcuts', () => TP.show('keys'))),
        sec('Files', h('div', { class: 'tp-note', text: 'Opens .xlsx, .xlsm, .xltx, .csv, .txt and XML Spreadsheet 2003 files; drop a file anywhere on the window to open it. Saves .xlsx, .xlsm, .xltx, .csv, .txt, .htm, .xml and PDF.' })));
    },
  });
  /* ---------- New Workbook ---------- */
  TP.register({
    id: 'new', title: 'New Workbook',
    render(b) {
      b.append(
        sec('New', link('new', 'Blank workbook', () => A().newWorkbook()), link('open', 'From existing workbook...', () => A().openDialog('.xlsx,.xlsm,.xltx,.xltm'))),
        sec('Templates', ...Object.keys(TPL).map((k) => link('sheetIcon', TPL[k].name, () => P.newFromTemplate(k)))),
        sec('', h('div', { class: 'tp-note', text: 'Templates open as new, unsaved workbooks.' })));
    },
  });
  /* ---------- Help ---------- */
  const HELP = [
    ['Entering data', 'Click a cell and type, then press Enter (moves down), Tab (moves right) or an arrow key. F2 edits the active cell; Alt+Enter starts a new line inside a cell. Ctrl+Enter fills the typed entry into every selected cell. Numbers, dates (3/14/2026, 14-Mar), times (9:30 AM), percentages (12%) and currency ($1,200) are recognised and formatted automatically. Start an entry with an apostrophe to keep it as text.'],
    ['Formulas', 'Start with = and use cell references (A1, $A$1, Sheet2!B3, named ranges, table references like Sales[Amount]). While typing a formula, click or drag on the sheet to insert references, and press F4 to cycle absolute and relative references. Ctrl+Shift+Enter enters an array formula; dynamic-array functions such as FILTER, SORT, UNIQUE and SEQUENCE spill their results. Insert ▸ Function (Shift+F3) lists more than 450 functions with argument help.'],
    ['AutoFill and series', 'Drag the fill handle (the square at the bottom-right of the selection) to copy cells or extend a series: 1, 2 → 3, 4…; Mon → Tue; Jan → Feb; Q1 → Q2; dates by day, weekday, month or year. Double-click the handle to fill down beside your data. Edit ▸ Fill ▸ Series gives full control; Ctrl+D and Ctrl+R fill down and right.'],
    ['Formatting cells', 'Format ▸ Cells (Ctrl+1) has Number, Alignment, Font, Border, Patterns and Protection tabs. Custom number formats use Excel codes such as #,##0.00;[Red](#,##0.00), 0%, d-mmm-yy, [h]:mm and "text"@. The Formatting toolbar sets fonts, alignment, Merge and Center, currency, percent, comma style, decimals, indents, borders, fill and font colours. The Format Painter copies formats.'],
    ['Rows, columns and sheets', 'Drag a header border to resize; double-click it to AutoFit. Right-click a header to insert, delete, hide or unhide. Window ▸ Freeze Panes keeps headings in view. Right-click a sheet tab to insert, delete, rename, move or copy sheets, or to set a tab colour; Ctrl+Page Up / Page Down switch sheets.'],
    ['Sorting and filtering', 'Data ▸ Sort sorts by up to three keys, ascending or descending, with or without a header row; the Sort buttons sort by the active column. Data ▸ Filter ▸ AutoFilter adds drop-down arrows to the header row: pick a value, (Top 10...), (Custom...), (Blanks) or (NonBlanks). Advanced Filter uses a criteria range and can copy unique records elsewhere. Data ▸ Subtotals inserts SUBTOTAL rows at each change in a column and outlines the list.'],
    ['Lists (tables)', 'Data ▸ List ▸ Create List (Ctrl+L) turns a range into a list with a header row, AutoFilter and an optional Total Row. Formulas can refer to it by name: =SUM(Table1[Amount]).'],
    ['Conditional formatting and validation', 'Format ▸ Conditional Formatting applies up to three conditions per range (cell value between, greater than…, or a formula), plus data bars, colour scales and icon sets read from newer workbooks. Data ▸ Validation restricts entries to whole numbers, decimals, lists, dates, times, text length or a custom formula, with input messages and error alerts.'],
    ['Charts and pictures', 'Insert ▸ Chart (F11) opens the Chart Wizard: column, bar, line, pie, XY scatter, area and doughnut charts, embedded in the sheet or on their own chart sheet. Charts follow their data as it changes. Insert ▸ Picture adds images and AutoShapes; drag to move, drag a handle to resize, double-click to format.'],
    ['Names, comments and hyperlinks', 'Type a name in the Name Box to name the selection, or use Insert ▸ Name ▸ Define (Ctrl+F3). Insert ▸ Comment (Shift+F2) attaches a note; cells with comments show a red triangle. Insert ▸ Hyperlink (Ctrl+K) links to web pages, e-mail addresses or places in the workbook.'],
    ['Formula auditing', 'Tools ▸ Formula Auditing traces precedents and dependents with arrows, evaluates a formula step by step and shows a Watch Window. Ctrl+` shows formulas instead of results. Error Checking walks through cells with errors.'],
    ['Printing and PDF', 'File ▸ Page Setup sets orientation, scaling (fit to N pages), paper size, margins, centering, headers and footers (&[Page], &[Pages], &[Date], &[File], &[Tab]), print area, rows to repeat and gridlines. View ▸ Page Break Preview shows the pages; Insert ▸ Page Break adds manual breaks. File ▸ Print and Print Preview produce a PDF you can print anywhere.'],
    ['Files and compatibility', 'Ledger reads and writes Office Open XML workbooks (.xlsx, macro-enabled .xlsm and templates), keeping formulas, styles, merged cells, comments, lists, conditional formats, validation, charts, pictures, defined names, page setup and macros. It also opens CSV and text files through the Text Import Wizard, opens password-protected workbooks and can save with a password (Save As ▸ Tools ▸ General Options). Binary .xls files must be converted to .xlsx first.'],
    ['Protection', 'Tools ▸ Protection ▸ Protect Sheet locks every cell whose Locked box (Format ▸ Cells ▸ Protection) is checked; choose what users may still do, and set an optional password. Protect Workbook locks the sheet structure.'],
    ['What-if analysis', 'Tools ▸ Goal Seek finds the input that makes a formula reach a target. Data ▸ Table builds one- and two-variable data tables. Tools ▸ Scenarios stores sets of input values to compare.'],
  ];
  TP.register({
    id: 'help', title: 'Ledger Help',
    render(b) {
      const q = h('input', { type: 'text', id: 'hp-q', value: P.helpQuery || '', placeholder: 'Search', style: 'flex:1;min-width:0' });
      const res = h('div');
      const draw = () => {
        res.textContent = '';
        const words = q.value.toLowerCase().split(/\s+/).filter(Boolean);
        const hit = (t) => words.every((w) => t.toLowerCase().includes(w));
        for (const [t, d] of HELP.filter(([t, d]) => !words.length || hit(t + ' ' + d))) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: t }), h('p', { text: d })));
        if (words.length) {
          const fns = L.fninfo ? L.fninfo.all().filter((f) => hit(f.name + ' ' + f.desc)).slice(0, 12) : [];
          if (fns.length) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: 'Functions' }), ...fns.map((f) => { const x = h('div', { class: 'help-fn' }, h('b', { text: L.fninfo.syntax(f.name) }), h('div', { text: f.desc })); x.addEventListener('click', () => L.dlg.insertFunction(f.name)); return x; })));
          const keys = KEYS.filter(([k, d]) => hit(k + ' ' + d));
          if (keys.length) res.appendChild(h('div', { class: 'help-topic' }, h('h4', { text: 'Keyboard shortcuts' }), h('table', { class: 'kbd-table' }, ...keys.map(([k, d]) => h('tr', null, h('td', { text: k }), h('td', { text: d }))))));
        }
        if (!res.childNodes.length) res.appendChild(h('div', { class: 'tp-note', text: 'No help topics match your search.' }));
      };
      q.addEventListener('input', draw);
      q.addEventListener('keydown', (e) => e.stopPropagation());
      b.append(h('div', { class: 'tp-row' }, q), res);
      draw();
      P.helpQuery = '';
    },
  });
  const KEYS = [
    ['Ctrl+N / O / S / P', 'New, Open, Save, Print'], ['F12', 'Save As'], ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'], ['F4', 'Repeat; in a formula, cycle $ references'],
    ['Ctrl+X / C / V', 'Cut, Copy, Paste'], ['Ctrl+Alt+V', 'Paste Special'], ['Ctrl+F / H', 'Find, Replace'], ['F5 / Ctrl+G', 'Go To'],
    ['F2', 'Edit the active cell'], ['Esc', 'Cancel an entry'], ['Alt+Enter', 'New line in a cell'], ['Ctrl+Enter', 'Fill the selection with the entry'], ['Ctrl+Shift+Enter', 'Enter an array formula'],
    ['Ctrl+D / Ctrl+R', 'Fill down / right'], ['Ctrl+; / Ctrl+Shift+:', 'Insert the date / time'], ["Ctrl+' / Ctrl+Shift+\"", 'Copy formula / value from the cell above'], ['Alt+=', 'AutoSum'], ['Shift+F3', 'Insert Function'], ['Ctrl+A (in a formula)', 'Function Arguments'], ['F9', 'Calculate now'],
    ['Arrow keys', 'Move one cell'], ['Ctrl+Arrow', 'Move to the edge of the data region'], ['Home / Ctrl+Home', 'Start of row / A1'], ['Ctrl+End', 'Last used cell'], ['End, Arrow', 'End mode jump'], ['Page Up / Down', 'Move one screen'], ['Alt+Page Up / Down', 'Move one screen left / right'], ['Ctrl+Page Up / Down', 'Previous / next sheet'],
    ['Shift+Arrow', 'Extend the selection'], ['Ctrl+Shift+Arrow', 'Extend to the edge of the data'], ['Ctrl+Space / Shift+Space', 'Select column / row'], ['Ctrl+A', 'Select the current region, then the whole sheet'], ['Ctrl+Shift+*', 'Select the current region'], ['Ctrl+/', 'Select the current array'],
    ['Ctrl+1', 'Format Cells'], ['Ctrl+B / I / U / 5', 'Bold, Italic, Underline, Strikethrough'], ['Ctrl+Shift+~ $ % ^ # @ !', 'General, Currency, Percent, Scientific, Date, Time, Number formats'], ['Ctrl+Shift+& / _', 'Outline border / remove borders'],
    ['Ctrl+9 / Ctrl+Shift+(', 'Hide / unhide rows'], ['Ctrl+0 / Ctrl+Shift+)', 'Hide / unhide columns'], ['Ctrl+- / Ctrl+Shift++', 'Delete / insert cells'], ['Shift+F11', 'Insert worksheet'], ['F11', 'Chart'], ['Shift+F2', 'Insert or edit comment'], ['Ctrl+K', 'Hyperlink'], ['Ctrl+F3 / F3', 'Define / paste name'], ['Ctrl+L', 'Create List'],
    ['Alt+Down', 'Pick from drop-down list'], ['Ctrl+`', 'Show formulas'], ['F7', 'Spelling'], ['Shift+F10', 'Shortcut menu'], ['Ctrl+F1', 'Task pane'], ['F1', 'Help'], ['Ctrl+Scroll', 'Zoom'],
  ];
  P.KEYS = KEYS;
  TP.register({
    id: 'keys', title: 'Keyboard Shortcuts',
    render(b) { b.appendChild(h('table', { class: 'kbd-table' }, ...KEYS.map(([k, d]) => h('tr', null, h('td', { text: k }), h('td', { text: d }))))); },
  });
  /* ---------- Office Clipboard ---------- */
  P.clips = [];
  P.addClip = function (clip) {
    P.clips.unshift({ clip, text: (clip.text || '').replace(/\t/g, '  ').slice(0, 200), time: Date.now() });
    if (P.clips.length > 24) P.clips.length = 24;
    L.bus.emit('clipboard');
  };
  TP.register({
    id: 'clipboard', title: 'Clipboard', live: true,
    render(b) {
      const items = P.clips;
      b.append(h('div', { class: 'tp-btns' }, ui.button('Paste All', () => { for (const it of items.slice().reverse()) { L.clip.set(it.clip); L.clip.pasteInternal({}); } }, { class: 'btn small' }), ui.button('Clear All', () => { items.length = 0; TP.refresh('clip'); }, { class: 'btn small' })),
        h('div', { class: 'tp-note', text: `${items.length} of 24 - Clipboard. Click an item to paste.` }));
      for (const it of items) {
        const el = h('button', { class: 'cb-item', type: 'button' }, h('span', { html: L.icons.get('sheetIcon') || '' }), h('span', { class: 'cb-t', text: it.text || '[cells]' }));
        el.addEventListener('click', () => { L.clip.set(it.clip); L.clip.pasteInternal({}); G().paint(); });
        b.appendChild(el);
      }
    },
  });

  /* ================= templates ================= */
  /** build a workbook from a compact description */
  function build(spec) {
    const wb = A().blankWorkbook();
    wb.sheets.splice(0);
    for (const s of spec.sheets) {
      const sh = wb.addSheet(s.name);
      const style = (base, d) => wb.styles.derive(base || 0, d);
      for (const [addr, val, st] of s.cells) {
        const p = F.parseCell(addr);
        const cell = sh.cell(p.r, p.c);
        if (typeof val === 'string' && val[0] === '=') { cell.f = F.toStore(val.slice(1)); cell.dirty = true; }
        else if (val !== null && val !== undefined) cell.v = val;
        if (st) cell.s = style(cell.s, st);
      }
      for (const [rg, st] of s.ranges || []) { const a = F.parseRange(rg); for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) { const cell = sh.cell(r, c); cell.s = style(cell.s, st); } }
      for (const [col, w] of Object.entries(s.cols || {})) sh.cols[F.colIndex(col)] = { w, custom: true };
      for (const m of s.merges || []) sh.merges.push(F.parseRange(m));
      if (s.freeze) sh.view.freeze = { r: s.freeze[0], c: s.freeze[1], top: { r: s.freeze[0], c: s.freeze[1] } };
      if (s.grid === false) sh.view.grid = false;
      if (s.print) Object.assign(sh.print, s.print);
      if (s.dv) for (const d of s.dv) sh.dv.push(Object.assign({ ranges: M.parseSqref(d.ref), allowBlank: true, showDrop: true, showInput: true, showError: true }, d));
      if (s.sel) sh.view.sel = { r: F.parseCell(s.sel).r, c: F.parseCell(s.sel).c, ranges: [F.parseRange(s.sel)] };
    }
    wb.active = 0;
    return wb;
  }
  const B = { font: { b: true } }, HDR = { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#336699') }, align: { h: 'center', v: 'center', wrap: true } };
  const TITLE = { font: { b: true, sz: 16, color: M.rgb('#1F3F7F') } };
  const MONEY = { nf: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' };
  const BOX = { border: { l: { style: 'thin', color: M.rgb('#808080') }, r: { style: 'thin', color: M.rgb('#808080') }, t: { style: 'thin', color: M.rgb('#808080') }, b: { style: 'thin', color: M.rgb('#808080') } } };
  const INPUT = { fill: { pattern: 'solid', fg: M.rgb('#FFFFCC') }, prot: { locked: false } };
  const TPL = {
    loan: {
      name: 'Loan Amortization',
      spec: () => {
        const cells = [
          ['A1', 'Loan Amortization Schedule', TITLE],
          ['A3', 'Loan amount', B], ['C3', 25000, Object.assign({}, MONEY, INPUT, BOX)],
          ['A4', 'Annual interest rate', B], ['C4', 0.065, Object.assign({ nf: '0.00%' }, INPUT, BOX)],
          ['A5', 'Loan period in years', B], ['C5', 5, Object.assign({}, INPUT, BOX)],
          ['A6', 'Payments per year', B], ['C6', 12, Object.assign({}, INPUT, BOX)],
          ['A7', 'Start date of loan', B], ['C7', 46113, Object.assign({ nf: 'm/d/yyyy' }, INPUT, BOX)],
          ['E3', 'Scheduled payment', B], ['G3', '=IF(C3*C4*C5*C6>0,-PMT(C4/C6,C5*C6,C3),"")', Object.assign({}, MONEY, BOX)],
          ['E4', 'Number of payments', B], ['G4', '=C5*C6', BOX],
          ['E5', 'Total interest', B], ['G5', '=SUM(F11:F370)', Object.assign({}, MONEY, BOX)],
          ['E6', 'Total cost of loan', B], ['G6', '=C3+G5', Object.assign({}, MONEY, BOX)],
          ['A10', 'Pmt No.', HDR], ['B10', 'Payment Date', HDR], ['C10', 'Beginning Balance', HDR], ['D10', 'Payment', HDR], ['E10', 'Principal', HDR], ['F10', 'Interest', HDR], ['G10', 'Ending Balance', HDR], ['H10', 'Cumulative Interest', HDR],
        ];
        for (let i = 0; i < 60; i++) {
          const r = 11 + i;
          cells.push(['A' + r, `=IF(ROW()-10<=$G$4,ROW()-10,"")`], ['B' + r, `=IF(A${r}="","",EDATE($C$7,(A${r}-1)*12/$C$6))`, { nf: 'm/d/yyyy' }],
            ['C' + r, i === 0 ? `=IF(A${r}="","",$C$3)` : `=IF(A${r}="","",G${r - 1})`, MONEY], ['D' + r, `=IF(A${r}="","",MIN($G$3,C${r}+F${r}))`, MONEY],
            ['E' + r, `=IF(A${r}="","",D${r}-F${r})`, MONEY], ['F' + r, `=IF(A${r}="","",C${r}*$C$4/$C$6)`, MONEY], ['G' + r, `=IF(A${r}="","",C${r}-E${r})`, MONEY], ['H' + r, i === 0 ? `=IF(A${r}="","",F${r})` : `=IF(A${r}="","",H${r - 1}+F${r})`, MONEY]);
        }
        return { sheets: [{ name: 'Loan', cells, cols: { A: 20, B: 13, C: 16, D: 13, E: 18, F: 13, G: 15, H: 14 }, freeze: [10, 0], print: { titleRows: null, fit: true, fitW: 1, fitH: null }, sel: 'C3' }] };
      },
    },
    expense: {
      name: 'Expense Statement',
      spec: () => {
        const cells = [
          ['A1', 'Expense Statement', TITLE], ['A3', 'Employee', B], ['B3', '', Object.assign({}, INPUT, BOX)], ['D3', 'Period from', B], ['E3', '', Object.assign({ nf: 'm/d/yyyy' }, INPUT, BOX)], ['F3', 'to', B], ['G3', '', Object.assign({ nf: 'm/d/yyyy' }, INPUT, BOX)],
          ['A5', 'Date', HDR], ['B5', 'Description', HDR], ['C5', 'Category', HDR], ['D5', 'Miles', HDR], ['E5', 'Mileage ($0.67/mi)', HDR], ['F5', 'Amount', HDR], ['G5', 'Total', HDR],
          ['A6', 46120, { nf: 'm/d/yyyy' }], ['B6', 'Client visit, Springfield'], ['C6', 'Travel'], ['D6', 84], ['F6', 18.5, MONEY],
          ['A7', 46121, { nf: 'm/d/yyyy' }], ['B7', 'Lunch with client'], ['C7', 'Meals'], ['F7', 46.2, MONEY],
          ['A8', 46124, { nf: 'm/d/yyyy' }], ['B8', 'Conference registration'], ['C8', 'Training'], ['F8', 395, MONEY],
          ['F20', 'Subtotal', B], ['G20', '=SUM(G6:G19)', Object.assign({ font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'double' } } }, MONEY)],
          ['F21', 'Less advances', B], ['G21', 0, Object.assign({}, MONEY, INPUT)], ['F22', 'Total due', B], ['G22', '=G20-G21', Object.assign({ font: { b: true } }, MONEY)],
          ['A24', 'Summary by category', B],
        ];
        for (let r = 6; r <= 19; r++) { cells.push(['E' + r, `=IF(D${r}="","",D${r}*0.67)`, MONEY], ['G' + r, `=IF(COUNT(E${r}:F${r})=0,"",SUM(E${r}:F${r}))`, MONEY]); if (r > 8) cells.push(['A' + r, null, { nf: 'm/d/yyyy' }], ['F' + r, null, MONEY]); }
        ['Travel', 'Meals', 'Lodging', 'Training', 'Supplies', 'Other'].forEach((c, i) => { cells.push(['A' + (25 + i), c], ['B' + (25 + i), `=SUMIF($C$6:$C$19,A${25 + i},$G$6:$G$19)`, MONEY]); });
        return { sheets: [{ name: 'Expenses', cells, cols: { A: 12, B: 30, C: 12, D: 8, E: 12, F: 12, G: 13 }, dv: [{ ref: 'C6:C19', type: 'list', f1: '"Travel,Meals,Lodging,Training,Supplies,Other"', prompt: 'Pick a category', promptTitle: 'Category' }], sel: 'B3' }] };
      },
    },
    invoice: {
      name: 'Sales Invoice',
      spec: () => {
        const cells = [
          ['A1', 'INVOICE', { font: { b: true, sz: 20, color: M.rgb('#336699') } }], ['E1', 'Invoice No.', B], ['F1', 1001, Object.assign({}, INPUT, BOX)], ['E2', 'Date', B], ['F2', '=TODAY()', Object.assign({ nf: 'mmmm d, yyyy' }, BOX)],
          ['A4', 'Bill to:', B], ['A5', 'Customer name', INPUT], ['A6', 'Street address', INPUT], ['A7', 'City, State ZIP', INPUT],
          ['A9', 'Qty', HDR], ['B9', 'Description', HDR], ['E9', 'Unit Price', HDR], ['F9', 'Line Total', HDR],
          ['A10', 4], ['B10', 'Widget, standard'], ['E10', 12.5, MONEY], ['A11', 1], ['B11', 'Installation service'], ['E11', 85, MONEY],
          ['E21', 'Subtotal', B], ['F21', '=SUM(F10:F20)', MONEY], ['E22', 'Sales tax rate', B], ['F22', 0.0725, Object.assign({ nf: '0.00%' }, INPUT)], ['E23', 'Sales tax', B], ['F23', '=ROUND(F21*F22,2)', MONEY],
          ['E24', 'TOTAL', { font: { b: true, sz: 12 } }], ['F24', '=F21+F23', Object.assign({ font: { b: true, sz: 12 }, border: { t: { style: 'thin' }, b: { style: 'double' } } }, MONEY)],
          ['A27', 'Make all checks payable to your company name. Thank you for your business!', { font: { i: true } }],
        ];
        for (let r = 10; r <= 20; r++) cells.push(['F' + r, `=IF(A${r}*E${r}=0,"",A${r}*E${r})`, MONEY]);
        return { sheets: [{ name: 'Invoice', cells, cols: { A: 8, B: 30, C: 8, D: 8, E: 14, F: 15 }, merges: ['B9:D9'], grid: false, sel: 'A5' }] };
      },
    },
    timecard: {
      name: 'Timecard',
      spec: () => {
        const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        const T = { nf: 'h:mm AM/PM' };
        const cells = [['A1', 'Weekly Timecard', TITLE], ['A3', 'Employee', B], ['B3', '', INPUT], ['A4', 'Week starting', B], ['B4', 46118, Object.assign({ nf: 'm/d/yyyy' }, INPUT)], ['D3', 'Hourly rate', B], ['E3', 22.5, Object.assign({}, MONEY, INPUT)],
          ['A6', 'Day', HDR], ['B6', 'Date', HDR], ['C6', 'Time in', HDR], ['D6', 'Lunch out', HDR], ['E6', 'Lunch in', HDR], ['F6', 'Time out', HDR], ['G6', 'Hours', HDR]];
        days.forEach((d, i) => {
          const r = 7 + i;
          cells.push(['A' + r, d], ['B' + r, `=$B$4+${i}`, { nf: 'm/d' }], ['C' + r, i < 5 ? 0.375 : null, T], ['D' + r, i < 5 ? 0.5 : null, T], ['E' + r, i < 5 ? 0.53125 : null, T], ['F' + r, i < 5 ? 0.7083333333 : null, T],
            ['G' + r, `=IF(COUNT(C${r}:F${r})<2,0,ROUND(((F${r}-C${r})-(E${r}-D${r}))*24,2))`, { nf: '0.00' }]);
        });
        cells.push(['F15', 'Total hours', B], ['G15', '=SUM(G7:G13)', { nf: '0.00', font: { b: true } }], ['F16', 'Regular (≤40)', B], ['G16', '=MIN(40,G15)', { nf: '0.00' }], ['F17', 'Overtime', B], ['G17', '=MAX(0,G15-40)', { nf: '0.00' }],
          ['F19', 'Total pay', B], ['G19', '=G16*E3+G17*E3*1.5', Object.assign({ font: { b: true } }, MONEY)]);
        return { sheets: [{ name: 'Timecard', cells, cols: { A: 12, B: 11, C: 11, D: 11, E: 11, F: 13, G: 11 }, sel: 'B3' }] };
      },
    },
    budget: {
      name: 'Personal Monthly Budget',
      spec: () => {
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const cols = 'BCDEFGHIJKLM';
        const N = { nf: '#,##0' };
        const cells = [['A1', 'Personal Monthly Budget', TITLE], ['A3', 'Income', HDR], ['N3', 'Year', HDR]];
        months.forEach((m, i) => cells.push([cols[i] + '3', m, HDR]));
        const income = [['Salary', 4200], ['Other income', 150]];
        const exp = [['Rent / mortgage', 1450], ['Utilities', 210], ['Groceries', 520], ['Transport', 260], ['Insurance', 180], ['Phone & internet', 95], ['Entertainment', 140], ['Savings', 400]];
        income.forEach(([n, v], k) => { const r = 4 + k; cells.push(['A' + r, n]); months.forEach((m, i) => cells.push([cols[i] + r, v, N])); cells.push(['N' + r, `=SUM(B${r}:M${r})`, N]); });
        cells.push(['A6', 'Total income', B]);
        for (const c of cols + 'N') cells.push([c + '6', `=SUM(${c}4:${c}5)`, Object.assign({ font: { b: true }, border: { t: { style: 'thin' } } }, N)]);
        cells.push(['A8', 'Expenses', HDR]);
        months.forEach((m, i) => cells.push([cols[i] + '8', m, HDR]));
        cells.push(['N8', 'Year', HDR]);
        exp.forEach(([n, v], k) => { const r = 9 + k; cells.push(['A' + r, n]); months.forEach((m, i) => cells.push([cols[i] + r, v, N])); cells.push(['N' + r, `=SUM(B${r}:M${r})`, N]); });
        const last = 8 + exp.length;
        cells.push(['A' + (last + 1), 'Total expenses', B]);
        for (const c of cols + 'N') cells.push([c + (last + 1), `=SUM(${c}9:${c}${last})`, Object.assign({ font: { b: true }, border: { t: { style: 'thin' } } }, N)]);
        cells.push(['A' + (last + 3), 'Net (income − expenses)', B]);
        for (const c of cols + 'N') cells.push([c + (last + 3), `=${c}6-${c}${last + 1}`, Object.assign({ font: { b: true }, nf: '#,##0;[Red]-#,##0' })]);
        return { sheets: [{ name: 'Budget', cells, cols: { A: 24 }, freeze: [3, 1], sel: 'B4' }] };
      },
    },
    tasks: {
      name: 'Task Tracker',
      spec: () => {
        const DAY = { nf: 'm/d/yyyy' };
        const cells = [
          ['A1', 'Task Tracker', TITLE],
          ['A2', 'Open', B], ['B2', '=COUNTIFS(A5:A104,"<>",E5:E104,"<>Done")', { font: { b: true } }],
          ['C2', 'Overdue', B], ['D2', '=COUNTIFS(A5:A104,"<>",E5:E104,"<>Done",D5:D104,"<"&TODAY())', { font: { b: true, color: M.rgb('#C00000') } }],
          ['A4', 'Task', HDR], ['B4', 'Owner', HDR], ['C4', 'Priority', HDR], ['D4', 'Due', HDR], ['E4', 'Status', HDR], ['F4', 'Days left', HDR], ['G4', 'Notes', HDR],
          ['A5', 'Book the venue'], ['B5', 'Sam'], ['C5', 'High'], ['D5', '=TODAY()+2', DAY], ['E5', 'In progress'],
          ['A6', 'Send invitations'], ['B6', 'Alex'], ['C6', 'Medium'], ['D6', '=TODAY()+7', DAY], ['E6', 'Not started'],
          ['A7', 'Order supplies'], ['B7', 'Sam'], ['C7', 'Low'], ['D7', '=TODAY()-1', DAY], ['E7', 'Waiting'], ['G7', 'Quote requested'],
          ['A8', 'Draft the budget'], ['B8', 'Alex'], ['C8', 'High'], ['D8', '=TODAY()-3', DAY], ['E8', 'Done'],
        ];
        for (let r = 5; r <= 104; r++) {
          cells.push(['F' + r, `=IF(OR(A${r}="",D${r}="",E${r}="Done"),"",D${r}-TODAY())`, { nf: '0;[Red]-0', align: { h: 'center' } }]);
          if (r > 8) cells.push(['D' + r, null, DAY]);
        }
        return { sheets: [{ name: 'Tasks', cells, cols: { A: 32, B: 12, C: 10, D: 11, E: 13, F: 10, G: 30 }, freeze: [4, 0],
          dv: [{ ref: 'C5:C104', type: 'list', f1: '"High,Medium,Low"' }, { ref: 'E5:E104', type: 'list', f1: '"Not started,In progress,Waiting,Done"', prompt: 'Pick a status', promptTitle: 'Status' }],
          sel: 'A9' }] };
      },
    },
    shopping: {
      name: 'Shopping List',
      spec: () => {
        const cells = [
          ['A1', 'Shopping List', TITLE],
          ['A2', 'Budget', B], ['B2', 80, Object.assign({}, MONEY, INPUT, BOX)], ['D2', 'List total', B], ['E2', '=SUM(F5:F44)', MONEY], ['D3', 'Still to buy', B], ['E3', '=SUMIF(A5:A44,"",F5:F44)', MONEY],
          ['F2', '=IF(B2="","",IF(E2>B2,"Over budget by "&TEXT(E2-B2,"$0.00"),TEXT(B2-E2,"$0.00")&" left"))', { font: { i: true } }],
          ['A4', '✓', HDR], ['B4', 'Item', HDR], ['C4', 'Category', HDR], ['D4', 'Qty', HDR], ['E4', 'Price each', HDR], ['F4', 'Total', HDR],
        ];
        const items = [['Apples', 'Produce', 6, 0.5], ['Milk', 'Dairy', 2, 1.29], ['Bread', 'Bakery', 1, 2.99], ['Eggs (dozen)', 'Dairy', 1, 3.49], ['Chicken breast', 'Meat & fish', 1, 7.5], ['Pasta', 'Pantry', 2, 1.19], ['Coffee', 'Pantry', 1, 8.99], ['Dish soap', 'Household', 1, 2.49]];
        items.forEach(([n, c, q, p], i) => { const r = 5 + i; cells.push(['B' + r, n], ['C' + r, c], ['D' + r, q], ['E' + r, p, MONEY]); });
        for (let r = 5; r <= 44; r++) {
          cells.push(['F' + r, `=IF(B${r}="","",IF(D${r}="",1,D${r})*E${r})`, MONEY], ['A' + r, null, { align: { h: 'center' }, font: { b: true, color: M.rgb('#1F8A4C') } }]);
          if (r >= 5 + items.length) cells.push(['E' + r, null, MONEY]);
        }
        return { sheets: [{ name: 'Shopping', cells, cols: { A: 5, B: 28, C: 14, D: 7, E: 12, F: 12 }, freeze: [4, 0],
          dv: [{ ref: 'A5:A44', type: 'list', f1: '"✓"', prompt: 'Tick what is in the basket', promptTitle: 'Bought' },
            { ref: 'C5:C44', type: 'list', f1: '"Produce,Dairy,Bakery,Meat & fish,Pantry,Frozen,Drinks,Household,Other"' }],
          sel: 'B13' }] };
      },
    },
    schedule: {
      name: 'Weekly Schedule',
      spec: () => {
        const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
        const cols = 'BCDEFGH';
        const SLOT = Object.assign({ align: { v: 'center', wrap: true } }, BOX);
        const cells = [['A1', 'Weekly Schedule', TITLE], ['A2', 'Week of', B], ['B2', '=TODAY()-WEEKDAY(TODAY(),3)', Object.assign({ nf: 'mmmm d, yyyy' }, INPUT)], ['A4', 'Time', HDR]];
        days.forEach((d, i) => cells.push([cols[i] + '4', d, HDR], [cols[i] + '5', `=$B$2+${i}`, { nf: 'mmm d', align: { h: 'center' }, font: { color: M.rgb('#595959') } }]));
        for (let k = 0; k <= 15; k++) {
          const r = 6 + k;
          cells.push(['A' + r, (7 + k) / 24, { nf: 'h:mm AM/PM', font: { b: true }, align: { v: 'center' } }]);
          for (const c of cols) cells.push([c + r, null, SLOT]);
        }
        const plan = [['B8', 'Team stand-up'], ['D8', 'Team stand-up'], ['F8', 'Team stand-up'], ['C11', 'Lunch with Sam'], ['E17', 'Swimming'], ['G10', 'Groceries'], ['H11', 'Family brunch'], ['B19', 'Piano lesson']];
        for (const [a, v] of plan) cells.push([a, v, Object.assign({ fill: { pattern: 'solid', fg: M.rgb('#DDEBF7') } }, SLOT)]);
        cells.push(['A23', 'Notes', B], ['B23', '', INPUT]);
        return { sheets: [{ name: 'Schedule', cells, cols: { A: 10, B: 15, C: 15, D: 15, E: 15, F: 15, G: 15, H: 15 }, freeze: [5, 1], print: { orientation: 'landscape' }, sel: 'C6' }] };
      },
    },
  };
  P.TEMPLATES = TPL;
  P.buildTemplate = (k) => build(TPL[k].spec());
  P.newFromTemplate = function (k) {
    const t = TPL[k];
    if (!t) return;
    const wb = build(t.spec());
    if (L.cf && L.cf.reset) L.cf.reset();
    A().prepare(wb);
    A().addBook(wb, t.name.replace(/\s+/g, '') + '1', { untitled: true });
  };
})(typeof window !== 'undefined' ? window : globalThis);
