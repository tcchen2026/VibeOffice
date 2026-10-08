/* Ledger — dialogs, part 2 (Data and Format): Sort, Validation, Subtotals, Text to Columns, Advanced Filter, Form,
 * Create List, Outline settings, Conditional Formatting, AutoFormat, Style, Goal Seek, Scenarios, Data Table,
 * Consolidate, PivotTable report. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, O = L.ops, F = L.formula, NF = L.numfmt, LY = L.layout, C = L.calc;
  const { h } = L;
  const ui = L.ui;
  const D = (L.dlg = L.dlg || {});
  const G = () => L.grid;
  const A = () => L.app;
  const sh = () => G().sheet();
  const wb = () => G().wb;
  const tryRun = (fn) => { try { fn(); } catch (e) { A().error(e); } G().paint(); G().syncObjects(); ui.refresh(); };
  const num = (v, d) => { const x = parseFloat(String(v).replace(/,/g, '')); return isNaN(x) ? d : x; };
  const val = (s, r, c) => C.cellValue(s, r, c);
  const textOf = (s, r, c) => L.csv.cellText(s, r, c);
  /** a reference box: type a range or pick one on the sheet */
  function refBox(value, width) {
    const inp = h('input', { type: 'text', value: value || '', style: `width:${width || 200}px`, class: 'refbox' });
    const pick = h('button', { type: 'button', class: 'ref-pick', 'aria-label': 'Select range', 'data-tip': 'Use the current selection', html: '<svg width="14" height="12" viewBox="0 0 14 12"><rect x=".5" y=".5" width="13" height="11" fill="#fff" stroke="#555"/><path d="M1 4h12M5 1v10" stroke="#999"/><rect x="5" y="4" width="8" height="7" fill="#c6d6f2"/></svg>' });
    pick.addEventListener('click', () => { const s = sh(); inp.value = F.quoteSheet(s.name) + '!' + F.absRangeName(G().range()); inp.dispatchEvent(new Event('change')); });
    const wrap = h('span', { class: 'refwrap' }, inp, pick);
    wrap.input = inp;
    return wrap;
  }
  D.refBox = refBox;
  /** "Sheet1!$A$1:$B$5" → {sheet, r1, c1, r2, c2} or null */
  function parseRef(text, single) {
    if (!text) return null;
    try {
      const ast = F.parse(String(text).trim().replace(/^=/, ''));
      const v = C.ev(ast, C.ctx(wb(), sh(), 0, 0));
      const a = v instanceof C.RefList ? v.list[0] : v;
      if (!(a instanceof C.Ref)) return null;
      if (single && (a.r1 !== a.r2 || a.c1 !== a.c2)) return null;
      return { sheet: a.sheet, r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 };
    } catch (e) { return null; }
  }
  D.parseRef = parseRef;
  const listRange = () => {
    const s = sh(), sel = G().sel(), rg = G().range();
    const t = wb().tableAt(s, sel.r, sel.c);
    if (t) return { r1: t.ref.r1, c1: t.ref.c1, r2: t.ref.r2 - (t.totals ? 1 : 0), c2: t.ref.c2, table: t };
    const single = G().ranges().length === 1 && rg.r1 === rg.r2 && rg.c1 === rg.c2;
    return single ? O.currentRegion(s, sel.r, sel.c) : O.clip(s, rg);
  };
  const colLabel = (s, rg, header, c) => { if (header) { const v = s.val(rg.r1, c); if (v != null && v !== '') return String(v); } return 'Column ' + F.colName(c); };

  /* ================================================================ Sort */
  D.sort = function () {
    const s = sh(), sel = G().sel();
    const rg = listRange();
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && s.get(rg.r1, rg.c1) == null) { ui.msg('Select a cell in the list before sorting.', { icon: 'warn' }); return; }
    let header = rg.table ? true : O.guessHeader(s, rg);
    const st = { keys: [{ c: Math.max(rg.c1, Math.min(rg.c2, sel.c)), desc: false }, { c: null, desc: false }, { c: null, desc: false }], caseSensitive: false, byCols: false, list: null };
    G().selectRange(rg, { r: sel.r, c: sel.c }, { noScroll: true });
    const body = h('div', { class: 'col' });
    const draw = () => {
      body.textContent = '';
      const opts = st.byCols ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => [rg.r1 + i, 'Row ' + (rg.r1 + i + 1)]) : Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => [rg.c1 + i, colLabel(s, rg, header, rg.c1 + i)]);
      st.keys.forEach((k, i) => {
        const selEl = ui.select((i ? [['', '(none)']] : []).concat(opts), k.c == null ? '' : k.c, (v) => { k.c = v === '' ? null : +v; });
        body.appendChild(ui.group(i === 0 ? 'Sort by' : 'Then by', h('div', { class: 'row', style: 'gap:12px;align-items:center' }, selEl, h('div', { class: 'col' }, ui.radio('sk' + i, '&Ascending', !k.desc, () => { k.desc = false; }), ui.radio('sk' + i, '&Descending', k.desc, () => { k.desc = true; })))));
      });
      body.appendChild(ui.group('My data range has', h('div', { class: 'row', style: 'gap:16px' }, ui.radio('sk-h', 'Header &row', header, () => { header = true; draw(); }), ui.radio('sk-h', 'No header ro&w', !header, () => { header = false; draw(); }))));
    };
    draw();
    const options = () => {
      const lists = [['', 'Normal']].concat(O.customLists.map((l, i) => [i, l.slice(0, 4).join(', ')]));
      ui.dialog({ title: 'Sort Options', width: 320, body: h('div', { class: 'col' }, ui.field('&First key sort order', ui.select(lists, st.list == null ? '' : st.list, (v) => { st.list = v === '' ? null : +v; })), ui.check('&Case sensitive', st.caseSensitive, (v) => { st.caseSensitive = v; }),
        ui.group('Orientation', ui.radio('so-o', 'Sort &top to bottom', !st.byCols, () => { st.byCols = false; st.keys.forEach((k) => { k.c = null; }); st.keys[0].c = sel.c; }), ui.radio('so-o', 'Sort &left to right', st.byCols, () => { st.byCols = true; st.keys.forEach((k) => { k.c = null; }); st.keys[0].c = sel.r; }))),
      buttons: [{ label: 'OK', primary: true, onClick: () => draw() }, { label: 'Cancel' }] });
    };
    ui.dialog({ title: 'Sort', width: 360, body, buttons: [{ label: '&Options...', onClick: () => { options(); return false; } }, { label: 'OK', primary: true, onClick: () => tryRun(() => {
      const keys = st.keys.filter((k) => k.c != null).map((k, i) => ({ index: st.byCols ? k.c - rg.r1 : k.c - rg.c1, desc: k.desc, list: i === 0 && st.list != null ? O.customLists[st.list] : undefined }));
      if (!keys.length) return;
      O.sort(s, { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, keys, { header, caseSensitive: st.caseSensitive, byCols: st.byCols });
    }) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Validation */
  D.validation = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const cur = L.cf.dvAt(s, sel.r, sel.c);
    const ranges = G().ranges().map((r) => O.clip(s, r)).map((r, i) => (G().isWholeCols(G().ranges()[i]) || G().isWholeRows(G().ranges()[i]) ? G().ranges()[i] : r));
    if (cur && ranges.some((rg) => { for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 300); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 50); c++) if (L.cf.dvAt(s, r, c) !== cur) return true; return false; })) {
      /* mixed: Excel asks before erasing */
    }
    const st = Object.assign({ type: 'none', op: 'between', allowBlank: true, showDrop: true, showInput: true, showError: true, errorStyle: 'stop', f1: '', f2: '', prompt: '', promptTitle: '', error: '', errorTitle: '' }, cur ? JSON.parse(JSON.stringify(cur)) : {});
    delete st.ranges;
    const shows = (st.f1 || '').replace(/^"(.*)"$/, '$1');
    const allow = ui.select([['none', 'Any value'], ['whole', 'Whole number'], ['decimal', 'Decimal'], ['list', 'List'], ['date', 'Date'], ['time', 'Time'], ['textLength', 'Text length'], ['custom', 'Custom']], st.type, (v) => { st.type = v; draw(); });
    const dataSel = ui.select([['between', 'between'], ['notBetween', 'not between'], ['equal', 'equal to'], ['notEqual', 'not equal to'], ['greaterThan', 'greater than'], ['lessThan', 'less than'], ['greaterThanOrEqual', 'greater than or equal to'], ['lessThanOrEqual', 'less than or equal to']], st.op, (v) => { st.op = v; draw(); });
    const v1 = h('input', { type: 'text', style: 'width:100%' }), v2 = h('input', { type: 'text', style: 'width:100%' });
    const toInput = (f) => { if (f == null || f === '') return ''; if (/^"/.test(f)) return f.slice(1, -1).replace(/""/g, '"'); if (/^-?[\d.]+(E[+-]?\d+)?$/i.test(f)) { const n = +f; if (st.type === 'date') return NF.text('m/d/yyyy', n); if (st.type === 'time') return NF.text('h:mm:ss AM/PM', n); return String(n); } return '=' + F.display(f); };
    v1.value = st.type === 'list' && /^"/.test(st.f1 || '') ? shows : toInput(st.f1);
    v2.value = toInput(st.f2);
    const blank = ui.check('Ignore &blank', st.allowBlank, (v) => { st.allowBlank = v; });
    const drop = ui.check('&In-cell dropdown', st.showDrop !== false, (v) => { st.showDrop = v; });
    const fields = h('div', { class: 'col' });
    function draw() {
      fields.textContent = '';
      const t = st.type;
      const ranged = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(t);
      if (ranged) fields.appendChild(ui.field('&Data:', dataSel));
      const two = ranged && (st.op === 'between' || st.op === 'notBetween');
      const lab1 = t === 'list' ? '&Source:' : t === 'custom' ? 'F&ormula:' : two ? (t === 'date' ? '&Start date:' : t === 'time' ? '&Start time:' : '&Minimum:') : t === 'date' ? '&Date:' : t === 'time' ? '&Time:' : t === 'textLength' ? '&Length:' : '&Value:';
      if (t !== 'none') fields.appendChild(ui.field(lab1, v1));
      if (two) fields.appendChild(ui.field(t === 'date' ? '&End date:' : t === 'time' ? '&End time:' : 'Ma&ximum:', v2));
      blank.hidden = t === 'none';
      drop.hidden = t !== 'list';
    }
    draw();
    const settings = h('div', { class: 'col' }, ui.group('Validation criteria', h('div', { class: 'row', style: 'gap:10px;align-items:flex-start' }, h('div', { class: 'col', style: 'flex:1' }, ui.field('&Allow:', allow), fields), h('div', { class: 'col' }, blank, drop))),
      ui.check('Apply these changes to all other cells with the same settings', false, (v) => { st.applyAll = v; }));
    const pt = h('input', { type: 'text', value: st.promptTitle || '', style: 'width:100%' }), pm = h('textarea', { rows: '4', style: 'width:100%' });
    pm.value = st.prompt || '';
    const input = h('div', { class: 'col' }, ui.check('&Show input message when cell is selected', st.showInput !== false, (v) => { st.showInput = v; }), h('div', { text: 'When cell is selected, show this input message:' }), ui.field('&Title:', pt), ui.field('&Input message:', pm));
    const et = h('input', { type: 'text', value: st.errorTitle || '', style: 'width:100%' }), em = h('textarea', { rows: '4', style: 'width:100%' });
    em.value = st.error || '';
    const errSel = ui.select([['stop', 'Stop'], ['warning', 'Warning'], ['information', 'Information']], st.errorStyle || 'stop', (v) => { st.errorStyle = v; });
    const error = h('div', { class: 'col' }, ui.check('&Show error alert after invalid data is entered', st.showError !== false, (v) => { st.showError = v; }), h('div', { text: 'When user enters invalid data, show this error alert:' }), h('div', { class: 'row', style: 'gap:10px;align-items:flex-start' }, ui.field('St&yle:', errSel), h('div', { class: 'col', style: 'flex:1' }, ui.field('&Title:', et), ui.field('&Error message:', em))));
    const ime = h('div', { class: 'col' }, ui.field('&Mode:', ui.select([['', 'No Control'], ['on', 'On'], ['off', 'Off (English mode)'], ['disabled', 'Disable']], st.ime || '', (v) => { st.ime = v || undefined; })));
    const toFormula = (text) => {
      const t = String(text).trim();
      if (t === '') return null;
      if (t[0] === '=') { F.parse(t.slice(1)); return F.toStore(t.slice(1)); }
      if (st.type === 'list') return '"' + t.replace(/"/g, '""') + '"';
      const p = NF.parseInput(t, { date1904: w.date1904 });
      if (p && typeof p.v === 'number') return String(p.v);
      throw new Error('The value you entered is not a valid number, date or time.');
    };
    ui.dialog({ title: 'Data Validation', width: 440, body: ui.tabs([{ label: 'Settings', body: settings }, { label: 'Input Message', body: input }, { label: 'Error Alert', body: error }, { label: 'IME Mode', body: ime }]),
      buttons: [{ label: '&Clear All', onClick: () => { tryRun(() => O.tx(w, 'Validation', () => { for (const rg of ranges) O.removeRangeRules(s, rg, 'dv'); })); } },
        { label: 'OK', primary: true, onClick: () => {
          let f1, f2;
          try {
            f1 = st.type === 'none' ? undefined : toFormula(v1.value);
            f2 = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(st.type) && (st.op === 'between' || st.op === 'notBetween') ? toFormula(v2.value) : undefined;
            if (st.type !== 'none' && f1 == null) throw new Error('You must enter a value in the ' + (st.type === 'list' ? 'Source' : 'Minimum') + ' box.');
            if (f2 === null) throw new Error('You must enter a value in the Maximum box.');
          } catch (e) { ui.msg(e.message, { icon: 'warn' }); return false; }
          const rule = { type: st.type, op: ['whole', 'decimal', 'date', 'time', 'textLength'].includes(st.type) ? st.op : undefined, allowBlank: st.allowBlank, showDrop: st.showDrop !== false, showInput: st.showInput !== false, showError: st.showError !== false, errorStyle: st.errorStyle, f1: f1 == null ? undefined : f1, f2: f2 == null ? undefined : f2, promptTitle: pt.value || undefined, prompt: pm.value || undefined, errorTitle: et.value || undefined, error: em.value || undefined, ime: st.ime };
          if (st.x14) { rule.x14 = true; rule.extKeep = st.extKeep; }
          tryRun(() => O.tx(w, 'Validation', () => {
            let target = ranges;
            if (st.applyAll && cur) target = target.concat(cur.ranges);
            for (const rg of target) O.removeRangeRules(s, rg, 'dv');
            if (st.type !== 'none' || rule.prompt || rule.promptTitle) O.setDV(s, s.dv.concat([Object.assign({ ranges: target.map((x) => Object.assign({}, x)) }, rule)]));
          }));
        } }, { label: 'Cancel' }] });
  };

  /* ================================================================ Subtotals */
  const SUBFN = [[9, 'Sum'], [3, 'Count'], [1, 'Average'], [4, 'Max'], [5, 'Min'], [6, 'Product'], [2, 'Count Nums'], [7, 'StdDev'], [8, 'StdDevp'], [10, 'Var'], [11, 'Varp']];
  D.subtotals = function () {
    const s = sh(), w = wb();
    const rg = listRange();
    if (rg.r2 <= rg.r1) { ui.msg('Ledger cannot determine which row in your list or selection contains column labels, which are required for this command.', { icon: 'warn' }); return; }
    const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
    const st = { at: rg.c1, fn: 9, add: new Set([rg.c2]), replace: true, pageBreak: false, below: true };
    const addBox = h('div', { class: 'lbox checks' }, ...cols.map((c) => ui.check(colLabel(s, rg, true, c), st.add.has(c), (v) => { if (v) st.add.add(c); else st.add.delete(c); })));
    const removeAll = () => tryRun(() => O.structural(s, 'Remove Subtotals', () => {
      for (let r = rg.r2; r > rg.r1; r--) { const row = s.rows[r]; if (row && row.cells.some((cl) => cl && cl.f != null && /^SUBTOTAL\(/i.test(cl.f))) O.insertLines(s, 'r', r, -1); }
      s.rows.forEach((row) => { if (row) { delete row.level; delete row.hidden; } });
      s.outline.levelRow = 0;
      LY.invalidate(s);
    }));
    ui.dialog({ title: 'Subtotal', width: 340, body: h('div', { class: 'col' },
      ui.field('&At each change in:', ui.select(cols.map((c) => [c, colLabel(s, rg, true, c)]), st.at, (v) => { st.at = +v; })),
      ui.field('&Use function:', ui.select(SUBFN, st.fn, (v) => { st.fn = +v; })),
      h('label', { text: 'Add subtotal to:' }), addBox,
      ui.check('Replace &current subtotals', true, (v) => { st.replace = v; }), ui.check('&Page break between groups', false, (v) => { st.pageBreak = v; }), ui.check('&Summary below data', true, (v) => { st.below = v; })),
    buttons: [{ label: '&Remove All', onClick: removeAll }, { label: 'OK', primary: true, onClick: () => tryRun(() => {
      if (!st.add.size) return;
      const fnName = SUBFN.find((x) => x[0] === st.fn)[1];
      O.structural(s, 'Subtotals', () => {
        let r2 = rg.r2;
        if (st.replace) { for (let r = r2; r > rg.r1; r--) { const row = s.rows[r]; if (row && row.cells.some((cl) => cl && cl.f != null && /^SUBTOTAL\(/i.test(cl.f))) { O.insertLines(s, 'r', r, -1); r2--; } } }
        /* groups: runs of equal values in the "at each change" column */
        const groups = [];
        let start = rg.r1 + 1;
        const key = (r) => textOf(s, r, st.at).toLowerCase();
        for (let r = rg.r1 + 2; r <= r2 + 1; r++) if (r > r2 || key(r) !== key(start)) { groups.push([start, r - 1]); start = r; }
        const bold = (r, c) => { const cl = s.get(r, c) || { v: null }; const n = Object.assign({}, cl); n.s = w.styles.derive(cl.s || 0, { font: { b: true } }); s.put(r, c, n); };
        /* insert from the bottom so earlier rows keep their numbers */
        const totalRows = [];
        for (let g = groups.length - 1; g >= 0; g--) {
          const [a, b] = groups[g];
          const at = st.below ? b + 1 : a;
          O.insertLines(s, 'r', at, 1);
          const ga = st.below ? a : a + 1, gb = st.below ? b : b + 1;
          s.put(at, st.at, { v: textOf(s, ga, st.at) + (st.fn === 9 ? ' Total' : ' ' + fnName) });
          bold(at, st.at);
          for (const c of st.add) { s.put(at, c, { f: `SUBTOTAL(${st.fn},${F.rangeName({ r1: ga, c1: c, r2: gb, c2: c })})`, v: null, dirty: true }); bold(at, c); }
          totalRows.push(at);
          for (let r = ga; r <= gb; r++) { const o = s.rowObj(r); o.level = 2; }
          s.rowObj(at).level = 1;
          if (st.pageBreak && g > 0) s.print.rowBreaks = (s.print.rowBreaks || []).concat([st.below ? at + 1 : at]).sort((x, y) => x - y);
        }
        /* grand total */
        const dataEnd = r2 + groups.length;
        const gt = st.below ? dataEnd + 1 : rg.r1 + 1;
        O.insertLines(s, 'r', gt, 1);
        const lastData = st.below ? dataEnd : dataEnd + 1;
        s.put(gt, st.at, { v: st.fn === 9 ? 'Grand Total' : 'Grand ' + fnName });
        bold(gt, st.at);
        for (const c of st.add) { s.put(gt, c, { f: `SUBTOTAL(${st.fn},${F.rangeName({ r1: st.below ? rg.r1 + 1 : rg.r1 + 2, c1: c, r2: lastData, c2: c })})`, v: null, dirty: true }); bold(gt, c); }
        s.outline.levelRow = 3;
        s.outline.below = st.below;
        if (s.maxR < gt) s.maxR = gt;
        LY.invalidate(s);
      });
    }) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Text to Columns */
  D.textToColumns = function () {
    const s = sh(), w = wb();
    const rg0 = O.clip(s, G().range());
    if (rg0.c1 !== rg0.c2) { ui.msg('Text to Columns can convert only one column at a time.', { icon: 'warn' }); return; }
    const rows = [];
    for (let r = rg0.r1; r <= Math.min(rg0.r2, s.maxR); r++) { const v = s.val(r, rg0.c1); rows.push(v == null ? '' : typeof v === 'string' ? v : textOf(s, r, rg0.c1)); }
    const text = rows.join('\n');
    const st = { fixed: !/[\t,;]/.test(text) && / {2,}/.test(text), delims: { tab: /\t/.test(text), semi: false, comma: /,/.test(text) && !/\t/.test(text), space: false, other: '' }, merge: false, quote: '"', breaks: null, formats: [], dest: F.cellName(rg0.r1, rg0.c1, true, true) };
    if (!st.delims.tab && !st.delims.comma && !st.fixed) st.delims.space = true;
    const delimStr = () => (st.delims.tab ? '\t' : '') + (st.delims.semi ? ';' : '') + (st.delims.comma ? ',' : '') + (st.delims.space ? ' ' : '') + (st.delims.other ? st.delims.other[0] : '');
    const parse = () => (st.fixed ? L.csv.parseFixed(text, st.breaks || (st.breaks = L.csv.guessBreaks(text))) : rows.map((t) => (L.csv.parse(t, { delims: delimStr(), quote: st.quote, merge: st.merge })[0] || [''])));
    const preview = h('div', { class: 'tiw-preview' });
    let selCol = 0;
    const drawPrev = () => {
      preview.textContent = '';
      const data = parse().slice(0, 12);
      const n = Math.max(1, ...data.map((r) => r.length));
      const tbl = h('table', { class: 'tiw-table' });
      const hr = h('tr');
      for (let c = 0; c < n; c++) { const th = h('th', { class: c === selCol ? 'on' : '', text: { general: 'General', text: 'Text', skip: 'Skip' }[st.formats[c] || 'general'] || st.formats[c] }); th.addEventListener('click', () => { selCol = c; drawPrev(); fmtSync(); }); hr.appendChild(th); }
      tbl.appendChild(hr);
      for (const r of data) { const tr = h('tr'); for (let c = 0; c < n; c++) tr.appendChild(h('td', { class: c === selCol ? 'on' : '', text: r[c] == null ? '' : r[c] })); tbl.appendChild(tr); }
      preview.appendChild(tbl);
    };
    const d = st.delims;
    const delimBox = h('div', { class: 'tiw-delims' }, ui.check('&Tab', d.tab, (v) => { d.tab = v; drawPrev(); }), ui.check('Se&micolon', d.semi, (v) => { d.semi = v; drawPrev(); }), ui.check('&Comma', d.comma, (v) => { d.comma = v; drawPrev(); }), ui.check('&Space', d.space, (v) => { d.space = v; drawPrev(); }), ui.field('&Other:', (() => { const i = h('input', { type: 'text', maxlength: '1', style: 'width:24px' }); i.addEventListener('input', () => { d.other = i.value; drawPrev(); }); return i; })()));
    const breaksIn = h('input', { type: 'text', value: '', style: 'width:200px', placeholder: 'e.g. 10, 25, 40' });
    breaksIn.addEventListener('change', () => { st.breaks = breaksIn.value.split(/[,\s]+/).map(Number).filter((x) => x > 0).sort((a, b) => a - b); drawPrev(); });
    const fmtRadios = [['general', '&General'], ['text', '&Text'], ['MDY', '&Date (MDY)'], ['DMY', 'Date (DM&Y)'], ['YMD', 'Date (&YMD)'], ['skip', 'Do not import column (s&kip)']].map(([k, l]) => ui.radio('ttc-f', l, k === 'general', () => { st.formats[selCol] = k; drawPrev(); }));
    const fmtSync = () => { const f = st.formats[selCol] || 'general'; ['general', 'text', 'MDY', 'DMY', 'YMD', 'skip'].forEach((k, i) => { fmtRadios[i].input.checked = f === k; }); };
    const dest = refBox(st.dest, 160);
    const typeBox = ui.group('Original data type', ui.radio('ttc-t', '&Delimited - Characters such as commas or tabs separate each field.', !st.fixed, () => { st.fixed = false; delimBox.hidden = false; brk.hidden = true; drawPrev(); }), ui.radio('ttc-t', 'Fixed &width - Fields are aligned in columns with spaces between each field.', st.fixed, () => { st.fixed = true; delimBox.hidden = true; brk.hidden = false; if (!st.breaks) st.breaks = L.csv.guessBreaks(text); breaksIn.value = st.breaks.join(', '); drawPrev(); }));
    const brk = ui.field('Column &breaks at characters:', breaksIn);
    delimBox.hidden = st.fixed; brk.hidden = !st.fixed;
    if (st.fixed) { st.breaks = L.csv.guessBreaks(text); breaksIn.value = st.breaks.join(', '); }
    drawPrev();
    ui.dialog({ title: 'Convert Text to Columns Wizard', width: 600, body: h('div', { class: 'col' }, typeBox, delimBox, brk, h('div', { class: 'row', style: 'gap:12px' }, ui.check('Treat consecutive delimiters as &one', false, (v) => { st.merge = v; drawPrev(); }), ui.field('Text &qualifier:', ui.select([['"', '"'], ["'", "'"], ['', '{none}']], '"', (v) => { st.quote = v; drawPrev(); }))), ui.group('Column data format', h('div', { class: 'ttc-fmts' }, ...fmtRadios)), ui.field('D&estination:', dest), h('div', { class: 'tiw-cap', text: 'Data preview:' }), preview),
      buttons: [{ label: '&Finish', primary: true, onClick: () => tryRun(() => {
        const to = parseRef(dest.input.value, false) || { sheet: s, r1: rg0.r1, c1: rg0.c1 };
        const data = parse();
        const dsh = to.sheet || s;
        const n = Math.max(...data.map((r) => r.length));
        /* overwriting other data needs a confirmation, as in Excel */
        let busy = false;
        for (let i = 0; i < data.length && !busy; i++) for (let c = 1; c < n; c++) { const cl = dsh.get(to.r1 + i, to.c1 + c); if (cl && cl.v != null && !(dsh === s && to.c1 + c === rg0.c1)) busy = true; }
        const go = () => O.tx(w, 'Text to Columns', () => {
          data.forEach((row, i) => {
            let c = 0;
            row.forEach((t, j) => {
              const f = st.formats[j] || 'general';
              if (f === 'skip') return;
              const x = L.csv.convert(t, f, { date1904: w.date1904 });
              const old = dsh.get(to.r1 + i, to.c1 + c);
              const cl = Object.assign({}, old || {});
              delete cl.f; delete cl.rt;
              cl.v = x.v == null ? null : x.v;
              if (x.f != null) { cl.f = x.f; cl.dirty = true; }
              if (x.nf) cl.s = w.styles.derive(cl.s || 0, { nf: x.nf });
              O.put(dsh, to.r1 + i, to.c1 + c, cl.v == null && cl.f == null && !cl.s ? null : cl);
              c++;
            });
          });
        });
        if (busy) ui.msg('Do you want to replace the contents of the destination cells?', { icon: 'question', buttons: ['OK', 'Cancel'] }).then((r) => { if (r === 0) tryRun(go); });
        else go();
      }) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Advanced Filter */
  D.advancedFilter = function () {
    const s = sh();
    const rg = listRange();
    let copy = false, unique = false;
    const list = refBox(F.quoteSheet(s.name) + '!' + F.absRangeName(rg), 220);
    const crit = refBox('', 220), to = refBox('', 220);
    to.hidden = true;
    ui.dialog({ title: 'Advanced Filter', width: 360, body: h('div', { class: 'col' }, ui.group('Action', ui.radio('af-a', '&Filter the list, in-place', true, () => { copy = false; to.input.disabled = true; }), ui.radio('af-a', 'C&opy to another location', false, () => { copy = true; to.input.disabled = false; })),
      ui.field('&List range:', list), ui.field('&Criteria range:', crit), ui.field('Copy &to:', (() => { to.hidden = false; to.input.disabled = true; return to; })()), ui.check('Unique &records only', false, (v) => { unique = v; })),
    buttons: [{ label: 'OK', primary: true, onClick: () => {
      const lr = parseRef(list.input.value), cr = crit.input.value.trim() ? parseRef(crit.input.value) : null;
      if (!lr) { ui.msg('The list range is not valid.', { icon: 'warn' }); return false; }
      if (crit.input.value.trim() && !cr) { ui.msg('The criteria range is not valid.', { icon: 'warn' }); return false; }
      const dest = copy ? parseRef(to.input.value) : null;
      if (copy && !dest) { ui.msg('The extract range is not valid.', { icon: 'warn' }); return false; }
      tryRun(() => {
        const n = L.filter.advanced(lr.sheet, { r1: lr.r1, c1: lr.c1, r2: Math.min(lr.r2, lr.sheet.maxR), c2: lr.c2 }, cr ? cr.sheet : null, cr ? { r1: cr.r1, c1: cr.c1, r2: cr.r2, c2: cr.c2 } : null, { unique, copyTo: dest ? { sheet: dest.sheet, r: dest.r1, c: dest.c1 } : null });
        A().status(`${n} of ${lr.r2 - lr.r1} records found`);
      });
    } }, { label: 'Cancel' }] });
  };

  /* ================================================================ Data Form */
  D.dataForm = function () {
    const s = sh(), w = wb();
    let rg = listRange();
    if (rg.r2 <= rg.r1 && !(s.get(rg.r1, rg.c1))) { ui.msg('No list was found. Select a cell within your list.', { icon: 'warn' }); return; }
    const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
    let rec = 1, criteria = null;
    const inputs = cols.map((c) => { const i = h('input', { type: 'text', style: 'width:100%' }); i.dataset.c = c; return i; });
    const count = h('div', { class: 'df-count' });
    const fieldsBox = h('div', { class: 'df-fields' }, ...cols.map((c, k) => ui.field(colLabel(s, rg, true, c) + ':', inputs[k])));
    const total = () => rg.r2 - rg.r1;
    const load = () => {
      if (criteria) return;
      const r = rg.r1 + rec;
      cols.forEach((c, k) => { const cl = s.get(r, c); inputs[k].value = rec > total() ? '' : cl && cl.f != null ? textOf(s, r, c) : cl && cl.v != null ? L.editor.cellText(s, r, c) : ''; inputs[k].readOnly = !!(cl && cl.f != null) && rec <= total(); });
      count.textContent = rec > total() ? 'New Record' : `${rec} of ${total()}`;
    };
    const save = () => {
      if (criteria || rec > total() + 1) return;
      const r = rg.r1 + rec;
      const changes = inputs.filter((i, k) => !i.readOnly && i.value !== (s.get(r, cols[k]) && s.get(r, cols[k]).f == null ? L.editor.cellText(s, r, cols[k]) : textOf(s, r, cols[k])));
      if (!changes.length) return;
      if (rec > total() && !inputs.some((i) => i.value)) return;
      tryRun(() => O.tx(w, 'Data Form', () => { for (const i of changes) O.enter(s, r, +i.dataset.c, i.value); }));
      if (rec > total()) { rg = Object.assign({}, rg, { r2: rg.r2 + 1 }); if (rg.table) { const t = rg.table; O.tx(w, 'Data Form', () => { const old = t.ref; t.ref = Object.assign({}, t.ref, { r2: t.ref.r2 + 1 }); w.undo.op(() => { t.ref = old; }, () => { t.ref = Object.assign({}, old, { r2: old.r2 + 1 }); }); }); } }
    };
    const matches = (r) => cols.every((c, k) => { const q = criteria[k]; if (!q) return true; const t = textOf(s, r, c); const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(q); const op = m[1] || '', x = m[2]; const n = num(x, null); const v = val(s, r, c); if (op && n != null && typeof v === 'number') return op === '<' ? v < n : op === '>' ? v > n : op === '<=' ? v <= n : op === '>=' ? v >= n : op === '<>' ? v !== n : v === n; return op === '<>' ? t.toLowerCase() !== x.toLowerCase() : t.toLowerCase().startsWith(x.toLowerCase()); });
    const find = (d) => { for (let k = rec + d; k >= 1 && k <= total(); k += d) if (!criteria || matches(rg.r1 + k)) return k; return null; };
    load();
    const btn = (label, fn) => ui.button(label, fn, { style: 'width:100%' });
    let critMode = null;
    const side = h('div', { class: 'col df-side' }, count,
      btn('Ne&w', () => { save(); rec = total() + 1; load(); inputs[0].focus(); }),
      btn('&Delete', () => { if (rec > total()) return; ui.msg('Displayed record will be permanently deleted', { icon: 'warn', buttons: ['OK', 'Cancel'] }).then((r) => { if (r !== 0) return; tryRun(() => O.shiftCells(s, { r1: rg.r1 + rec, c1: rg.c1, r2: rg.r1 + rec, c2: rg.c2 }, 'up', false)); rg = Object.assign({}, rg, { r2: rg.r2 - 1 }); rec = Math.min(rec, Math.max(1, total())); load(); }); }),
      btn('&Restore', () => load()),
      btn('Find &Prev', () => { save(); const k = find(-1); if (k) { rec = k; load(); } }),
      btn('Find &Next', () => { save(); const k = find(1); if (k) { rec = k; load(); } }),
      critMode = btn('&Criteria', () => {
        if (!criteria) { save(); criteria = inputs.map(() => ''); inputs.forEach((i) => { i.value = ''; i.readOnly = false; }); count.textContent = 'Criteria'; critMode.textContent = 'Form'; }
        else { criteria = inputs.map((i) => i.value.trim()); const k = find(1) || find(-1) || rec; const crit = criteria; criteria = null; rec = k; load(); criteria = crit; criteria = crit.some(Boolean) ? crit : null; critMode.textContent = 'Criteria'; }
      }));
    inputs.forEach((i) => i.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); save(); if (rec <= total()) rec++; load(); } }));
    ui.dialog({ title: s.name, width: 460, body: h('div', { class: 'row df', style: 'gap:12px;align-items:flex-start' }, fieldsBox, side), buttons: [{ label: 'Cl&ose', primary: true, onClick: () => save() }] });
  };

  /* ================================================================ Create List */
  D.createList = function () {
    const s = sh(), w = wb();
    if (A().tableAt()) { ui.msg('A list cannot overlap another list.', { icon: 'warn' }); return; }
    const rg = listRange();
    const ref = refBox(F.absRangeName(rg), 200);
    let header = O.guessHeader(s, rg) || rg.r2 > rg.r1 && typeof s.val(rg.r1, rg.c1) === 'string';
    ui.dialog({ title: 'Create List', width: 320, body: h('div', { class: 'col' }, h('div', { text: 'Where is the data for your list?' }), ref, ui.check('&My list has headers', header, (v) => { header = v; })),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const p = parseRef(ref.input.value);
        if (!p) { ui.msg('Reference is not valid.', { icon: 'warn' }); return false; }
        tryRun(() => {
          const t = D.makeTable(p.sheet, { r1: p.r1, c1: p.c1, r2: Math.min(p.r2, Math.max(p.r1 + 1, p.sheet.maxR)), c2: p.c2 }, header);
          G().selectRange(t.ref, { r: t.ref.r1 + 1, c: t.ref.c1 });
        });
      } }, { label: 'Cancel' }] });
  };
  /** turn a range into a list (table); adds a header row when the data has none */
  D.makeTable = function (s, rg, header, styleName) {
    const w = s.wb;
    let t;
    O.tx(w, 'Create List', () => {
      if (!header) { O.insertLines(s, 'r', rg.r1, 1); rg = Object.assign({}, rg, { r2: rg.r2 + 1 }); for (let c = rg.c1; c <= rg.c2; c++) O.put(s, rg.r1, c, { v: 'Column' + (c - rg.c1 + 1) }); }
      let n = 1; while (w.findTable('Table' + n) || w.findName('Table' + n)) n++;
      const id = Math.max(0, ...w.tables.map((x) => x.id || 0)) + 1;
      const used = new Set();
      const columns = [];
      for (let c = rg.c1; c <= rg.c2; c++) { let name = String(s.val(rg.r1, c) == null || s.val(rg.r1, c) === '' ? 'Column' + (c - rg.c1 + 1) : textOf(s, rg.r1, c)); let k = 2; const base = name; while (used.has(name.toLowerCase())) name = base + k++; used.add(name.toLowerCase()); columns.push({ name, id: c - rg.c1 + 1 }); if (s.val(rg.r1, c) !== name) O.put(s, rg.r1, c, Object.assign({}, s.get(rg.r1, c) || {}, { v: name, f: undefined })); }
      t = { id, name: 'Table' + n, sheet: s, ref: Object.assign({}, rg), header: true, totals: false, columns, style: { name: styleName || 'TableStyleMedium2', rowStripes: true, colStripes: false, first: false, last: false }, autoFilter: true, filter: { cols: [] } };
      w.tables.push(t);
      w.undo.op(() => { w.tables = w.tables.filter((x) => x !== t); }, () => { if (!w.tables.includes(t)) w.tables.push(t); });
      O.markStructural();
    });
    return t;
  };

  /* ================================================================ Outline settings */
  D.outlineSettings = function () {
    const s = sh();
    const o = Object.assign({}, s.outline);
    ui.dialog({ title: 'Settings', width: 300, body: h('div', { class: 'col' }, ui.group('Direction', ui.check('Summary rows &below detail', o.below !== false, (v) => { o.below = v; }), ui.check('Summary columns to ri&ght of detail', o.right !== false, (v) => { o.right = v; })), ui.check('&Automatic styles', !!o.applyStyles, (v) => { o.applyStyles = v; })),
      buttons: [{ label: '&Create', onClick: () => A().autoOutline() }, { label: 'OK', primary: true, onClick: () => tryRun(() => { const old = s.outline; O.tx(s.wb, 'Outline Settings', () => { s.outline = Object.assign({}, s.outline, o); s.wb.undo.op(() => { s.outline = old; }, () => { s.outline = Object.assign({}, old, o); }); }); }) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Conditional Formatting (Excel 2003) */
  const CF_OPS = [['between', 'between'], ['notBetween', 'not between'], ['equal', 'equal to'], ['notEqual', 'not equal to'], ['greaterThan', 'greater than'], ['lessThan', 'less than'], ['greaterThanOrEqual', 'greater than or equal to'], ['lessThanOrEqual', 'less than or equal to']];
  D.conditionalFormat = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const ranges = G().ranges().map((r) => O.clip(s, r));
    const top = ranges[0];
    const editable = (r) => r.type === 'cellIs' || r.type === 'expression';
    /* the conditions that apply to the active cell */
    const existing = [];
    for (const cf of s.cf) if (cf.ranges.some((rg) => M.rangeContains(rg, sel.r, sel.c))) for (const r of cf.rules) if (editable(r)) existing.push({ rule: r, anchor: cf.ranges[0] });
    existing.sort((a, b) => (a.rule.priority || 0) - (b.rule.priority || 0));
    const fromStored = (f, anchor) => (f == null ? '' : /^"/.test(f) ? f.slice(1, -1).replace(/""/g, '"') : /^-?[\d.]+(E[+-]?\d+)?$/i.test(f) ? f : '=' + F.translate(F.display(f), sel.r - anchor.r1, sel.c - anchor.c1));
    const conds = existing.slice(0, 3).map(({ rule, anchor }) => ({ keep: JSON.parse(JSON.stringify(rule)), kind: rule.type === 'expression' ? 'formula' : 'value', op: rule.op || 'between', v1: fromStored(rule.f && rule.f[0], anchor), v2: fromStored(rule.f && rule.f[1], anchor), dxf: rule.dxf != null && w.dxfs[rule.dxf] ? JSON.parse(JSON.stringify(w.dxfs[rule.dxf])) : {}, stop: rule.stop }));
    if (!conds.length) conds.push({ kind: 'value', op: 'between', v1: '', v2: '', dxf: {} });
    const body = h('div', { class: 'col' });
    const prevCss = (d) => { const f = d.font || {}; const css = []; if (f.b) css.push('font-weight:bold'); if (f.i) css.push('font-style:italic'); if (f.u || f.strike) css.push('text-decoration:' + (f.u ? 'underline ' : '') + (f.strike ? 'line-through' : '')); if (f.color) css.push('color:' + M.colorHex(w, f.color)); if (d.fill && (d.fill.fg || d.fill.bg)) css.push('background:' + M.colorHex(w, d.fill.fg || d.fill.bg)); if (d.border) css.push('outline:1px solid #000;outline-offset:-3px'); return css.join(';'); };
    function draw() {
      body.textContent = '';
      conds.forEach((cd, i) => {
        const kindSel = ui.select([['value', 'Cell Value Is'], ['formula', 'Formula Is']], cd.kind, (v) => { cd.kind = v; draw(); });
        const opSel = ui.select(CF_OPS, cd.op, (v) => { cd.op = v; draw(); });
        const a = h('input', { type: 'text', value: cd.v1, style: 'width:' + (cd.kind === 'formula' ? '300px' : cd.op === 'between' || cd.op === 'notBetween' ? '110px' : '240px') });
        a.addEventListener('input', () => { cd.v1 = a.value; });
        const b = h('input', { type: 'text', value: cd.v2, style: 'width:110px' });
        b.addEventListener('input', () => { cd.v2 = b.value; });
        const two = cd.kind === 'value' && (cd.op === 'between' || cd.op === 'notBetween');
        const prev = h('div', { class: 'cf-prev', style: prevCss(cd.dxf), text: Object.keys(cd.dxf).length ? 'AaBbCcYyZz' : 'No Format Set' });
        const fmtBtn = ui.button('&Format...', () => D.cfFormat(cd.dxf, (d2) => { cd.dxf = d2; draw(); }));
        body.appendChild(ui.group('Condition ' + (i + 1), h('div', { class: 'row', style: 'gap:6px;align-items:center;flex-wrap:wrap' }, kindSel, cd.kind === 'value' ? opSel : null, a, two ? h('span', { text: 'and' }) : null, two ? b : null),
          h('div', { class: 'row', style: 'gap:8px;align-items:center;margin-top:4px' }, h('span', { class: 'hint', text: 'Preview of format to use when condition is true:' }), prev, fmtBtn)));
      });
    }
    draw();
    const toStore = (t) => {
      t = String(t).trim();
      if (t === '') return null;
      if (t[0] === '=') { const rel = F.translate(t.slice(1), top.r1 - sel.r, top.c1 - sel.c); F.parse(rel); return F.toStore(rel); }
      const p = NF.parseInput(t, { date1904: w.date1904 });
      if (p && typeof p.v === 'number') return String(p.v);
      if (p && typeof p.v === 'boolean') return p.v ? 'TRUE' : 'FALSE';
      return '"' + t.replace(/"/g, '""') + '"';
    };
    ui.dialog({ title: 'Conditional Formatting', width: 560, body,
      buttons: [{ label: '&Add >>', onClick: () => { if (conds.length < 3) { conds.push({ kind: 'value', op: 'between', v1: '', v2: '', dxf: {} }); draw(); } return false; } },
        { label: '&Delete...', onClick: () => { const checks = conds.map((c, i) => ui.check('Condition ' + (i + 1), false, null)); ui.dialog({ title: 'Delete Conditional Format', width: 260, body: ui.group('Select condition(s) to delete:', ...checks), buttons: [{ label: 'OK', primary: true, onClick: () => { for (let i = checks.length - 1; i >= 0; i--) if (checks[i].input.checked) conds.splice(i, 1); if (!conds.length) conds.push({ kind: 'value', op: 'between', v1: '', v2: '', dxf: {} }); draw(); } }, { label: 'Cancel' }] }); return false; } },
        { label: 'OK', primary: true, onClick: () => {
          const rules = [];
          try {
            for (const cd of conds) {
              if (cd.kind === 'formula') { if (!String(cd.v1).trim()) continue; const f = toStore(cd.v1[0] === '=' ? cd.v1 : '=' + cd.v1); rules.push({ ...cd.keep, type: 'expression', op: undefined, f: [f], dxf: cd.dxf }); }
              else { const f1 = toStore(cd.v1); if (f1 == null) continue; const two = cd.op === 'between' || cd.op === 'notBetween'; const f2 = two ? toStore(cd.v2) : null; if (two && f2 == null) throw new Error('You must enter a value in both boxes for between conditions.'); rules.push({ ...cd.keep, type: 'cellIs', op: cd.op, f: two ? [f1, f2] : [f1], dxf: cd.dxf }); }
            }
          } catch (e) { ui.msg(e.message || 'The formula contains an error.', { icon: 'warn' }); return false; }
          tryRun(() => O.tx(w, 'Conditional Formatting', () => {
            for (const rg of ranges) O.removeRangeRules(s, rg, 'cf', editable);
            if (!rules.length) return;
            let prio = Math.max(0, ...s.cf.flatMap((x) => x.rules.map((r) => r.priority || 0)));
            const out = rules.map((r) => { const dx = w.dxfs.push(r.dxf) - 1; return { ...r, dxf: dx, priority: ++prio }; });
            /* Excel 2003 conditions come before newer rules: renumber so these are first */
            const others = s.cf.flatMap((x) => x.rules).sort((a, b) => (a.priority || 0) - (b.priority || 0));
            out.forEach((r, i) => { r.priority = i + 1; });
            const shifted = s.cf.map((x) => Object.assign({}, x, { rules: x.rules.map((r) => Object.assign({}, r, { priority: out.length + 1 + others.indexOf(r) })) }));
            O.setCF(s, shifted.concat([{ ranges: ranges.map((r) => Object.assign({}, r)), rules: out }]));
          }));
        } }, { label: 'Cancel' }] });
  };
  /** the reduced Format Cells used by conditional formats: font style / colour, borders, fill */
  D.cfFormat = function (dxf, done) {
    const w = wb();
    const d = JSON.parse(JSON.stringify(dxf || {}));
    const f = d.font || {};
    const styleSel = ui.select([['', '(unchanged)'], ['r', 'Regular'], ['i', 'Italic'], ['b', 'Bold'], ['bi', 'Bold Italic']], f.b == null && f.i == null ? '' : (f.b ? 'b' : '') + (f.i ? 'i' : '') || 'r', (v) => { if (!v) { delete f.b; delete f.i; } else { f.b = /b/.test(v); f.i = /i/.test(v); } });
    const uSel = ui.select([['', '(unchanged)'], ['none', 'None'], ['single', 'Single'], ['double', 'Double']], f.u || '', (v) => { if (!v) delete f.u; else f.u = v === 'none' ? undefined : v; });
    const col = (get, set, mode) => { const b = h('button', { type: 'button', class: 'swatch-btn wide', style: `background:${get() || 'transparent'}` }, h('span', { text: get() ? '' : 'Automatic' }), h('span', { class: 'dd-arrow' })); b.addEventListener('click', () => ui.colorMenu(b, { mode, grid: true }, (c) => { if (c.auto || c.none) { set(null); b.style.background = 'transparent'; b.firstChild.textContent = c.none ? 'No Color' : 'Automatic'; } else if (typeof c === 'string') { set(c); b.style.background = c; b.firstChild.textContent = ''; } })); return b; };
    const fontTab = h('div', { class: 'col' }, ui.field('Font st&yle:', styleSel), ui.field('&Underline:', uSel), ui.field('&Color:', col(() => (f.color ? M.colorHex(w, f.color) : null), (c) => { if (c) f.color = M.rgb(c); else delete f.color; }, 'font')), ui.check('Stri&kethrough', !!f.strike, (v) => { f.strike = v || undefined; }));
    const fill = d.fill || {};
    const patTab = h('div', { class: 'col' }, ui.field('Cell shading color:', col(() => (fill.fg || fill.bg ? M.colorHex(w, fill.fg || fill.bg) : null), (c) => { if (c) { fill.pattern = 'solid'; fill.fg = M.rgb(c); fill.bg = M.rgb(c); } else { delete fill.pattern; delete fill.fg; delete fill.bg; } }, 'fill')));
    const bd = d.border || {};
    const bTab = h('div', { class: 'col' }, ui.check('&Outline', !!(bd.t && bd.b && bd.l && bd.r), (v) => { for (const k of ['t', 'b', 'l', 'r']) if (v) bd[k] = { style: 'thin' }; else delete bd[k]; }), ...[['t', 'Top'], ['b', 'Bottom'], ['l', 'Left'], ['r', 'Right']].map(([k, l]) => ui.check(l, !!bd[k], (v) => { if (v) bd[k] = { style: 'thin' }; else delete bd[k]; })));
    ui.dialog({ title: 'Format Cells', width: 340, body: ui.tabs([{ label: 'Font', body: fontTab }, { label: 'Border', body: bTab }, { label: 'Patterns', body: patTab }]),
      buttons: [{ label: '&Clear', onClick: () => { done({}); } }, { label: 'OK', primary: true, onClick: () => { const out = {}; const fo = {}; for (const k in f) if (f[k] !== undefined) fo[k] = f[k]; if (Object.keys(fo).length) out.font = fo; if (fill.pattern) out.fill = fill; if (Object.keys(bd).length) out.border = bd; done(out); } }, { label: 'Cancel' }] });
  };

  /* ================================================================ AutoFormat */
  const DARK = M.rgb('#000080'), MID = M.rgb('#C0C0C0'), PALE = M.rgb('#FFFFCC');
  const AUTOFMT = {
    'Simple': { head: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'thin' } } }, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'thin' } } }, body: {} },
    'Classic 1': { head: { font: { i: true }, border: { b: { style: 'thin' } } }, total: { border: { t: { style: 'thin' }, b: { style: 'double' } } }, first: { font: { b: true } }, body: {} },
    'Classic 2': { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#800080') } }, first: { fill: { pattern: 'solid', fg: MID }, font: { b: true } }, total: { font: { b: true }, border: { t: { style: 'thin' } } }, body: {} },
    'Classic 3': { head: { font: { b: true, i: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: DARK } }, body: { fill: { pattern: 'solid', fg: MID } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin' } } } },
    'Accounting 1': { head: { font: { b: true, i: true }, border: { b: { style: 'thin' } } }, body: { nf: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'double' } }, nf: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' } },
    'Accounting 2': { head: { font: { b: true }, border: { t: { style: 'medium' }, b: { style: 'thin' } } }, body: { nf: '#,##0.00_);(#,##0.00)' }, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'medium' } }, nf: '"$"#,##0.00_);\\("$"#,##0.00\\)' } },
    'Accounting 3': { head: { font: { b: true, i: true }, border: { b: { style: 'double' } } }, first: { font: { b: true } }, body: { nf: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }, total: { font: { b: true }, border: { t: { style: 'thin' } }, nf: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' } },
    'Accounting 4': { head: { font: { b: true }, border: { b: { style: 'thin' } } }, body: { nf: '#,##0.00_);(#,##0.00)' }, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'double' } }, nf: '"$"#,##0.00_);\\("$"#,##0.00\\)' } },
    'Colorful 1': { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#008080') } }, first: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#008080') } }, body: { fill: { pattern: 'solid', fg: M.rgb('#CCFFFF') } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: M.rgb('#99CCFF') } } },
    'Colorful 2': { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#993366') } }, body: { fill: { pattern: 'solid', fg: PALE } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: M.rgb('#FFCC99') } } },
    'Colorful 3': { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#000000') } }, first: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#000000') } }, body: { fill: { pattern: 'solid', fg: M.rgb('#FFCC00') } }, total: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#000000') } } },
    'List 1': { head: { font: { b: true }, border: { b: { style: 'medium' } } }, band: { fill: { pattern: 'solid', fg: M.rgb('#E3E3E3') } }, body: {}, total: { font: { b: true }, border: { t: { style: 'thin' } } } },
    'List 2': { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#333399') } }, band: { fill: { pattern: 'solid', fg: M.rgb('#CCCCFF') } }, body: {}, total: { font: { b: true }, border: { t: { style: 'double' } } } },
    'List 3': { head: { font: { b: true, i: true }, border: { t: { style: 'thin' }, b: { style: 'thin' } } }, band: { fill: { pattern: 'solid', fg: M.rgb('#CCFFCC') } }, body: {}, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'thin' } } } },
    '3D Effects 1': { head: { font: { b: true }, fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin', color: M.rgb('#FFFFFF') }, l: { style: 'thin', color: M.rgb('#FFFFFF') }, b: { style: 'thin', color: M.rgb('#808080') }, r: { style: 'thin', color: M.rgb('#808080') } } }, body: { fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin', color: M.rgb('#FFFFFF') }, l: { style: 'thin', color: M.rgb('#FFFFFF') }, b: { style: 'thin', color: M.rgb('#808080') }, r: { style: 'thin', color: M.rgb('#808080') } } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: MID } } },
    '3D Effects 2': { head: { font: { b: true }, fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin', color: M.rgb('#808080') }, l: { style: 'thin', color: M.rgb('#808080') }, b: { style: 'thin', color: M.rgb('#FFFFFF') }, r: { style: 'thin', color: M.rgb('#FFFFFF') } } }, body: { fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin', color: M.rgb('#808080') }, l: { style: 'thin', color: M.rgb('#808080') }, b: { style: 'thin', color: M.rgb('#FFFFFF') }, r: { style: 'thin', color: M.rgb('#FFFFFF') } } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: MID } } },
    'None': null,
  };
  D.autoFormat = function () {
    const s = sh(), w = wb();
    const rg = listRange();
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) { ui.msg('Cannot detect a table around the active cell.\n\nSelect the range you want to AutoFormat.', { icon: 'warn' }); return; }
    let pick = 'Simple';
    const parts = { nf: true, border: true, font: true, fill: true, align: true, size: true };
    const grid = h('div', { class: 'afmt-grid' });
    const sample = (name) => {
      const spec = AUTOFMT[name];
      const tbl = h('table', { class: 'afmt-sample' });
      const data = [['', 'Jan', 'Feb', 'Mar', 'Total'], ['East', 7, 7, 5, 19], ['West', 6, 4, 7, 17], ['South', 8, 7, 9, 24], ['Total', 21, 18, 21, 60]];
      data.forEach((row, r) => { const tr = h('tr'); row.forEach((x, c) => { const role = !spec ? {} : r === 0 ? spec.head : r === data.length - 1 ? spec.total : c === 0 && spec.first ? spec.first : spec.band && r % 2 === 0 ? Object.assign({}, spec.body, spec.band) : spec.body; const td = h('td', { text: String(x) }); const css = []; const f = role.font || {}; if (f.b) css.push('font-weight:bold'); if (f.i) css.push('font-style:italic'); if (f.color) css.push('color:' + M.colorHex(w, f.color)); if (role.fill) css.push('background:' + M.colorHex(w, role.fill.fg)); const b = role.border || {}; for (const [k, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) if (b[k]) css.push(`border-${side}:${b[k].style === 'medium' ? 2 : b[k].style === 'double' ? 3 : 1}px ${b[k].style === 'double' ? 'double' : 'solid'} ${M.colorHex(w, b[k].color, '#000')}`); td.style.cssText = css.join(';'); if (typeof x === 'number') td.style.textAlign = 'right'; tr.appendChild(td); }); tbl.appendChild(tr); });
      return tbl;
    };
    for (const name of Object.keys(AUTOFMT)) { const b = h('button', { type: 'button', class: 'afmt-item' + (name === pick ? ' on' : '') }, sample(name), h('div', { text: name })); b.addEventListener('click', () => { pick = name; L.$$('.afmt-item', grid).forEach((x) => x.classList.toggle('on', x === b)); }); grid.appendChild(b); }
    const optBox = h('div', { class: 'row', style: 'gap:10px;flex-wrap:wrap' }, ...[['nf', '&Number'], ['border', '&Border'], ['font', '&Font'], ['fill', '&Patterns'], ['align', '&Alignment'], ['size', '&Width/Height']].map(([k, l]) => ui.check(l, true, (v) => { parts[k] = v; })));
    optBox.hidden = true;
    ui.dialog({ title: 'AutoFormat', width: 560, body: h('div', { class: 'col' }, grid, optBox), buttons: [{ label: '&Options...', onClick: () => { optBox.hidden = !optBox.hidden; return false; } }, { label: 'OK', primary: true, onClick: () => tryRun(() => {
      const spec = AUTOFMT[pick];
      O.tx(w, 'AutoFormat', () => {
        for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
          const cl = s.get(r, c) || { v: null };
          const isTotal = r === rg.r2 && rg.r2 > rg.r1 + 1 && /total/i.test(String(s.val(r, rg.c1) || ''));
          let role = !spec ? null : r === rg.r1 ? spec.head : isTotal ? spec.total : c === rg.c1 && spec.first ? spec.first : spec.band && (r - rg.r1) % 2 === 0 ? Object.assign({}, spec.body, spec.band) : spec.body;
          const base = w.styles.get(cl.s || 0);
          const d = { };
          if (parts.font) d.font = role && role.font ? Object.assign({ b: undefined, i: undefined, color: undefined }, role.font) : { b: undefined, i: undefined, color: undefined };
          if (parts.fill) d.fill = role && role.fill ? role.fill : null;
          if (parts.border) d.border = role && role.border ? Object.assign({ t: null, b: null, l: null, r: null }, role.border) : { t: null, b: null, l: null, r: null };
          if (parts.nf && role && role.nf && typeof s.val(r, c) === 'number') d.nf = role.nf;
          if (parts.align && r === rg.r1 && c > rg.c1 && role) d.align = { h: 'center' };
          void base;
          const ns = w.styles.derive(cl.s || 0, d);
          if (ns !== (cl.s || 0)) O.put(s, r, c, Object.assign({}, cl, { s: ns }));
        }
        if (parts.size) O.autofitCols(s, Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i), { r1: rg.r1, r2: rg.r2 });
      });
    }) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Style */
  const BUILTIN_STYLES = { Normal: 0, Comma: 3, 'Comma [0]': 6, Currency: 4, 'Currency [0]': 7, Percent: 5, Hyperlink: 8, 'Followed Hyperlink': 9 };
  D.builtinStyle = function (w, name) {
    const base = w.styles.get(0);
    switch (name) {
      case 'Comma': return M.mergeStyle(base, { nf: NF.BUILTIN[43] });
      case 'Comma [0]': return M.mergeStyle(base, { nf: NF.BUILTIN[41] });
      case 'Currency': return M.mergeStyle(base, { nf: NF.BUILTIN[44] });
      case 'Currency [0]': return M.mergeStyle(base, { nf: NF.BUILTIN[42] });
      case 'Percent': return M.mergeStyle(base, { nf: '0%' });
      case 'Hyperlink': return M.mergeStyle(base, { font: { u: 'single', color: { theme: 10 } } });
      case 'Followed Hyperlink': return M.mergeStyle(base, { font: { u: 'single', color: { theme: 11 } } });
      default: return base;
    }
  };
  D.style = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const cur = LY.styleOf(s, sel.r, sel.c, s.get(sel.r, sel.c));
    const names = Array.from(new Set(Object.keys(BUILTIN_STYLES).concat(w.cellStyles.filter((c) => !c.hidden).map((c) => c.name))));
    let name = cur.xs || 'Normal';
    const nameIn = h('input', { type: 'text', value: name, list: 'style-names', style: 'width:220px' });
    const dl = h('datalist', { id: 'style-names' }, ...names.map((n) => h('option', { value: n })));
    const styleOf = (n) => { const cs = w.cellStyles.find((c) => c.name === n); return cs ? w.styles.get(cs.style) : D.builtinStyle(w, n); };
    const inc = { nf: true, align: true, font: true, border: true, fill: true, prot: true };
    const descBox = h('div', { class: 'col' });
    const describe = () => {
      const st = styleOf(nameIn.value) || cur;
      descBox.textContent = '';
      const lines = [['nf', 'Number', st.nf || 'General'], ['align', 'Alignment', (st.align && st.align.h) || 'General, Bottom Aligned'], ['font', 'Font', `${LY.fontName(w, st.font || {})} ${(st.font && st.font.sz) || 10}${st.font && st.font.b ? ' Bold' : ''}${st.font && st.font.i ? ' Italic' : ''}`], ['border', 'Border', st.border ? 'Borders' : 'No Borders'], ['fill', 'Patterns', st.fill ? 'Shaded' : 'No Shading'], ['prot', 'Protection', st.prot && st.prot.locked === false ? 'Unlocked' : 'Locked']];
      for (const [k, l, v] of lines) descBox.appendChild(h('div', { class: 'row', style: 'gap:6px' }, ui.check(l, inc[k], (x) => { inc[k] = x; }), h('span', { class: 'hint', text: v })));
    };
    nameIn.addEventListener('input', describe);
    describe();
    const fromCell = () => { const st = JSON.parse(JSON.stringify(cur)); delete st.xs; return st; };
    const save = (n, st) => {
      const idx = w.styles.add(st);
      const list = w.cellStyles.filter((c) => c.name !== n);
      const bi = BUILTIN_STYLES[n];
      list.push({ name: n, style: idx, builtinId: bi, customBuiltin: bi != null ? true : undefined });
      const old = w.cellStyles;
      O.tx(w, 'Style', () => { w.cellStyles = list; w.undo.op(() => { w.cellStyles = old; }, () => { w.cellStyles = list; }); });
    };
    ui.dialog({ title: 'Style', width: 420, body: h('div', { class: 'col' }, ui.field('&Style name:', nameIn), dl, ui.group('Style includes', descBox)),
      buttons: [
        { label: '&Add', onClick: () => { const n = nameIn.value.trim(); if (!n) return false; save(n, fromCell()); describe(); return false; } },
        { label: '&Delete', onClick: () => { const n = nameIn.value.trim(); if (n === 'Normal') { ui.msg('The Normal style cannot be deleted.', { icon: 'warn' }); return false; } const old = w.cellStyles; const list = old.filter((c) => c.name !== n); O.tx(w, 'Style', () => { w.cellStyles = list; w.undo.op(() => { w.cellStyles = old; }, () => { w.cellStyles = list; }); }); nameIn.value = 'Normal'; describe(); return false; } },
        { label: '&Modify...', onClick: () => {
          const n = nameIn.value.trim() || 'Normal';
          const base = styleOf(n) || cur;
          D.formatCells(0, { style: base, title: 'Format Cells - ' + n, apply: (newSt) => {
            if (n === 'Normal') {
              /* the default style: every cell without a style of its own follows it */
              const old = w.styles.list[0];
              const st = Object.freeze(Object.assign({}, newSt));
              O.tx(w, 'Modify Style', () => { const set = (x) => { w.styles.map.delete(JSON.stringify(w.styles.list[0])); w.styles.list[0] = x; w.styles.map.set(JSON.stringify(x), 0); w._mdw = null; for (const s2 of w.sheets) LY.invalidate(s2); }; set(st); w.undo.op(() => set(old), () => set(st)); O.markStructural(); });
            } else save(n, newSt);
            describe();
          } });
          return false;
        } },
        { label: 'OK', primary: true, onClick: () => tryRun(() => {
          const n = nameIn.value.trim() || 'Normal';
          let st = styleOf(n);
          if (!w.cellStyles.some((c) => c.name === n) && BUILTIN_STYLES[n] == null) { save(n, fromCell()); st = fromCell(); }
          O.format(s, G().ranges(), (cs) => {
            const d = {};
            if (inc.nf) d.nf = st.nf;
            if (inc.align) d.align = st.align;
            if (inc.font) d.font = Object.assign({ b: undefined, i: undefined, u: undefined, color: undefined }, st.font);
            if (inc.border) d.border = st.border ? Object.assign({ t: null, b: null, l: null, r: null }, st.border) : { t: null, b: null, l: null, r: null };
            if (inc.fill) d.fill = st.fill;
            if (inc.prot) d.prot = st.prot;
            d.xs = n === 'Normal' ? undefined : n;
            void cs;
            return d;
          }, 'Style');
        }) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Goal Seek */
  D.goalSeek = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const setC = refBox(F.cellName(sel.r, sel.c, true, true), 120), toV = h('input', { type: 'text', style: 'width:120px' }), byC = refBox('', 120);
    ui.dialog({ title: 'Goal Seek', width: 300, body: h('div', { class: 'col' }, ui.field('S&et cell:', setC), ui.field('To &value:', toV), ui.field('By &changing cell:', byC)),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const t = parseRef(setC.input.value, true), ch = parseRef(byC.input.value, true), goal = num(toV.value, NaN);
        if (!t || !(t.sheet.get(t.r1, t.c1) && t.sheet.get(t.r1, t.c1).f != null)) { ui.msg('Cell must contain a formula.', { icon: 'warn' }); return false; }
        if (!ch || (ch.sheet.get(ch.r1, ch.c1) && ch.sheet.get(ch.r1, ch.c1).f != null)) { ui.msg('Cell must contain a value.', { icon: 'warn' }); return false; }
        if (isNaN(goal)) { ui.msg('Invalid value.', { icon: 'warn' }); return false; }
        const res = D.solveGoal(w, t, ch, goal);
        const before = res.start;
        tryRun(() => { const cell = ch.sheet.get(ch.r1, ch.c1); if (cell) cell.v = before; C.changed(w, [{ sh: ch.sheet, r: ch.r1, c: ch.c1 }]); O.tx(w, 'Goal Seek', () => O.setValue(ch.sheet, ch.r1, ch.c1, res.x, true)); });
        ui.dialog({ title: 'Goal Seek Status', width: 300, body: h('div', { class: 'col' }, h('div', { text: `Goal Seeking with Cell ${F.cellName(t.r1, t.c1)} ${res.ok ? 'found a solution.' : 'may not have found a solution.'}` }), h('table', { class: 'kv' }, h('tr', null, h('td', { text: 'Target value:' }), h('td', { text: String(goal) })), h('tr', null, h('td', { text: 'Current value:' }), h('td', { text: L.csv.cellText(t.sheet, t.r1, t.c1) })))),
          buttons: [{ label: 'OK', primary: true }, { label: 'Cancel', onClick: () => A().undo() }] });
      } }, { label: 'Cancel' }] });
  };
  /** secant / bisection search for x with f(x) = goal; leaves the model as it was */
  D.solveGoal = function (w, t, ch, goal) {
    const cell = ch.sheet.get(ch.r1, ch.c1) || ch.sheet.cell(ch.r1, ch.c1);
    const start = typeof cell.v === 'number' ? cell.v : 0;
    const f = (x) => { cell.v = x; C.changed(w, [{ sh: ch.sheet, r: ch.r1, c: ch.c1 }]); const v = C.cellValue(t.sheet, t.r1, t.c1); return typeof v === 'number' ? v - goal : NaN; };
    let x0 = start, x1 = start === 0 ? 0.01 : start * 1.01;
    let f0 = f(x0), f1 = f(x1), best = { x: x0, e: Math.abs(f0) };
    let ok = false;
    for (let i = 0; i < 100; i++) {
      if (isFinite(f1) && Math.abs(f1) < best.e) best = { x: x1, e: Math.abs(f1) };
      if (Math.abs(f1) <= 0.001 * Math.max(1, Math.abs(goal)) * 1e-3 || Math.abs(f1) < 1e-9) { ok = true; break; }
      if (!isFinite(f0) || !isFinite(f1) || f1 === f0) { x1 = x1 * 2 + 1; f1 = f(x1); continue; }
      const x2 = x1 - f1 * (x1 - x0) / (f1 - f0);
      x0 = x1; f0 = f1; x1 = x2; f1 = f(x1);
    }
    if (!ok && best.e < 0.001 * Math.max(1, Math.abs(goal))) ok = true;
    cell.v = start;
    C.changed(w, [{ sh: ch.sheet, r: ch.r1, c: ch.c1 }]);
    return { x: ok ? x1 : best.x, ok, start };
  };

  /* ================================================================ Scenarios */
  D.scenarios = function () {
    const s = sh(), w = wb();
    if (!s.scenarios) s.scenarios = [];
    const box = h('div');
    let cur = 0;
    const draw = () => { box.textContent = ''; const lb = D.list(s.scenarios.map((x, i) => [i, x.name]), s.scenarios.length ? cur : null, (i) => { cur = i; info.textContent = s.scenarios[i] ? `Changing cells: ${s.scenarios[i].cells.map((c) => F.cellName(c.r, c.c, true, true)).join(',')}\nComment: ${s.scenarios[i].comment || ''}` : ''; }); lb.ondbl = () => show(); box.appendChild(s.scenarios.length ? lb : h('div', { class: 'tp-note', text: 'No Scenarios defined. Choose Add to add scenarios.' })); };
    const info = h('div', { class: 'hint', style: 'white-space:pre-line' });
    const show = () => { const sc = s.scenarios[cur]; if (!sc) return; tryRun(() => O.tx(w, 'Show Scenario', () => { for (const c of sc.cells) O.enter(s, c.r, c.c, String(c.v)); })); };
    const edit = (sc) => {
      const nm = h('input', { type: 'text', value: sc ? sc.name : '', style: 'width:100%' });
      const cells = refBox(sc ? sc.cells.map((c) => F.cellName(c.r, c.c, true, true)).join(',') : F.absRangeName(G().range()), 220);
      const cm = h('textarea', { rows: '3', style: 'width:100%' });
      cm.value = sc ? sc.comment || '' : 'Created by ' + (A().opts.userName || '') + ' on ' + new Date().toLocaleDateString();
      ui.dialog({ title: sc ? 'Edit Scenario' : 'Add Scenario', width: 360, body: h('div', { class: 'col' }, ui.field('Scenario &name:', nm), ui.field('Changing &cells:', cells), ui.field('Co&mment:', cm)),
        buttons: [{ label: 'OK', primary: true, onClick: () => {
          const refs = [];
          for (const part of cells.input.value.split(',')) { const p = parseRef(part); if (!p) { ui.msg('Reference is not valid.', { icon: 'warn' }); return false; } for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) refs.push({ r, c }); }
          if (!nm.value.trim()) return false;
          /* Scenario Values: one box per changing cell */
          const ins = refs.map((x) => h('input', { type: 'text', value: (sc && (sc.cells.find((q) => q.r === x.r && q.c === x.c) || {}).v) != null ? String(sc.cells.find((q) => q.r === x.r && q.c === x.c).v) : L.editor.cellText(s, x.r, x.c), style: 'width:120px' }));
          ui.dialog({ title: 'Scenario Values', width: 300, body: h('div', { class: 'col' }, h('div', { text: 'Enter values for each of the changing cells.' }), ...refs.map((x, i) => ui.field(F.cellName(x.r, x.c, true, true) + ':', ins[i]))),
            buttons: [{ label: 'OK', primary: true, onClick: () => {
              const item = { name: nm.value.trim(), comment: cm.value, cells: refs.map((x, i) => { const p = NF.parseInput(ins[i].value, {}); return { r: x.r, c: x.c, v: p && typeof p.v === 'number' ? p.v : ins[i].value }; }) };
              const old = s.scenarios.slice();
              const list = sc ? s.scenarios.map((q) => (q === sc ? item : q)) : s.scenarios.concat([item]);
              O.tx(w, 'Scenario', () => { s.scenarios = list; w.undo.op(() => { s.scenarios = old; }, () => { s.scenarios = list; }); });
              cur = list.indexOf(item); draw();
            } }, { label: 'Cancel' }] });
        } }, { label: 'Cancel' }] });
    };
    const summary = () => {
      if (!s.scenarios.length) return;
      tryRun(() => {
        const name = w.sheetByName('Scenario Summary') ? w.nextSheetName('Scenario Summary ') : 'Scenario Summary';
        const ns = O.addSheet(w, name, w.sheets.indexOf(s) + 1);
        O.tx(w, 'Scenario Summary', () => {
          const allCells = [];
          for (const sc of s.scenarios) for (const c of sc.cells) if (!allCells.some((x) => x.r === c.r && x.c === c.c)) allCells.push(c);
          const B = { font: { b: true } };
          const put = (r, c, v, st) => { const cell = { v }; if (st) cell.s = w.styles.derive(0, st); O.put(ns, r, c, cell); };
          put(0, 1, 'Scenario Summary', { font: { b: true, sz: 12 } });
          put(2, 2, 'Current Values:', B);
          s.scenarios.forEach((sc, i) => put(2, 3 + i, sc.name, B));
          put(3, 1, 'Changing Cells:', B);
          allCells.forEach((c, k) => { put(4 + k, 1, F.cellName(c.r, c.c, true, true)); put(4 + k, 2, s.val(c.r, c.c)); s.scenarios.forEach((sc, i) => { const x = sc.cells.find((q) => q.r === c.r && q.c === c.c); put(4 + k, 3 + i, x ? x.v : s.val(c.r, c.c)); }); });
          const base = 5 + allCells.length;
          put(base, 1, 'Result Cells:', B);
          /* result cells: the formulas that depend on the changing cells */
          const results = [];
          for (const c of allCells) { const deps = []; w.graph.dependents(s, c.r, c.c, deps); for (const d of deps) if (d.sh === s && !results.some((x) => x.r === d.r && x.c === d.c)) results.push({ r: d.r, c: d.c }); }
          results.slice(0, 30).forEach((rc, k) => {
            put(base + 1 + k, 1, F.cellName(rc.r, rc.c, true, true)); put(base + 1 + k, 2, val(s, rc.r, rc.c));
            s.scenarios.forEach((sc, i) => {
              const saved = sc.cells.map((x) => { const cl = s.get(x.r, x.c); return [cl, cl ? cl.v : null]; });
              sc.cells.forEach((x) => { const cl = s.get(x.r, x.c) || s.cell(x.r, x.c); cl.v = x.v; });
              C.changed(w, sc.cells.map((x) => ({ sh: s, r: x.r, c: x.c })));
              put(base + 1 + k, 3 + i, val(s, rc.r, rc.c));
              saved.forEach(([cl, v]) => { if (cl) cl.v = v; });
              C.changed(w, sc.cells.map((x) => ({ sh: s, r: x.r, c: x.c })));
            });
          });
          ns.cols[1] = { w: 16, custom: true }; ns.cols[2] = { w: 15, custom: true };
          for (let i = 0; i < s.scenarios.length; i++) ns.cols[3 + i] = { w: 13, custom: true };
        });
        w.active = w.sheets.indexOf(ns);
        A().renderTabs();
      });
    };
    draw();
    ui.dialog({ title: 'Scenario Manager', width: 420, body: h('div', { class: 'col' }, h('label', { text: 'Sc&enarios:' }), box, info),
      buttons: [{ label: '&Show', onClick: () => { show(); return false; } }, { label: '&Add...', onClick: () => { edit(null); return false; } }, { label: '&Delete', onClick: () => { const old = s.scenarios.slice(); const list = old.filter((_, i) => i !== cur); O.tx(w, 'Scenario', () => { s.scenarios = list; w.undo.op(() => { s.scenarios = old; }, () => { s.scenarios = list; }); }); cur = 0; draw(); return false; } }, { label: '&Edit...', onClick: () => { edit(s.scenarios[cur]); return false; } }, { label: 'S&ummary...', onClick: () => { summary(); } }, { label: 'Close', primary: true }] });
  };

  /* ================================================================ Data Table (what-if) */
  D.dataTable = function () {
    const s = sh();
    const rg = O.clip(s, G().range());
    if (rg.r2 <= rg.r1 && rg.c2 <= rg.c1) { ui.msg('Select the table range: the formula and input values must be included.', { icon: 'warn' }); return; }
    const rowIn = refBox('', 120), colIn = refBox('', 120);
    ui.dialog({ title: 'Table', width: 300, body: h('div', { class: 'col' }, ui.field('&Row input cell:', rowIn), ui.field('&Column input cell:', colIn)),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const ri = rowIn.input.value.trim() ? parseRef(rowIn.input.value, true) : null, ci = colIn.input.value.trim() ? parseRef(colIn.input.value, true) : null;
        if (!ri && !ci) { ui.msg('Input cell reference is not valid.', { icon: 'warn' }); return false; }
        tryRun(() => D.buildDataTable(s, rg, ri, ci));
      } }, { label: 'Cancel' }] });
  };
  /** fill a one- or two-variable data table (values; the TABLE() master keeps it recalculable on save) */
  D.buildDataTable = function (s, rg, ri, ci) {
    const w = s.wb;
    const two = ri && ci;
    const res = { r1: rg.r1 + 1, c1: rg.c1 + 1, r2: rg.r2, c2: rg.c2 };
    const inputs = [ri, ci].filter(Boolean).map((p) => { const cl = p.sheet.get(p.r1, p.c1) || p.sheet.cell(p.r1, p.c1); return { p, cl, v: cl.v }; });
    const setIn = (p, v) => { const cl = p.sheet.get(p.r1, p.c1) || p.sheet.cell(p.r1, p.c1); cl.v = v; C.changed(w, [{ sh: p.sheet, r: p.r1, c: p.c1 }]); };
    const out = [];
    try {
      for (let r = res.r1; r <= res.r2; r++) for (let c = res.c1; c <= res.c2; c++) {
        let fr, fc;
        if (two) { setIn(ri, val(s, rg.r1, c)); setIn(ci, val(s, r, rg.c1)); fr = rg.r1; fc = rg.c1; }
        else if (ci) { setIn(ci, val(s, r, rg.c1)); fr = rg.r1; fc = c; }
        else { setIn(ri, val(s, rg.r1, c)); fr = r; fc = rg.c1; }
        out.push({ r, c, v: val(s, fr, fc) });
      }
    } finally { for (const x of inputs) { x.cl.v = x.v; C.changed(w, [{ sh: x.p.sheet, r: x.p.r1, c: x.p.c1 }]); } }
    O.tx(w, 'Table', () => {
      for (const x of out) { const old = s.get(x.r, x.c); const n = Object.assign({}, old || {}); delete n.f; n.v = M.isErr(x.v) ? x.v : x.v; O.put(s, x.r, x.c, n); }
      const master = Object.assign({}, s.get(res.r1, res.c1) || {});
      master.dt = { ref: F.rangeName(res), dt2D: !!two, dtr: !!(ri && !ci), r1: ri ? F.cellName(ri.r1, ri.c1) : F.cellName(ci.r1, ci.c1), r2: two ? F.cellName(ci.r1, ci.c1) : undefined };
      O.put(s, res.r1, res.c1, master);
    });
  };

  /* ================================================================ Consolidate */
  const CONS_FN = [['SUM', 'Sum'], ['COUNT', 'Count'], ['AVERAGE', 'Average'], ['MAX', 'Max'], ['MIN', 'Min'], ['PRODUCT', 'Product'], ['COUNTNUMS', 'Count Nums'], ['STDEV', 'StdDev'], ['STDEVP', 'StdDevp'], ['VAR', 'Var'], ['VARP', 'Varp']];
  D.consolidate = function () {
    const s = sh(), w = wb(), sel = G().sel();
    let fn = 'SUM', top = false, left = false;
    const refs = (w._consRefs || []).slice();
    const refIn = refBox('', 220);
    const box = h('div');
    let cur = -1;
    const draw = () => { box.textContent = ''; box.appendChild(D.list(refs.map((x, i) => [i, x]), cur >= 0 ? cur : null, (i) => { cur = i; refIn.input.value = refs[i]; }, { class: 'lbox short' })); };
    draw();
    ui.dialog({ title: 'Consolidate', width: 380, body: h('div', { class: 'col' }, ui.field('&Function:', ui.select(CONS_FN, fn, (v) => { fn = v; })), ui.field('&Reference:', refIn), h('div', { class: 'row', style: 'gap:6px' }, ui.button('&Add', () => { const t = refIn.input.value.trim(); if (!parseRef(t)) { ui.msg('Reference is not valid.', { icon: 'warn' }); return; } if (!refs.includes(t)) refs.push(t); draw(); }), ui.button('&Delete', () => { if (cur >= 0) { refs.splice(cur, 1); cur = -1; draw(); } })), h('label', { text: 'All r&eferences:' }), box, ui.group('Use labels in', ui.check('&Top row', false, (v) => { top = v; }), ui.check('&Left column', false, (v) => { left = v; }))),
      buttons: [{ label: 'OK', primary: true, onClick: () => tryRun(() => {
        w._consRefs = refs.slice();
        const srcs = refs.map((t) => parseRef(t)).filter(Boolean);
        if (!srcs.length) return;
        const rowKeys = [], colKeys = [];
        const data = new Map();
        for (const p of srcs) {
          const r0 = p.r1 + (top ? 1 : 0), c0 = p.c1 + (left ? 1 : 0);
          for (let r = r0; r <= Math.min(p.r2, p.sheet.maxR); r++) for (let c = c0; c <= Math.min(p.c2, p.sheet.maxC); c++) {
            const rk = left ? String(L.csv.cellText(p.sheet, r, p.c1)) : String(r - r0);
            const ck = top ? String(L.csv.cellText(p.sheet, p.r1, c)) : String(c - c0);
            if (!rowKeys.includes(rk)) rowKeys.push(rk);
            if (!colKeys.includes(ck)) colKeys.push(ck);
            const v = val(p.sheet, r, c);
            const k = rk + '\u0001' + ck;
            if (!data.has(k)) data.set(k, []);
            data.get(k).push(v);
          }
        }
        const agg = (vals) => { const n = vals.filter((x) => typeof x === 'number'); switch (fn) { case 'COUNT': return vals.filter((x) => x != null && x !== '').length; case 'COUNTNUMS': return n.length; case 'AVERAGE': return n.length ? n.reduce((a, b) => a + b, 0) / n.length : 0; case 'MAX': return n.length ? Math.max(...n) : 0; case 'MIN': return n.length ? Math.min(...n) : 0; case 'PRODUCT': return n.reduce((a, b) => a * b, 1); case 'STDEV': case 'STDEVP': case 'VAR': case 'VARP': { const m = n.reduce((a, b) => a + b, 0) / (n.length || 1); const ss = n.reduce((a, b) => a + (b - m) * (b - m), 0); const vv = ss / Math.max(1, fn.endsWith('P') ? n.length : n.length - 1); return /STDEV/.test(fn) ? Math.sqrt(vv) : vv; } default: return n.reduce((a, b) => a + b, 0); } };
        O.tx(w, 'Consolidate', () => {
          const r0 = sel.r + (top ? 1 : 0), c0 = sel.c + (left ? 1 : 0);
          if (top) colKeys.forEach((k, j) => O.put(s, sel.r, c0 + j, { v: k }));
          if (left) rowKeys.forEach((k, i) => O.put(s, r0 + i, sel.c, { v: k }));
          rowKeys.forEach((rk, i) => colKeys.forEach((ck, j) => { const vals = data.get(rk + '\u0001' + ck); if (vals) O.put(s, r0 + i, c0 + j, { v: agg(vals) }); }));
        });
      }) }, { label: 'Close' }] });
  };

  /* ================================================================ PivotTable report (static summary) */
  const PIV_FN = [['sum', 'Sum'], ['count', 'Count'], ['average', 'Average'], ['max', 'Max'], ['min', 'Min'], ['countNums', 'Count Nums']];
  D.pivotTable = function () {
    const s = sh(), w = wb();
    const rg = listRange();
    if (rg.r2 <= rg.r1) { ui.msg('The PivotTable field name is not valid. To create a PivotTable report, you must use data that is organized as a list with labeled columns.', { icon: 'warn' }); return; }
    const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => [rg.c1 + i, colLabel(s, rg, true, rg.c1 + i)]);
    const numericCol = cols.find(([c]) => typeof val(s, rg.r1 + 1, c) === 'number');
    const st = { row: cols[0][0], col: '', data: numericCol ? numericCol[0] : cols[cols.length - 1][0], fn: numericCol ? 'sum' : 'count', page: '', pageVal: '', where: 'new' };
    const src = refBox(F.quoteSheet(s.name) + '!' + F.absRangeName(rg), 220);
    const opt = (k, none) => ui.select((none ? [['', '(none)']] : []).concat(cols), st[k], (v) => { st[k] = v === '' ? '' : +v; });
    ui.dialog({ title: 'PivotTable and PivotChart Wizard', width: 420, body: h('div', { class: 'col' },
      ui.field('&Range:', src),
      ui.group('Layout', ui.field('&Row field:', opt('row')), ui.field('&Column field:', opt('col', true)), h('div', { class: 'row', style: 'gap:8px' }, ui.field('&Data field:', opt('data')), ui.field('Summarize by:', ui.select(PIV_FN, st.fn, (v) => { st.fn = v; }))), ui.field('&Page field:', opt('page', true))),
      ui.group('Where do you want to put the PivotTable report?', ui.radio('pv-w', '&New worksheet', true, () => { st.where = 'new'; }), ui.radio('pv-w', '&Existing worksheet (at the active cell)', false, () => { st.where = 'here'; })),
      h('div', { class: 'hint', text: 'Ledger builds the report as a formatted summary of values; rebuild it after the source data changes.' })),
    buttons: [{ label: '&Finish', primary: true, onClick: () => {
      const p = parseRef(src.input.value);
      if (!p) { ui.msg('Reference is not valid.', { icon: 'warn' }); return false; }
      tryRun(() => D.buildPivot(p, st));
    } }, { label: 'Cancel' }] });
  };
  D.buildPivot = function (p, st) {
    const w = wb(), s = p.sheet;
    const r2 = Math.min(p.r2, s.maxR);
    const rowKeys = [], colKeys = [];
    const cells = new Map(), rowTot = new Map(), colTot = new Map(), all = [];
    const keyOf = (r, c) => { const t = L.csv.cellText(s, r, c); return t === '' ? '(blank)' : t; };
    const sortKey = (r, c) => { const v = val(s, r, c); return v; };
    const rowSort = new Map(), colSort = new Map();
    for (let r = p.r1 + 1; r <= r2; r++) {
      if (st.page !== '' && st.pageVal && keyOf(r, st.page) !== st.pageVal) continue;
      const rk = keyOf(r, st.row), ck = st.col === '' ? '' : keyOf(r, st.col);
      if (!rowSort.has(rk)) { rowSort.set(rk, sortKey(r, st.row)); rowKeys.push(rk); }
      if (!colSort.has(ck)) { colSort.set(ck, st.col === '' ? '' : sortKey(r, st.col)); colKeys.push(ck); }
      const v = val(s, r, st.data);
      const k = rk + '\u0001' + ck;
      for (const [map, kk] of [[cells, k], [rowTot, rk], [colTot, ck]]) { if (!map.has(kk)) map.set(kk, []); map.get(kk).push(v); }
      all.push(v);
    }
    const cmp = (m) => (a, b) => O.compareValues(m.get(a), m.get(b), {});
    rowKeys.sort(cmp(rowSort)); colKeys.sort(cmp(colSort));
    const agg = (vals) => { const n = vals.filter((x) => typeof x === 'number'); switch (st.fn) { case 'count': return vals.filter((x) => x != null && x !== '').length; case 'countNums': return n.length; case 'average': return n.length ? n.reduce((a, b) => a + b, 0) / n.length : M.ERR.DIV0; case 'max': return n.length ? Math.max(...n) : 0; case 'min': return n.length ? Math.min(...n) : 0; default: return n.reduce((a, b) => a + b, 0); } };
    const fnLabel = PIV_FN.find((x) => x[0] === st.fn)[1];
    const head = (c) => String(s.val(p.r1, c) == null ? F.colName(c) : s.val(p.r1, c));
    let dst, r0, c0;
    O.tx(w, 'PivotTable', () => {
      if (st.where === 'new') { dst = O.addSheet(w, w.nextSheetName('Sheet'), w.sheets.indexOf(s)); r0 = 2; c0 = 0; }
      else { dst = sh(); r0 = G().sel().r; c0 = G().sel().c; }
      const BOLD = w.styles.derive(0, { font: { b: true } });
      const HEAD = w.styles.derive(0, { font: { b: true }, fill: { pattern: 'solid', fg: M.rgb('#C0C0C0') }, border: { t: { style: 'thin' }, b: { style: 'thin' } } });
      const TOT = w.styles.derive(0, { font: { b: true }, fill: { pattern: 'solid', fg: M.rgb('#E3E3E3') }, border: { t: { style: 'thin' } } });
      const put = (r, c, v, sid) => O.put(dst, r, c, { v, s: sid || undefined });
      if (st.page !== '') { put(r0 - 2 < 0 ? r0 : r0 - 2, c0, head(st.page), BOLD); put(r0 - 2 < 0 ? r0 : r0 - 2, c0 + 1, st.pageVal || '(All)'); }
      put(r0, c0, fnLabel + ' of ' + head(st.data), HEAD);
      if (st.col !== '') put(r0, c0 + 1, head(st.col), HEAD);
      const hr = r0 + 1;
      put(hr, c0, head(st.row), HEAD);
      const showCols = st.col === '' ? [''] : colKeys;
      showCols.forEach((ck, j) => put(hr, c0 + 1 + j, st.col === '' ? 'Total' : ck, HEAD));
      if (st.col !== '') put(hr, c0 + 1 + showCols.length, 'Grand Total', HEAD);
      rowKeys.forEach((rk, i) => {
        put(hr + 1 + i, c0, rk);
        showCols.forEach((ck, j) => { const vals = cells.get(rk + '\u0001' + ck); if (vals) put(hr + 1 + i, c0 + 1 + j, agg(vals)); });
        if (st.col !== '') put(hr + 1 + i, c0 + 1 + showCols.length, agg(rowTot.get(rk)), BOLD);
      });
      const gr = hr + 1 + rowKeys.length;
      put(gr, c0, 'Grand Total', TOT);
      showCols.forEach((ck, j) => put(gr, c0 + 1 + j, agg(colTot.get(ck) || []), TOT));
      if (st.col !== '') put(gr, c0 + 1 + showCols.length, agg(all), TOT);
      O.autofitCols(dst, Array.from({ length: showCols.length + 2 }, (_, i) => c0 + i));
    });
    if (st.where === 'new') { w.active = w.sheets.indexOf(dst); A().renderTabs(); }
  };
})(typeof window !== 'undefined' ? window : globalThis);
