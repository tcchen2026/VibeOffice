/* Ledger — AutoFilter (Data ▸ Filter): the Excel 2003 drop-down list, Custom AutoFilter, Top 10, Advanced Filter helpers. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, O = L.ops, LY = L.layout, NF = L.numfmt, C = L.calc;
  const { h } = L;
  const ui = L.ui;
  const FL = (L.filter = {});
  const G = () => L.grid;
  const sh = () => G().sheet();

  /** displayed text of a cell, as filters compare it */
  const textOf = (s, r, c) => (L.csv ? L.csv.cellText(s, r, c) : String(s.val(r, c) == null ? '' : s.val(r, c)));
  const valOf = (s, r, c) => { const cl = s.get(r, c); if (!cl) return null; return cl.dirty ? C.cellValue(s, r, c) : cl.v; };
  const hasCriteria = (fc) => !!(fc && (fc.values || fc.custom || fc.top10 || fc.dynamic || fc.color || fc.icon || fc.blank || fc.dates));

  /* ------------------------------------------------------------ criteria */
  const wild = (pat) => new RegExp('^' + String(pat).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '[\\s\\S]').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'i');
  function customTest(cond, v, text) {
    const op = cond.op || 'equal';
    const raw = cond.val == null ? '' : String(cond.val);
    const num = raw.trim() !== '' && !isNaN(+raw) ? +raw : null;
    if (op === 'equal' || op === 'notEqual') {
      let eq;
      if (raw === '' || raw === ' ' && op === 'notEqual') eq = text === '' || (raw === ' ' ? false : text === '');
      if (raw === ' ' && op === 'notEqual') return text !== '';
      if (/[*?]/.test(raw)) eq = wild(raw).test(text);
      else if (num != null && typeof v === 'number') eq = Math.abs(v - num) < 1e-12 * Math.max(1, Math.abs(num)) || text === raw;
      else eq = text.toLowerCase() === raw.toLowerCase();
      return op === 'equal' ? eq : !eq;
    }
    /* comparisons: numbers compare as numbers, text as text */
    let cmp;
    if (num != null && typeof v === 'number') cmp = v - num;
    else if (num != null) return false;
    else { if (typeof v === 'number') return false; const a = text.toLowerCase(), b = raw.toLowerCase(); cmp = a < b ? -1 : a > b ? 1 : 0; }
    return op === 'greaterThan' ? cmp > 0 : op === 'greaterThanOrEqual' ? cmp >= 0 : op === 'lessThan' ? cmp < 0 : op === 'lessThanOrEqual' ? cmp <= 0 : true;
  }
  const serialNow = () => NF.jsDateToSerial(new Date(), G().wb && G().wb.date1904);
  function dynamicTest(dyn, v, stats) {
    const t = dyn.type;
    if (t === 'aboveAverage') return typeof v === 'number' && v > stats.avg;
    if (t === 'belowAverage') return typeof v === 'number' && v < stats.avg;
    if (typeof v !== 'number') return false;
    const today = Math.floor(serialNow());
    const p = NF.serialToParts(v, false);
    const tp = NF.serialToParts(today, false);
    const d = Math.floor(v);
    switch (t) {
      case 'today': return d === today;
      case 'yesterday': return d === today - 1;
      case 'tomorrow': return d === today + 1;
      case 'thisWeek': { const s = today - ((today + 6) % 7); return d >= s && d < s + 7; }
      case 'lastWeek': { const s = today - ((today + 6) % 7) - 7; return d >= s && d < s + 7; }
      case 'nextWeek': { const s = today - ((today + 6) % 7) + 7; return d >= s && d < s + 7; }
      case 'thisMonth': return p.y === tp.y && p.m === tp.m;
      case 'lastMonth': { const m = tp.m === 1 ? 12 : tp.m - 1, y = tp.m === 1 ? tp.y - 1 : tp.y; return p.y === y && p.m === m; }
      case 'nextMonth': { const m = tp.m === 12 ? 1 : tp.m + 1, y = tp.m === 12 ? tp.y + 1 : tp.y; return p.y === y && p.m === m; }
      case 'thisYear': return p.y === tp.y;
      case 'lastYear': return p.y === tp.y - 1;
      case 'nextYear': return p.y === tp.y + 1;
      case 'yearToDate': return p.y === tp.y && d <= today;
      default: {
        let m;
        if ((m = /^M(\d+)$/.exec(t))) return p.m === +m[1];
        if ((m = /^Q(\d)$/.exec(t))) return Math.ceil(p.m / 3) === +m[1];
        if (dyn.val != null && dyn.maxVal != null) return v >= dyn.val && v < dyn.maxVal;
        return true;
      }
    }
  }
  function dateGroupTest(list, v) {
    if (typeof v !== 'number') return false;
    const p = NF.serialToParts(v, false);
    return list.some((g) => {
      if (p.y !== g.y) return false;
      if (g.g === 'year') return true;
      if (p.m !== g.m) return false;
      if (g.g === 'month') return true;
      if (p.d !== g.d) return false;
      if (g.g === 'day') return true;
      if (p.H !== g.H) return false;
      if (g.g === 'hour') return true;
      if (p.M !== g.M) return false;
      if (g.g === 'minute') return true;
      return p.S === g.S;
    });
  }
  /** row → passes column criterion fc? */
  function makeTester(s, ref, fc) {
    const c = ref.c1 + fc.col;
    const r1 = ref.r1 + 1, r2 = ref.r2;
    let stats = null, topSet = null;
    if (fc.top10 || (fc.dynamic && /Average/.test(fc.dynamic.type))) {
      const nums = [];
      for (let r = r1; r <= r2; r++) { const v = valOf(s, r, c); if (typeof v === 'number') nums.push(v); }
      stats = { avg: nums.reduce((a, b) => a + b, 0) / (nums.length || 1) };
      if (fc.top10) {
        const sorted = nums.slice().sort((a, b) => (fc.top10.top ? b - a : a - b));
        const n = fc.top10.percent ? Math.max(1, Math.floor((nums.length * fc.top10.val) / 100)) : Math.max(1, Math.round(fc.top10.val));
        const cut = sorted[Math.min(n, sorted.length) - 1];
        topSet = { cut, top: fc.top10.top };
      }
    }
    const valueSet = fc.values ? new Set(fc.values.map((x) => String(x).toLowerCase())) : null;
    return (r) => {
      const v = valOf(s, r, c);
      const text = v == null ? '' : textOf(s, r, c);
      if (fc.values || fc.blank || fc.dates) {
        if (text === '' || v == null) return !!fc.blank;
        if (valueSet && valueSet.has(text.toLowerCase())) return true;
        if (valueSet && typeof v === 'number' && valueSet.has(String(v))) return true;
        if (fc.dates && dateGroupTest(fc.dates, v)) return true;
        return false;
      }
      if (fc.custom) {
        const res = fc.custom.list.map((cond) => customTest(cond, v, text));
        return fc.custom.and ? res.every(Boolean) : res.some(Boolean);
      }
      if (fc.top10) { if (typeof v !== 'number' || topSet.cut == null) return false; return topSet.top ? v >= topSet.cut : v <= topSet.cut; }
      if (fc.dynamic) return dynamicTest(fc.dynamic, v, stats);
      if (fc.color) {
        const dxf = G().wb.dxfs[fc.color.dxf];
        const st = LY.styleOf(s, r, c, s.get(r, c));
        if (!dxf) return true;
        if (fc.color.cell) { const want = dxf.fill && M.colorHex(G().wb, dxf.fill.fg || dxf.fill.bg); const have = st.fill && M.colorHex(G().wb, st.fill.fg); return (want || null) === (have || null); }
        const want = dxf.font && M.colorHex(G().wb, dxf.font.color, '#000000'); const have = M.colorHex(G().wb, st.font && st.font.color, '#000000'); return want === have;
      }
      return true;
    };
  }
  /** hide rows that fail any criterion of filter f over range ref (header row excluded) */
  FL.applyRows = function (s, ref, f) {
    const testers = (f.cols || []).filter(hasCriteria).map((fc) => makeTester(s, ref, fc));
    const last = Math.min(ref.r2, Math.max(s.maxR, ref.r1));
    for (let r = ref.r1 + 1; r <= last; r++) {
      const ok = testers.every((t) => t(r));
      const row = s.rows[r];
      if (ok) { if (row && row.filtered) { delete row.hidden; delete row.filtered; } }
      else { const o = s.rowObj(r); o.hidden = true; o.filtered = true; }
    }
    LY.invalidate(s);
  };
  /** the filter (sheet AutoFilter or table) that contains column c of the active cell */
  function filterAt(s, r, c) {
    const t = G().wb.tableAt(s, r, c);
    if (t && t.autoFilter !== false && t.header !== false) return { ref: t.ref, f: t.filter || (t.filter = { cols: [] }), table: t };
    if (s.autoFilter && s.autoFilter.ref && M.rangeContains(s.autoFilter.ref, r, c)) return { ref: s.autoFilter.ref, f: s.autoFilter, table: null };
    return null;
  }
  /** change a filter column's criterion and re-filter (undoable) */
  FL.setCriterion = function (x, col, crit) {
    const s = sh(), wb = G().wb;
    const f = x.table ? x.table.filter || (x.table.filter = { cols: [] }) : s.autoFilter;
    const old = JSON.parse(JSON.stringify(f));
    const apply = (nf) => { if (x.table) x.table.filter = nf; else s.autoFilter = nf; };
    O.structural(s, 'Filter', () => {
      const nf = JSON.parse(JSON.stringify(f));
      nf.cols = nf.cols.filter((q) => q.col !== col);
      if (crit) nf.cols.push(Object.assign({ col }, crit));
      nf.cols.sort((a, b) => a.col - b.col);
      apply(nf);
      const ref = x.table ? x.table.ref : nf.ref;
      FL.applyRows(s, x.table ? { r1: ref.r1, c1: ref.c1, r2: x.table.totals ? ref.r2 - 1 : ref.r2, c2: ref.c2 } : ref, nf);
      const now = JSON.parse(JSON.stringify(nf));
      if (x.table) wb.undo.op(() => apply(JSON.parse(JSON.stringify(old))), () => apply(JSON.parse(JSON.stringify(now))));
    });
    G().paint(); G().syncObjects();
    L.app.updateStatus && L.app.updateStatus();
    const shown = countVisible(s, x);
    if (L.app.status) L.app.status(crit ? `${shown.vis} of ${shown.total} records found` : '');
  };
  function countVisible(s, x) {
    const ref = x.table ? x.table.ref : (x.table ? null : s.autoFilter && s.autoFilter.ref) || x.ref;
    let vis = 0, total = 0;
    const last = Math.min(ref.r2 - (x.table && x.table.totals ? 1 : 0), s.maxR);
    for (let r = ref.r1 + 1; r <= last; r++) { total++; const row = s.rows[r]; if (!(row && row.hidden)) vis++; }
    return { vis, total };
  }

  /* ------------------------------------------------------------ commands */
  FL.active = function () {
    const s = G().sheet();
    if (s.autoFilter && s.autoFilter.cols && s.autoFilter.cols.some(hasCriteria)) return true;
    return G().wb.tables.some((t) => t.sheet === s && t.filter && t.filter.cols.some(hasCriteria));
  };
  FL.toggle = function () {
    const s = sh(), wb = G().wb, sel = G().sel();
    const t = wb.tableAt(s, sel.r, sel.c);
    if (t) {
      const was = t.autoFilter !== false;
      O.structural(s, 'AutoFilter', () => {
        if (was && t.filter) { t.filter.cols = []; FL.applyRows(s, t.ref, t.filter); }
        t.autoFilter = !was;
        wb.undo.op(() => { t.autoFilter = was; }, () => { t.autoFilter = !was; });
      });
      G().paint(); G().syncObjects(); return;
    }
    if (s.autoFilter) {
      O.structural(s, 'AutoFilter', () => {
        s.rows.forEach((row) => { if (row && row.filtered) { delete row.hidden; delete row.filtered; } });
        s.autoFilter = null;
        LY.invalidate(s);
      });
      G().paint(); G().syncObjects(); L.app.status && L.app.status('');
      return;
    }
    let rg = O.clip(s, G().range());
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    if (single) rg = O.currentRegion(s, sel.r, sel.c);
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && !(s.get(rg.r1, rg.c1) && s.get(rg.r1, rg.c1).v != null)) {
      ui.msg('No list was found. Select a single cell within your list, and then click the command again.', { icon: 'warn' });
      return;
    }
    if (single && !O.guessHeader(s, rg) && rg.r1 > 0) { /* Excel uses the region's first row as headers regardless */ }
    O.structural(s, 'AutoFilter', () => { s.autoFilter = { ref: rg, cols: [] }; });
    G().paint(); G().syncObjects();
  };
  FL.showAll = function () {
    const s = sh();
    O.structural(s, 'Show All', () => {
      if (s.autoFilter) s.autoFilter.cols = s.autoFilter.cols.filter((q) => !hasCriteria(q));
      for (const t of G().wb.tables) if (t.sheet === s && t.filter) t.filter.cols = [];
      s.rows.forEach((row) => { if (row && row.filtered) { delete row.hidden; delete row.filtered; } });
      LY.invalidate(s);
    });
    G().paint(); G().syncObjects(); L.app.status && L.app.status('');
  };

  /* ------------------------------------------------------------ the drop-down */
  /** distinct entries of column c among rows visible under the other columns' criteria */
  function entries(s, x, c) {
    const ref = x.table ? { r1: x.table.ref.r1, c1: x.table.ref.c1, r2: x.table.ref.r2 - (x.table.totals ? 1 : 0), c2: x.table.ref.c2 } : x.ref;
    const col = c - ref.c1;
    const others = (x.f.cols || []).filter((q) => q.col !== col && hasCriteria(q)).map((q) => makeTester(s, ref, q));
    const nums = new Map(), texts = new Map();
    let blanks = false;
    const last = Math.min(ref.r2, s.maxR);
    for (let r = ref.r1 + 1; r <= last; r++) {
      const row = s.rows[r];
      if (row && row.hidden && !row.filtered) continue;
      if (!others.every((t) => t(r))) continue;
      const v = valOf(s, r, c);
      const t = v == null ? '' : textOf(s, r, c);
      if (t === '') { blanks = true; continue; }
      if (typeof v === 'number') { if (!nums.has(t)) nums.set(t, v); }
      else if (!texts.has(t.toLowerCase())) texts.set(t.toLowerCase(), t);
      if (nums.size + texts.size > 10000) break;
    }
    const list = Array.from(nums.entries()).sort((a, b) => a[1] - b[1]).map((e) => e[0]).concat(Array.from(texts.values()).sort((a, b) => C.strCmp ? C.strCmp(a, b) : a.localeCompare(b)));
    return { list, blanks, ref, col };
  }
  FL.open = function (btn, x, c) {
    const s = sh();
    if (G().guardProtect && G().guardProtect('autoFilter', true)) return;
    const { list, blanks, ref, col } = entries(s, x, c);
    const fc = (x.f.cols || []).find((q) => q.col === col);
    const cur = fc && fc.values && fc.values.length === 1 && !fc.blank ? fc.values[0] : fc && fc.blank && !(fc.values && fc.values.length) ? '(Blanks)' : fc && fc.custom && fc.custom.list.length === 1 && fc.custom.list[0].op === 'notEqual' && fc.custom.list[0].val === ' ' ? '(NonBlanks)' : !hasCriteria(fc) ? '(All)' : null;
    const items = ['(Sort Ascending)', '(Sort Descending)', '(All)', '(Top 10...)', '(Custom...)'].concat(list.slice(0, 1000));
    if (list.length > 1000) items.push('…');
    if (blanks) items.push('(Blanks)', '(NonBlanks)');
    const r = btn.getBoundingClientRect();
    const box = h('div', { class: 'af-list', role: 'listbox', tabindex: '-1' });
    const q = G().cellRect(ref.r1, c);
    box.style.minWidth = Math.max(q.w + 2, 120) + 'px';
    let act = -1;
    const rows = items.map((t, i) => {
      const el = h('div', { class: 'af-item' + (t === cur ? ' cur' : ''), role: 'option', text: t });
      el.addEventListener('pointerenter', () => setAct(i));
      el.addEventListener('click', () => choose(t));
      box.appendChild(el);
      return el;
    });
    const setAct = (i) => { if (rows[act]) rows[act].classList.remove('act'); act = i; if (rows[act]) { rows[act].classList.add('act'); rows[act].scrollIntoView({ block: 'nearest' }); } };
    function close() { box.remove(); document.removeEventListener('pointerdown', outside, true); document.removeEventListener('keydown', keys, true); G().focus(); }
    function outside(e) { if (!box.contains(e.target) && e.target !== btn) close(); }
    function keys(e) {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); }
      else if (e.key === 'ArrowDown') { e.preventDefault(); e.stopPropagation(); setAct(Math.min(rows.length - 1, act + 1)); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); e.stopPropagation(); setAct(Math.max(0, act - 1)); }
      else if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); if (act >= 0) choose(items[act]); }
      else if (e.key.length === 1) { const k = e.key.toLowerCase(); const i = items.findIndex((t, j) => j > act && t.toLowerCase().startsWith(k)); const i2 = i >= 0 ? i : items.findIndex((t) => t.toLowerCase().startsWith(k)); if (i2 >= 0) setAct(i2); e.preventDefault(); e.stopPropagation(); }
    }
    function choose(t) {
      close();
      const dataRef = { r1: ref.r1, c1: ref.c1, r2: Math.min(ref.r2, Math.max(s.maxR, ref.r1)), c2: ref.c2 };
      if (t === '(Sort Ascending)' || t === '(Sort Descending)') { try { O.sort(s, dataRef, [{ index: col, desc: t === '(Sort Descending)' }], { header: true }); } catch (e) { L.app.error(e); } G().paint(); return; }
      if (t === '(All)') { FL.setCriterion(x, col, null); return; }
      if (t === '(Top 10...)') { FL.top10Dialog().then((o) => { if (o) FL.setCriterion(x, col, { top10: o }); }); return; }
      if (t === '(Custom...)') { FL.customDialog(s, x, c, fc).then((o) => { if (o) FL.setCriterion(x, col, { custom: o }); }); return; }
      if (t === '(Blanks)') { FL.setCriterion(x, col, { blank: true, values: [] }); return; }
      if (t === '(NonBlanks)') { FL.setCriterion(x, col, { custom: { and: false, list: [{ op: 'notEqual', val: ' ' }] } }); return; }
      if (t === '…') return;
      FL.setCriterion(x, col, { values: [t] });
    }
    document.body.appendChild(box);
    const bh = box.getBoundingClientRect();
    let left = Math.min(r.right - bh.width, window.innerWidth - bh.width - 2);
    left = Math.max(2, Math.min(left, q.x + L.$('#gridwrap').getBoundingClientRect().left));
    let top = r.bottom;
    if (top + bh.height > window.innerHeight - 4) top = Math.max(4, r.top - bh.height);
    Object.assign(box.style, { left: left + 'px', top: top + 'px' });
    setTimeout(() => { document.addEventListener('pointerdown', outside, true); document.addEventListener('keydown', keys, true); const ci = items.indexOf(cur); setAct(ci >= 0 ? ci : 0); }, 0);
  };

  /* ------------------------------------------------------------ dialogs */
  const OPS = [['equal', 'equals'], ['notEqual', 'does not equal'], ['greaterThan', 'is greater than'], ['greaterThanOrEqual', 'is greater than or equal to'], ['lessThan', 'is less than'], ['lessThanOrEqual', 'is less than or equal to'],
    ['begins', 'begins with'], ['notBegins', 'does not begin with'], ['ends', 'ends with'], ['notEnds', 'does not end with'], ['contains', 'contains'], ['notContains', 'does not contain']];
  /** stored {op,val} → dialog operator + text */
  function toDialog(cond) {
    let op = cond.op || 'equal', v = cond.val == null ? '' : String(cond.val);
    if (op === 'equal' || op === 'notEqual') {
      const neg = op === 'notEqual';
      if (/^\*[^*?]+\*$/.test(v)) return { op: neg ? 'notContains' : 'contains', v: v.slice(1, -1) };
      if (/^[^*?]+\*$/.test(v)) return { op: neg ? 'notBegins' : 'begins', v: v.slice(0, -1) };
      if (/^\*[^*?]+$/.test(v)) return { op: neg ? 'notEnds' : 'ends', v: v.slice(1) };
    }
    return { op, v };
  }
  function fromDialog(op, v) {
    switch (op) {
      case 'begins': return { op: 'equal', val: v + '*' };
      case 'notBegins': return { op: 'notEqual', val: v + '*' };
      case 'ends': return { op: 'equal', val: '*' + v };
      case 'notEnds': return { op: 'notEqual', val: '*' + v };
      case 'contains': return { op: 'equal', val: '*' + v + '*' };
      case 'notContains': return { op: 'notEqual', val: '*' + v + '*' };
      default: {
        /* numbers typed with formatting are stored as plain numbers */
        const p = NF.parseInput(v, {});
        return { op, val: p && typeof p.v === 'number' && /[\d]/.test(v) ? String(p.v) : v };
      }
    }
  }
  FL.customDialog = function (s, x, c, fc) {
    const ent = entries(s, x, c).list;
    const conds = fc && fc.custom ? fc.custom.list.map(toDialog) : fc && fc.values && fc.values.length === 1 ? [{ op: 'equal', v: fc.values[0] }] : [];
    const mk = (i) => {
      const cur = conds[i] || { op: i === 0 ? 'equal' : '', v: '' };
      const opSel = ui.select([['', '']].concat(OPS), cur.op, null, { style: 'width:200px' });
      const inp = h('input', { type: 'text', value: cur.v, list: 'af-vals', style: 'width:220px' });
      return { opSel, inp, row: h('div', { class: 'row', style: 'gap:8px;margin:4px 0' }, opSel, inp) };
    };
    const a = mk(0), b = mk(1);
    const dl = h('datalist', { id: 'af-vals' }, ...ent.slice(0, 500).map((v) => h('option', { value: v })));
    const and = ui.radio('af-andor', '&And', !(fc && fc.custom && !fc.custom.and) || !(fc && fc.custom), null);
    const or = ui.radio('af-andor', '&Or', !!(fc && fc.custom && !fc.custom.and && fc.custom.list.length > 1), null);
    const head = s.val(x.ref.r1, c);
    const body = h('div', { class: 'col' }, h('div', { text: 'Show rows where:' }), h('b', { text: head == null ? L.formula.colName(c) : String(head) }), a.row, h('div', { class: 'row', style: 'gap:16px;margin-left:24px' }, and, or), b.row, dl,
      h('div', { class: 'hint', text: 'Use ? to represent any single character\nUse * to represent any series of characters', style: 'white-space:pre-line;margin-top:6px' }));
    let res = null;
    return ui.dialog({ title: 'Custom AutoFilter', body, width: 480, buttons: [{ label: 'OK', primary: true, onClick: () => {
      const list = [];
      if (a.opSel.value) list.push(fromDialog(a.opSel.value, a.inp.value));
      if (b.opSel.value) list.push(fromDialog(b.opSel.value, b.inp.value));
      if (list.length) res = { and: and.input.checked, list };
    } }, { label: 'Cancel' }] }).done.then(() => res);
  };
  FL.top10Dialog = function () {
    const tb = ui.select([['top', 'Top'], ['bottom', 'Bottom']], 'top');
    const n = ui.spin({ value: 10, min: 1, max: 500, step: 1, dec: 0 });
    const kind = ui.select([['items', 'Items'], ['percent', 'Percent']], 'items');
    let res = null;
    return ui.dialog({ title: 'Top 10 AutoFilter', width: 330, body: h('div', { class: 'col' }, h('div', { text: 'Show' }), h('div', { class: 'row', style: 'gap:8px' }, tb, n, kind)), buttons: [{ label: 'OK', primary: true, onClick: () => { res = { top: tb.value === 'top', percent: kind.value === 'percent', val: Math.round(n.get()) }; } }, { label: 'Cancel' }] }).done.then(() => res);
  };

  /* ------------------------------------------------------------ advanced filter */
  /**
   * Advanced Filter: list range with header row, criteria range (header + rows; rows OR, columns AND).
   * opts: {copyTo: {sheet, r, c} | null, unique: bool}
   */
  FL.advanced = function (s, listRg, critSheet, critRg, opts) {
    opts = opts || {};
    const heads = [];
    for (let c = listRg.c1; c <= listRg.c2; c++) heads.push(String(s.val(listRg.r1, c) == null ? '' : s.val(listRg.r1, c)).toLowerCase());
    let pass = () => true;
    if (critRg && C.dbSelect) {
      const ctx = C.ctx(s.wb, s, listRg.r1, listRg.c1);
      const res = C.dbSelect(ctx, new C.Ref(s, listRg.r1, listRg.c1, listRg.r2, listRg.c2), undefined, new C.Ref(critSheet, critRg.r1, critRg.c1, critRg.r2, critRg.c2));
      if (M.isErr(res)) throw new Error('The extract range has a missing or illegal field name.');
      const ok = new Set(res.rows);
      pass = (r) => ok.has(r);
    }
    const keep = [];
    const seen = new Set();
    for (let r = listRg.r1 + 1; r <= listRg.r2; r++) {
      if (!pass(r)) continue;
      if (opts.unique) { const k = []; for (let c = listRg.c1; c <= listRg.c2; c++) k.push(textOf(s, r, c).toLowerCase()); const key = k.join('\u0001'); if (seen.has(key)) continue; seen.add(key); }
      keep.push(r);
    }
    if (opts.copyTo) {
      const dst = opts.copyTo.sheet || s;
      O.tx(s.wb, 'Advanced Filter', () => {
        const rowsOut = [listRg.r1].concat(keep);
        rowsOut.forEach((r, k) => { for (let c = listRg.c1; c <= listRg.c2; c++) { const cl = s.get(r, c); const v = valOf(s, r, c); O.put(dst, opts.copyTo.r + k, opts.copyTo.c + c - listRg.c1, v == null && !(cl && cl.s) ? null : { v, s: cl ? cl.s : undefined }); } });
      });
      return keep.length;
    }
    O.structural(s, 'Advanced Filter', () => {
      const set = new Set(keep);
      for (let r = listRg.r1 + 1; r <= listRg.r2; r++) {
        if (set.has(r)) { const row = s.rows[r]; if (row && row.filtered) { delete row.hidden; delete row.filtered; } }
        else { const o = s.rowObj(r); o.hidden = true; o.filtered = true; }
      }
      s.filterMode = true;
      LY.invalidate(s);
    });
    return keep.length;
  };
  FL.filterAt = filterAt;
})(typeof window !== 'undefined' ? window : globalThis);
