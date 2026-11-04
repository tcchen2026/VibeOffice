/* Ledger — dialogs, part 1: Format Cells, rows / columns / cells / sheets, Find and Replace, Go To,
 * Paste Special, Series, Save As, Properties, Zoom, Symbol, comments, Options, Protection, Hyperlink, About. */
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
  /** a list box (common ui.listBox) with Ledger's names: select(v), value() (null when nothing is chosen), index(),
   *  ondbl; o.class keeps its height ('lbox short', 'lbox fnlist'…), o.translate 'no': its items are data */
  const list = (items, value, onPick, o) => {
    o = o || {};
    const el = ui.listBox(items, value, { cls: o.class || 'lbox', translate: o.translate, onChange: onPick ? (v, i) => onPick(v, i) : null, onDbl: (v, i) => { if (el.ondbl) el.ondbl(v, i); } });
    el.select = (v) => { el.set(v); return el.index(); };
    el.value = () => (el.index() >= 0 ? el.get() : null);
    return el;
  };
  D.list = list;

  /* ================================================================ Format Cells */
  /* the categories are also NF.category's names: English here, shown translated by list() */
  /* [category, label]: the categories are also NF.category's names; three need their own words in some languages */
  const CATS = [['General', L.N('General|number format')], L.N('Number'), L.N('Currency'), L.N('Accounting'), L.N('Date'), L.N('Time'), L.N('Percentage'), L.N('Fraction'), L.N('Scientific'), ['Text', L.N('Text|number format')], L.N('Special'), ['Custom', L.N('Custom|number format')]];
  const SYMBOLS = [['$', '"$"'], ['€', '"€"'], ['£', '"£"'], ['¥', '"¥"'], [L.N('None'), '']];
  const DATE_TYPES = ['m/d/yyyy', 'dddd, mmmm dd, yyyy', 'm/d', 'm/d/yy', 'mm/dd/yy', 'd-mmm', 'd-mmm-yy', 'dd-mmm-yy', 'mmm-yy', 'mmmm-yy', 'mmmm d, yyyy', 'm/d/yy h:mm AM/PM', 'm/d/yy h:mm', 'mmmmm', 'mmmmm-yy', 'd-mmm-yyyy'];
  const TIME_TYPES = ['h:mm:ss AM/PM', 'h:mm', 'h:mm AM/PM', 'h:mm:ss', 'mm:ss.0', '[h]:mm:ss', 'm/d/yy h:mm AM/PM', 'm/d/yy h:mm'];
  const FRACTIONS = [['# ?/?', 'Up to one digit (1/4)'], ['# ??/??', 'Up to two digits (21/25)'], ['# ???/???', 'Up to three digits (312/943)'], ['# ?/2', 'As halves (1/2)'], ['# ?/4', 'As quarters (2/4)'], ['# ?/8', 'As eighths (4/8)'], ['# ??/16', 'As sixteenths (8/16)'], ['# ?/10', 'As tenths (3/10)'], ['# ??/100', 'As hundredths (30/100)']];
  const SPECIAL = [['00000', 'Zip Code'], ['00000-0000', 'Zip Code + 4'], ['[<=9999999]###-####;(###) ###-####', 'Phone Number'], ['000-00-0000', 'Social Security Number']];
  function numCode(dec, sep, neg) {
    const base = (sep ? '#,##0' : '0') + (dec ? '.' + '0'.repeat(dec) : '');
    return [base, base + ';[Red]' + base, base + '_);\\(' + base + '\\)', base + '_);[Red]\\(' + base + '\\)'][neg || 0];
  }
  function curCode(dec, sym, neg) {
    const b = sym + '#,##0' + (dec ? '.' + '0'.repeat(dec) : '');
    return [b, b + ';[Red]' + b, b + '_);\\(' + b + '\\)', b + '_);[Red]\\(' + b + '\\)'][neg || 0];
  }
  function accCode(dec, sym) {
    const d = dec ? '.' + '0'.repeat(dec) : '';
    const s = sym ? sym + '* ' : '* ';
    return `_(${s}#,##0${d}_);_(${s}\\(#,##0${d}\\);_(${s}"-"${dec ? '?'.repeat(dec) : ''}_);_(@_)`;
  }
  /** what the current code looks like in terms of the dialog's options */
  function analyse(code) {
    const cat = NF.category(code);
    const o = { cat, dec: Math.min(30, NF.decimals(code)), sep: /#,##0/.test(code), neg: 0, sym: '"$"' };
    const sec = code.split(';');
    if (sec.length > 1) { const red = /\[Red\]/i.test(sec[1]), paren = /\\\(|\(/.test(sec[1]); o.neg = paren ? (red ? 3 : 2) : red ? 1 : 0; }
    const sm = /"([$€£¥])"/.exec(code);
    if (sm) o.sym = '"' + sm[1] + '"'; else if (cat === 'Accounting' && !/\$/.test(code)) o.sym = '';
    return o;
  }
  D.formatCells = function (tab, target) {
    const s = sh(), w = wb(), sel = G().sel();
    const cell = target ? null : s.get(sel.r, sel.c);
    const st = target ? target.style : LY.styleOf(s, sel.r, sel.c, cell);
    const v = target ? 1234.5 : cell ? (cell.dirty ? C.cellValue(s, sel.r, sel.c) : cell.v) : null;
    const changed = {};
    const out = { nf: st.nf || 'General', align: Object.assign({}, st.align || {}), font: Object.assign({}, st.font || {}), fill: st.fill ? JSON.parse(JSON.stringify(st.fill)) : null, prot: Object.assign({ locked: true }, st.prot || {}) };
    /* ---------- Number ---------- */
    const an = analyse(out.nf);
    const numBody = h('div', { class: 'fc-num' });
    const sample = h('div', { class: 'fc-sample', translate: 'no' });
    const opts = h('div', { class: 'fc-opts' });
    const desc = h('div', { class: 'fc-desc' });
    const showSample = () => { let t = ''; try { t = v == null ? '' : NF.format(out.nf, v, { date1904: w.date1904 }).text; } catch (e) { t = ''; } sample.textContent = t; };
    const setCode = (code) => { out.nf = code; changed.nf = true; showSample(); };
    let catState = an.cat;
    const catList = list(CATS, an.cat, (c) => { catState = c; drawOpts(true); });
    function drawOpts(fromUser) {
      opts.textContent = '';
      const c = catState;
      const decSpin = (d, fn) => ui.spin({ value: d, min: 0, max: 30, step: 1, dec: 0, onChange: (x) => fn(Math.round(x)) });
      if (c === 'General') { desc.textContent = L.t('General format cells have no specific number format.'); if (fromUser) setCode('General'); }
      else if (c === 'Number') {
        let dec = an.cat === 'Number' ? an.dec : 2, sep = an.cat === 'Number' ? an.sep : false, neg = an.cat === 'Number' ? an.neg : 0;
        const negList = list([0, 1, 2, 3].map((k) => [k, k]), neg, (k) => { neg = k; upd(); }, { class: 'lbox short' });
        const drawNeg = () => { Array.from(negList.children).forEach((r, k) => { const code = numCode(dec, sep, k); r.textContent = NF.text(code, -1234.1); r.style.color = /Red/.test(code) ? '#FF0000' : ''; }); };
        const upd = () => { setCode(numCode(dec, sep, neg)); drawNeg(); };
        opts.append(ui.field('&Decimal places:', decSpin(dec, (x) => { dec = x; upd(); })), ui.check('&Use 1000 Separator (,)', sep, (x) => { sep = x; upd(); }), h('label', { text: 'Negative numbers:' }), negList);
        drawNeg();
        desc.textContent = L.t('Number is used for general display of numbers. Currency and Accounting offer specialized formatting for monetary value.');
        if (fromUser) upd();
      } else if (c === 'Currency') {
        let dec = an.cat === 'Currency' ? an.dec : 2, sym = an.cat === 'Currency' ? an.sym : '"$"', neg = an.cat === 'Currency' ? an.neg : 3;
        const negList = list([0, 1, 2, 3].map((k) => [k, k]), neg, (k) => { neg = k; upd(); }, { class: 'lbox short' });
        const drawNeg = () => { Array.from(negList.children).forEach((r, k) => { const code = curCode(dec, sym, k); r.textContent = NF.text(code, -1234.1); r.style.color = /Red/.test(code) ? '#FF0000' : ''; }); };
        const upd = () => { setCode(curCode(dec, sym, neg)); drawNeg(); };
        opts.append(ui.field('&Decimal places:', decSpin(dec, (x) => { dec = x; upd(); })), ui.field('&Symbol:', ui.select(SYMBOLS.map(([l, cde]) => [cde, l]), sym, (x) => { sym = x; upd(); })), h('label', { text: 'Negative numbers:' }), negList);
        drawNeg();
        desc.textContent = L.t('Currency formats are used for general monetary values. Use Accounting formats to align decimal points in a column.');
        if (fromUser) upd();
      } else if (c === 'Accounting') {
        let dec = an.cat === 'Accounting' ? an.dec : 2, sym = an.cat === 'Accounting' ? an.sym : '"$"';
        const upd = () => setCode(accCode(dec, sym));
        opts.append(ui.field('&Decimal places:', decSpin(dec, (x) => { dec = x; upd(); })), ui.field('&Symbol:', ui.select(SYMBOLS.map(([l, cde]) => [cde, l]), sym, (x) => { sym = x; upd(); })));
        desc.textContent = L.t('Accounting formats line up the currency symbols and decimal points in a column.');
        if (fromUser) upd();
      } else if (c === 'Date' || c === 'Time') {
        const types = c === 'Date' ? DATE_TYPES : TIME_TYPES;
        const sv = typeof v === 'number' ? v : 37329.5649;
        const lb = list(types.map((t) => [t, (t === 'm/d/yyyy' || t === 'h:mm:ss AM/PM' ? '*' : '') + NF.text(t, sv, { date1904: w.date1904 })]), types.includes(out.nf) ? out.nf : null, (t) => setCode(t), { translate: 'no' });
        opts.append(h('label', { text: 'Type:' }), lb, ui.field('&Locale (location):', ui.select([['en-US', 'English (United States)']], 'en-US')));
        desc.textContent = c === 'Date' ? L.t('Date formats display date and time serial numbers as date values. Formats that begin with an asterisk (*) respond to changes in regional date and time settings.') : L.t('Time formats display date and time serial numbers as time values.');
        if (fromUser && !types.includes(out.nf)) { setCode(types[0]); lb.select(types[0]); }
      } else if (c === 'Percentage' || c === 'Scientific') {
        let dec = an.cat === c ? an.dec : 2;
        const upd = () => setCode(c === 'Percentage' ? '0' + (dec ? '.' + '0'.repeat(dec) : '') + '%' : '0' + (dec ? '.' + '0'.repeat(dec) : '') + 'E+00');
        opts.append(ui.field('&Decimal places:', decSpin(dec, (x) => { dec = x; upd(); })));
        desc.textContent = c === 'Percentage' ? L.t('Percentage formats multiply the cell value by 100 and display the result with a percent symbol.') : '';
        if (fromUser) upd();
      } else if (c === 'Fraction') {
        const lb = list(FRACTIONS, FRACTIONS.some((f) => f[0] === out.nf) ? out.nf : null, (t) => setCode(t));
        opts.append(h('label', { text: 'Type:' }), lb);
        if (fromUser && !FRACTIONS.some((f) => f[0] === out.nf)) { setCode(FRACTIONS[0][0]); lb.select(FRACTIONS[0][0]); }
      } else if (c === 'Text') { desc.textContent = L.t('Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered.'); if (fromUser) setCode('@'); }
      else if (c === 'Special') {
        const lb = list(SPECIAL, SPECIAL.some((f) => f[0] === out.nf) ? out.nf : null, (t) => setCode(t));
        opts.append(h('label', { text: 'Type:' }), lb, ui.field('&Locale (location):', ui.select([['en-US', 'English (United States)']], 'en-US')));
        desc.textContent = L.t('Special formats are useful for tracking list and database values.');
        if (fromUser && !SPECIAL.some((f) => f[0] === out.nf)) { setCode(SPECIAL[0][0]); lb.select(SPECIAL[0][0]); }
      } else {
        const used = new Set();
        for (const x of w.styles.list) if (x.nf && x.nf !== 'General') used.add(x.nf);
        const codes = ['General', '0', '0.00', '#,##0', '#,##0.00', '#,##0_);(#,##0)', '#,##0_);[Red](#,##0)', '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)', '"$"#,##0_);\\("$"#,##0\\)', '"$"#,##0_);[Red]\\("$"#,##0\\)', '"$"#,##0.00_);\\("$"#,##0.00\\)', '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)', '0%', '0.00%', '0.00E+00', '##0.0E+0', '# ?/?', '# ??/??', 'm/d/yyyy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'm/d/yyyy h:mm', 'mm:ss', 'mm:ss.0', '@', '[h]:mm:ss', NF.BUILTIN[41], NF.BUILTIN[42], NF.BUILTIN[43], NF.BUILTIN[44]];
        for (const u of used) if (!codes.includes(u)) codes.push(u);
        const inp = h('input', { type: 'text', value: out.nf, style: 'width:100%' });
        const lb = list(codes, codes.includes(out.nf) ? out.nf : null, (t) => { inp.value = t; setCode(t); }, { translate: 'no' });
        inp.addEventListener('input', () => { try { NF.parse(inp.value); setCode(inp.value || 'General'); } catch (e) { /* keep */ } });
        const del = ui.button('&Delete', () => {
          const code = lb.value();
          if (!code || NF.builtinId(code) >= 0 || code === 'General') return;
          /* cells using the deleted format return to General */
          O.tx(w, 'Delete Number Format', () => { for (const s2 of w.sheets) s2.rows.forEach((row, r) => { if (row) row.cells.forEach((cl, c) => { if (cl && cl.s && w.styles.get(cl.s).nf === code) { const n = Object.assign({}, cl); n.s = w.styles.derive(cl.s, { nf: 'General' }); O.put(s2, r, c, n); } }); }); });
          drawOpts(false);
        });
        opts.append(ui.field('&Type:', inp), lb, h('div', { class: 'row' }, del));
        desc.textContent = L.t('Type the number format code, using one of the existing codes as a starting point.');
        if (fromUser && !changed.nf) setCode(out.nf);
      }
    }
    numBody.append(h('div', { class: 'fc-cats' }, h('label', { text: 'Category:' }), catList), h('div', { class: 'fc-right' }, ui.group('Sample', sample), opts, desc));
    drawOpts(false);
    showSample();
    /* ---------- Alignment ---------- */
    const al = out.align;
    const hSel = ui.select([['general', 'General'], ['left', 'Left (Indent)'], ['center', 'Center'], ['right', 'Right (Indent)'], ['fill', 'Fill'], ['justify', 'Justify'], ['centerContinuous', 'Center Across Selection'], ['distributed', 'Distributed (Indent)']], al.h || 'general', (x) => { al.h = x === 'general' ? undefined : x; changed.align = true; indent.setDisabled(!/left|right|distributed/.test(x)); });
    const vSel = ui.select([['top', 'Top'], ['center', 'Center'], ['bottom', 'Bottom'], ['justify', 'Justify'], ['distributed', 'Distributed']], al.v || 'bottom', (x) => { al.v = x === 'bottom' ? undefined : x; changed.align = true; });
    const indent = ui.spin({ value: al.indent || 0, min: 0, max: 15, step: 1, dec: 0, onChange: (x) => { al.indent = Math.round(x) || undefined; changed.align = true; } });
    indent.setDisabled(!/left|right|distributed/.test(al.h || ''));
    const rot0 = al.rot === 255 ? 0 : al.rot > 90 ? 90 - al.rot : al.rot || 0;
    const rotSpin = ui.spin({ value: rot0, min: -90, max: 90, step: 1, dec: 0, fmt: (x) => L.t('{0} Degrees', Math.round(x)), onChange: (x) => { const d = Math.round(x); al.rot = d < 0 ? 90 - d : d || undefined; vert.classList.remove('on'); changed.align = true; drawDial(); } });
    const vert = h('button', { type: 'button', class: 'fc-vert' + (al.rot === 255 ? ' on' : ''), html: Array.from(L.t('Text')).map(L.esc).join('<br>') });
    vert.addEventListener('click', () => { const on = !vert.classList.contains('on'); vert.classList.toggle('on', on); al.rot = on ? 255 : undefined; changed.align = true; drawDial(); });
    const dial = h('div', { class: 'fc-dial' });
    const drawDial = () => {
      const deg = al.rot === 255 ? 0 : al.rot > 90 ? 90 - al.rot : al.rot || 0;
      dial.innerHTML = `<svg width="110" height="130" viewBox="0 0 110 130">${Array.from({ length: 13 }, (_, i) => { const a = (-90 + i * 15) * Math.PI / 180; return `<circle cx="${20 + 70 * Math.cos(a)}" cy="${65 - 70 * Math.sin(a)}" r="${i % 6 === 0 ? 2.5 : 1.5}" fill="${i % 6 === 0 ? '#000' : '#888'}"/>`; }).join('')}<line x1="20" y1="65" x2="${20 + 64 * Math.cos(deg * Math.PI / 180)}" y2="${65 - 64 * Math.sin(deg * Math.PI / 180)}" stroke="#c00" stroke-width="1.5"/><text x="26" y="69" font-size="11" transform="rotate(${-deg} 20 65)" fill="${al.rot === 255 ? '#aaa' : '#000'}">${L.esc(L.t('Text'))} —</text></svg>`;
    };
    dial.addEventListener('click', (e) => { const r = dial.getBoundingClientRect(); const x = e.clientX - r.left - 20, y = 65 - (e.clientY - r.top); let d = Math.round(Math.atan2(y, Math.max(0, x)) * 180 / Math.PI); d = Math.max(-90, Math.min(90, d)); rotSpin.set(d); al.rot = d < 0 ? 90 - d : d || undefined; vert.classList.remove('on'); changed.align = true; drawDial(); });
    drawDial();
    const merged = !!LY.mergeAt(s, sel.r, sel.c);
    let mergeState = null;
    const alBody = h('div', { class: 'fc-align' },
      h('div', { class: 'col', style: 'flex:1' },
        ui.group('Text alignment', h('div', { class: 'row', style: 'gap:8px;align-items:flex-end' }, ui.field('&Horizontal:', hSel), ui.field('&Indent:', indent)), ui.field('&Vertical:', vSel), ui.check('&Justify distributed', !!al.justLast, (x) => { al.justLast = x || undefined; changed.align = true; })),
        ui.group('Text control', ui.check('&Wrap text', !!al.wrap, (x) => { al.wrap = x || undefined; changed.align = true; }), ui.check('Shrin&k to fit', !!al.shrink, (x) => { al.shrink = x || undefined; changed.align = true; }), ui.check('&Merge cells', merged, (x) => { mergeState = x; })),
        ui.group('Right-to-left', ui.field('&Text direction:', ui.select([['ctx', 'Context'], ['ltr', 'Left-to-Right'], ['rtl', 'Right-to-Left']], al.rtl ? 'rtl' : al.ltr ? 'ltr' : 'ctx', (x) => { al.rtl = x === 'rtl' || undefined; al.ltr = x === 'ltr' || undefined; changed.align = true; })))),
      ui.group('Orientation', h('div', { class: 'row' }, vert, dial), ui.field('', rotSpin)));
    /* ---------- Font ---------- */
    const f = out.font;
    /* the three lists with a box above each, as in Excel (common ui.listBox): picking or typing changes the font */
    const fontList = ui.listBox(L.FONT_LIST, LY.fontName(w, f), { input: true, height: 96, id: 'fc-font', translate: 'no', onChange: (x) => { if (!String(x).trim()) return; f.name = String(x); delete f.scheme; changed.font = true; prev(); } });
    const styles = [['r', L.N('Regular')], ['i', L.N('Italic')], ['b', L.N('Bold')], ['bi', L.N('Bold Italic')]];
    const curStyle = (f.b ? 'b' : '') + (f.i ? 'i' : '') || 'r';
    const styleList = ui.listBox(styles, curStyle, { input: true, height: 96, id: 'fc-style', onChange: (x) => { if (!/^(r|i|b|bi)$/.test(x)) return; f.b = /b/.test(x) || undefined; f.i = /i/.test(x) || undefined; changed.font = true; prev(); } });
    const sizeList = ui.listBox(L.SIZE_LIST.map(String), String(f.sz || 10), { input: true, height: 96, id: 'fc-size', translate: 'no', onChange: (x) => { const n = num(x, null); if (n >= 1 && n <= 409) { f.sz = n; changed.font = true; prev(); } } });
    const uSel = ui.select([['none', 'None'], ['single', 'Single'], ['double', 'Double'], ['singleAccounting', 'Single Accounting'], ['doubleAccounting', 'Double Accounting']], f.u || 'none', (x) => { f.u = x === 'none' ? undefined : x; changed.font = true; prev(); });
    const fcol = ui.colorButton(f.color && !f.color.auto ? M.colorHex(w, f.color, '#000000') : 'auto', { mode: 'font', grid: true, onChange: (c) => { f.color = c === 'auto' ? undefined : M.rgb(c); changed.font = true; prev(); } });
    const normal = ui.check('&Normal font', false, (x) => { if (x) { Object.keys(f).forEach((k) => delete f[k]); Object.assign(f, JSON.parse(JSON.stringify(w.styles.get(0).font))); fontList.set(LY.fontName(w, f)); sizeList.set(String(f.sz)); styleList.set((f.b ? 'b' : '') + (f.i ? 'i' : '') || 'r'); changed.font = true; prev(); } });
    const fprev = h('div', { class: 'fc-fprev', text: 'AaBbCcYyZz' });
    function prev() {
      fprev.style.font = LY.cssFont(Object.assign({}, f, { name: LY.fontName(w, f) }), 1.3);
      fprev.style.color = M.colorHex(w, f.color, '#000000');
      fprev.style.textDecoration = [f.u ? 'underline' : '', f.strike ? 'line-through' : ''].join(' ').trim() || 'none';
      fprev.style.textDecorationStyle = /double/.test(f.u || '') ? 'double' : 'solid';
      fprev.style.verticalAlign = f.vert === 'superscript' ? 'super' : f.vert === 'subscript' ? 'sub' : '';
    }
    prev();
    const fontBody = h('div', { class: 'fc-font' },
      h('div', { class: 'fc-f3' }, h('div', { class: 'col' }, h('label', { for: 'fc-font', html: ui.ampHTML('&Font:') }), fontList), h('div', { class: 'col' }, h('label', { for: 'fc-style', html: ui.ampHTML('Font st&yle:') }), styleList), h('div', { class: 'col' }, h('label', { for: 'fc-size', html: ui.ampHTML('&Size:') }), sizeList)),
      h('div', { class: 'row', style: 'gap:12px;align-items:flex-end' }, ui.field('&Underline:', uSel), ui.field('&Color:', fcol), normal),
      h('div', { class: 'row', style: 'gap:12px;align-items:stretch' },
        ui.group('Effects', ui.check('Stri&kethrough', !!f.strike, (x) => { f.strike = x || undefined; changed.font = true; prev(); }), ui.check('Sup&erscript', f.vert === 'superscript', (x) => { f.vert = x ? 'superscript' : undefined; changed.font = true; prev(); }), ui.check('Su&bscript', f.vert === 'subscript', (x) => { f.vert = x ? 'subscript' : undefined; changed.font = true; prev(); })),
        ui.group('Preview', fprev)));
    /* ---------- Border ---------- */
    const rg = G().range();
    const multiR = rg.r1 !== rg.r2, multiC = rg.c1 !== rg.c2;
    const b0 = st.border || {};
    const edges = { top: b0.t || null, bottom: b0.b || null, left: b0.l || null, right: b0.r || null, insideH: null, insideV: null, diagUp: b0.du ? b0.d : null, diagDown: b0.dd ? b0.d : null };
    if (multiR || multiC) { const right = LY.styleOf(s, rg.r1, rg.c2, s.get(rg.r1, rg.c2)).border || {}, bottom = LY.styleOf(s, rg.r2, rg.c1, s.get(rg.r2, rg.c1)).border || {}; edges.right = right.r || null; edges.bottom = bottom.b || null; if (multiR) edges.insideH = b0.b || null; if (multiC) edges.insideV = b0.r || null; }
    const touched = {};
    let line = { style: 'thin' };
    const LINES = ['hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double'];
    const lineBox = h('div', { class: 'fc-lines' });
    const lineSvg = (k) => { const w2 = 52, y = 7; const dash = { hair: '1 1', dotted: '2 2', dashDotDot: '6 2 2 2 2 2', dashDot: '6 2 2 2', dashed: '4 2', mediumDashDotDot: '6 2 2 2 2 2', slantDashDot: '8 2 3 2', mediumDashDot: '6 2 2 2', mediumDashed: '6 3' }[k] || ''; const sw = /medium|slant/.test(k) ? 2 : k === 'thick' ? 3 : 1; return k === 'double' ? `<svg width="${w2}" height="14"><line x1="2" y1="5" x2="${w2 - 2}" y2="5" stroke="#000"/><line x1="2" y1="9" x2="${w2 - 2}" y2="9" stroke="#000"/></svg>` : `<svg width="${w2}" height="14"><line x1="2" y1="${y}" x2="${w2 - 2}" y2="${y}" stroke="#000" stroke-width="${sw}" ${dash ? `stroke-dasharray="${dash}"` : ''}/></svg>`; };
    const LINE_NAMES = { none: L.N('None'), hair: L.N('Hairline'), dotted: L.N('Dotted'), dashDotDot: 'Dash dot dot', dashDot: 'Dash dot', dashed: L.N('Dashed'), thin: L.N('Thin'), mediumDashDotDot: 'Medium dash dot dot', slantDashDot: 'Slanted dash dot', mediumDashDot: 'Medium dash dot', mediumDashed: 'Medium dashed', medium: L.N('Medium'), thick: L.N('Thick'), double: L.N('Double') };
    for (const k of ['none'].concat(LINES)) { const b = h('button', { type: 'button', class: 'fc-line' + (k === 'thin' ? ' on' : ''), html: k === 'none' ? `<span style="font-size:10px">${L.esc(L.t('None'))}</span>` : lineSvg(k), 'aria-label': LINE_NAMES[k] }); b.addEventListener('click', () => { L.$$('.fc-line', lineBox).forEach((x) => x.classList.remove('on')); b.classList.add('on'); line = k === 'none' ? null : Object.assign({}, line || {}, { style: k }); }); lineBox.appendChild(b); }
    let lineColor = null;
    const lcol = ui.colorButton('auto', { mode: 'font', grid: true, onChange: (c) => { lineColor = c === 'auto' ? null : M.rgb(c); } });
    const prevBox = h('div', { class: 'fc-bprev' });
    const curLine = () => (line ? Object.assign({ style: line.style }, lineColor ? { color: lineColor } : {}) : null);
    const toggle = (k) => { const cur = edges[k]; const ln = curLine(); edges[k] = cur && ln && cur.style === ln.style && JSON.stringify(cur.color || null) === JSON.stringify(ln.color || null) ? null : ln; touched[k] = true; drawB(); };
    function drawB() {
      const W2 = 150, H2 = 110, x0 = 15, y0 = 10, x1 = W2 - 15, y1 = H2 - 10, mx = (x0 + x1) / 2, my = (y0 + y1) / 2, word = L.esc(L.t('Text'));
      const ln = (e, xa, ya, xb, yb) => { if (!e) return ''; const c = M.colorHex(w, e.color, '#000000'); const sw = /medium|slant/.test(e.style) ? 2 : e.style === 'thick' ? 3 : 1; const dash = /dash|dot|hair/i.test(e.style) ? 'stroke-dasharray="3 2"' : ''; return e.style === 'double' ? `<line x1="${xa}" y1="${ya}" x2="${xb}" y2="${yb}" stroke="${c}" stroke-width="3"/><line x1="${xa}" y1="${ya}" x2="${xb}" y2="${yb}" stroke="#fff" stroke-width="1"/>` : `<line x1="${xa}" y1="${ya}" x2="${xb}" y2="${yb}" stroke="${c}" stroke-width="${sw}" ${dash}/>`; };
      prevBox.innerHTML = `<svg width="${W2}" height="${H2}" viewBox="0 0 ${W2} ${H2}"><rect x="${x0}" y="${y0}" width="${x1 - x0}" height="${y1 - y0}" fill="#fff"/>` +
        ['M' + x0 + ',' + (y0 - 6) + 'v4', 'M' + (x0 - 6) + ',' + y0 + 'h4', 'M' + x1 + ',' + (y0 - 6) + 'v4', 'M' + (x1 + 2) + ',' + y0 + 'h4', 'M' + x0 + ',' + (y1 + 2) + 'v4', 'M' + (x0 - 6) + ',' + y1 + 'h4', 'M' + x1 + ',' + (y1 + 2) + 'v4', 'M' + (x1 + 2) + ',' + y1 + 'h4'].map((d) => `<path d="${d}" stroke="#888"/>`).join('') +
        `<text x="${multiC ? (x0 + mx) / 2 : mx}" y="${multiR ? (y0 + my) / 2 + 4 : my + 4}" font-size="11" text-anchor="middle" fill="#666">${word}</text>` + (multiC ? `<text x="${(mx + x1) / 2}" y="${multiR ? (y0 + my) / 2 + 4 : my + 4}" font-size="11" text-anchor="middle" fill="#666">${word}</text>` : '') + (multiR ? `<text x="${multiC ? (x0 + mx) / 2 : mx}" y="${(my + y1) / 2 + 4}" font-size="11" text-anchor="middle" fill="#666">${word}</text>` : '') +
        ln(edges.top, x0, y0, x1, y0) + ln(edges.bottom, x0, y1, x1, y1) + ln(edges.left, x0, y0, x0, y1) + ln(edges.right, x1, y0, x1, y1) + (multiR ? ln(edges.insideH, x0, my, x1, my) : '') + (multiC ? ln(edges.insideV, mx, y0, mx, y1) : '') + ln(edges.diagUp, x0, y1, x1, y0) + ln(edges.diagDown, x0, y0, x1, y1) + '</svg>';
    }
    prevBox.addEventListener('click', (e) => {
      const r = prevBox.getBoundingClientRect(); const x = e.clientX - r.left, y = e.clientY - r.top;
      const W2 = 150, H2 = 110;
      const d = { top: Math.abs(y - 10), bottom: Math.abs(y - (H2 - 10)), left: Math.abs(x - 15), right: Math.abs(x - (W2 - 15)) };
      if (multiR) d.insideH = Math.abs(y - H2 / 2); if (multiC) d.insideV = Math.abs(x - W2 / 2);
      const k = Object.keys(d).sort((a, b) => d[a] - d[b])[0];
      if (d[k] < 12) toggle(k);
    });
    drawB();
    const bbtn = (k, icon, tip, en) => { const b = h('button', { type: 'button', class: 'fc-bb', html: L.icons.get(icon), 'data-tip': tip, 'aria-label': tip }); b.disabled = en === false; b.addEventListener('click', () => toggle(k)); return b; };
    const preset = (label, icon, fn, en) => { const b = h('button', { type: 'button', class: 'fc-preset', html: L.icons.get(icon) + `<span>${L.esc(L.t(label))}</span>` }); b.disabled = en === false; b.addEventListener('click', fn); return b; };
    const borderBody = h('div', { class: 'fc-border' },
      h('div', { class: 'col' }, h('label', { text: 'Presets' }), h('div', { class: 'row', style: 'gap:16px' },
        preset(L.N('None'), 'bordersNone', () => { for (const k in edges) { edges[k] = null; touched[k] = true; } drawB(); }),
        preset(L.N('Outline|border'), 'bordersOut', () => { for (const k of ['top', 'bottom', 'left', 'right']) { edges[k] = curLine(); touched[k] = true; } drawB(); }),
        preset(L.N('Inside'), 'bordersInside', () => { if (multiR) { edges.insideH = curLine(); touched.insideH = true; } if (multiC) { edges.insideV = curLine(); touched.insideV = true; } drawB(); }, multiR || multiC)),
      h('label', { text: 'Border' }),
      h('div', { class: 'fc-bgrid' }, h('div', { class: 'col' }, bbtn('top', 'bordersTop', L.N('Top')), bbtn('insideH', 'bordersInsideH', 'Inside horizontal', multiR), bbtn('bottom', 'bordersBottom', L.N('Bottom'))), prevBox),
      h('div', { class: 'row', style: 'gap:4px;margin-left:34px' }, bbtn('diagUp', 'bordersDiagUp', 'Diagonal up'), bbtn('left', 'bordersLeft', L.N('Left')), bbtn('insideV', 'bordersInsideV', 'Inside vertical', multiC), bbtn('right', 'bordersRight', L.N('Right')), bbtn('diagDown', 'bordersDiagDown', 'Diagonal down')),
      h('div', { class: 'hint', text: 'The selected border style can be applied by clicking the presets, preview diagram or the buttons above.' })),
      ui.group('Line', h('label', { text: 'Style:' }), lineBox, ui.field('&Color:', lcol)));
    /* ---------- Patterns ---------- */
    const fill = out.fill;
    let fillColor = fill && fill.pattern === 'solid' ? M.colorHex(w, fill.fg) : fill && fill.bg ? M.colorHex(w, fill.bg) : null;
    let pattern = fill && fill.pattern && fill.pattern !== 'solid' ? fill.pattern : null;
    let patColor = fill && fill.pattern && fill.pattern !== 'solid' ? M.colorHex(w, fill.fg, '#000000') : '#000000';
    const sampleP = h('div', { class: 'fc-psample' });
    const showP = () => { sampleP.style.background = fillColor || '#fff'; sampleP.style.backgroundImage = pattern ? `url("${patternURL(pattern, patColor, fillColor || '#FFFFFF')}")` : 'none'; };
    const pal = h('div', { class: 'fc-pal' });
    pal.appendChild(h('button', { type: 'button', class: 'cp-wide', text: 'No Color', onclick: () => { fillColor = null; changed.fill = true; showP(); } }));
    const grid = h('div', { class: 'cp-grid' });
    for (const c of L.color.STANDARD) grid.appendChild(h('button', { type: 'button', class: 'cp-sw', style: `background:${c}`, 'aria-label': ui.colorName(c), 'data-tip': ui.colorName(c), onclick: () => { fillColor = c; changed.fill = true; showP(); } }));
    pal.appendChild(grid);
    const PATTERNS = ['solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis'];
    /* each pattern's name, as Excel's tips name them */
    const PAT_NAMES = { solid: L.N('Solid|pattern'), darkGray: L.N('75% Gray'), mediumGray: L.N('50% Gray'), lightGray: L.N('25% Gray'), gray125: L.N('12.5% Gray'), gray0625: L.N('6.25% Gray'), darkHorizontal: 'Horizontal Stripe', darkVertical: 'Vertical Stripe', darkDown: 'Reverse Diagonal Stripe', darkUp: 'Diagonal Stripe', darkGrid: 'Diagonal Crosshatch', darkTrellis: 'Thick Diagonal Crosshatch',
      lightHorizontal: 'Thin Horizontal Stripe', lightVertical: 'Thin Vertical Stripe', lightDown: 'Thin Reverse Diagonal Stripe', lightUp: 'Thin Diagonal Stripe', lightGrid: 'Thin Horizontal Crosshatch', lightTrellis: 'Thin Diagonal Crosshatch' };
    const patBtn = h('button', { type: 'button', class: 'swatch-btn wide', text: 'Pattern' }, h('span', { class: 'dd-arrow' }));
    patBtn.addEventListener('click', () => {
      const r = patBtn.getBoundingClientRect();
      ui.openMenu([{ custom: (close) => {
        const g2 = h('div', { class: 'pat-grid' });
        for (const p of PATTERNS) { const b = h('button', { type: 'button', style: `background:#fff url("${patternURL(p, '#000', '#fff')}")`, 'aria-label': PAT_NAMES[p], 'data-tip': PAT_NAMES[p], 'data-pattern': p }); b.addEventListener('click', () => { close(); pattern = p === 'solid' ? null : p; changed.fill = true; showP(); }); g2.appendChild(b); }
        const g3 = h('div', { class: 'cp-grid' });
        for (const c of L.color.STANDARD) g3.appendChild(h('button', { type: 'button', class: 'cp-sw', style: `background:${c}`, 'aria-label': ui.colorName(c), 'data-tip': ui.colorName(c), onclick: () => { close(); patColor = c; changed.fill = true; showP(); } }));
        return h('div', null, g2, h('div', { class: 'cp-sep' }), h('div', { class: 'tp-note', text: 'Pattern color:' }), g3);
      } }], { left: r.left, bottom: r.bottom });
    });
    showP();
    const patBody = h('div', { class: 'fc-pat' }, ui.group('Cell shading', h('label', { text: 'Color:' }), pal, h('div', { class: 'row', style: 'gap:8px;align-items:center' }, h('span', { text: 'Pattern:' }), patBtn)), ui.group('Sample', sampleP));
    /* ---------- Protection ---------- */
    const pr = out.prot;
    const protBody = h('div', { class: 'col' }, ui.check('&Locked', pr.locked !== false, (x) => { pr.locked = x; changed.prot = true; }), ui.check('H&idden', !!pr.hidden, (x) => { pr.hidden = x || undefined; changed.prot = true; }),
      h('p', { class: 'hint', text: 'Locking cells or hiding formulas has no effect unless the worksheet is protected. To protect the worksheet, choose Protection from the Tools menu, and then choose Protect Sheet. A password is optional.' }));
    const tabs = ui.tabs([{ label: 'Number|Format Cells tab', body: numBody }, { label: 'Alignment', body: alBody }, { label: 'Font', body: fontBody }, { label: 'Border', body: borderBody }, { label: 'Patterns', body: patBody }, { label: 'Protection', body: protBody }], tab || 0);
    ui.dialog({
      title: target && target.title ? target.title : 'Format Cells', body: tabs, width: 520,
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (target) {
          /* modifying a named style: build the whole style */
          const d = { nf: out.nf, align: Object.keys(al).some((k) => al[k] !== undefined) ? Object.assign({}, al) : null, font: Object.assign({}, f), fill: fillColor || pattern ? (pattern ? { pattern, fg: M.rgb(patColor), bg: fillColor ? M.rgb(fillColor) : M.rgb('#FFFFFF') } : { pattern: 'solid', fg: M.rgb(fillColor), bg: { indexed: 64 } }) : null, prot: pr.locked === false || pr.hidden ? { locked: pr.locked === false ? false : undefined, hidden: pr.hidden || undefined } : null };
          const b = {};
          if (edges.top) b.t = edges.top; if (edges.bottom) b.b = edges.bottom; if (edges.left) b.l = edges.left; if (edges.right) b.r = edges.right;
          if (edges.diagUp || edges.diagDown) { b.d = edges.diagUp || edges.diagDown; if (edges.diagUp) b.du = true; if (edges.diagDown) b.dd = true; }
          d.border = Object.keys(b).length ? b : null;
          const nst = M.mergeStyle(st, d);
          nst.border = d.border; nst.align = d.align; nst.fill = d.fill; nst.prot = d.prot;
          for (const k of Object.keys(nst.font || {})) if (nst.font[k] === undefined) delete nst.font[k];
          delete nst.xs;
          target.apply(nst);
          return;
        }
        tryRun(() => {
          O.tx(w, 'Format Cells', () => {
            const delta = {};
            if (changed.nf) delta.nf = out.nf;
            if (changed.align) { const a = {}; for (const k of ['h', 'v', 'wrap', 'indent', 'rot', 'shrink', 'rtl', 'ltr', 'justLast']) a[k] = al[k]; delta.align = a; }
            if (changed.font) { const fo = Object.assign({}, f); for (const k of ['b', 'i', 'u', 'strike', 'vert', 'color']) if (fo[k] === undefined) fo[k] = undefined; delta.font = fo; }
            if (changed.fill) delta.fill = fillColor || pattern ? (pattern ? { pattern, fg: M.rgb(patColor), bg: fillColor ? M.rgb(fillColor) : M.rgb('#FFFFFF') } : { pattern: 'solid', fg: M.rgb(fillColor), bg: { indexed: 64 } }) : null;
            if (changed.prot) delta.prot = { locked: pr.locked === false ? false : undefined, hidden: pr.hidden || undefined };
            if (Object.keys(delta).length) O.format(s, G().ranges(), (stl) => {
              const d2 = Object.assign({}, delta);
              if (d2.font) d2.font = Object.assign({}, d2.font);
              if (d2.align) { d2.align = Object.assign({}, d2.align); }
              void stl;
              return d2;
            }, 'Format Cells');
            const spec = {};
            for (const k in touched) spec[k] = edges[k];
            if (Object.keys(spec).length) O.borders(s, G().ranges(), spec, 'Format Cells');
            if (mergeState === true && !merged) O.merge(s, G().ranges(), {});
            if (mergeState === false && merged) O.unmerge(s, G().ranges());
            if (changed.align && (al.wrap || al.rot) || changed.font) for (const r of A().selRows()) O.autoRow(s, r);
          });
        });
      } }, { label: 'Cancel' }],
    });
  };
  /** a small tile showing a fill pattern (solid: the pattern colour) */
  function patternURL(p, fg, bg) {
    if (!L.render.CELL_PATTERNS[p]) return `data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4"><rect width="4" height="4" fill="${fg}"/></svg>`)}`;
    return L.render.patternURL(p, fg, bg);
  }
  D.patternURL = patternURL;

  /* ================================================================ rows, columns, cells */
  const sizeDialog = (title, label, value, min, max, apply) => {
    const inp = h('input', { type: 'text', value: String(value), style: 'width:80px' });
    ui.dialog({ title, width: 260, body: ui.field(label, inp), buttons: [{ label: 'OK', primary: true, onClick: () => { const v = num(inp.value, NaN); if (!(v >= min && v <= max)) { ui.msg(L.t('Enter a number between {0} and {1}.', min, max), { icon: 'warn' }); return false; } tryRun(() => apply(v)); } }, { label: 'Cancel' }] });
  };
  D.rowHeight = function () {
    const s = sh(), rows = A().selRows();
    const pts = rows.map((r) => M.rowPt(s, r));
    sizeDialog('Row Height', 'Row &height:', pts.every((p) => p === pts[0]) ? pts[0] : '', 0, 409, (v) => O.setRowHeight(s, rows, v));
  };
  D.colWidth = function () {
    const s = sh(), cols = A().selCols();
    const mdw = M.mdw(s.wb);
    const ch = cols.map((c) => M.pxToChars(M.colPx(s, c), mdw));
    sizeDialog('Column Width', '&Column width:', ch.every((x) => x === ch[0]) ? ch[0] : '', 0, 255, (v) => O.setColWidth(s, cols, v === 0 ? 0 : M.pxToWidth(M.charsToPx(v, mdw), mdw)));
  };
  D.standardWidth = function () {
    const s = sh(), mdw = M.mdw(s.wb);
    sizeDialog('Standard Width', '&Standard column width:', M.pxToChars(M.defaultColPx(s), mdw), 0, 255, (v) => {
      const old = { d: s.defColW, b: s.baseColW };
      O.tx(s.wb, 'Standard Width', () => { s.defColW = M.pxToWidth(M.charsToPx(v, mdw), mdw); s.baseColW = Math.round(v); LY.invalidate(s); s.wb.undo.op(() => { s.defColW = old.d; s.baseColW = old.b; LY.invalidate(s); }, () => { s.defColW = M.pxToWidth(M.charsToPx(v, mdw), mdw); s.baseColW = Math.round(v); LY.invalidate(s); }); });
    });
  };
  const shiftDialog = (title, items, apply, def) => {
    let pick = def || items[0][0];
    const radios = items.map(([k, label]) => ui.radio('shift', label, k === pick, () => { pick = k; }));
    ui.dialog({ title, width: 240, body: ui.group(title.replace(/\.\.\.$/, ''), ...radios), buttons: [{ label: 'OK', primary: true, onClick: () => tryRun(() => apply(pick)) }, { label: 'Cancel' }] });
  };
  D.insertCells = function () {
    if (G().guardProtect('insertRows')) return;
    shiftDialog(L.N('Insert'), [['right', 'Shift cells r&ight'], ['down', 'Shift cells &down'], ['row', 'Entire &row'], ['col', 'Entire &column']], (k) => {
      if (k === 'row') A().insertLines('r'); else if (k === 'col') A().insertLines('c');
      else O.shiftCells(sh(), O.clip(sh(), G().range()), k, true);
    }, 'down');
  };
  D.deleteCells = function () {
    if (G().guardProtect('deleteRows')) return;
    shiftDialog(L.N('Delete'), [['left', 'Shift cells &left'], ['up', 'Shift cells &up'], ['row', 'Entire &row'], ['col', 'Entire &column']], (k) => {
      if (k === 'row') A().deleteLines('r'); else if (k === 'col') A().deleteLines('c');
      else O.shiftCells(sh(), O.clip(sh(), G().range()), k, false);
    });
  };
  D.groupRowsCols = function (d) {
    return new Promise((res) => {
      let pick = 'r';
      ui.dialog({ title: d > 0 ? 'Group' : 'Ungroup', width: 220, body: ui.group(d > 0 ? 'Group' : 'Ungroup', ui.radio('grp', '&Rows', true, () => { pick = 'r'; }), ui.radio('grp', '&Columns', false, () => { pick = 'c'; })), buttons: [{ label: 'OK', primary: true, onClick: () => res(pick) }, { label: 'Cancel', onClick: () => res(null) }] }).done.then(() => res(null));
    });
  };

  /* ================================================================ sheets */
  D.insertSheetDlg = function () {
    const items = [['ws', L.N('Worksheet')], ['chart', L.N('Chart')]].concat(Object.keys(L.panes.TEMPLATES).map((k) => ['tpl:' + k, L.panes.TEMPLATES[k].name]));
    let pick = 'ws';
    const lb = list(items, 'ws', (v) => { pick = v; });
    const d = ui.dialog({ title: 'Insert', width: 360, body: ui.tabs([{ label: 'General', body: h('div', { class: 'col' }, lb) }, { label: 'Spreadsheet Solutions', body: h('div', { class: 'tp-note', text: 'Templates are listed on the General tab.' }) }]), buttons: [{ label: 'OK', primary: true, onClick: () => go() }, { label: 'Cancel' }] });
    lb.ondbl = () => { d.close(0); go(); };
    function go() {
      if (pick === 'ws') A().insertSheet();
      else if (pick === 'chart') L.chartWizard.open({ sheet: true });
      else {
        /* copy the template's sheet into this workbook */
        const t = L.panes.buildTemplate(pick.slice(4));
        const src = t.sheets[0];
        const w = wb();
        tryRun(() => {
          const s = O.addSheet(w, w.sheetByName(src.name) ? w.nextSheetName(src.name) : src.name, w.active);
          O.structural(s, L.N('Insert'), () => {
            const remap = new Map();
            const sid = (i) => { if (!i) return i; if (!remap.has(i)) remap.set(i, w.styles.add(JSON.parse(JSON.stringify(t.styles.get(i))))); return remap.get(i); };
            src.rows.forEach((row, r) => { if (row) row.cells.forEach((cl, c) => { if (cl) { const n = Object.assign({}, cl, { s: sid(cl.s) }); if (n.f != null) n.dirty = true; s.put(r, c, n); } }); });
            s.cols = src.cols.map((x) => (x ? Object.assign({}, x) : x));
            s.merges = src.merges.slice(); s.dv = JSON.parse(JSON.stringify(src.dv)); s.view.freeze = src.view.freeze; s.view.grid = src.view.grid;
          });
          w.active = w.sheets.indexOf(s);
          A().renderTabs();
        });
      }
    }
  };
  D.moveCopySheet = function () {
    const w = wb();
    const sheets = A().selectedSheets();
    const books = A().books.map((b, i) => [i, b.name]);
    let book = A().cur, before = w.sheets.indexOf(sheets[0]) + 1, copy = false;
    const lb = h('div');
    const drawList = () => { lb.textContent = ''; const tw = A().books[book].wb; lb.appendChild(list(tw.sheets.map((s, i) => [i, s.name]).concat([[tw.sheets.length, L.t('(move to end)')]]), Math.min(before, tw.sheets.length), (v) => { before = v; }, { translate: 'no' })); };
    drawList();
    ui.dialog({ title: 'Move or Copy', width: 320, body: h('div', { class: 'col' }, h('div', { text: 'Move selected sheets' }), ui.field('&To book:', ui.select(books, book, (v) => { book = +v; before = A().books[book].wb.sheets.length; drawList(); }, { translate: 'no' })), h('label', { text: '&Before sheet:' }), lb, ui.check('&Create a copy', false, (v) => { copy = v; })),
      buttons: [{ label: 'OK', primary: true, onClick: () => tryRun(() => {
        const tw = A().books[book].wb;
        if (tw === w) {
          let at = before;
          for (const s of sheets) { if (copy) { const c = O.copySheet(w, s, at); at = w.sheets.indexOf(c) + 1; w.active = w.sheets.indexOf(c); } else { O.moveSheet(w, s, at); at = w.sheets.indexOf(s) + 1; } }
        } else {
          /* another open workbook: copy the cells, styles and settings across */
          for (const s of sheets) {
            const ns = O.addSheet(tw, tw.sheetByName(s.name) ? tw.nextSheetName(s.name) : s.name, before);
            const remap = new Map();
            const sid = (i) => { if (!i) return i; if (!remap.has(i)) remap.set(i, tw.styles.add(JSON.parse(JSON.stringify(w.styles.get(i))))); return remap.get(i); };
            s.rows.forEach((row, r) => { if (row) { const nr = ns.rowObj(r); Object.assign(nr, row, { cells: [] }); if (row.s) nr.s = sid(row.s); row.cells.forEach((cl, c) => { if (cl) { const n = Object.assign({}, cl, { s: sid(cl.s), _gnode: null }); if (n.f != null) n.dirty = true; ns.put(r, c, n); } }); } });
            ns.cols = s.cols.map((x) => (x ? Object.assign({}, x, x.s ? { s: sid(x.s) } : {}) : x));
            ns.merges = s.merges.map((m) => Object.assign({}, m));
            ns.print = JSON.parse(JSON.stringify(s.print)); ns.view = JSON.parse(JSON.stringify(s.view)); ns.tabColor = s.tabColor;
            O.recalc(tw, null, true);
            if (!copy) O.deleteSheet(w, s);
          }
        }
        A().renderTabs();
      }) }, { label: 'Cancel' }] });
  };
  D.unhideSheet = function () {
    const w = wb();
    const hidden = w.sheets.map((s, i) => [i, s]).filter(([, s]) => s.state === 'hidden');
    const lb = list(hidden.map(([i, s]) => [i, s.name]), hidden.length ? hidden[0][0] : null, null, { translate: 'no' });
    const d = ui.dialog({ title: 'Unhide', width: 280, body: h('div', { class: 'col' }, h('label', { text: '&Unhide sheet:' }), lb), buttons: [{ label: 'OK', primary: true, onClick: () => go() }, { label: 'Cancel' }] });
    lb.ondbl = () => { d.close(0); go(); };
    function go() { const i = lb.value(); if (i == null) return; tryRun(() => { O.sheetProp(w, w.sheets[i], 'state', 'visible', 'Unhide Sheet'); w.active = i; A().renderTabs(); }); }
  };
  D.tabColor = function () {
    const w = wb(), sheets = A().selectedSheets();
    let pick = sheets[0].tabColor ? M.colorHex(w, sheets[0].tabColor) : null;
    const grid = h('div', { class: 'cp-grid' });
    const none = h('button', { type: 'button', class: 'cp-wide', text: 'No Color' });
    const mark = () => { L.$$('.cp-sw', grid).forEach((b) => b.classList.toggle('on', b.dataset.c === pick)); none.classList.toggle('on', !pick); };
    none.addEventListener('click', () => { pick = null; mark(); });
    for (const c of L.color.STANDARD) { const b = h('button', { type: 'button', class: 'cp-sw', style: `background:${c}`, 'data-tip': ui.colorName(c), 'aria-label': ui.colorName(c) }); b.dataset.c = c; b.addEventListener('click', () => { pick = c; mark(); }); grid.appendChild(b); }
    mark();
    ui.dialog({ title: 'Format Tab Color', width: 260, body: h('div', { class: 'col' }, ui.group('Tab Color', none, grid)), buttons: [{ label: 'OK', primary: true, onClick: () => tryRun(() => { O.tx(w, 'Tab Color', () => { for (const s of sheets) O.sheetProp(w, s, 'tabColor', pick ? M.rgb(pick) : null); }); A().renderTabs(); }) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Find and Replace */
  let findState = { what: '', with: '', within: 'sheet', by: 'rows', lookIn: 'formulas', matchCase: false, whole: false, more: false };
  D.find = function (mode) {
    const w = wb();
    const what = h('input', { type: 'text', value: findState.what, style: 'width:100%' });
    const withIn = h('input', { type: 'text', value: findState.with, style: 'width:100%' });
    const withRow = ui.field('R&eplace with:', withIn);
    const optBox = h('div', { class: 'fr-opts' },
      h('div', { class: 'col' }, ui.field('Wit&hin:', ui.select([['sheet', 'Sheet'], ['workbook', 'Workbook']], findState.within, (v) => { findState.within = v; })), ui.field('&Search:', ui.select([['rows', 'By Rows'], ['cols', 'By Columns']], findState.by, (v) => { findState.by = v; })), ui.field('&Look in:', ui.select([['formulas', 'Formulas'], ['values', 'Values'], ['comments', 'Comments']], findState.lookIn, (v) => { findState.lookIn = v; }))),
      h('div', { class: 'col' }, ui.check('Match &case', findState.matchCase, (v) => { findState.matchCase = v; }), ui.check('Match entire cell c&ontents', findState.whole, (v) => { findState.whole = v; })));
    optBox.hidden = !findState.more;
    const results = h('div', { class: 'fr-results', hidden: true });
    const status = h('div', { class: 'hint' });
    const moreBtn = ui.button(findState.more ? 'Options <<' : 'Op&tions >>', () => { findState.more = !findState.more; optBox.hidden = !findState.more; moreBtn.innerHTML = ui.ampHTML(findState.more ? 'Options <<' : 'Op&tions >>'); });
    const sync = () => { findState.what = what.value; findState.with = withIn.value; };
    const query = () => { sync(); const selRg = G().range(); const multi = G().ranges().length > 1 || selRg.r1 !== selRg.r2 || selRg.c1 !== selRg.c2; return Object.assign({}, findState, { sheet: sh(), range: findState.within === 'sheet' && multi ? selRg : null }); };
    const goTo = (hit) => { const i = w.sheets.indexOf(hit.sh); if (i !== w.active) A().activateSheet(i); G().select(hit.r, hit.c); };
    const findNext = (back) => {
      const q = query();
      if (!q.what) return null;
      const hits = O.findAll(w, q);
      if (!hits.length) { ui.msg("Ledger cannot find the data you're searching for.", { icon: 'info' }); return null; }
      const s = G().sel(), cur = { si: w.active, r: s.r, c: s.c };
      const key = (x) => [w.sheets.indexOf(x.sh), q.by === 'cols' ? x.c : x.r, q.by === 'cols' ? x.r : x.c];
      const ck = [cur.si, q.by === 'cols' ? cur.c : cur.r, q.by === 'cols' ? cur.r : cur.c];
      const cmp = (a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
      let hit = back ? hits.slice().reverse().find((x) => cmp(key(x), ck) < 0) || hits[hits.length - 1] : hits.find((x) => cmp(key(x), ck) > 0) || hits[0];
      goTo(hit);
      return hit;
    };
    const findAllBtn = () => {
      const q = query();
      const hits = O.findAll(w, q);
      results.hidden = false;
      results.textContent = '';
      const tbl = h('table', { class: 'fr-table' }, h('tr', null, ...[L.N('Book|workbook'), L.N('Sheet'), L.N('Name'), L.N('Cell'), L.N('Value'), L.N('Formula')].map((t) => h('th', { text: t }))));
      for (const x of hits.slice(0, 2000)) {
        const cl = x.sh.get(x.r, x.c);
        const nm = w.names.find((n) => { try { const a = F.parseRange(String(n.ref).split('!').pop().replace(/\$/g, '')); return a && M.rangeContains(a, x.r, x.c); } catch (e) { return false; } });
        const tr = h('tr', null, h('td', { translate: 'no', text: A().books[A().cur].name }), h('td', { translate: 'no', text: x.sh.name }), h('td', { translate: 'no', text: nm ? nm.name : '' }), h('td', { translate: 'no', text: F.cellName(x.r, x.c, true, true) }), h('td', { translate: 'no', text: L.csv.cellText(x.sh, x.r, x.c) }), h('td', { translate: 'no', text: cl && cl.f != null ? '=' + F.display(cl.f) : '' }));
        tr.addEventListener('click', () => { L.$$('tr', tbl).forEach((t) => t.classList.remove('on')); tr.classList.add('on'); goTo(x); });
        tbl.appendChild(tr);
      }
      results.appendChild(tbl);
      status.textContent = L.t('{0} cell(s) found', hits.length);
    };
    let tabsEl, d = null;
    const findTab = h('div', { class: 'col' }), repTab = h('div', { class: 'col' });
    const form = h('div', { class: 'col' }, ui.field('Fi&nd what:', what), withRow, h('div', { class: 'row', style: 'justify-content:flex-end' }, moreBtn), optBox);
    tabsEl = ui.tabs([{ label: 'Fi&nd', body: findTab, onShow: () => { withRow.hidden = true; findTab.appendChild(form); setBtns(false); } }, { label: 'Re&place', body: repTab, onShow: () => { withRow.hidden = false; repTab.appendChild(form); setBtns(true); } }], mode === 'replace' ? 1 : 0);
    d = ui.dialog({
      title: 'Find and Replace', width: 520, modeless: true, body: h('div', { class: 'col' }, tabsEl, results, status),
      buttons: [
        { label: 'Replace &All', onClick: () => { const q = query(); const hits = O.findAll(w, Object.assign({}, q, { lookIn: 'formulas' })); tryRun(() => { const n = O.replaceAll(w, q, hits); ui.msg(n === 1 ? L.t('Ledger has completed its search and has made {0} replacement.', n) : L.t('Ledger has completed its search and has made {0} replacements.', n), { icon: 'info' }); }); return false; } },
        { label: '&Replace', onClick: () => { const q = query(); const s = G().sel(); const here = O.findAll(w, Object.assign({}, q, { lookIn: 'formulas', within: 'sheet', range: { r1: s.r, c1: s.c, r2: s.r, c2: s.c } })); if (here.length) tryRun(() => O.replaceAll(w, q, here)); findNext(false); return false; } },
        { label: 'F&ind All', onClick: () => { findAllBtn(); return false; } },
        { label: '&Find Next', primary: true, onClick: () => { findNext(false); return false; } },
        { label: 'Close' },
      ],
    });
    function setBtns(rep) { if (!d) return; d.buttons[0].hidden = !rep; d.buttons[1].hidden = !rep; }
    setTimeout(() => { setBtns(mode === 'replace'); what.focus(); what.select(); }, 0);
    what.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); findNext(e.shiftKey); } });
  };

  /* ================================================================ Go To */
  D.goTo = function () {
    const w = wb();
    const names = w.names.filter((n) => !n.hidden && !/^_xlnm\./i.test(n.name) && (n.scope == null || n.scope === w.active)).map((n) => n.name).concat(w.tables.map((t) => t.name)).concat(A().lastGoTo || []);
    const ref = h('input', { type: 'text', style: 'width:100%' });
    const lb = list(Array.from(new Set(names)), null, (v) => { ref.value = v; }, { translate: 'no' });
    const d = ui.dialog({ title: 'Go To', width: 320, body: h('div', { class: 'col' }, h('label', { html: ui.ampHTML('&Go to:') }), lb, ui.field('&Reference:', ref)),
      buttons: [{ label: '&Special...', onClick: () => { d.close(null); D.goToSpecial(); return false; } }, { label: 'OK', primary: true, onClick: () => { if (!A().goToRef(ref.value.trim())) { ui.msg('Reference is not valid.', { icon: 'warn' }); return false; } } }, { label: 'Cancel' }] });
    lb.ondbl = (v) => { d.close(0); A().goToRef(v); };
  };
  D.goToSpecial = function () {
    const st = { kind: 'constants', types: { n: true, t: true, l: true, e: true }, arr: 'row' };
    const sub = (k) => h('div', { class: 'gs-sub' }, ...[['n', L.N('Numbers')], ['t', L.N('Text')], ['l', L.N('Logicals')], ['e', L.N('Errors')]].map(([t, l]) => ui.check(l, true, (v) => { st.types[t] = v; })));
    const r = (k, label, extra) => h('div', null, ui.radio('gts', label, k === 'constants', () => { st.kind = k; }), extra || null);
    const body = h('div', { class: 'gs-grid' },
      h('div', { class: 'col' }, r('comments', 'Co&mments'), r('constants', 'C&onstants'), r('formulas', '&Formulas', sub('f')), r('blanks', '&Blanks'), r('region', 'Current &region'), r('array', 'Current &array'), r('objects', 'O&bjects')),
      h('div', { class: 'col' }, r('rowdiff', 'Ro&w differences'), r('coldiff', 'Col&umn differences'), r('prec', '&Precedents'), r('dep', '&Dependents'), r('last', 'La&st cell'), r('visible', 'Visible cells onl&y'), r('cf', 'Condi&tional formats'), r('dv', 'Data &validation')));
    ui.dialog({ title: 'Go To Special', width: 400, body, buttons: [{ label: 'OK', primary: true, onClick: () => {
      const s = sh(), sel = G().sel();
      const scope = (() => { const rg = G().range(); const single = G().ranges().length === 1 && rg.r1 === rg.r2 && rg.c1 === rg.c2; return single ? { r1: 0, c1: 0, r2: Math.max(0, s.maxR), c2: Math.max(0, s.maxC) } : O.clip(s, rg); })();
      const out = [];
      const add = (r, c) => out.push({ r1: r, c1: c, r2: r, c2: c });
      const typeOk = (v) => (typeof v === 'number' && st.types.n) || (typeof v === 'string' && st.types.t) || (typeof v === 'boolean' && st.types.l) || (M.isErr(v) && st.types.e);
      switch (st.kind) {
        case 'comments': for (const cm of s.comments.values()) if (M.rangeContains(scope, cm.r, cm.c)) add(cm.r, cm.c); break;
        case 'constants': s.each(scope.r1, scope.c1, scope.r2, scope.c2, (cl, rr, cc) => { if (cl.f == null && !cl.am && cl.v != null && cl.v !== '' && typeOk(cl.v)) add(rr, cc); }); break;
        case 'formulas': s.each(scope.r1, scope.c1, scope.r2, scope.c2, (cl, rr, cc) => { if ((cl.f != null || cl.am) && typeOk(cl.dirty ? C.cellValue(s, rr, cc) : cl.v)) add(rr, cc); }); break;
        case 'blanks': for (let rr = scope.r1; rr <= scope.r2; rr++) for (let cc = scope.c1; cc <= scope.c2 && out.length < 50000; cc++) { const cl = s.get(rr, cc); if (!cl || (cl.v == null && cl.f == null)) add(rr, cc); } break;
        case 'region': out.push(O.currentRegion(s, sel.r, sel.c)); break;
        case 'array': { const a = O.arrayAt(s, sel.r, sel.c); if (a) out.push(a); break; }
        case 'last': { const e = s.contentEnd(); add(e.r, e.c); break; }
        case 'visible': { const g = LY.geo(s); for (const rg of G().ranges()) { let start = null; const rr = O.clip(s, rg); for (let row = rr.r1; row <= rr.r2 + 1; row++) { const vis = row <= rr.r2 && g.rows.size(row) > 0; if (vis && start == null) start = row; if (!vis && start != null) { out.push({ r1: start, c1: rg.c1, r2: row - 1, c2: rg.c2 }); start = null; } } } break; }
        case 'cf': for (const cf of s.cf) out.push(...cf.ranges); break;
        case 'dv': for (const d of s.dv) out.push(...d.ranges); break;
        case 'objects': if (s.drawings.length) G().selectObject(s.drawings[0]); return;
        case 'prec': case 'dep': {
          const cell = s.get(sel.r, sel.c);
          if (st.kind === 'prec') { if (cell && cell.f != null) { const p = C.precedents(w, s, C.astOf(cell), sel.r, sel.c); for (const a of p.refs) if (a.sheet === s) out.push({ r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 }); } }
          else { const deps = []; w.graph.dependents(s, sel.r, sel.c, deps); for (const n of deps) if (n.sh === s) add(n.r, n.c); }
          break;
        }
        case 'rowdiff': case 'coldiff': {
          const rg = O.clip(s, G().range());
          const txt = (rr, cc) => { const cl = s.get(rr, cc); return cl ? (cl.f != null ? F.display(cl.f) : String(cl.v)) : ''; };
          if (st.kind === 'rowdiff') for (let rr = rg.r1; rr <= rg.r2; rr++) { const base = s.get(rr, sel.c); for (let cc = rg.c1; cc <= rg.c2; cc++) { if (cc === sel.c) continue; const cl = s.get(rr, cc); if (base && base.f != null && cl && cl.f != null) { if (F.translate(F.display(base.f), 0, cc - sel.c) !== F.display(cl.f)) add(rr, cc); } else if (txt(rr, cc) !== txt(rr, sel.c)) add(rr, cc); } }
          else for (let cc = rg.c1; cc <= rg.c2; cc++) { const base = s.get(sel.r, cc); for (let rr = rg.r1; rr <= rg.r2; rr++) { if (rr === sel.r) continue; const cl = s.get(rr, cc); if (base && base.f != null && cl && cl.f != null) { if (F.translate(F.display(base.f), rr - sel.r, 0) !== F.display(cl.f)) add(rr, cc); } else if (txt(rr, cc) !== txt(sel.r, cc)) add(rr, cc); } }
          break;
        }
        default: break;
      }
      if (!out.length) { ui.msg('No cells were found.', { icon: 'info' }); return false; }
      G().selectRange(out[0], { r: out[0].r1, c: out[0].c1 });
      const s2 = G().sel();
      s2.ranges = out.slice(0, 5000);
      s2.active = 0;
      G().changed();
    } }, { label: 'Cancel' }] });
  };

  /* ================================================================ Paste Special */
  D.pasteSpecial = function () {
    const st = { what: 'all', op: 'none', skipBlanks: false, transpose: false };
    const what = [['all', '&All'], ['formulas', '&Formulas'], ['values', '&Values'], ['formats', 'Forma&ts'], ['comments', '&Comments'], ['validation', 'Validatio&n'], ['noBorders', 'All e&xcept borders'], ['colWidths', 'Column &widths'], ['formulasNF', 'Formulas and number f&ormats'], ['valuesNF', 'Values and number for&mats']];
    const ops = [['none', 'N&one'], ['add', 'A&dd'], ['sub', '&Subtract'], ['mul', 'M&ultiply'], ['div', 'D&ivide']];
    const half = (arr, name, key) => { const radios = arr.map(([k, l]) => ui.radio(name, l, k === st[key], () => { st[key] = k; })); const m = Math.ceil(radios.length / 2); return h('div', { class: 'ps-cols' }, h('div', { class: 'col' }, ...radios.slice(0, m)), h('div', { class: 'col' }, ...radios.slice(m))); };
    ui.dialog({ title: 'Paste Special', width: 440, body: h('div', { class: 'col' }, ui.group('Paste', half(what, 'ps-what', 'what')), ui.group('Operation', half(ops, 'ps-op', 'op')), h('div', { class: 'row', style: 'gap:24px' }, ui.check('Skip &blanks', false, (v) => { st.skipBlanks = v; }), ui.check('Transpos&e', false, (v) => { st.transpose = v; }))),
      buttons: [{ label: 'Paste &Link', onClick: () => tryRun(() => L.clip.pasteInternal({ link: true })) }, { label: 'OK', primary: true, onClick: () => tryRun(() => L.clip.pasteInternal(st)) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Series, Fill Across */
  D.series = function () {
    const s = sh(), rg = O.clip(s, G().range());
    const st = { rows: rg.c2 - rg.c1 > rg.r2 - rg.r1, type: 'linear', unit: 'day', trend: false };
    const first = s.get(rg.r1, rg.c1);
    if (first && typeof first.v === 'number' && NF.isDate(LY.styleOf(s, rg.r1, rg.c1, first).nf || 'General')) st.type = 'date';
    const step = h('input', { type: 'text', value: '1', style: 'width:80px' }), stop = h('input', { type: 'text', value: '', style: 'width:80px' });
    const units = [['day', '&Day'], ['weekday', '&Weekday'], ['month', '&Month'], ['year', '&Year']].map(([k, l]) => ui.radio('ser-u', l, k === 'day', () => { st.unit = k; }));
    ui.dialog({ title: 'Series', width: 380, body: h('div', { class: 'col' },
      h('div', { class: 'row', style: 'gap:10px;align-items:flex-start' },
        ui.group('Series in', ui.radio('ser-in', '&Rows', st.rows, () => { st.rows = true; }), ui.radio('ser-in', '&Columns', !st.rows, () => { st.rows = false; })),
        ui.group('Type', ...[['linear', '&Linear'], ['growth', '&Growth'], ['date', 'Da&te'], ['auto', 'Auto&Fill']].map(([k, l]) => ui.radio('ser-t', l, k === st.type, () => { st.type = k; }))),
        ui.group('Date unit', ...units)),
      ui.check('T&rend', false, (v) => { st.trend = v; }),
      h('div', { class: 'row', style: 'gap:10px' }, ui.field('&Step value:', step), ui.field('St&op value:', stop))),
    buttons: [{ label: 'OK', primary: true, onClick: () => tryRun(() => {
      if (st.type === 'auto') { const src = st.rows ? { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c1 } : { r1: rg.r1, c1: rg.c1, r2: rg.r1, c2: rg.c2 }; O.autoFill(s, src, rg, 'auto'); return; }
      O.fillSeries(s, rg, { rows: st.rows, type: st.type, unit: st.unit, step: num(step.value, 1), stop: stop.value.trim() ? num(stop.value, null) : null, trend: st.trend });
    }) }, { label: 'Cancel' }] });
  };
  D.fillAcross = function () {
    const sheets = A().selectedSheets();
    let what = 'all';
    ui.dialog({ title: 'Fill Across Worksheets', width: 260, body: ui.group('Fill', ...[['all', '&All'], ['contents', '&Contents'], ['formats', '&Formats']].map(([k, l], i) => ui.radio('fa', l, i === 0, () => { what = k; }))), buttons: [{ label: 'OK', primary: true, onClick: () => tryRun(() => {
      const src = sh(), w = wb();
      O.tx(w, 'Fill Across Worksheets', () => {
        for (const rg0 of G().ranges()) {
          const rg = O.clip(src, rg0);
          for (const t of sheets) {
            if (t === src) continue;
            for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
              const a = src.get(r, c), b = t.get(r, c);
              const n = Object.assign({}, b || { v: null });
              if (what !== 'formats') { n.v = a ? a.v : null; if (a && a.f != null) { n.f = a.f; n.dirty = true; } else delete n.f; }
              if (what !== 'contents') n.s = a ? a.s : undefined;
              O.put(t, r, c, n.v == null && n.f == null && !n.s ? null : n);
            }
          }
        }
      });
    }) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Save As */
  /** Save As: the type and name, then the workbook written (A.saveAs); whether it was (false on Cancel). The desktop
   *  app asks with the system's Save dialog (L.file) */
  D.saveAs = async function () {
    const b = A().books[A().cur];
    const base = b.name.replace(/\.[^.]+$/, '');
    const types = L.file.types.map((t) => [t.id, t.label]);
    let type = b.type === 'xlsm' || b.type === 'xltx' || b.type === 'xltm' || b.type === 'csv' || b.type === 'txt' ? b.type : 'xlsx';
    if (L.file.native()) {
      const c = await L.file.choose('saveAs', type, base);
      return c ? A().saveAs(c.type, c.file, 'chosen') : false;
    }
    let saving = null;
    const name = h('input', { type: 'text', value: base, style: 'width:100%' });
    const pw = { open: b.password || '', backup: false };
    /* the Tools button's menu, under the button */
    const tools = (button) => {
      const r = button.getBoundingClientRect();
      ui.openMenu([{ label: '&General Options...', run: () => {
        const p1 = h('input', { type: 'password', value: pw.open, autocomplete: 'new-password' });
        ui.dialog({ title: 'Save Options', width: 320, body: h('div', { class: 'col' }, ui.group('File sharing', ui.field('Password to &open:', p1)), h('div', { class: 'hint', text: 'Workbooks saved with a password are encrypted (AES-256) and open in Excel 2010 and later.' })), buttons: [{ label: 'OK', primary: true, onClick: () => {
          if (!p1.value) { pw.open = ''; return; }
          const p2 = h('input', { type: 'password', autocomplete: 'new-password' });
          ui.dialog({ title: 'Confirm Password', width: 300, body: h('div', { class: 'col' }, ui.field('Reenter password to proceed.', p2), h('div', { class: 'hint', text: 'Caution: If you lose or forget the password, it cannot be recovered.' })), buttons: [{ label: 'OK', primary: true, onClick: () => { if (p2.value !== p1.value) { ui.msg("Confirmation password is not identical.", { icon: 'warn' }); return false; } pw.open = p1.value; } }, { label: 'Cancel' }] });
        } }, { label: 'Cancel' }] });
      } }], { left: r.left, bottom: r.bottom });
    };
    return ui.dialog({ title: 'Save As', width: 460, body: h('div', { class: 'col' }, ui.field('File &name:', name), ui.field('Save as &type:', ui.select(types, type, (v) => { type = v; })), h('div', { class: 'tp-note', text: 'Your browser asks where to save the file, or saves it to the Downloads folder.' })),
      buttons: [{ label: 'Too&ls', onClick: (close, button) => { tools(button); return false; } }, { label: '&Save', primary: true, onClick: () => {
        const ext = { xmlss: 'xml', html: 'htm' }[type] || type;
        const n = (name.value.trim() || base).replace(/\.(xlsx|xlsm|xltx|xltm|xml|csv|txt|htm|html|pdf)$/i, '') + '.' + ext;
        b.password = pw.open || undefined;
        saving = A().saveAs(type, n);
      } }, { label: 'Cancel' }] }).done.then(() => saving || false);
  };

  /* ================================================================ Properties */
  /** File ▸ Properties (common ui.propertiesDialog); a custom property's type is its vt element (lpwstr, i4, bool…) */
  D.properties = function () {
    const w = wb(), p = w.props;
    const fmt = (x) => (x ? new Date(x).toLocaleString() : '');
    let cells = 0, formulas = 0;
    for (const s of w.sheets) s.rows.forEach((row) => { if (row) row.cells.forEach((cl) => { if (cl && (cl.v != null || cl.f != null)) cells++; if (cl && cl.f != null) formulas++; }); });
    const VT = { lpwstr: 'text', bstr: 'text', i4: 'number', i8: 'number', int: 'number', r4: 'number', r8: 'number', decimal: 'number', bool: 'bool', filetime: 'date', date: 'date' };
    const toVt = (c) => (c.type === 'number' ? (Number.isInteger(c.value) ? 'i4' : 'r8') : c.type === 'bool' ? 'bool' : c.type === 'date' ? 'filetime' : 'lpwstr');
    const toValue = (c) => (c.type === 'bool' ? (c.value ? 'true' : 'false') : c.type === 'date' ? new Date(c.value).toISOString().replace(/\.\d{3}Z$/, 'Z') : String(c.value));
    ui.propertiesDialog({
      name: A().books[A().cur].name,
      fields: { title: p.title, subject: p.subject, author: p.creator, manager: p.manager, company: p.company, category: p.category, keywords: p.keywords, comments: p.description },
      statistics: [['Created:', fmt(p.created)], ['Modified:', fmt(p.modified)], ['Last saved by:', p.lastModifiedBy || ''], ['Application:', p.application || '']],
      counts: [['Sheets:', w.sheets.length], ['Cells with data:', cells], ['Formulas:', formulas], ['Defined names:', w.names.filter((n) => !n.hidden).length], ['Lists (tables):', (w.tables || []).length]],
      contents: [[L.t('Worksheets'), w.sheets.map((s) => (s.state !== 'visible' ? L.t('{0} (hidden)', s.name) : s.name))], [L.t('Named Ranges'), w.names.filter((n) => !n.hidden).map((n) => n.name)]],
      custom: (p.custom || []).map((c) => { const type = VT[c.type] || 'text'; return { name: c.name, type, value: type === 'number' ? Number(c.value) : type === 'bool' ? /^(1|true)$/i.test(c.value) : type === 'date' ? String(c.value).slice(0, 10) : c.value, was: c }; }),
      apply: ({ fields, custom }) => {
        const KEY = { author: 'creator', comments: 'description' };
        const old = Object.assign({}, p), now = Object.assign({}, p);
        for (const [k, v] of Object.entries(fields)) now[KEY[k] || k] = v;
        /* an unchanged custom property stays as it was read */
        if (custom) now.custom = custom.map((c) => c.was || { name: c.name, type: toVt(c), value: toValue(c) });
        O.tx(w, L.N('Properties'), () => { Object.assign(p, now); w.undo.op(() => Object.assign(p, old), () => Object.assign(p, now)); });
      },
    });
  };

  /* ================================================================ Zoom, Symbol */
  D.zoom = function () {
    const z = sh().view.zoom || 100;
    const PRESETS = [[200, '&200%'], [100, '&100%'], [75, '&75%'], [50, '5&0%'], [25, '2&5%']];
    let pick = PRESETS.some(([v]) => v === z) ? z : 'custom';
    const custom = ui.spin({ value: z, min: 10, max: 400, step: 5, dec: 0, unit: '%', id: 'zm-custom', onChange: () => { pick = 'custom'; customR.input.checked = true; } });
    const customR = ui.radio('zoom', '&Custom:', pick === 'custom', () => { pick = 'custom'; });
    const opts = PRESETS.map(([v, l]) => ui.radio('zoom', l, v === pick, () => { pick = v; custom.set(v); }));
    ui.dialog({ title: 'Zoom', width: 240, body: ui.group('Magnification', h('div', { class: 'col', style: 'gap:2px' }, ...opts, ui.radio('zoom', '&Fit selection', false, () => { pick = 'fit'; }), h('div', { class: 'row', style: 'gap:6px' }, customR, custom))),
      buttons: [{ label: 'OK', primary: true, onClick: () => { if (pick === 'fit') A().zoomToSelection(); else G().setZoom(pick === 'custom' ? Math.round(custom.get()) : pick); } }, { label: 'Cancel' }] });
  };
  /** Insert ▸ Symbol (common ui.symbolDialog): into the cell being edited, or the active cell's text */
  D.symbol = () => ui.symbolDialog({
    fonts: L.FONT_LIST,
    insert: (c) => { if (!L.editor.active) L.editor.begin('edit'); L.editor.insertText(c.ch); },
  });

  /* ================================================================ comments */
  /** Excel edits a comment in place: a yellow note beside the cell */
  D.comment = function (r, c, cm) {
    const s = sh(), w = wb();
    const q = G().cellRect(r, c);
    const wrapR = L.$('#gridwrap').getBoundingClientRect();
    const author = (A().opts.userName || 'Ledger User');
    const box = h('div', { class: 'cmt-edit' });
    const ta = h('textarea', { 'aria-label': 'Comment', spellcheck: 'true' });
    const initialText = cm ? L.threads.editText(cm) || '' : author + ':\n';
    ta.value = initialText;
    box.appendChild(ta);
    Object.assign(box.style, { left: (wrapR.left + q.x + q.w + 12) + 'px', top: Math.max(wrapR.top + 2, wrapR.top + q.y - 6) + 'px', width: (cm && cm.w ? cm.w * 96 / 72 : 150) + 'px', height: (cm && cm.h ? cm.h * 96 / 72 : 80) + 'px' });
    document.body.appendChild(box);
    ta.focus();
    if (!cm) ta.setSelectionRange(ta.value.length, ta.value.length);
    let done = false;
    const finish = (save) => {
      if (done) return; done = true;
      box.remove();
      document.removeEventListener('pointerdown', outside, true);
      if (save) {
        const text = ta.value;
        const same = cm && initialText === text;
        if (!same) tryRun(() => O.tx(w, cm ? 'Edit Comment' : 'Insert Comment', () => O.setComment(s, r, c, L.threads.editNote(Object.assign({}, cm || { author, visible: false }, { text, runs: undefined, w: box.offsetWidth * 0.75 || undefined, h: box.offsetHeight * 0.75 || undefined }), text))));
      }
      G().paint(); G().syncObjects(); G().focus();
    };
    const outside = (e) => { if (!box.contains(e.target)) finish(true); };
    setTimeout(() => document.addEventListener('pointerdown', outside, true), 0);
    ta.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') { e.preventDefault(); finish(true); } });
  };

  /* ================================================================ Options */
  D.options = function (tabName) {
    const o = A().opts, w = wb(), s = sh();
    const v = s.view;
    const ck = (label, get, set) => ui.check(label, get(), set);
    const viewTab = h('div', { class: 'col' },
      ui.group('Show', ck('&Startup Task Pane', () => o.taskOpen !== false, (x) => { o.taskOpen = x; }), ck('Task pane f&ollows the selected object', () => o.paneFollows !== false, (x) => { o.paneFollows = x; }), ck('&Formula bar', () => o.formulaBar !== false, (x) => { o.formulaBar = x; }), ck('St&atus bar', () => o.statusBar !== false, (x) => { o.statusBar = x; })),
      ui.group('Comments', ui.radio('opt-cm', '&None', !w._showComments, () => { w._showComments = false; }), ui.radio('opt-cm', 'Comment &indicator only', !w._showComments, () => { w._showComments = false; }), ui.radio('opt-cm', 'Comment && indi&cator', !!w._showComments, () => { w._showComments = true; })),
      ui.group('Window options', h('div', { class: 'opt-cols' },
        ck('Page brea&ks', () => !!s._pages, (x) => { if (x) L.print.computeBreaks(s); else s._pages = null; }), ck('Fo&rmulas', () => !!v.formulas, (x) => { v.formulas = x; }), ck('&Gridlines', () => v.grid !== false, (x) => { v.grid = x; }), ck('Row && column h&eaders', () => v.headings !== false, (x) => { v.headings = x; }),
        ck('&Zero values', () => v.zeros !== false, (x) => { v.zeros = x; }), ck('Outline s&ymbols', () => v.outlineSymbols !== false, (x) => { v.outlineSymbols = x; }), ck('Hori&zontal scroll bar', () => !w.hideHScroll, (x) => { w.hideHScroll = !x; }), ck('&Vertical scroll bar', () => !w.hideVScroll, (x) => { w.hideVScroll = !x; }), ck('She&et tabs', () => !w.hideTabs, (x) => { w.hideTabs = !x; }))));
    const calc = w.calcPr;
    const iterN = h('input', { type: 'text', value: String(calc.iterateCount || 100), style: 'width:60px' }), iterD = h('input', { type: 'text', value: String(calc.iterateDelta || 0.001), style: 'width:60px' });
    const calcTab = h('div', { class: 'col' },
      ui.group('Calculation', ui.radio('opt-calc', '&Automatic', calc.mode !== 'manual' && calc.mode !== 'autoNoTable', () => { calc.mode = 'auto'; }), ui.radio('opt-calc', 'Automatic except &tables', calc.mode === 'autoNoTable', () => { calc.mode = 'autoNoTable'; }), ui.radio('opt-calc', '&Manual', calc.mode === 'manual', () => { calc.mode = 'manual'; }), ui.button('Calc &Now (F9)', () => A().calculateNow())),
      ui.group('Iteration', ck('&Iteration', () => !!calc.iterate, (x) => { calc.iterate = x; }), h('div', { class: 'row', style: 'gap:10px' }, ui.field('Maximum iteratio&ns:', iterN), ui.field('Maximum &change:', iterD))),
      ui.group('Workbook options', ck('&Precision as displayed', () => calc.fullPrecision === false, (x) => { calc.fullPrecision = !x; }), ck(L.N('&1904 date system'), () => !!w.date1904, (x) => { w.date1904 = x; }), ck('Accept la&bels in formulas', () => false, () => {})));
    const dir = ui.select([['down', 'Down'], ['right', 'Right'], ['up', 'Up'], ['left', 'Left']], o.enterDir || 'down', (x) => { o.enterDir = x; });
    const editTab = h('div', { class: 'col' },
      ck('Edit directly in c&ell', () => o.editInCell !== false, (x) => { o.editInCell = x; }), ck('Allow cell &drag and drop', () => o.dragDrop !== false, (x) => { o.dragDrop = x; }),
      h('div', { class: 'row', style: 'gap:8px' }, ck('&Move selection after Enter', () => o.moveAfterEnter !== false, (x) => { o.moveAfterEnter = x; }), ui.field('D&irection:', dir)),
      ck('Enable Auto&Complete for cell values', () => o.autoComplete !== false, (x) => { o.autoComplete = x; }), ck('Extend data range &formats and formulas', () => o.extendFormats !== false, (x) => { o.extendFormats = x; }), ck('Enable automatic percent entry', () => o.autoPercent !== false, (x) => { o.autoPercent = x; }));
    const uname = h('input', { type: 'text', value: o.userName || '', style: 'width:200px' });
    const nSheets = ui.spin({ value: o.sheetsInNew || 3, min: 1, max: 255, step: 1, dec: 0 });
    const dfont = ui.select(L.FONT_LIST, o.defaultFont || 'Arial', null, { translate: 'no' }), dsize = ui.select(L.SIZE_LIST, o.defaultSize || 10);
    const look = ui.lookField();
    const genTab = h('div', { class: 'col' }, look && look.el, ck('R1C1 reference st&yle', () => !!w.r1c1, (x) => { w.r1c1 = x; }), ck('&Recently used file list', () => true, () => {}),
      ui.field('Sheets in new work&book:', nSheets), h('div', { class: 'row', style: 'gap:8px' }, ui.field('Standard fon&t:', dfont), ui.field('Si&ze:', dsize)), ui.field('User &name:', uname));
    const lists = O.customLists;
    const listBox = h('div');
    const listText = h('textarea', { rows: '8', style: 'width:100%', translate: 'no' });
    let curList = -1;
    const drawLists = () => { listBox.textContent = ''; listBox.appendChild(list([[-1, L.t('NEW LIST')]].concat(lists.map((l, i) => [i, l.slice(0, 4).join(', ') + (l.length > 4 ? ', …' : '')])), curList, (i) => { curList = i; listText.value = i >= 0 ? lists[i].join('\n') : ''; listText.readOnly = i >= 0 && i < 4; }, { translate: 'no' })); };
    drawLists();
    const listsTab = h('div', { class: 'row', style: 'gap:10px;align-items:flex-start' }, h('div', { class: 'col', style: 'flex:1' }, h('label', { text: 'Custom lists:' }), listBox), h('div', { class: 'col', style: 'flex:1' }, h('label', { text: 'List entries:' }), listText),
      h('div', { class: 'col' }, ui.button('&Add', () => { const items = listText.value.split(/[\n,]+/).map((x) => x.trim()).filter(Boolean); if (!items.length) return; if (curList >= 4) lists[curList] = items; else lists.push(items); L.store.set('customLists', lists.slice(4)); curList = -1; listText.value = ''; drawLists(); }), ui.button('&Delete', () => { if (curList >= 4) { lists.splice(curList, 1); L.store.set('customLists', lists.slice(4)); curList = -1; listText.value = ''; drawLists(); } })));
    const dictSel = ui.select([['en_US', 'English (U.S.)'], ['en_GB', 'English (U.K.)']], o.dictLang === 'en_GB' ? 'en_GB' : 'en_US', (v) => { o.dictLang = v; });
    const spellTab = h('div', { class: 'col' }, ui.field('&Dictionary language:', dictSel), ck('Ignore words in &UPPERCASE', () => o.ignoreUpper !== false, (x) => { o.ignoreUpper = x; }), ck('Ignore words with &numbers', () => o.ignoreNum !== false, (x) => { o.ignoreNum = x; }), ck('Ignore &Internet and file addresses', () => o.ignoreInternet !== false, (x) => { o.ignoreInternet = x; }), ui.button('&AutoCorrect Options...', () => D.autoCorrect()));
    const errTab = h('div', { class: 'col' }, ck('&Enable background error checking', () => o.bgErrors !== false, (x) => { o.bgErrors = x; }), ck('Evaluates to error &value', () => true, () => {}), ck('Number stored as &text', () => o.errNumText !== false, (x) => { o.errNumText = x; }), ck('&Inconsistent formula in region', () => o.errInconsistent !== false, (x) => { o.errInconsistent = x; }));
    /* Security: the password to open this workbook (applied when it is next saved as .xlsx / .xlsm) */
    const book = A().books[A().cur];
    const pwOpen = h('input', { type: 'password', value: (book && book.password) || '', autocomplete: 'new-password', style: 'width:180px' });
    const securityTab = h('div', { class: 'col' },
      ui.group('File encryption settings for this workbook', ui.field('Password to &open:', pwOpen)),
      h('div', { class: 'hint', text: 'The workbook is encrypted with AES-256 the next time you save it as an Excel workbook, the way Excel 2007 and later do. Leave the box empty to save it without a password. If you lose or forget the password, it cannot be recovered.' }),
      ui.group('Privacy options', ck('&Remove personal information from file properties on save', () => !!(w.extra && (w.extra.removePersonal || w.extra.filterPrivacy)), (x) => { w.extra = w.extra || {}; w.extra.removePersonal = x; w.extra.filterPrivacy = x; })));
    const tabs = [{ label: 'View', body: viewTab }, { label: 'Calculation', body: calcTab }, { label: 'Edit', body: editTab }, { label: 'General', body: genTab }, { label: 'Custom Lists', body: listsTab }, { label: 'Error Checking', body: errTab }, { label: 'Spelling', body: spellTab }, { label: 'Security', body: securityTab }];
    const idx = Math.max(0, tabs.findIndex((t) => t.label === tabName));
    ui.dialog({ title: 'Options', width: 540, body: ui.tabs(tabs, idx), buttons: [{ label: 'OK', primary: true, onClick: () => {
      calc.iterateCount = Math.max(1, Math.min(32767, parseInt(iterN.value, 10) || 100));
      calc.iterateDelta = num(iterD.value, 0.001);
      o.userName = uname.value; o.sheetsInNew = Math.round(nSheets.get()); o.defaultFont = dfont.value; o.defaultSize = +dsize.value;
      if (book && (pwOpen.value || '') !== (book.password || '')) { book.password = pwOpen.value || undefined; book.dirty = true; }
      A().applyOpts();
      if (look) look.apply();
      if (calc.mode !== 'manual') A().calculateNow();
      G().paint(true); G().syncObjects(true); A().renderTabs();
    } }, { label: 'Cancel' }] });
  };
  D.autoCorrect = function () {
    const AC = L.autocorrect, o = A().opts;
    const ck = (label, key, def) => ui.check(label, o[key] !== undefined ? o[key] !== false : def, (x) => { o[key] = x; });
    const from = h('input', { type: 'text', style: 'width:120px' }), to = h('input', { type: 'text', style: 'width:220px' });
    const box = h('div');
    const draw = () => { box.textContent = ''; const all = Array.from(AC.list().entries()).sort((a, b) => a[0].localeCompare(b[0])); box.appendChild(list(all.map(([a, b]) => [a, a + '  →  ' + b]), null, (k) => { from.value = k; to.value = AC.list().get(k); }, { class: 'lbox ac-list', translate: 'no' })); };
    draw();
    ui.dialog({ title: 'AutoCorrect: English (U.S.)', width: 460, body: ui.tabs([{ label: 'AutoCorrect', body: h('div', { class: 'col' },
      ck('Correct TWo INitial CApitals', 'acTwoCaps', true), ck('Capitalize first letter of &sentences', 'acSentence', false), ck('Capitalize &names of days', 'acDays', true), ck('Correct accidental use of cAPS &LOCK key', 'acCapsLock', true), ck('Replace &text as you type', 'acReplace', true),
      h('div', { class: 'row', style: 'gap:8px' }, ui.field('&Replace:', from), ui.field('&With:', to)), box,
      h('div', { class: 'row', style: 'gap:8px;justify-content:flex-end' }, ui.button('&Add', () => { if (from.value) { AC.setEntry(from.value, to.value); draw(); } }), ui.button('&Delete', () => { if (from.value) { AC.setEntry(from.value, null); from.value = to.value = ''; draw(); } }))) }]),
    buttons: [{ label: 'OK', primary: true, onClick: () => A().saveOpts() }, { label: 'Cancel' }] });
  };
  D.autocorrect = D.autoCorrect;

  /* ================================================================ protection */
  /** Excel's legacy 16-bit password hash (sheetProtection/@password) */
  D.legacyHash = function (pw) {
    let hash = 0;
    const chars = Array.from(String(pw)).map((c) => c.charCodeAt(0) & 0x7fff);
    for (let i = chars.length - 1; i >= 0; i--) { hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff); hash ^= chars[i]; }
    hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
    hash ^= chars.length;
    hash ^= 0xce4b;
    return hash.toString(16).toUpperCase().padStart(4, '0');
  };
  /** check a password against stored protection attributes; true when it unlocks */
  async function checkPw(attrs, pw) {
    if (!attrs) return true;
    if (attrs.password) return D.legacyHash(pw) === String(attrs.password).toUpperCase();
    if (attrs.hashValue && attrs.algorithmName) {
      const alg = { 'SHA-512': 'SHA-512', 'SHA-256': 'SHA-256', 'SHA-384': 'SHA-384', 'SHA-1': 'SHA-1' }[attrs.algorithmName];
      if (!alg) return false;
      const b64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
      const salt = b64(attrs.saltValue || '');
      const enc = new Uint8Array(Array.from(String(pw)).flatMap((ch) => { const c = ch.charCodeAt(0); return [c & 255, c >> 8]; }));
      let hbuf = await L.sha.digest(alg, new Uint8Array([...salt, ...enc]));   /* common/sha.js: also over plain http */
      const n = +attrs.spinCount || 0;
      for (let i = 0; i < n; i++) { const it = new Uint8Array(4); new DataView(it.buffer).setUint32(0, i, true); hbuf = await L.sha.digest(alg, new Uint8Array([...hbuf, ...it])); }
      return btoa(String.fromCharCode(...hbuf)) === attrs.hashValue;
    }
    return !attrs.password && !attrs.hashValue;
  }
  D.checkPw = checkPw;
  const askPw = (title) => new Promise((res) => {
    const p = h('input', { type: 'password', autocomplete: 'off' });
    let v = null;
    ui.dialog({ title, width: 280, body: ui.field('&Password:', p), buttons: [{ label: 'OK', primary: true, onClick: () => { v = p.value; } }, { label: 'Cancel' }] }).done.then(() => res(v));
  });
  D.protectSheet = async function () {
    const s = sh(), w = wb();
    if (s.protection) {
      const a = s.protection.attrs || {};
      if (a.password || a.hashValue) {
        const pw = await askPw('Unprotect Sheet');
        if (pw == null) return;
        if (!(await checkPw(a, pw))) { ui.msg('The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.', { icon: 'warn' }); return; }
      }
      tryRun(() => O.setProtection(s, null));
      return;
    }
    const allow = { selectLockedCells: true, selectUnlockedCells: true, formatCells: false, formatColumns: false, formatRows: false, insertColumns: false, insertRows: false, insertHyperlinks: false, deleteColumns: false, deleteRows: false, sort: false, autoFilter: false, pivotTables: false, objects: false, scenarios: false };
    const labels = { selectLockedCells: 'Select locked cells', selectUnlockedCells: 'Select unlocked cells', formatCells: 'Format cells', formatColumns: 'Format columns', formatRows: 'Format rows', insertColumns: 'Insert columns', insertRows: 'Insert rows', insertHyperlinks: 'Insert hyperlinks', deleteColumns: 'Delete columns', deleteRows: 'Delete rows', sort: L.N('Sort'), autoFilter: 'Use AutoFilter', pivotTables: 'Use PivotTable reports', objects: 'Edit objects', scenarios: 'Edit scenarios' };
    const pw = h('input', { type: 'password', autocomplete: 'new-password', style: 'width:100%' });
    const box = h('div', { class: 'lbox checks' }, ...Object.keys(allow).map((k) => ui.check(labels[k], allow[k], (x) => { allow[k] = x; })));
    ui.dialog({ title: 'Protect Sheet', width: 340, body: h('div', { class: 'col' }, ui.check('&Protect worksheet and contents of locked cells', true, null), ui.field('Password to unprotect sheet:', pw), h('label', { text: 'Allow all users of this worksheet to:' }), box),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const finish = () => {
          /* attributes in the file mean "this action is protected" (1), so allowed actions are written as 0 */
          const attrs = { sheet: '1', objects: allow.objects ? '0' : '1', scenarios: allow.scenarios ? '0' : '1' };
          for (const k of ['formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'sort', 'autoFilter', 'pivotTables']) attrs[k] = allow[k] ? '0' : '1';
          if (!allow.selectLockedCells) attrs.selectLockedCells = '1';
          if (!allow.selectUnlockedCells) attrs.selectUnlockedCells = '1';
          if (pw.value) attrs.password = D.legacyHash(pw.value);
          tryRun(() => O.setProtection(s, { attrs }));
        };
        if (pw.value) {
          const p2 = h('input', { type: 'password', autocomplete: 'new-password' });
          ui.dialog({ title: 'Confirm Password', width: 300, body: h('div', { class: 'col' }, ui.field('&Reenter password to proceed.', p2), h('div', { class: 'hint', text: 'Caution: If you lose or forget the password, it cannot be recovered.' })), buttons: [{ label: 'OK', primary: true, onClick: () => { if (p2.value !== pw.value) { ui.msg('Confirmation password is not identical.', { icon: 'warn' }); return false; } finish(); } }, { label: 'Cancel' }] });
        } else finish();
        void w;
      } }, { label: 'Cancel' }] });
  };
  D.protectWorkbook = async function () {
    const w = wb();
    if (w.protection) {
      const a = w.protection.attrs || {};
      const legacy = a.workbookPassword ? { password: a.workbookPassword } : a.workbookHashValue ? { hashValue: a.workbookHashValue, saltValue: a.workbookSaltValue, spinCount: a.workbookSpinCount, algorithmName: a.workbookAlgorithmName } : null;
      if (legacy) { const pw = await askPw('Unprotect Workbook'); if (pw == null) return; if (!(await checkPw(legacy, pw))) { ui.msg('The password you supplied is not correct.', { icon: 'warn' }); return; } }
      const old = w.protection;
      O.tx(w, 'Unprotect Workbook', () => { w.protection = null; w.undo.op(() => { w.protection = old; }, () => { w.protection = null; }); });
      ui.refresh();
      return;
    }
    let structure = true, windows = false;
    const pw = h('input', { type: 'password', autocomplete: 'new-password' });
    ui.dialog({ title: 'Protect Workbook', width: 300, body: h('div', { class: 'col' }, ui.group('Protect workbook for', ui.check('&Structure', true, (x) => { structure = x; }), ui.check('&Windows', false, (x) => { windows = x; })), ui.field('&Password (optional):', pw)),
      buttons: [{ label: 'OK', primary: true, onClick: () => { const attrs = {}; if (pw.value) attrs.workbookPassword = D.legacyHash(pw.value); const p = { attrs, structure, windows }; O.tx(w, 'Protect Workbook', () => { w.protection = p; w.undo.op(() => { w.protection = null; }, () => { w.protection = p; }); }); ui.refresh(); } }, { label: 'Cancel' }] });
  };

  /* ================================================================ Hyperlink */
  /** Insert ▸ Hyperlink (common ui.hyperlinkDialog) on the active cell: Text to display is the cell's text, and only
   *  a changed text replaces the cell's value (a number stays a number) */
  D.hyperlink = function () {
    const s = sh(), w = wb(), sel = G().sel();
    if (G().guardProtect('insertHyperlinks')) return;
    const cur = s.links.find((l) => M.rangeContains(l.ref, sel.r, sel.c));
    const cell = s.get(sel.r, sel.c);
    const shown = cur && cur.display != null ? cur.display : L.csv.cellText(s, sel.r, sel.c);
    const loc = cur && cur.location ? String(cur.location) : '';
    const named = loc && !/!/.test(loc) ? 'name:' + loc : null;
    const places = [{ label: L.t('Cell Reference'), group: true }].concat(w.sheets.map((x) => ({ value: x.name, label: x.name })));
    const names = w.names.filter((n) => !n.hidden && !/^_xlnm/.test(n.name));
    if (names.length) places.push({ label: L.t('Defined Names'), group: true }, ...names.map((n) => ({ value: 'name:' + n.name, label: n.name })));
    const sheetOf = loc && /!/.test(loc) ? loc.slice(0, loc.lastIndexOf('!')).replace(/^'|'$/g, '').replace(/''/g, "'") : s.name;
    ui.hyperlinkDialog({
      link: cur ? { url: cur.target, place: loc ? named || sheetOf : undefined, tip: cur.tooltip } : null,
      text: shown, textLocked: !!(cell && cell.f != null), places, defaultPlace: s.name,
      ref: { label: 'Type the cell refere&nce:', value: loc && /!/.test(loc) ? loc.slice(loc.lastIndexOf('!') + 1) : 'A1' },
      remove: () => tryRun(() => O.clear(s, [{ r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c }], 'hyperlinks')),
      apply: (r) => tryRun(() => {
        const l = { ref: { r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c } };
        if (r.kind === 'place') { const v = r.place != null ? String(r.place) : s.name; l.location = v.startsWith('name:') ? v.slice(5) : F.quoteSheet(v) + '!' + (r.ref || 'A1'); }
        else l.target = r.url;
        const changed = r.text != null && r.text !== shown;
        l.display = (r.text != null ? r.text : shown) || l.target || l.location;
        if (r.tip) l.tooltip = r.tip;
        O.tx(w, L.N('Hyperlink'), () => {
          O.setLinks(s, s.links.filter((x) => !M.rangeContains(x.ref, sel.r, sel.c)).concat([l]));
          const c0 = s.get(sel.r, sel.c);
          if (!c0 || c0.f == null) {
            if (!c0 || c0.v == null || changed) O.setValue(s, sel.r, sel.c, l.display, true);
            const c2 = s.get(sel.r, sel.c);
            const stl = w.cellStyles.find((x) => x.name === 'Hyperlink');
            const n = Object.assign({}, c2);
            n.s = w.styles.derive(c2.s || 0, stl ? { font: Object.assign({}, w.styles.get(stl.style).font) } : { font: { u: 'single', color: M.rgb('#0000FF') } });
            O.put(s, sel.r, sel.c, n);
          }
        });
      }),
    });
  };

  /* ================================================================ misc */
  D.links = function () {
    const w = wb();
    const rows = (w.externalLinks || []).map((e, i) => h('tr', null, h('td', { translate: 'no', text: decodeURIComponent(String(e.relTarget || '').split(/[\\/]/).pop()) }), h('td', { text: 'Worksheet' }), h('td', { text: 'A' }), h('td', { text: 'Values cached in this workbook' })));
    ui.dialog({ title: 'Edit Links', width: 520, body: h('div', { class: 'col' }, h('table', { class: 'fr-table' }, h('tr', null, h('th', { text: 'Source' }), h('th', { text: 'Type' }), h('th', { text: 'Update' }), h('th', { text: 'Status' })), ...rows), h('div', { class: 'hint', text: 'Formulas that refer to other workbooks use the values Excel saved with this file. Break Link replaces those formulas with their values.' })),
      buttons: [{ label: '&Break Link', onClick: () => { tryRun(() => { O.tx(w, 'Break Links', () => { for (const s of w.sheets) s.rows.forEach((row, r) => { if (row) row.cells.forEach((cl, c) => { if (cl && cl.f != null && /\[\d+\]/.test(cl.f)) O.setValue(s, r, c, cl.v, true); }); }); }); w.externalLinks = []; }); } }, { label: 'Close', primary: true }] });
  };
  D.customize = () => ui.customize({ toolbars: ['standard', 'formatting', 'borders', 'chart', 'drawing', 'auditing', 'reviewing', 'picture', 'list'], opts: A().opts, save: () => A().saveOpts(), shortcuts: A().SHORTCUTS });
  D.about = () => ui.about({ name: 'Ledger', text: [
    L.N('A spreadsheet that runs entirely in your browser. Workbooks are read and written as Office Open XML (.xlsx, .xlsm), CSV and text; nothing is uploaded.'),
    'The calculation engine implements more than 450 worksheet functions, array and dynamic-array formulas, structured references and iterative calculation.'] });
})(typeof window !== 'undefined' ? window : globalThis);
