/* Quire — dialogs, part 1: shared controls and the Format menu dialogs. */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const { h } = L;
  const ui = L.ui;
  const G = (L.dlg = L.dlg || {});
  const doc = () => D.doc;
  const A = () => L.app;
  let uid = 0;
  const nid = (p) => (p || 'd') + ++uid;

  /* ================= shared controls ================= */
  const UNIT_LABEL = { in: '"', cm: ' cm', mm: ' mm', pt: ' pt', pi: ' pi' };
  const UNIT_PT = { in: 72, cm: 28.3465, mm: 2.83465, pt: 1, pi: 12 };
  /** length spinner showing the user's unit, value in points */
  G.len = function (pt, o) {
    o = o || {};
    const unit = o.unit || D.unit || 'in';
    const k = UNIT_PT[unit];
    const step = o.step != null ? o.step : unit === 'in' ? 0.1 : unit === 'cm' ? 0.1 : unit === 'mm' ? 1 : unit === 'pt' ? 1 : 0.5;
    const dec = unit === 'in' || unit === 'cm' ? 2 : unit === 'pt' ? 1 : 1;
    const fmt = (v) => (v == null || v === '' ? (o.autoLabel || '') : L.round(v / k, dec) + UNIT_LABEL[unit]);
    const id = o.id || nid('len');
    const inp = h('input', { type: 'text', id, class: 'spin-in', value: fmt(pt), autocomplete: 'off', inputmode: 'decimal' });
    let val = pt;
    const parse = () => { const s = inp.value.trim(); if (!s && o.autoLabel) return null; if (o.autoLabel && s.toLowerCase() === o.autoLabel.toLowerCase()) return null; const v = D.parseLen(s, unit); return v == null || isNaN(v) ? val : v; };
    const set = (v, quiet) => { if (v != null) v = L.clamp(v, o.min == null ? -1e6 : o.min, o.max == null ? 1e6 : o.max); val = v; inp.value = fmt(v); if (!quiet && o.onChange) o.onChange(v); };
    const up = h('button', { type: 'button', class: 'spin-b up', tabindex: '-1', 'aria-label': 'Increase', html: '<span></span>' });
    const dn = h('button', { type: 'button', class: 'spin-b dn', tabindex: '-1', 'aria-label': 'Decrease', html: '<span></span>' });
    up.addEventListener('click', () => set((parse() || 0) + step * k));
    dn.addEventListener('click', () => set((parse() || 0) - step * k));
    inp.addEventListener('change', () => set(parse()));
    inp.addEventListener('keydown', (e) => { if (e.key === 'ArrowUp') { e.preventDefault(); up.click(); } if (e.key === 'ArrowDown') { e.preventDefault(); dn.click(); } });
    const wrap = h('span', { class: 'spin' }, inp, h('span', { class: 'spin-bs' }, up, dn));
    wrap.input = inp;
    wrap.get = () => parse();
    wrap.set = (v) => set(v, true);
    wrap.setDisabled = (dis) => { inp.disabled = dis; up.disabled = dis; dn.disabled = dis; wrap.classList.toggle('dis', dis); };
    return wrap;
  };
  /** plain number spinner */
  G.num = function (v, o) {
    o = o || {};
    const s = ui.spin(Object.assign({ value: v == null ? 0 : v, step: 1, dec: 0, min: 0 }, o, { id: o.id || nid('num') }));
    if (v == null && o.autoLabel) s.input.value = o.autoLabel;
    return s;
  };
  /** colour picker button. value: 'auto' | 'RRGGBB' | null(no colour) */
  G.color = function (value, o) {
    o = o || {};
    let val = value;
    const sw = h('span', { class: 'cpk-sw' });
    const lbl = h('span', { class: 'cpk-l' });
    const btn = h('button', { type: 'button', class: 'cpk', id: o.id || nid('col') }, sw, lbl, h('span', { class: 'dd-arrow' }));
    const show = () => { sw.style.background = val && val !== 'auto' ? '#' + val : o.autoColor || (val === 'auto' ? '#000' : 'transparent'); lbl.textContent = val === 'auto' ? (o.autoText || 'Automatic') : !val ? (o.noneText || 'No Color') : ui.colorName('#' + val); };
    show();
    btn.addEventListener('click', () => ui.colorMenu(btn, { mode: o.mode || 'font', automatic: o.automatic !== false }, (v) => {
      if (v && v.auto) val = 'auto';
      else if (v && v.none) val = null;
      else if (typeof v === 'string') val = v.replace('#', '').toUpperCase();
      show();
      if (o.onChange) o.onChange(val);
    }));
    btn.get = () => val;
    btn.set = (v) => { val = v; show(); };
    return btn;
  };
  /** list box with optional text input on top (Word's font/size/style pickers) */
  G.listBox = function (items, value, o) {
    o = o || {};
    const box = h('div', { class: 'listbox', style: o.height ? `height:${o.height}px` : '' });
    const inp = o.input ? h('input', { type: 'text', class: 'listbox-in', id: o.id || nid('lb'), value: value == null ? '' : String(value), autocomplete: 'off' }) : null;
    let list = items.slice();
    let cur = value;
    const render = () => {
      L.clear(box);
      for (const it of list) {
        const v = typeof it === 'object' ? it.value : it;
        const el = h('div', { class: 'li' + (String(v) === String(cur) ? ' on' : ''), text: typeof it === 'object' ? it.label : String(it) });
        if (typeof it === 'object' && it.style) el.style.cssText += it.style;
        if (typeof it === 'object' && it.group) { el.className = 'li-group'; box.appendChild(el); continue; }
        el.addEventListener('click', () => { pick(v); });
        el.addEventListener('dblclick', () => { pick(v); if (o.onDbl) o.onDbl(v); });
        box.appendChild(el);
      }
      const on = box.querySelector('.li.on');
      if (on) setTimeout(() => on.scrollIntoView({ block: 'nearest' }), 0);
    };
    const pick = (v, fromInput) => { cur = v; if (inp && !fromInput) inp.value = String(v); for (const el of box.querySelectorAll('.li')) el.classList.toggle('on', el.textContent === (list.find((x) => String(typeof x === 'object' ? x.value : x) === String(v)) || {}).label || el.textContent === String(v)); if (o.onChange) o.onChange(v); };
    if (inp) inp.addEventListener('input', () => {
      const v = inp.value;
      const hit = list.find((x) => String(typeof x === 'object' ? x.label : x).toLowerCase().startsWith(v.toLowerCase()));
      cur = v;
      for (const el of box.querySelectorAll('.li')) el.classList.toggle('on', !!hit && el.textContent === String(typeof hit === 'object' ? hit.label : hit));
      const on = box.querySelector('.li.on');
      if (on) on.scrollIntoView({ block: 'nearest' });
      if (o.onChange) o.onChange(v);
    });
    render();
    const wrap = h('div', { class: 'col', style: 'gap:2px' }, inp, box);
    wrap.get = () => (inp ? inp.value : cur);
    wrap.set = (v) => { cur = v; if (inp) inp.value = v == null ? '' : String(v); render(); };
    wrap.setItems = (it) => { list = it; render(); };
    wrap.input = inp;
    return wrap;
  };
  /** labelled field */
  G.f = (label, ctl, cls) => ui.field(label, ctl, cls ? { cls } : null);
  G.row = (...kids) => h('div', { class: 'row' }, ...kids);
  G.col = (...kids) => h('div', { class: 'col' }, ...kids);
  G.radios = function (name, options, value, onChange) {
    const box = h('div', { class: 'col', style: 'gap:2px' });
    const els = options.map(([v, l]) => { const r = ui.radio(name, l, String(v) === String(value), () => { box.value = v; if (onChange) onChange(v); }); box.appendChild(r); return r; });
    box.value = value;
    box.get = () => box.value;
    box.set = (v) => { box.value = v; els.forEach((r, i) => { r.input.checked = String(options[i][0]) === String(v); }); };
    return box;
  };
  const FONTS = () => Array.from(new Set(L.FONT_LIST.concat(Object.values(doc().styles).map((s) => s.rPr && s.rPr.font).filter(Boolean), [doc().defaults.rPr.font]))).sort((a, b) => a.localeCompare(b));
  G.FONTS = FONTS;
  /** mixed values across the selection (undefined when mixed) */
  const uni = (k) => E.uniformRun(k);
  /** apply props and remember the dialog choices */
  function preview(el, r) {
    el.style.cssText = R.runCSS(Object.assign({}, r, { sz: Math.min(r.sz || 12, 30) }), 1) + ';white-space:nowrap;overflow:hidden';
  }

  /* ================= Font ================= */
  const UNDERLINES = [['none', '(none)'], ['single', 'Single'], ['words', 'Words only'], ['double', 'Double'], ['thick', 'Thick'], ['dotted', 'Dotted'], ['dottedHeavy', 'Thick dotted'], ['dash', 'Dash'], ['dashedHeavy', 'Thick dash'], ['dashLong', 'Long dash'], ['dotDash', 'Dot dash'], ['dotDotDash', 'Dot dot dash'], ['wave', 'Wave'], ['wavyDouble', 'Double wave'], ['wavyHeavy', 'Heavy wave']];
  G.UNDERLINES = UNDERLINES;
  /**
   * opts: {target: rPr object to edit (style dialogs) — returns via onOK(newRPr)}
   */
  G.font = function (opts) {
    opts = opts || {};
    const base = opts.target ? Object.assign({}, opts.base || {}, opts.target) : E.curRun();
    const val = (k) => (opts.target ? base[k] : uni(k));
    const st = { font: val('font'), sz: val('sz'), b: val('b'), i: val('i'), u: val('u'), uColor: val('uColor'), color: val('color'), strike: val('strike'), dstrike: val('dstrike'), vert: val('vert'), shadow: val('shadow'), outline: val('outline'), emboss: val('emboss'), imprint: val('imprint'), smallCaps: val('smallCaps'), caps: val('caps'), hidden: val('hidden'), w: val('w'), spacing: val('spacing'), position: val('position'), kern: val('kern'), effect: val('effect') };
    const changed = new Set();
    const set = (k, v) => { st[k] = v; changed.add(k); upd(); };
    const fontLB = G.listBox(FONTS(), st.font || '', { input: true, height: 110, id: 'fd-font', onChange: (v) => set('font', v) });
    const styleName = () => (st.b && st.i ? 'Bold Italic' : st.b ? 'Bold' : st.i ? 'Italic' : st.b === undefined || st.i === undefined ? '' : 'Regular');
    const styleLB = G.listBox(['Regular', 'Italic', 'Bold', 'Bold Italic'], styleName(), { input: true, height: 110, id: 'fd-style', onChange: (v) => { const s = String(v).toLowerCase(); st.b = /bold/.test(s); st.i = /italic/.test(s); changed.add('b'); changed.add('i'); upd(); } });
    const sizeLB = G.listBox(L.SIZE_LIST.map(String), st.sz != null ? String(L.round(st.sz, 1)) : '', { input: true, height: 110, id: 'fd-size', onChange: (v) => { const n = parseFloat(v); if (n > 0) set('sz', L.clamp(Math.round(n * 2) / 2, 1, 1638)); } });
    const color = G.color(st.color == null ? 'auto' : st.color, { id: 'fd-color', onChange: (v) => set('color', v === 'auto' ? undefined : v) });
    const ul = ui.select(UNDERLINES, st.u || 'none', (v) => set('u', v === 'none' ? (opts.target ? undefined : 'none') : v), { id: 'fd-ul' });
    const ulColor = G.color(st.uColor || 'auto', { id: 'fd-ulc', onChange: (v) => set('uColor', v === 'auto' ? undefined : v) });
    const fx = [['strike', 'Stri&kethrough'], ['dstrike', 'Double strikethrou&gh'], ['sup', 'Su&perscript'], ['sub', 'Su&bscript'], ['shadow', 'S&hadow'], ['outline', '&Outline'], ['emboss', '&Emboss'], ['imprint', 'En&grave'], ['smallCaps', 'S&mall caps'], ['caps', '&All caps'], ['hidden', '&Hidden']];
    const fxEls = {};
    const fxBox = h('div', { class: 'grid3' });
    for (const [k, l] of fx) {
      const cur = k === 'sup' ? st.vert === 'superscript' : k === 'sub' ? st.vert === 'subscript' : !!st[k];
      const c = ui.check(l, cur, (v) => {
        if (k === 'sup' || k === 'sub') { st.vert = v ? (k === 'sup' ? 'superscript' : 'subscript') : undefined; changed.add('vert'); fxEls[k === 'sup' ? 'sub' : 'sup'].input.checked = false; }
        else { st[k] = v || (opts.target ? undefined : false); changed.add(k); if (v && k === 'strike') { st.dstrike = undefined; fxEls.dstrike.input.checked = false; changed.add('dstrike'); } if (v && k === 'dstrike') { st.strike = undefined; fxEls.strike.input.checked = false; changed.add('strike'); } if (v && k === 'caps') { st.smallCaps = undefined; fxEls.smallCaps.input.checked = false; changed.add('smallCaps'); } if (v && k === 'smallCaps') { st.caps = undefined; fxEls.caps.input.checked = false; changed.add('caps'); } if (v && k === 'emboss') { st.imprint = undefined; fxEls.imprint.input.checked = false; changed.add('imprint'); } if (v && k === 'imprint') { st.emboss = undefined; fxEls.emboss.input.checked = false; changed.add('emboss'); } }
        upd();
      });
      if (k !== 'sup' && k !== 'sub' && st[k] === undefined && !opts.target && E.uniformRun(k) === undefined && !E.collapsed()) c.input.indeterminate = true;
      fxEls[k] = c;
      fxBox.appendChild(c);
    }
    const pv = h('div', { class: 'preview-box font-preview' }, h('span', { text: '' }));
    const pvCap = h('div', { class: 'prev-cap', text: '' });
    const sample = () => { if (opts.target) return opts.sample || 'Sample Text'; const t = E.collapsed() ? '' : O.textRange(...E.range(), ' ').trim().slice(0, 40); return t || (st.font || 'Times New Roman'); };
    function upd() {
      const r = {};
      for (const k in st) if (st[k] !== undefined && st[k] !== false) r[k] = st[k];
      if (r.color === 'auto') delete r.color;
      if (r.u === 'none') delete r.u;
      preview(pv.firstChild, r);
      pv.firstChild.textContent = sample();
      pvCap.textContent = st.font && !L.FONT_LIST.includes(st.font) ? 'This font is not installed; the closest available font is used on screen.' : 'This is a TrueType font. This font will be used on both printer and screen.';
    }
    upd();
    /* character spacing */
    const scale = ui.select([200, 150, 100, 90, 80, 66, 50, 33].map((x) => [x, x + '%']), st.w || 100, (v) => set('w', +v === 100 ? undefined : +v), { id: 'fd-scale' });
    const spKind = ui.select([['normal', 'Normal'], ['exp', 'Expanded'], ['cond', 'Condensed']], !st.spacing ? 'normal' : st.spacing > 0 ? 'exp' : 'cond', (v) => { if (v === 'normal') { spBy.set(0); set('spacing', undefined); } else { const n = Math.abs(spBy.get() || 1); spBy.set(n); set('spacing', v === 'exp' ? n : -n); } }, { id: 'fd-spk' });
    const spBy = ui.spin({ value: Math.abs(st.spacing || 0), step: 0.1, dec: 1, unit: ' pt', min: 0, max: 1584, id: 'fd-spby', onChange: (v) => { if (spKind.value === 'normal' && v) spKind.value = 'exp'; set('spacing', spKind.value === 'cond' ? -v : v || undefined); } });
    const posKind = ui.select([['normal', 'Normal'], ['raised', 'Raised'], ['lowered', 'Lowered']], !st.position ? 'normal' : st.position > 0 ? 'raised' : 'lowered', (v) => { if (v === 'normal') { posBy.set(0); set('position', undefined); } else { const n = Math.abs(posBy.get() || 3); posBy.set(n); set('position', v === 'raised' ? n : -n); } }, { id: 'fd-posk' });
    const posBy = ui.spin({ value: Math.abs(st.position || 0), step: 1, dec: 1, unit: ' pt', min: 0, max: 1584, id: 'fd-posby', onChange: (v) => { if (posKind.value === 'normal' && v) posKind.value = 'raised'; set('position', posKind.value === 'lowered' ? -v : v || undefined); } });
    const kern = ui.check('&Kerning for fonts:', !!st.kern, (v) => set('kern', v ? kernAt.get() || 14 : undefined));
    const kernAt = ui.spin({ value: st.kern || 14, step: 1, dec: 1, unit: ' pt', min: 1, id: 'fd-kern', onChange: (v) => { if (kern.input.checked) set('kern', v); } });
    const anim = ui.select([['', '(none)'], ['blinkBackground', 'Blinking Background'], ['lights', 'Las Vegas Lights'], ['antsBlack', 'Marching Black Ants'], ['antsRed', 'Marching Red Ants'], ['shimmer', 'Shimmer'], ['sparkle', 'Sparkle Text']], st.effect || '', (v) => set('effect', v || undefined), { id: 'fd-anim', size: 7, style: 'width:100%' });
    const tabs = ui.tabs([
      { label: 'Fo&nt', body: h('div', null, G.row(h('div', { class: 'col', style: 'flex:2;min-width:150px' }, h('label', { for: 'fd-font', html: '<u>F</u>ont:' }), fontLB), h('div', { class: 'col', style: 'flex:1.2;min-width:100px' }, h('label', { for: 'fd-style', html: 'Font st<u>y</u>le:' }), styleLB), h('div', { class: 'col', style: 'flex:0.7;min-width:60px' }, h('label', { for: 'fd-size', html: '<u>S</u>ize:' }), sizeLB)),
        G.row(G.f('Font &color:', color), G.f('&Underline style:', ul), G.f('Underline color:', ulColor)), ui.group('Effects', fxBox)) },
      { label: 'Cha&racter Spacing', body: h('div', { class: 'col' }, G.f('S&cale:', scale, 'wide'), G.row(G.f('&Spacing:', spKind, 'wide'), G.f('&By:', spBy, 'narrow')), G.row(G.f('&Position:', posKind, 'wide'), G.f('B&y:', posBy, 'narrow')), G.row(kern, G.f('Poi&nts and above', kernAt))) },
      { label: 'Te&xt Effects', body: h('div', { class: 'col' }, h('label', { for: 'fd-anim', text: 'Animations:' }), anim, h('div', { class: 'tp-note', text: 'Animated text effects are kept in the file; they display as plain text, as in later versions of Word.' })) },
    ], opts.tab || 0);
    const body = h('div', null, tabs, h('div', { class: 'cd-cap', text: 'Preview' }), pv, pvCap);
    ui.dialog({
      title: 'Font', width: 470, body,
      buttons: [
        { label: '&Default...', onClick: () => { setDefault(); return false; } },
        { label: 'OK', primary: true, onClick: () => apply() },
        { label: 'Cancel' },
      ],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
    function result() { const out = {}; for (const k of changed) out[k] = st[k] === false ? (opts.target ? undefined : false) : st[k]; if ('u' in out && out.u === 'none' && opts.target) out.u = undefined; return out; }
    function apply() {
      const r = result();
      if (opts.onOK) { opts.onOK(r); return; }
      if (!changed.size) return;
      E.formatRun(r, 'Font');
    }
    async function setDefault() {
      const ok = await ui.msg(`Do you want to change the default font to ${st.font || 'Times New Roman'}, ${st.sz || 12} pt?\nThis change will affect all new documents based on the NORMAL template.`, { icon: 'question', buttons: ['&Yes', '&No'] });
      if (ok !== 0) return;
      const d = doc();
      E.edit('Default Font', () => { D.touchKey(d, 'defaults'); d.defaults.rPr = Object.assign({}, d.defaults.rPr, st.font ? { font: st.font } : {}, st.sz ? { sz: st.sz } : {}); D.touchKey(d, 'styles'); if (d.styles.Normal) { d.styles.Normal.rPr = Object.assign({}, d.styles.Normal.rPr); delete d.styles.Normal.rPr.font; delete d.styles.Normal.rPr.sz; } D.stylesChanged(); return E.sel; });
      L.store.set('defaultFont', { font: st.font, sz: st.sz });
    }
  };

  /* ================= Paragraph ================= */
  G.paragraph = function (opts) {
    opts = opts || {};
    const pp = opts.target ? Object.assign({}, opts.base || {}, opts.target) : E.curPara();
    const paras = opts.target ? [] : E.selectedParas();
    const same = (get) => { if (opts.target || paras.length < 2) return true; const v0 = JSON.stringify(get(D.pProps(doc(), paras[0]))); return paras.every((p) => JSON.stringify(get(D.pProps(doc(), p))) === v0); };
    const st = { jc: pp.jc || 'left', l: (pp.ind && pp.ind.l) || 0, r: (pp.ind && pp.ind.r) || 0, fl: (pp.ind && pp.ind.fl) || 0, b: (pp.sp && pp.sp.b) || 0, a: (pp.sp && pp.sp.a) || 0, bAuto: !!(pp.sp && pp.sp.bAuto), aAuto: !!(pp.sp && pp.sp.aAuto), line: (pp.sp && pp.sp.line) || 1, rule: (pp.sp && pp.sp.rule) || 'auto', outline: pp.outline, contextual: !!pp.contextual, widow: pp.widow !== false, keepNext: !!pp.keepNext, keepLines: !!pp.keepLines, pageBreakBefore: !!pp.pageBreakBefore, suppressLineNumbers: !!pp.suppressLineNumbers, suppressAutoHyphens: !!pp.suppressAutoHyphens };
    const changed = new Set();
    const set = (k, v) => { st[k] = v; changed.add(k); drawPv(); };
    const align = ui.select([['left', 'Left'], ['center', 'Centered'], ['right', 'Right'], ['both', 'Justified'], ['distribute', 'Distributed']], st.jc, (v) => set('jc', v), { id: 'pd-al' });
    const lvl = ui.select([['', 'Body text']].concat([1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => [String(n - 1), 'Level ' + n])), st.outline == null || st.outline >= 9 ? '' : String(st.outline), (v) => set('outline', v === '' ? 9 : +v), { id: 'pd-lvl' });
    const left = G.len(st.l, { id: 'pd-l', min: -1584, onChange: (v) => set('l', v) });
    const right = G.len(st.r, { id: 'pd-r', min: -1584, onChange: (v) => set('r', v) });
    const special = ui.select([['none', '(none)'], ['first', 'First line'], ['hang', 'Hanging']], st.fl > 0 ? 'first' : st.fl < 0 ? 'hang' : 'none', (v) => { if (v === 'none') { by.set(0); set('fl', 0); } else { const n = Math.abs(by.get() || 36); by.set(n); set('fl', v === 'first' ? n : -n); } }, { id: 'pd-sp' });
    const by = G.len(Math.abs(st.fl), { id: 'pd-by', min: 0, onChange: (v) => { if (special.value === 'none' && v) special.value = 'first'; set('fl', special.value === 'hang' ? -v : v); } });
    if (!same((p) => p.ind)) { left.input.value = ''; right.input.value = ''; }
    const before = ui.spin({ value: st.b, step: 6, dec: 1, unit: ' pt', min: 0, max: 1584, id: 'pd-b', onChange: (v) => { set('b', v); st.bAuto = false; changed.add('bAuto'); } });
    const after = ui.spin({ value: st.a, step: 6, dec: 1, unit: ' pt', min: 0, max: 1584, id: 'pd-a', onChange: (v) => { set('a', v); st.aAuto = false; changed.add('aAuto'); } });
    if (st.bAuto) before.input.value = 'Auto';
    if (st.aAuto) after.input.value = 'Auto';
    const lsKind = st.rule === 'exact' ? 'exact' : st.rule === 'atLeast' ? 'atLeast' : st.line === 1 ? 'single' : st.line === 1.5 ? '1.5' : st.line === 2 ? 'double' : 'multiple';
    const ls = ui.select([['single', 'Single'], ['1.5', '1.5 lines'], ['double', 'Double'], ['atLeast', 'At least'], ['exact', 'Exactly'], ['multiple', 'Multiple']], lsKind, (v) => {
      if (v === 'single') { st.line = 1; st.rule = 'auto'; at.set(''); }
      else if (v === '1.5') { st.line = 1.5; st.rule = 'auto'; }
      else if (v === 'double') { st.line = 2; st.rule = 'auto'; }
      else if (v === 'multiple') { st.line = st.rule === 'auto' && st.line > 0 ? st.line : 3; st.rule = 'auto'; at.set(st.line); }
      else { st.rule = v; st.line = st.rule === 'auto' ? 12 : Math.max(st.line > 6 ? st.line : 12, 1); at.set(st.line); }
      changed.add('line'); changed.add('rule'); drawPv();
    }, { id: 'pd-ls' });
    const at = ui.spin({ value: lsKind === 'single' || lsKind === '1.5' || lsKind === 'double' ? 0 : st.line, step: lsKind === 'multiple' ? 0.5 : 1, dec: 2, unit: '', min: 0, id: 'pd-at', fmt: (v) => (ls.value === 'atLeast' || ls.value === 'exact' ? L.round(v, 1) + ' pt' : v ? L.round(v, 2) + ' li' : ''), onChange: (v) => { if (ls.value === 'single' || ls.value === '1.5' || ls.value === 'double') ls.value = 'multiple'; st.line = v; st.rule = ls.value === 'atLeast' || ls.value === 'exact' ? ls.value : 'auto'; changed.add('line'); changed.add('rule'); drawPv(); } });
    if (lsKind === 'single' || lsKind === '1.5' || lsKind === 'double') at.input.value = '';
    const ctxChk = ui.check("&Don't add space between paragraphs of the same style", st.contextual, (v) => set('contextual', v));
    const pv = h('div', { class: 'para-preview', style: 'width:100%;height:110px' });
    function drawPv() {
      L.clear(pv);
      const k = 0.35;
      const sh = (len, dark, extra) => { const el = h(dark ? 'b' : 'i', { style: `width:${len}%;${extra || ''}` }); return el; };
      for (let i = 0; i < 3; i++) pv.appendChild(sh(100 - (i === 2 ? 30 : 0)));
      const box = h('div', { style: `display:flex;flex-direction:column;align-items:${{ center: 'center', right: 'flex-end' }[st.jc] || 'stretch'};margin:${Math.max(2, st.b * k)}px ${Math.max(0, st.r * k)}px ${Math.max(2, st.a * k)}px ${Math.max(0, st.l * k)}px;gap:${st.rule === 'auto' ? Math.max(1, (st.line - 1) * 5 + 2) : 2}px` });
      for (let i = 0; i < 5; i++) box.appendChild(sh(i === 4 ? 55 : 100, true, i === 0 ? `margin-left:${st.fl * k}px;width:calc(${st.jc === 'both' ? '100%' : '92%'} - ${Math.max(0, st.fl * k)}px)` : ''));
      pv.appendChild(box);
      for (let i = 0; i < 3; i++) pv.appendChild(sh(100));
    }
    drawPv();
    const pageChecks = [['widow', '&Widow/Orphan control'], ['keepNext', 'Keep with ne&xt'], ['keepLines', '&Keep lines together'], ['pageBreakBefore', 'Page &break before'], ['suppressLineNumbers', '&Suppress line numbers'], ['suppressAutoHyphens', "&Don't hyphenate"]].map(([k, l]) => ui.check(l, st[k], (v) => set(k, v)));
    const tabs = ui.tabs([
      { label: '&Indents and Spacing', body: h('div', null, ui.group('General', G.row(G.f('Ali&gnment:', align), G.f('&Outline level:', lvl))), ui.group('Indentation', G.row(h('div', { class: 'col' }, G.f('&Left:', left, 'narrow'), G.f('&Right:', right, 'narrow')), h('div', { class: 'col' }, G.f('&Special:', special, 'narrow'), G.f('B&y:', by, 'narrow')))), ui.group('Spacing', G.row(h('div', { class: 'col' }, G.f('&Before:', before, 'narrow'), G.f('Aft&er:', after, 'narrow')), h('div', { class: 'col' }, G.f('Li&ne spacing:', ls, 'wide'), G.f('&At:', at, 'wide'))), ctxChk), ui.group('Preview', pv)) },
      { label: 'Line and &Page Breaks', body: h('div', null, ui.group('Pagination', ...pageChecks.slice(0, 4)), ui.group('Formatting exceptions', ...pageChecks.slice(4))) },
    ], opts.tab || 0);
    ui.dialog({
      title: 'Paragraph', width: 470, body: tabs,
      buttons: [{ label: '&Tabs...', onClick: () => { apply(); setTimeout(() => G.tabs(opts), 0); } }, { label: 'OK', primary: true, onClick: () => apply() }, { label: 'Cancel' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
    function apply() {
      if (!changed.size) return;
      const fn = (pPr) => {
        if (changed.has('jc')) pPr.jc = st.jc === 'left' && !opts.target ? 'left' : st.jc;
        if (changed.has('outline')) pPr.outline = st.outline;
        if (changed.has('l') || changed.has('r') || changed.has('fl')) { pPr.ind = Object.assign({}, pPr.ind); if (changed.has('l')) pPr.ind.l = L.round(st.l, 2); if (changed.has('r')) pPr.ind.r = L.round(st.r, 2); if (changed.has('fl')) pPr.ind.fl = L.round(st.fl, 2); }
        if (['b', 'a', 'line', 'rule', 'bAuto', 'aAuto'].some((k) => changed.has(k))) {
          pPr.sp = Object.assign({}, pPr.sp);
          if (changed.has('b')) { pPr.sp.b = st.b; delete pPr.sp.bAuto; }
          if (changed.has('a')) { pPr.sp.a = st.a; delete pPr.sp.aAuto; }
          if (changed.has('line') || changed.has('rule')) { pPr.sp.line = st.line; pPr.sp.rule = st.rule; }
        }
        for (const k of ['contextual', 'keepNext', 'keepLines', 'pageBreakBefore', 'suppressLineNumbers', 'suppressAutoHyphens']) if (changed.has(k)) pPr[k] = st[k] || (opts.target ? undefined : false);
        if (changed.has('widow')) pPr.widow = st.widow;
      };
      if (opts.onOK) { const t = {}; fn(t); opts.onOK(t); return; }
      E.formatPara(fn, 'Paragraph');
    }
  };

  /* ================= Tabs ================= */
  G.tabs = function (opts) {
    opts = opts || {};
    const d = doc();
    const pp = opts.target ? Object.assign({}, opts.base || {}, opts.target) : E.curPara();
    let tabs = (pp.tabs || []).filter((t) => t.al !== 'clear' && t.al !== 'num').map((t) => Object.assign({}, t));
    const cleared = [];
    const pos = h('input', { type: 'text', id: 'td-pos', style: 'width:100%' });
    const list = G.listBox(tabs.map((t) => D.fmtLen(t.pos)), null, { height: 90, onChange: (v) => { const t = tabs.find((x) => D.fmtLen(x.pos) === v); if (t) { pos.value = v; al.set(t.al === 'start' ? 'left' : t.al === 'end' ? 'right' : t.al); lead.set(t.leader || 'none'); } } });
    const def = G.len(d.settings.defTab || 36, { id: 'td-def', min: 1 });
    const al = G.radios('td-al', [['left', '&Left'], ['center', '&Center'], ['right', '&Right'], ['decimal', '&Decimal'], ['bar', '&Bar']], 'left');
    const lead = G.radios('td-lead', [['none', '&1 None'], ['dot', '&2 .......'], ['hyphen', '&3 -------'], ['underscore', '&4 ______']], 'none');
    const clearInfo = h('div', { class: 'tp-note' });
    const refresh = () => { tabs.sort((a, b) => a.pos - b.pos); list.setItems(tabs.map((t) => D.fmtLen(t.pos))); clearInfo.textContent = cleared.length ? 'Tab stops to be cleared: ' + cleared.map((p) => D.fmtLen(p)).join(', ') : ''; };
    const setTab = () => { const p = D.parseLen(pos.value); if (p == null || isNaN(p)) return false; const i = tabs.findIndex((t) => Math.abs(t.pos - p) < 0.5); const t = { pos: L.round(p, 2), al: al.get(), leader: lead.get() === 'none' ? undefined : lead.get() }; if (i >= 0) tabs[i] = t; else tabs.push(t); refresh(); return true; };
    const body = h('div', null, G.row(h('div', { class: 'col', style: 'flex:1' }, h('label', { for: 'td-pos', html: '<u>T</u>ab stop position:' }), pos, list), h('div', { class: 'col', style: 'flex:1' }, G.f('De&fault tab stops:', def), clearInfo)),
      G.row(ui.group('Alignment', al), ui.group('Leader', lead)),
      h('div', { class: 'tp-btns' }, ui.button('&Set', () => setTab()), ui.button('Cl&ear', () => { const p = D.parseLen(pos.value); if (p == null) return; tabs = tabs.filter((t) => Math.abs(t.pos - p) > 0.5); cleared.push(p); refresh(); }), ui.button('Clear &All', () => { for (const t of tabs) cleared.push(t.pos); tabs = []; for (const t of D.pProps(d, E.sel ? E.sel.f.p : D.firstPara(d.main.blocks)).tabs || []) cleared.push(t.pos); refresh(); })));
    ui.dialog({
      title: 'Tabs', width: 380, body, focus: pos,
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (pos.value.trim() && !tabs.some((t) => D.fmtLen(t.pos) === pos.value.trim())) setTab();
        const dt = def.get();
        if (dt && Math.abs(dt - (d.settings.defTab || 36)) > 0.1) E.edit('Default Tabs', () => { D.touchKey(d, 'settings'); d.settings.defTab = dt; return E.sel; });
        const fn = (pPr, p) => {
          const styleTabs = p ? (D.styleProps(d, pPr.style || 'Normal').pPr.tabs || []) : (opts.base && opts.base.tabs) || [];
          const out = tabs.slice();
          for (const c of cleared) if (styleTabs.some((t) => Math.abs(t.pos - c) < 0.5) && !out.some((t) => Math.abs(t.pos - c) < 0.5)) out.push({ pos: c, al: 'clear' });
          pPr.tabs = out.length ? out.sort((a, b) => a.pos - b.pos) : undefined;
        };
        if (opts.onOK) { const t = {}; fn(t, null); opts.onOK(t); return; }
        E.formatPara(fn, 'Tabs');
      } }, { label: 'Cancel' }],
    });
  };

  /* ================= Borders and Shading ================= */
  const BORDER_STYLES = [['single', 'Single'], ['dotted', 'Dotted'], ['dashSmallGap', 'Small dash'], ['dashed', 'Dashed'], ['dotDash', 'Dot dash'], ['dotDotDash', 'Dot dot dash'], ['double', 'Double'], ['triple', 'Triple'], ['thinThickSmallGap', 'Thin-thick'], ['thickThinSmallGap', 'Thick-thin'], ['thinThickThinSmallGap', 'Thin-thick-thin'], ['wave', 'Wave'], ['doubleWave', 'Double wave'], ['dashDotStroked', 'Dash dot stroked'], ['threeDEmboss', '3-D emboss'], ['threeDEngrave', '3-D engrave'], ['outset', 'Outset'], ['inset', 'Inset']];
  const WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
  const SHD_PATTERNS = [['clear', 'Clear'], ['solid', 'Solid (100%)'], ['pct5', '5%'], ['pct10', '10%'], ['pct12', '12.5%'], ['pct15', '15%'], ['pct20', '20%'], ['pct25', '25%'], ['pct30', '30%'], ['pct35', '35%'], ['pct37', '37.5%'], ['pct40', '40%'], ['pct45', '45%'], ['pct50', '50%'], ['pct55', '55%'], ['pct60', '60%'], ['pct62', '62.5%'], ['pct65', '65%'], ['pct70', '70%'], ['pct75', '75%'], ['pct80', '80%'], ['pct85', '85%'], ['pct87', '87.5%'], ['pct90', '90%'], ['pct95', '95%'], ['horzStripe', 'Dk Horizontal'], ['vertStripe', 'Dk Vertical'], ['reverseDiagStripe', 'Dk Dwn Diagonal'], ['diagStripe', 'Dk Up Diagonal'], ['horzCross', 'Dk Grid'], ['diagCross', 'Dk Trellis'], ['thinHorzStripe', 'Lt Horizontal'], ['thinVertStripe', 'Lt Vertical'], ['thinReverseDiagStripe', 'Lt Dwn Diagonal'], ['thinDiagStripe', 'Lt Up Diagonal'], ['thinHorzCross', 'Lt Grid'], ['thinDiagCross', 'Lt Trellis']];
  G.BORDER_STYLES = BORDER_STYLES;
  function borderPanel(kind, init, o) {
    /* kind: 'para' | 'cell' | 'table' | 'page' */
    const sides = kind === 'page' ? ['top', 'bottom', 'left', 'right'] : kind === 'para' ? ['top', 'bottom', 'left', 'right', 'between'] : ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'];
    const bd = {};
    for (const s of sides) if (init && init[s] && R.hasBorder(init[s])) bd[s] = Object.assign({}, init[s]);
    const first = Object.values(bd)[0] || { val: 'single', sz: 0.5, color: 'auto' };
    let pen = { val: first.val, sz: first.sz || 0.5, color: first.color || 'auto' };
    const styleBox = G.listBox(BORDER_STYLES.map(([v, l]) => ({ value: v, label: l })), pen.val, { height: 120, onChange: (v) => { pen.val = v; } });
    const color = G.color(pen.color, { id: nid('bdc'), onChange: (v) => { pen.color = v || 'auto'; } });
    const width = ui.select(WIDTHS.map((w) => [w, w + ' pt']), pen.sz, (v) => { pen.sz = +v; }, { id: nid('bdw') });
    const pvBox = h('div', { class: 'bd-preview' });
    const inner = h('div', { class: 'bd-box' });
    const text = h('div', { class: 'bd-text' }, h('i'), h('i'), h('i'), h('i'), h('i'));
    pvBox.append(inner, text);
    const sideBtn = {};
    const pos = { top: 'left:2px;top:30px', bottom: 'left:2px;bottom:30px', left: 'left:30px;bottom:2px', right: 'right:30px;bottom:2px', between: 'left:2px;top:64px', insideH: 'left:2px;top:64px', insideV: 'left:64px;bottom:2px' };
    const ico = { top: 'bordersTop', bottom: 'bordersBottom', left: 'bordersLeft', right: 'bordersRight', between: 'bordersInH', insideH: 'bordersInH', insideV: 'bordersInV' };
    for (const s of sides) {
      const b = h('button', { type: 'button', class: 'bd-btn', style: pos[s], 'aria-label': s, 'data-tip': s, html: L.icons.get(ico[s]) });
      b.addEventListener('click', () => { if (bd[s]) delete bd[s]; else bd[s] = Object.assign({}, pen, { space: s === 'left' || s === 'right' ? 4 : 1 }); setting.set('custom'); draw(); });
      sideBtn[s] = b;
      pvBox.appendChild(b);
    }
    const draw = () => {
      const css = (s) => (bd[s] ? R.borderCSS(bd[s]) : 'none');
      inner.style.borderTop = css('top'); inner.style.borderBottom = css('bottom'); inner.style.borderLeft = css('left'); inner.style.borderRight = css('right');
      text.style.borderTop = 'none';
      inner.style.boxShadow = o.shadow ? '3px 3px 0 #555' : '';
      const mid = bd.between || bd.insideH;
      text.style.cssText = `position:absolute;left:30px;right:30px;top:34px;display:flex;flex-direction:column;gap:5px;${mid ? 'border-bottom:' + R.borderCSS(mid) + ';padding-bottom:6px' : ''}${bd.insideV ? ';border-right:' + R.borderCSS(bd.insideV) : ''}`;
      for (const s of sides) sideBtn[s].classList.toggle('on', !!bd[s]);
    };
    const setting = G.radios(nid('bds'), [['none', 'N&one'], ['box', 'Bo&x'], ['shadow', 'Sh&adow'], ['3d', '3-&D'], ['custom', 'C&ustom']].concat(kind === 'cell' || kind === 'table' ? [['all', 'A&ll'], ['grid', 'Gri&d']] : []), Object.keys(bd).length ? 'custom' : 'none', (v) => {
      for (const s of Object.keys(bd)) delete bd[s];
      o.shadow = false;
      if (v === 'none') { draw(); return; }
      const outer = ['top', 'bottom', 'left', 'right'];
      const b0 = Object.assign({}, pen);
      if (v === '3d') { b0.val = 'threeDEmboss'; }
      for (const s of outer) bd[s] = Object.assign({}, b0);
      if (v === 'shadow') o.shadow = true;
      if (v === 'all') for (const s of ['insideH', 'insideV']) bd[s] = Object.assign({}, pen);
      if (v === 'grid') { for (const s of outer) bd[s] = Object.assign({}, pen, { sz: Math.max(pen.sz, 1.5) }); bd.insideH = Object.assign({}, pen, { sz: 0.5 }); bd.insideV = Object.assign({}, pen, { sz: 0.5 }); }
      draw();
    });
    draw();
    const el = G.row(ui.group('Setting:', setting), h('div', { class: 'col', style: 'flex:1;min-width:140px' }, h('label', { text: 'St&yle:' }), styleBox, G.f('&Color:', color), G.f('&Width:', width)), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), h('div', { class: 'tp-note', text: 'Click on diagram or use buttons to apply borders' }), pvBox));
    el.get = () => { const out = {}; for (const s of sides) out[s] = bd[s] ? Object.assign({}, bd[s], o.shadow ? { shadow: true } : {}) : null; return out; };
    return el;
  }
  G.borders = function (tab, opts) {
    opts = opts || {};
    const d = doc();
    const inTbl = L.tables && L.tables.inTable();
    const cellMode = inTbl && (E.cellSel || L.tables.wholeCellsSelected() || E.collapsed());
    const pp = E.curPara();
    let target = opts.target ? 'style' : cellMode ? 'cell' : 'para';
    const ci = inTbl ? L.tables.cur() : null;
    const init = target === 'cell' ? (ci.cell.tcPr.borders || R.tblProps(d, ci.tbl).tp.borders) : opts.target ? (opts.target.borders || {}) : pp.borders || {};
    const pnl = borderPanel(target === 'cell' ? 'cell' : 'para', init, {});
    const applyTo = ui.select(inTbl ? [['cell', 'Cell'], ['table', 'Table'], ['para', 'Paragraph'], ['text', 'Text']] : [['para', 'Paragraph'], ['text', 'Text']], target === 'style' ? 'para' : target, (v) => { target = v; }, { id: 'bd-apply' });
    const sect = A().curSect();
    const pb = sect.sect.borders || {};
    const pagePnl = borderPanel('page', pb, {});
    const pageApply = ui.select([['all', 'Whole document'], ['section', 'This section'], ['firstPage', 'This section - First page only'], ['notFirstPage', 'This section - All except first page']], pb.display || 'section', null, { id: 'bd-papply' });
    const shdCur = target === 'cell' ? ci.cell.tcPr.shd || {} : opts.target ? opts.target.shd || {} : pp.shd || {};
    let fill = shdCur.fill && shdCur.fill !== 'auto' ? shdCur.fill.toUpperCase() : null;
    const fillGrid = h('div', { class: 'cp-grid', style: 'grid-template-columns:repeat(8,18px)' });
    const fillName = h('div', { class: 'tp-note' });
    const showFill = () => { fillName.textContent = fill ? ui.colorName('#' + fill) : 'No Fill'; L.$$('.cp-sw', fillGrid).forEach((b) => b.classList.toggle('on', b.dataset.c === fill)); pvShd.style.background = fill ? '#' + fill : '#fff'; };
    for (const c of L.color.STANDARD) { const b = h('button', { class: 'cp-sw', type: 'button', style: `background:${c}`, 'data-c': c.replace('#', '').toUpperCase(), 'data-tip': ui.colorName(c), 'aria-label': ui.colorName(c) }); b.addEventListener('click', () => { fill = b.dataset.c; showFill(); }); fillGrid.appendChild(b); }
    const pvShd = h('div', { class: 'preview-box', style: 'width:150px;height:90px' });
    const patt = ui.select(SHD_PATTERNS, shdCur.val && shdCur.val !== 'nil' ? shdCur.val : 'clear', null, { id: 'bd-patt' });
    const pattColor = G.color(shdCur.color && shdCur.color !== 'auto' ? shdCur.color : 'auto', { id: 'bd-pcol' });
    const shdApply = ui.select(inTbl ? [['cell', 'Cell'], ['table', 'Table'], ['para', 'Paragraph'], ['text', 'Text']] : [['para', 'Paragraph'], ['text', 'Text']], target === 'style' ? 'para' : target, null, { id: 'bd-sapply' });
    showFill();
    const tabs = ui.tabs([
      { label: '&Borders', body: h('div', null, pnl, opts.target ? null : G.f('Appl&y to:', applyTo)) },
      { label: '&Page Border', body: h('div', null, pagePnl, G.f('Apply t&o:', pageApply)) },
      { label: '&Shading', body: h('div', null, G.row(h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Fill' }), ui.button('No Fill', () => { fill = null; showFill(); }, { class: 'btn small' }), fillGrid, fillName, ui.button('&More Colors...', () => ui.moreColors(fill ? '#' + fill : null, (c) => { if (c) { fill = c.replace('#', '').toUpperCase(); showFill(); } }), { class: 'btn small' })), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), pvShd)), ui.group('Patterns', G.row(G.f('St&yle:', patt), G.f('&Color:', pattColor))), opts.target ? null : G.f('Apply t&o:', shdApply)) },
    ], tab === 'page' ? 1 : tab === 'shading' ? 2 : 0);
    ui.dialog({
      title: 'Borders and Shading', width: 540, body: tabs,
      buttons: [{ label: '&Horizontal Line...', onClick: () => { A().insertHorizontalLine(); } }, { label: 'OK', primary: true, onClick: () => apply() }, { label: 'Cancel' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
    function apply() {
      const bd = pnl.get();
      const clean = (o) => { const out = {}; for (const k in o) if (o[k]) out[k] = o[k]; return Object.keys(out).length ? out : undefined; };
      const shd = fill || patt.value !== 'clear' ? { val: patt.value, fill: fill || 'auto', color: pattColor.get() || 'auto' } : undefined;
      if (opts.onOK) { opts.onOK({ borders: clean(bd), shd }); return; }
      E.edit('Borders and Shading', () => {
        /* borders */
        if (target === 'para' || target === 'style') O.setParaProps(E.selectedParas(), (pPr) => { pPr.borders = clean(bd); });
        else if (target === 'text' && !E.collapsed()) { const [a, b] = E.range(); O.setRunProps(a, b, { border: bd.top || bd.left || bd.bottom || bd.right || undefined }); }
        else if (target === 'cell' || target === 'table') {
          const rg = L.tables.range();
          D.touchTbl(rg.tbl);
          if (target === 'table') { rg.tbl.tblPr.borders = Object.assign({}, rg.tbl.tblPr.borders || {}); for (const k of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']) rg.tbl.tblPr.borders[k] = bd[k] || { val: 'nil' }; for (const r of rg.tbl.rows) for (const c of r.cells) delete c.tcPr.borders; }
          else {
            const map = rg.map;
            for (let r = rg.r0; r <= rg.r1; r++) for (const m of map[r]) {
              if (!rg.cells.includes(m.cell)) continue;
              const cb = {};
              cb.top = r === rg.r0 ? bd.top : bd.insideH; cb.bottom = r + (m.rowspan || 1) - 1 >= rg.r1 ? bd.bottom : bd.insideH;
              cb.left = m.c0 === rg.c0 ? bd.left : bd.insideV; cb.right = m.c0 + m.span - 1 >= rg.c1 ? bd.right : bd.insideV;
              for (const k in cb) if (!cb[k]) cb[k] = { val: 'nil' };
              m.cell.tcPr.borders = cb;
            }
          }
        }
        /* shading */
        const sa = shdApply.value;
        if (sa === 'para') O.setParaProps(E.selectedParas(), (pPr) => { pPr.shd = shd; });
        else if (sa === 'text' && !E.collapsed()) { const [a, b] = E.range(); O.setRunProps(a, b, { shd }); }
        else if (sa === 'cell' || sa === 'table') { const rg = L.tables.range(); D.touchTbl(rg.tbl); const cells = sa === 'table' ? rg.tbl.rows.flatMap((r) => r.cells) : rg.cells; for (const c of cells) { if (shd) c.tcPr.shd = shd; else delete c.tcPr.shd; } }
        /* page borders */
        const pbd = pagePnl.get();
        const pclean = clean(pbd);
        const holders = pageApply.value === 'all' ? D.sections(d).map((s) => ({ holder: s.endPara || d })) : [{ holder: sect.holder }];
        for (const hh of holders) { D.touchKey(hh.holder, 'sect'); const sc = hh.holder === d ? d.sect : hh.holder.sect; sc.borders = pclean ? Object.assign({}, pclean, { display: pageApply.value === 'firstPage' || pageApply.value === 'notFirstPage' ? pageApply.value : undefined, offsetFrom: (sc.borders && sc.borders.offsetFrom) || 'text' }) : null; }
        return E.sel;
      });
    }
  };

  /* ================= Bullets and Numbering ================= */
  const BULLET_GALLERY = [['•', 'Symbol', ''], ['o', 'Courier New', 'o'], ['▪', 'Wingdings', ''], ['❖', 'Wingdings', ''], ['➢', 'Wingdings', ''], ['✓', 'Wingdings', ''], ['◆', 'Wingdings', '']];
  const NUM_GALLERY = [['decimal', '%1.'], ['decimal', '%1)'], ['upperRoman', '%1.'], ['upperLetter', '%1.'], ['lowerLetter', '%1)'], ['lowerLetter', '%1.'], ['lowerRoman', '%1.']];
  G.bullets = function (tab) {
    const p = E.sel ? E.sel.f.p : null;
    const info = p && L.lists ? L.lists.info(p) : null;
    let choice = null;
    let restart = false;
    const gal = (kind) => {
      const g = h('div', { class: 'gal-grid' });
      const none = h('button', { type: 'button', class: 'none', text: 'None' });
      none.addEventListener('click', () => { pick(g, none, { none: true }); });
      none.addEventListener('dblclick', () => okBtn.click());
      g.appendChild(none);
      const items = kind === 'bullet' ? BULLET_GALLERY : kind === 'number' ? NUM_GALLERY : D.OUTLINE_PRESETS.slice(0, 7);
      items.forEach((it) => {
        const b = h('button', { type: 'button' });
        if (kind === 'bullet') for (let i = 0; i < 3; i++) b.appendChild(h('div', null, h('span', { text: it[0], style: 'font-size:16px;width:12px' }), h('i')));
        else if (kind === 'number') for (let i = 1; i <= 3; i++) b.appendChild(h('div', null, h('span', { text: it[1].replace('%1', D.fmtNum(i, it[0])) }), h('i')));
        else { const lv = (k) => { const [fmt, t] = it.lv(k); return t.replace(/%(\d)/g, (m, x) => D.fmtNum(1, +x - 1 === k ? fmt : it.lv(+x - 1)[0])); }; for (let k = 0; k < 3; k++) b.appendChild(h('div', { class: 'l' + k }, h('span', { text: lv(k) }), h('i'))); }
        b.addEventListener('click', () => pick(g, b, { kind, it }));
        b.addEventListener('dblclick', () => okBtn.click());
        g.appendChild(b);
      });
      return g;
    };
    const pick = (g, b, c) => { L.$$('button', g).forEach((x) => x.classList.toggle('on', x === b)); choice = c; };
    const gB = gal('bullet'), gN = gal('number'), gO = gal('outline');
    const numOpts = G.radios('bn-num', [['restart', 'R&estart numbering'], ['continue', 'C&ontinue previous list']], 'continue', (v) => { restart = v === 'restart'; });
    const abs = () => {
      if (!choice || choice.none) return null;
      if (choice.custom) return choice.custom;
      if (choice.kind === 'bullet') return D.makeBulletAbs(choice.it[0], choice.it[1], choice.it[2]);
      if (choice.kind === 'number') return D.makeNumberAbs(choice.it[0], choice.it[1]);
      return D.makeOutlineAbs(choice.it);
    };
    const tabs = ui.tabs([{ label: '&Bulleted', body: gB }, { label: '&Numbered', body: h('div', null, gN, ui.group('List numbering', numOpts)) }, { label: 'O&utline Numbered', body: gO }], tab != null ? tab : info ? (info.fmt === 'bullet' ? 0 : 1) : 0);
    let okBtn;
    const dlg = ui.dialog({
      title: 'Bullets and Numbering', width: 450, body: tabs,
      buttons: [
        { label: '&Reset', onClick: () => { L.lists.remove(); return true; } },
        { label: 'Customi&ze...', onClick: () => { const a = abs() || (info ? L.clone(D.numDef(doc(), info.numId).abs) : D.makeBulletAbs('•', 'Symbol', '')); dlg.close(null); G.customizeList(a, info ? info.lvl : 0); return false; } },
        { label: 'OK', primary: true, onClick: () => { if (!choice) return; if (choice.none) { L.lists.remove(); return; } const a = abs(); if (choice.kind === 'outline' && choice.it.headings) applyHeadingNumbering(a); else L.lists.apply(a, { restart, newList: !info || restart }); } },
        { label: 'Cancel' },
      ],
    });
    okBtn = dlg.buttons[2];
  };
  function applyHeadingNumbering(abs) {
    /* outline numbering linked to Heading 1–9 styles */
    const d = doc();
    E.edit('Outline Numbered', () => {
      D.touchKey(d, 'numbering'); D.touchKey(d, 'styles');
      const nid2 = D.addNum(d, abs);
      for (let i = 1; i <= 9; i++) { const id = 'Heading' + i; if (!d.styles[id]) d.styles[id] = D.builtinStyles()[id]; d.styles[id].pPr = Object.assign({}, d.styles[id].pPr, { num: { id: nid2, lvl: i - 1 } }); }
      D.stylesChanged();
      return E.sel;
    });
  }
  /** Customize Bulleted / Numbered / Outline Numbered List */
  G.customizeList = function (abs, lvl) {
    abs = L.clone(abs);
    let cur = lvl || 0;
    const isBullet = abs.levels[cur].fmt === 'bullet';
    const multi = abs.multi === 'multilevel';
    const lvList = G.listBox([1, 2, 3, 4, 5, 6, 7, 8, 9].map(String), String(cur + 1), { height: 140, onChange: (v) => { save(); cur = +v - 1; load(); } });
    const numText = h('input', { type: 'text', id: 'cl-text', style: 'width:140px' });
    const numStyle = ui.select([['decimal', '1, 2, 3, ...'], ['upperRoman', 'I, II, III, ...'], ['lowerRoman', 'i, ii, iii, ...'], ['upperLetter', 'A, B, C, ...'], ['lowerLetter', 'a, b, c, ...'], ['ordinal', '1st, 2nd, 3rd, ...'], ['cardinalText', 'One, Two, Three, ...'], ['ordinalText', 'First, Second, Third, ...'], ['decimalZero', '01, 02, 03, ...'], ['bullet', 'Bullet'], ['none', '(none)']], abs.levels[cur].fmt, () => { save(); drawPv(); }, { id: 'cl-style' });
    const start = G.num(abs.levels[cur].start || 1, { min: 0, id: 'cl-start', onChange: () => drawPv() });
    const bulletCh = h('button', { type: 'button', class: 'btn', style: 'min-width:40px;font-size:16px' });
    const align = ui.select([['left', 'Left'], ['center', 'Centered'], ['right', 'Right']], abs.levels[cur].jc || 'left', () => drawPv(), { id: 'cl-al' });
    const alignedAt = G.len(0, { id: 'cl-at', min: 0, onChange: () => drawPv() });
    const tabAfter = G.len(36, { id: 'cl-tab', min: 0 });
    const indentAt = G.len(36, { id: 'cl-ind', min: 0, onChange: () => drawPv() });
    const follow = ui.select([['tab', 'Tab character'], ['space', 'Space'], ['nothing', 'Nothing']], abs.levels[cur].suff || 'tab', null, { id: 'cl-follow' });
    const linkStyle = ui.select([['', '(no style)']].concat([1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => ['Heading' + i, 'Heading ' + i])), abs.levels[cur].pStyle || '', null, { id: 'cl-link' });
    const legal = ui.check('&Legal style numbering', !!abs.levels[cur].isLgl, null);
    const pv = h('div', { class: 'para-preview', style: 'height:150px;width:180px;font:9px Arial,Arimo,sans-serif;gap:5px' });
    let bullet = { ch: abs.levels[cur].glyph || '•', font: (abs.levels[cur].rPr && abs.levels[cur].rPr.font) || 'Symbol', raw: abs.levels[cur].text };
    let fontR = Object.assign({}, abs.levels[cur].rPr || {});
    bulletCh.addEventListener('click', () => G.symbol({ pick: (ch, font) => { bullet = { ch, font: font || 'Symbol', raw: ch }; bulletCh.textContent = ch; drawPv(); } }));
    function load() {
      const lv = abs.levels[cur];
      numText.value = (lv.text || '').replace(/%(\d)/g, (m, x) => `{${x}}`);
      numStyle.value = lv.fmt;
      start.set(lv.start == null ? 1 : lv.start);
      align.value = lv.jc || 'left';
      const ind = lv.ind || { l: 36 * (cur + 1), fl: -18 };
      alignedAt.set((ind.l || 0) + (ind.fl || 0));
      indentAt.set(ind.l || 0);
      tabAfter.set(lv.tabPos != null ? lv.tabPos : ind.l || 0);
      follow.value = lv.suff || 'tab';
      linkStyle.value = lv.pStyle || '';
      legal.input.checked = !!lv.isLgl;
      bullet = { ch: lv.glyph || L.mapSymbolChar(lv.text || '•', lv.rPr && lv.rPr.font) || '•', font: (lv.rPr && lv.rPr.font) || 'Symbol', raw: lv.text };
      fontR = Object.assign({}, lv.rPr || {});
      bulletCh.textContent = bullet.ch;
      bulletRow.hidden = numStyle.value !== 'bullet';
      numRow.hidden = numStyle.value === 'bullet';
      drawPv();
    }
    function save() {
      const lv = abs.levels[cur];
      lv.fmt = numStyle.value;
      if (lv.fmt === 'bullet') { lv.text = bullet.raw || bullet.ch; lv.glyph = bullet.ch; lv.rPr = Object.assign({}, fontR, { font: bullet.font }); }
      else { lv.text = numText.value.replace(/\{(\d)\}/g, '%$1'); delete lv.glyph; lv.rPr = Object.assign({}, fontR); if (lv.rPr.font && /symbol|wingdings/i.test(lv.rPr.font)) delete lv.rPr.font; }
      lv.start = start.get();
      lv.jc = align.value;
      const at = alignedAt.get() || 0, ind = indentAt.get() || 0;
      lv.ind = { l: L.round(ind, 2), fl: L.round(at - ind, 2) };
      const tp = tabAfter.get();
      if (tp != null && Math.abs(tp - ind) > 0.5) lv.tabPos = L.round(tp, 2); else delete lv.tabPos;
      lv.suff = follow.value;
      if (linkStyle.value) lv.pStyle = linkStyle.value; else delete lv.pStyle;
      if (legal.input.checked) lv.isLgl = true; else delete lv.isLgl;
    }
    function drawPv() {
      save();
      L.clear(pv);
      const counters = [0, 0, 0, 0, 0, 0, 0, 0, 0];
      const seq = multi ? [0, 1, 2, 1, 0, 1] : [cur, cur, cur, cur, cur];
      for (const k of seq) {
        counters[k]++; for (let j = k + 1; j < 9; j++) counters[j] = 0;
        const lv = abs.levels[k];
        const lbl = lv.fmt === 'bullet' ? lv.glyph || '•' : (lv.text || '').replace(/%(\d)/g, (m, x) => { const j = +x - 1; return D.fmtNum((counters[j] || 1) + (abs.levels[j].start || 1) - 1, lv.isLgl && j < k ? 'decimal' : abs.levels[j].fmt); });
        const ind = lv.ind || {};
        pv.appendChild(h('div', { style: `display:flex;gap:4px;align-items:center;margin-left:${((ind.l || 0) + (ind.fl || 0)) * 0.3}px` }, h('span', { text: lbl, style: 'flex:none;min-width:10px' }), h('i', { style: 'flex:1;height:3px;background:#999' })));
      }
    }
    const bulletRow = G.row(G.f('Bullet character:', bulletCh), ui.button('&Font...', () => G.font({ target: fontR, sample: bullet.ch, onOK: (r) => { Object.assign(fontR, r); drawPv(); } }), { class: 'btn small' }));
    const numRow = h('div', { class: 'col' }, G.row(G.f('Number f&ormat:', numText), ui.button('&Font...', () => G.font({ target: fontR, sample: '1.', onOK: (r) => { Object.assign(fontR, r); drawPv(); } }), { class: 'btn small' })), G.row(G.f('&Number style:', numStyle), G.f('&Start at:', start)), multi ? G.row(G.f('Lin&k level to style:', linkStyle), legal) : null);
    load();
    ui.dialog({
      title: isBullet ? 'Customize Bulleted List' : multi ? 'Customize Outline Numbered List' : 'Customize Numbered List', width: 520,
      body: G.row(multi ? h('div', { class: 'col', style: 'width:60px' }, h('label', { text: 'Le&vel' }), lvList) : null, h('div', { class: 'col', style: 'flex:1;min-width:250px' }, bulletRow, numRow, ui.group('Number position', G.row(G.f('Alignment:', align), G.f('&Aligned at:', alignedAt))), ui.group('Text position', G.row(G.f('Ta&b space after:', tabAfter), G.f('&Indent at:', indentAt)), G.f('Follow number with:', follow))), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), pv)),
      buttons: [{ label: 'OK', primary: true, onClick: () => { save(); L.lists.apply(abs, {}); } }, { label: 'Cancel' }],
    });
  };

  /* ================= Columns ================= */
  G.columns = function () {
    const d = doc();
    const s = A().curSect();
    const sect = s.sect;
    const cols = Object.assign({ n: 1, space: 36, eq: true, w: [], sep: false }, sect.cols || {});
    const bodyW = sect.pgW - sect.ml - sect.mr - (sect.gutter || 0);
    const n = G.num(cols.n, { min: 1, max: 13, id: 'cd-n', onChange: () => { cols.n = n.get(); build(); } });
    const eq = ui.check('&Equal column width', cols.eq !== false, (v) => { cols.eq = v; build(); });
    const sep = ui.check('Line &between', !!cols.sep, (v) => { cols.sep = v; drawPv(); });
    const tbl = h('div', { class: 'col' });
    const widths = [];
    const presetBox = h('div', { class: 'preset-grid' });
    const presets = [['one', 'O&ne', 1, true], ['two', 'T&wo', 2, true], ['three', '&Three', 3, true], ['left', '&Left', 2, false, [0.35, 0.65]], ['right', '&Right', 2, false, [0.65, 0.35]]];
    for (const [k, l, cn, e, fr] of presets) {
      const b = h('button', { type: 'button' }, h('div', { class: 'colsample', style: 'width:50px;height:40px;padding:3px;gap:3px' }), h('span', { html: l.replace(/&(.)/, '<u>$1</u>') }));
      const smp = b.firstChild;
      const fracs = fr || Array.from({ length: cn }, () => 1 / cn);
      for (const f of fracs) { const c = h('div', { style: `flex:${f}` }); for (let i = 0; i < 4; i++) c.appendChild(h('i')); smp.appendChild(c); }
      b.addEventListener('click', () => { cols.n = cn; cols.eq = e; n.set(cn); eq.input.checked = e; if (fr) { const avail = bodyW - 36; cols.w = fr.map((x) => ({ w: L.round(avail * x, 2), space: 36 })); } build(); L.$$('button', presetBox).forEach((x) => x.classList.toggle('on', x === b)); });
      presetBox.appendChild(b);
    }
    const pv = h('div', { class: 'colsample', style: 'width:110px;height:130px' });
    function drawPv() {
      L.clear(pv);
      for (let i = 0; i < cols.n; i++) { const c = h('div', { style: `flex:${cols.eq !== false || !widths[i] ? 1 : widths[i].w.get() || 1}` }); for (let k = 0; k < 14; k++) c.appendChild(h('i')); pv.appendChild(c); if (cols.sep && i < cols.n - 1) pv.appendChild(h('span', { style: 'border-left:1px solid #000;flex:none' })); }
    }
    function build() {
      L.clear(tbl);
      widths.length = 0;
      const nn = Math.max(1, cols.n);
      const eqW = (bodyW - (cols.space || 36) * (nn - 1)) / nn;
      tbl.appendChild(h('div', { class: 'grid3', style: 'font-weight:bold' }, h('span', { text: 'Col #:' }), h('span', { text: 'Width:' }), h('span', { text: 'Spacing:' })));
      for (let i = 0; i < Math.min(nn, 4); i++) {
        const cw = cols.eq !== false ? eqW : (cols.w[i] && cols.w[i].w) || eqW;
        const sp = cols.eq !== false ? cols.space : (cols.w[i] && cols.w[i].space) != null ? cols.w[i].space : 36;
        const w = G.len(cw, { min: 36, id: 'cd-w' + i, onChange: () => drawPv() });
        const spc = G.len(sp, { min: 0, id: 'cd-s' + i });
        if (cols.eq !== false && i > 0) { w.setDisabled(true); spc.setDisabled(true); }
        if (i === nn - 1) spc.setDisabled(true);
        widths.push({ w, spc });
        tbl.appendChild(h('div', { class: 'grid3' }, h('span', { text: i + 1 + ':' }), w, spc));
      }
      drawPv();
    }
    build();
    const applyTo = ui.select([['section', 'This section'], ['all', 'Whole document'], ['forward', 'This point forward']], D.sections(d).length > 1 ? 'section' : 'all', null, { id: 'cd-apply' });
    const newCol = ui.check('Start new col&umn', false, null);
    ui.dialog({
      title: 'Columns', width: 470,
      body: h('div', null, ui.group('Presets', presetBox), G.row(G.f('&Number of columns:', n, 'wide'), sep), ui.group('Width and spacing', tbl, eq), G.row(G.f('&Apply to:', applyTo), newCol, h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), pv))),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const nn = Math.max(1, n.get());
        const out = { n: nn, space: widths[0] ? widths[0].spc.get() || 36 : 36, eq: eq.input.checked, sep: sep.input.checked, w: [] };
        if (!out.eq) out.w = widths.map((x, i) => ({ w: L.round(x.w.get(), 2), space: i < nn - 1 ? L.round(x.spc.get() || 0, 2) : 0 }));
        if (nn === 1) { out.w = []; out.eq = true; }
        E.edit('Columns', () => {
          if (applyTo.value === 'all') { for (const sc of D.sections(d)) { const hd = sc.endPara || d; D.touchKey(hd, 'sect'); (hd === d ? d.sect : hd.sect).cols = L.clone(out); } }
          else if (applyTo.value === 'forward') {
            /* continuous section break at the caret's paragraph start */
            const pos = E.sel.f;
            const tb = D.topBlock(d, pos.p);
            const prevI = tb.i - 1;
            const cont = D.touchList(d.main);
            let endP = cont.blocks[prevI];
            if (!endP || endP.t !== 'p') { endP = D.para(); cont.blocks.splice(tb.i, 0, endP); }
            D.touch(endP);
            endP.sect = L.clone(D.sectOfBlock(d, Math.max(0, prevI)).sect);
            const after = A().curSect();
            D.touchKey(after.holder, 'sect');
            const sc = after.holder === d ? d.sect : after.holder.sect;
            sc.cols = L.clone(out); sc.type = 'continuous';
          } else { D.touchKey(s.holder, 'sect'); (s.holder === d ? d.sect : s.holder.sect).cols = L.clone(out); }
          return E.sel;
        });
        if (newCol.input.checked) E.enter('column');
      } }, { label: 'Cancel' }],
    });
  };

  /* ================= Page Setup ================= */
  G.pageSetup = function (tab) {
    const d = doc();
    const s = A().curSect();
    const sect = L.clone(s.sect);
    const mt = G.len(sect.mt, { id: 'ps-t', min: -1584 }), mb = G.len(sect.mb, { id: 'ps-b', min: -1584 }), ml = G.len(sect.ml, { id: 'ps-l', min: 0 }), mr = G.len(sect.mr, { id: 'ps-r', min: 0 });
    const gutter = G.len(sect.gutter || 0, { id: 'ps-g', min: 0 });
    const gutPos = ui.select([['left', 'Left'], ['top', 'Top']], d.settings.gutterAtTop ? 'top' : 'left', null, { id: 'ps-gp' });
    let orient = sect.orient || (sect.pgW > sect.pgH ? 'landscape' : 'portrait');
    const orBtns = h('div', { class: 'preset-grid' });
    const orB = (k, l) => { const b = h('button', { type: 'button', class: orient === k ? 'on' : '' }, h('div', { class: 'ps-page', style: k === 'landscape' ? 'width:60px;height:46px' : 'width:46px;height:60px' }), h('span', { html: l })); b.addEventListener('click', () => { if (orient !== k) { orient = k; const w = pw.get(), hh = ph.get(); pw.set(Math.min(w, hh) === w && k === 'landscape' ? hh : w); ph.set(Math.min(w, hh) === w && k === 'landscape' ? w : hh); if (k === 'portrait' && w > hh) { pw.set(hh); ph.set(w); } } L.$$('button', orBtns).forEach((x) => x.classList.toggle('on', x === b)); drawPv(); }); return b; };
    orBtns.append(orB('portrait', 'P<u>o</u>rtrait'), orB('landscape', 'Land<u>s</u>cape'));
    const multi = ui.select([['normal', 'Normal'], ['mirror', 'Mirror margins'], ['twoOnOne', '2 pages per sheet'], ['bookFold', 'Book fold']], d.settings.mirror ? 'mirror' : d.settings.bookFold ? 'bookFold' : 'normal', () => { const m = multi.value === 'mirror' || multi.value === 'bookFold'; L.$('label[for="ps-l"]').textContent = m ? 'Inside:' : 'Left:'; L.$('label[for="ps-r"]').textContent = m ? 'Outside:' : 'Right:'; drawPv(); }, { id: 'ps-multi' });
    const paperName = D.paperName(sect);
    const paper = ui.select(D.PAPER.map((p) => [p[0], p[0]]).concat([['Custom size', 'Custom size']]), paperName, (v) => { const p = D.PAPER.find((x) => x[0] === v); if (!p) return; const land = orient === 'landscape'; pw.set(land ? p[2] : p[1]); ph.set(land ? p[1] : p[2]); drawPv(); }, { id: 'ps-paper' });
    const pw = G.len(sect.pgW, { id: 'ps-w', min: 72, max: 1584, onChange: () => { paper.value = D.paperName({ pgW: pw.get(), pgH: ph.get() }); drawPv(); } });
    const ph = G.len(sect.pgH, { id: 'ps-h', min: 72, max: 1584, onChange: () => { paper.value = D.paperName({ pgW: pw.get(), pgH: ph.get() }); drawPv(); } });
    const start = ui.select([['continuous', 'Continuous'], ['nextColumn', 'New column'], ['nextPage', 'New page'], ['evenPage', 'Even page'], ['oddPage', 'Odd page']], sect.type || 'nextPage', null, { id: 'ps-start' });
    const evenOdd = ui.check('Different &odd and even', !!d.settings.evenOdd, null);
    const firstPg = ui.check('Different first &page', !!sect.titlePg, null);
    const hdr = G.len(sect.hdr, { id: 'ps-hdr', min: 0 }), ftr = G.len(sect.ftr, { id: 'ps-ftr', min: 0 });
    const vAlign = ui.select([['top', 'Top'], ['center', 'Center'], ['both', 'Justified'], ['bottom', 'Bottom']], sect.vAlign || 'top', null, { id: 'ps-va' });
    const multiSect = D.sections(d).length > 1;
    const applyTo = ui.select(multiSect ? [['section', 'This section'], ['forward', 'This point forward'], ['all', 'Whole document']] : [['all', 'Whole document'], ['forward', 'This point forward']], multiSect ? 'section' : 'all', null, { id: 'ps-apply' });
    let lnNum = sect.lnNum ? Object.assign({}, sect.lnNum) : null;
    const pv = h('div', { class: 'ps-page', style: 'margin:6px auto' });
    function drawPv() {
      const W = pw.get() || sect.pgW, H = ph.get() || sect.pgH;
      const k = 90 / Math.max(W, H);
      pv.style.width = W * k + 'px'; pv.style.height = H * k + 'px';
      L.clear(pv);
      const l = (ml.get() || 0) * k, r = (mr.get() || 0) * k, t = Math.abs(mt.get() || 0) * k, b = Math.abs(mb.get() || 0) * k;
      for (let y = t; y < H * k - b - 2; y += 4) pv.appendChild(h('i', { style: `top:${y}px;left:${l + (gutter.get() || 0) * k}px;right:${r}px` }));
    }
    drawPv();
    for (const c of [mt, mb, ml, mr, gutter]) c.input.addEventListener('change', drawPv);
    const pvBox = () => h('div', { class: 'col', style: 'align-items:center' }, h('div', { class: 'cd-cap', text: 'Preview' }), pv, G.f('Apply to:', applyTo));
    const tabs = ui.tabs([
      { label: '&Margins', body: h('div', null, ui.group('Margins', G.row(h('div', { class: 'col' }, G.f('&Top:', mt, 'narrow'), G.f('&Left:', ml, 'narrow'), G.f('&Gutter:', gutter, 'narrow')), h('div', { class: 'col' }, G.f('&Bottom:', mb, 'narrow'), G.f('&Right:', mr, 'narrow'), G.f('Gutter pos&ition:', gutPos)))), ui.group('Orientation', orBtns), ui.group('Pages', G.f('&Multiple pages:', multi))) , onShow: () => setTimeout(() => drawPv(), 0) },
      { label: '&Paper', body: h('div', null, G.f('Pape&r size:', paper), G.row(G.f('W&idth:', pw), G.f('H&eight:', ph)), ui.group('Paper source', h('div', { class: 'tp-note', text: 'Default tray (Automatically Select). Paper source is chosen in the print dialog of the program that prints the PDF.' }))) },
      { label: '&Layout', body: h('div', null, ui.group('Section', G.f('Section sta&rt:', start)), ui.group('Headers and footers', evenOdd, firstPg, G.row(G.f('From edge: &Header:', hdr, 'wide'), G.f('&Footer:', ftr, 'narrow'))), ui.group('Page', G.f('&Vertical alignment:', vAlign)), h('div', { class: 'tp-btns' }, ui.button('Line &Numbers...', () => lineNumbers()), ui.button('&Borders...', () => G.borders('page')))) },
    ], tab === 'paper' ? 1 : tab === 'layout' ? 2 : 0);
    function lineNumbers() {
      const on = ui.check('Add &line numbering', !!lnNum, null);
      const startAt = G.num(lnNum ? lnNum.start || 1 : 1, { min: 1, id: 'ln-start' });
      const from = G.len(lnNum && lnNum.distance != null ? lnNum.distance : null, { id: 'ln-from', autoLabel: 'Auto', min: 0 });
      const countBy = G.num(lnNum ? lnNum.countBy || 1 : 1, { min: 1, max: 100, id: 'ln-by' });
      const restart = G.radios('ln-r', [['newPage', 'Restart each &page'], ['newSection', 'Restart each &section'], ['continuous', '&Continuous']], lnNum ? lnNum.restart || 'newPage' : 'newPage');
      ui.dialog({ title: 'Line Numbers', width: 300, body: h('div', { class: 'col' }, on, G.f('Start &at:', startAt), G.f('From &text:', from), G.f('Count &by:', countBy), ui.group('Numbering:', restart)), buttons: [{ label: 'OK', primary: true, onClick: () => { lnNum = on.input.checked ? { start: startAt.get(), distance: from.get(), countBy: countBy.get(), restart: restart.get() } : null; } }, { label: 'Cancel' }] });
    }
    ui.dialog({
      title: 'Page Setup', width: 440, body: h('div', null, tabs, pvBox()),
      buttons: [{ label: '&Default...', onClick: () => { setDefault(); return false; } }, { label: 'OK', primary: true, onClick: () => apply() }, { label: 'Cancel' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
    function values() {
      return { mt: mt.get(), mb: mb.get(), ml: ml.get(), mr: mr.get(), gutter: gutter.get() || 0, pgW: pw.get(), pgH: ph.get(), orient, type: start.value, titlePg: firstPg.input.checked, hdr: hdr.get(), ftr: ftr.get(), vAlign: vAlign.value, lnNum };
    }
    async function setDefault() {
      const r = await ui.msg('Do you want to change the default settings for page setup?\nThis change will affect all new documents.', { icon: 'question', buttons: ['&Yes', '&No'] });
      if (r !== 0) return;
      const v = values();
      L.store.set('defaultSect', { mt: v.mt, mb: v.mb, ml: v.ml, mr: v.mr, gutter: v.gutter, pgW: v.pgW, pgH: v.pgH, orient: v.orient, hdr: v.hdr, ftr: v.ftr });
    }
    function apply() {
      const v = values();
      E.edit('Page Setup', () => {
        D.touchKey(d, 'settings');
        d.settings.evenOdd = evenOdd.input.checked;
        d.settings.mirror = multi.value === 'mirror' || multi.value === 'bookFold';
        d.settings.bookFold = multi.value === 'bookFold' || undefined;
        d.settings.gutterAtTop = gutPos.value === 'top' || undefined;
        const assign = (hd) => { D.touchKey(hd, 'sect'); const sc = hd === d ? d.sect : hd.sect; Object.assign(sc, v); };
        if (applyTo.value === 'all') for (const sc of D.sections(d)) assign(sc.endPara || d);
        else if (applyTo.value === 'forward') {
          const tb = D.topBlock(d, E.sel.f.p);
          const cont = D.touchList(d.main);
          let endP = cont.blocks[tb.i - 1];
          if (tb.i === 0) { for (const sc of D.sections(d)) assign(sc.endPara || d); return E.sel; }
          if (!endP || endP.t !== 'p') { endP = D.para(); cont.blocks.splice(tb.i, 0, endP); }
          D.touch(endP);
          endP.sect = L.clone(D.sectOfBlock(d, tb.i - 1).sect);
          endP.sect.type = endP.sect.type || 'nextPage';
          d._idxDirty = true;
          const after = A().curSect();
          assign(after.holder);
          (after.holder === d ? d.sect : after.holder.sect).type = 'nextPage';
        } else assign(s.holder);
        return E.sel;
      });
    }
  };

  /* ================= Drop Cap ================= */
  G.dropCap = function () {
    const p = E.sel && E.sel.f.p;
    if (!p) return;
    const pp = D.pProps(doc(), p);
    const dc = pp.dropCap || {};
    let pos = dc.type || 'none';
    const box = h('div', { class: 'preset-grid' });
    const mk = (k, l, svg) => { const b = h('button', { type: 'button', class: pos === k ? 'on' : '' }, h('span', { html: svg }), h('span', { html: l })); b.addEventListener('click', () => { pos = k; L.$$('button', box).forEach((x) => x.classList.toggle('on', x === b)); }); return b; };
    const lines = '<path d="M6 8h40M6 14h40M6 20h40M6 26h40M6 32h40" stroke="#999"/>';
    box.append(mk('none', '<u>N</u>one', `<svg width="52" height="40">${lines}<text x="6" y="12" font-size="10" font-family="Times">W</text></svg>`), mk('drop', '<u>D</u>ropped', `<svg width="52" height="40"><path d="M22 8h24M22 14h24M22 20h24M6 26h40M6 32h40" stroke="#999"/><text x="4" y="22" font-size="22" font-family="Times" font-weight="bold">W</text></svg>`), mk('margin', 'In <u>m</u>argin', `<svg width="52" height="40"><path d="M14 8h34M14 14h34M14 20h34M14 26h34M14 32h34" stroke="#999"/><text x="0" y="22" font-size="20" font-family="Times" font-weight="bold">W</text></svg>`));
    const font = ui.select(FONTS(), (D.rProps(doc(), p, p.runs.find((x) => x.t === 'text')) || {}).font || 'Times New Roman', null, { id: 'dc-font' });
    const nl = G.num(dc.lines || 3, { min: 1, max: 10, id: 'dc-lines' });
    const dist = G.len(dc.dist || 0, { min: 0, id: 'dc-dist' });
    ui.dialog({
      title: 'Drop Cap', width: 300, body: h('div', { class: 'col' }, ui.group('Position', box), ui.group('Options', G.f('&Font:', font), G.f('&Lines to drop:', nl, 'wide'), G.f('Distance from te&xt:', dist, 'wide'))),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const d = doc();
        E.edit('Drop Cap', () => {
          /* Word moves the first letter into its own framed paragraph */
          const prev = D.prevPara(d, p);
          const isDropPara = !!pp.dropCap;
          if (pos === 'none') {
            if (isDropPara) { D.touch(p); delete p.pPr.dropCap; for (const it of p.runs) if (it.rPr) { delete it.rPr.sz; delete it.rPr.font; } const nx = D.nextPara(d, p); if (nx) { O.joinNext(p); } }
            return E.sel;
          }
          let target = p;
          if (!isDropPara) {
            const t = D.ptext(p);
            const m = /^\s*\S/.exec(t);
            if (!m) return E.sel;
            const cut = m[0].length;
            const np = O.splitPara(D.pos(p, cut));
            void np;
            target = p;
          }
          D.touch(target);
          const r = D.rProps(d, target, target.runs.find((x) => x.t === 'text'));
          target.pPr.dropCap = { type: pos, lines: nl.get(), dist: dist.get() || 0 };
          for (const it of target.runs) if (it.t === 'text') it.rPr = Object.assign({}, it.rPr, { font: font.value, sz: L.round((r.sz || 12) * nl.get() * 1.13, 1), position: undefined });
          void prev;
          return D.pos(D.nextPara(d, target) || target, 0);
        });
      } }, { label: 'Cancel' }],
    });
  };

  /* ================= Change Case ================= */
  G.changeCase = function () {
    const r = G.radios('cc', [['sentence', '&Sentence case.'], ['lower', '&lowercase'], ['upper', '&UPPERCASE'], ['title', '&Title Case'], ['toggle', 't&OGGLE cASE']], 'sentence');
    ui.dialog({ title: 'Change Case', width: 230, body: r, buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (!E.sel) return;
      let [a, b] = E.range();
      if (D.eqPos(a, b)) { const w = O.wordAt(a); if (!w) return; [a, b] = w; }
      E.edit('Change Case', () => { O.changeCase(a, b, r.get()); return { a, f: b }; });
    } }, { label: 'Cancel' }] });
  };

  /* ================= Printed Watermark ================= */
  G.watermark = function () {
    const d = doc();
    const wm = d.watermark || {};
    let media = wm.type === 'picture' ? wm.media : null;
    const kind = G.radios('wm-k', [['none', 'N&o watermark'], ['picture', 'P&icture watermark'], ['text', 'Te&xt watermark']], wm.type || 'none');
    const pic = ui.button('Select &Picture...', async () => { const f = (await L.pickFiles('image/*'))[0]; if (!f) return; media = L.media.add(new Blob([await L.readAsArrayBuffer(f)], { type: f.type }), f.name); picName.textContent = f.name; kind.set('picture'); });
    const picName = h('span', { class: 'tp-note', text: media ? 'Picture selected' : '' });
    const scale = ui.select([['auto', 'Auto'], [5, '500%'], [2, '200%'], [1.5, '150%'], [1, '100%'], [0.5, '50%']], wm.scale || 'auto', null, { id: 'wm-scale' });
    const washout = ui.check('&Washout', wm.washout !== false, null);
    const text = h('input', { type: 'text', id: 'wm-text', list: 'wm-list', value: wm.text || 'ASAP', style: 'width:200px' });
    const dl = h('datalist', { id: 'wm-list' }, ...['ASAP', 'ATTORNEY-CLIENT PRIVILEGE', 'CONFIDENTIAL', 'COPY', 'DO NOT COPY', 'DRAFT', 'ORIGINAL', 'PERSONAL', 'SAMPLE', 'TOP SECRET', 'URGENT'].map((x) => h('option', { value: x })));
    const font = ui.select(FONTS(), wm.font || 'Times New Roman', null, { id: 'wm-font' });
    const size = ui.select([['auto', 'Auto'], 36, 40, 44, 48, 54, 60, 66, 72, 80, 90, 96, 105, 120, 144].map((x) => (Array.isArray(x) ? x : [x, String(x)])), wm.size || 'auto', null, { id: 'wm-size' });
    const color = G.color(wm.color || 'C0C0C0', { id: 'wm-col', automatic: false });
    const semi = ui.check('&Semitransparent', wm.semi !== false, null);
    const layout = G.radios('wm-l', [['diagonal', '&Diagonal'], ['horizontal', 'H&orizontal']], wm.layout || 'diagonal');
    ui.dialog({
      title: 'Printed Watermark', width: 400,
      body: h('div', { class: 'col' }, kind, ui.group('Picture', G.row(pic, picName), G.row(G.f('Sca&le:', scale), washout)), ui.group('Text', dl, G.f('&Text:', text), G.f('&Font:', font), G.f('&Size:', size), G.f('&Color:', color), semi, G.f('Layout:', layout))),
      buttons: [{ label: '&Apply', onClick: () => { apply(); return false; } }, { label: 'OK', primary: true, onClick: () => apply() }, { label: 'Cancel' }],
    });
    function apply() {
      const k = kind.get();
      E.edit('Printed Watermark', () => {
        D.touchKey(d, 'watermark');
        if (k === 'none') d.watermark = null;
        else if (k === 'picture') d.watermark = media ? { type: 'picture', media, scale: scale.value === 'auto' ? null : +scale.value, washout: washout.input.checked } : null;
        else d.watermark = { type: 'text', text: text.value || 'DRAFT', font: font.value, size: size.value === 'auto' ? 0 : +size.value, color: color.get() || 'C0C0C0', semi: semi.input.checked, layout: layout.get() };
        return E.sel;
      });
    }
  };

  /* ================= Theme ================= */
  const THEMES = [
    ['(No Theme)', null],
    ['Blends', { body: 'Trebuchet MS', head: 'Trebuchet MS', hc: '003366', bg: 'EAF1DD' }], ['Blueprint', { body: 'Verdana', head: 'Verdana', hc: '1F497D', bg: 'DBE5F1' }], ['Capsules', { body: 'Arial', head: 'Arial Black', hc: '993300', bg: 'FFF8E1' }],
    ['Compass', { body: 'Book Antiqua', head: 'Book Antiqua', hc: '663300', bg: 'F5F0E1' }], ['Edge', { body: 'Tahoma', head: 'Tahoma', hc: '333399', bg: null }], ['Expedition', { body: 'Book Antiqua', head: 'Book Antiqua', hc: '4F6228', bg: 'F2F2E6' }],
    ['Industrial', { body: 'Verdana', head: 'Arial Black', hc: '595959', bg: 'F2F2F2' }], ['Journal', { body: 'Georgia', head: 'Georgia', hc: '7F0000', bg: 'FDFAF2' }], ['Network', { body: 'Verdana', head: 'Verdana', hc: '00557F', bg: null }],
    ['Rice Paper', { body: 'Garamond', head: 'Garamond', hc: '5C4033', bg: 'FBF6E9' }], ['Sumi Painting', { body: 'Palatino Linotype', head: 'Palatino Linotype', hc: '262626', bg: 'FAFAF5' }], ['Willow', { body: 'Georgia', head: 'Trebuchet MS', hc: '336600', bg: 'F3F8EC' }],
  ];
  G.theme = function () {
    let pick = null;
    const pv = h('div', { class: 'preview-box', style: 'width:260px;height:180px;flex-direction:column;align-items:flex-start;justify-content:flex-start;padding:12px;gap:6px' });
    const draw = (t) => { L.clear(pv); pv.style.background = t && t.bg ? '#' + t.bg : '#fff'; pv.append(h('div', { text: 'Heading 1 style', style: `font:bold 18px ${L.fontStack(t ? t.head : 'Arial')};color:#${t ? t.hc : '000'}` }), h('div', { text: 'Heading 2 style', style: `font:bold 14px ${L.fontStack(t ? t.head : 'Arial')};color:#${t ? t.hc : '000'}` }), h('div', { text: 'Regular Text Sample — the quick brown fox jumps over the lazy dog.', style: `font:12px ${L.fontStack(t ? t.body : 'Times New Roman')}` }), h('div', { html: '<u style="color:#00f">Hyperlink</u>', style: 'font-size:12px' })); };
    const lb = G.listBox(THEMES.map((x) => x[0]), '(No Theme)', { height: 200, onChange: (v) => { pick = THEMES.find((x) => x[0] === v); draw(pick && pick[1]); } });
    draw(null);
    ui.dialog({ title: 'Theme', width: 470, body: G.row(h('div', { class: 'col', style: 'width:150px' }, h('label', { text: 'Choose a Theme:' }), lb), pv), buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (!pick) return;
      const t = pick[1];
      const d = doc();
      E.edit('Theme', () => {
        D.touchKey(d, 'styles'); D.touchKey(d, 'bg');
        const base = D.builtinStyles();
        for (let i = 1; i <= 3; i++) { const id = 'Heading' + i; d.styles[id] = d.styles[id] || base[id]; d.styles[id].rPr = Object.assign({}, d.styles[id].rPr, { font: t ? t.head : base[id].rPr.font, color: t ? t.hc : undefined }); }
        d.styles.Normal.rPr = Object.assign({}, d.styles.Normal.rPr, { font: t ? t.body : undefined });
        if (!t) delete d.styles.Normal.rPr.font;
        d.bg = t ? t.bg : null;
        D.stylesChanged();
        return E.sel;
      });
    } }, { label: 'Cancel' }] });
  };

  /* ================= Text Direction ================= */
  G.textDirection = function () {
    const it = L.drawing && L.drawing.selectedItem();
    const r = G.radios('tdir', [['horz', 'Horizontal'], ['vert', 'Vertical (top to bottom)'], ['vert270', 'Vertical (bottom to top)']], (it && it.vert) || 'horz');
    ui.dialog({ title: 'Text Direction - Text Box', width: 280, body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Text direction applies to text boxes and table cells.' }), r), buttons: [{ label: 'OK', primary: true, onClick: () => { if (it && it.tb) L.drawing.modifySelected('Text Direction', (x) => { x.vert = r.get() === 'horz' ? undefined : r.get(); }); else if (L.tables && L.tables.inTable()) L.tables.forCells('Text Direction', (c) => { const v = r.get(); if (v === 'horz') delete c.tcPr.textDir; else c.tcPr.textDir = v === 'vert' ? 'tbRl' : 'btLr'; }); else A().status('Select a text box or table cell first.'); } }, { label: 'Cancel' }] });
  };

  /* ================= Style (New / Modify) ================= */
  G.style = function (id) {
    const d = doc();
    const isNew = !id;
    const src = isNew ? null : d.styles[id] || D.builtinStyles()[id];
    const curP = E.sel ? E.sel.f.p : null;
    const st = isNew ? { name: 'Style' + (Object.keys(d.styles).length + 1), type: 'paragraph', basedOn: curP ? curP.pPr.style || 'Normal' : 'Normal', next: null, pPr: {}, rPr: {} } : L.clone(src);
    if (isNew && curP) { const pp = L.clone(curP.pPr); delete pp.style; delete pp.num; delete pp.sect; st.pPr = pp; const it = curP.runs.find((x) => x.t === 'text'); if (it) { st.rPr = L.clone(it.rPr); delete st.rPr.style; delete st.rPr.link; delete st.rPr.ins; delete st.rPr.del; } }
    const name = h('input', { type: 'text', id: 'sd-name', value: isNew ? st.name : D.styleDisplayName(st), style: 'width:220px' });
    if (!isNew && st.builtin) name.disabled = true;
    const type = ui.select([['paragraph', 'Paragraph'], ['character', 'Character'], ['table', 'Table'], ['numbering', 'List']], st.type, (v) => { st.type = v; refreshLists(); drawPv(); }, { id: 'sd-type', disabled: !isNew });
    const paraStyles = () => Object.keys(d.styles).filter((k) => d.styles[k].type === (st.type === 'character' ? 'character' : st.type) && k !== id).map((k) => [k, D.styleDisplayName(d.styles[k])]).sort((a, b) => a[1].localeCompare(b[1]));
    const based = ui.select([['', '(no style)']].concat(paraStyles()), st.basedOn || '', (v) => { st.basedOn = v || null; drawPv(); }, { id: 'sd-based' });
    const next = ui.select(paraStyles().concat(isNew ? [['__self', '(this style)']] : []), st.next || (isNew ? '__self' : id), (v) => { st.next = v; }, { id: 'sd-next' });
    function refreshLists() { L.clear(based); for (const [v, l] of [['', '(no style)']].concat(paraStyles())) based.appendChild(h('option', { value: v, text: l, selected: v === (st.basedOn || '') })); nextRow.hidden = st.type !== 'paragraph'; }
    const nextRow = G.f('&Style for following paragraph:', next, 'wide');
    /* quick formatting row */
    const eff = () => { const b = st.basedOn ? (st.type === 'character' ? D.charStyleProps(d, st.basedOn) : D.styleProps(d, st.basedOn).rPr) : d.defaults.rPr; return Object.assign({}, b, st.rPr); };
    const font = ui.select(FONTS(), eff().font || 'Times New Roman', (v) => { st.rPr.font = v; drawPv(); }, { id: 'sd-font' });
    const size = ui.select(L.SIZE_LIST.map((x) => [x, String(x)]), eff().sz || 12, (v) => { st.rPr.sz = +v; drawPv(); }, { id: 'sd-size' });
    const tog = (k, icon, tip) => { const b = h('button', { type: 'button', class: 'tb-btn' + (eff()[k] ? ' on' : ''), 'data-tip': tip, 'aria-label': tip, html: L.icons.get(icon) }); b.addEventListener('click', () => { st.rPr[k] = eff()[k] ? (k === 'u' ? 'none' : false) : k === 'u' ? 'single' : true; b.classList.toggle('on'); drawPv(); }); return b; };
    const col = G.color(eff().color || 'auto', { id: 'sd-col', onChange: (v) => { st.rPr.color = v === 'auto' ? undefined : v; drawPv(); } });
    const jc = (v, icon, tip) => { const b = h('button', { type: 'button', class: 'tb-btn', 'data-tip': tip, 'aria-label': tip, html: L.icons.get(icon) }); b.addEventListener('click', () => { st.pPr.jc = v; drawPv(); }); return b; };
    const fmtRow = h('div', { class: 'row', style: 'align-items:center;gap:3px' }, font, size, tog('b', 'bold', 'Bold'), tog('i', 'italic', 'Italic'), tog('u', 'underline', 'Underline'), col, jc('left', 'alignLeft', 'Align Left'), jc('center', 'alignCenter', 'Center'), jc('right', 'alignRight', 'Align Right'), jc('both', 'justify', 'Justify'));
    const pv = h('div', { class: 'preview-box', style: 'height:100px;flex-direction:column;align-items:stretch;justify-content:center;padding:6px 10px;gap:4px' });
    const desc = h('div', { class: 'tp-note', style: 'min-height:32px' });
    function drawPv() {
      L.clear(pv);
      const r = eff();
      const pp = st.type === 'paragraph' ? D.mergePPr(st.basedOn ? D.styleProps(d, st.basedOn).pPr : {}, st.pPr) : {};
      pv.append(h('div', { style: 'height:3px;background:#ccc' }), h('div', { text: 'Sample Text Sample Text Sample Text Sample Text', style: R.runCSS(Object.assign({}, r, { sz: Math.min(r.sz || 12, 24) }), 1) + `;text-align:${{ center: 'center', right: 'right', both: 'justify' }[pp.jc] || 'left'};white-space:nowrap;overflow:hidden` }), h('div', { style: 'height:3px;background:#ccc' }));
      const parts = [];
      if (st.basedOn) parts.push((d.styles[st.basedOn] ? D.styleDisplayName(d.styles[st.basedOn]) : st.basedOn) + ' +');
      for (const [k, v] of Object.entries(st.rPr)) if (v !== undefined) parts.push(k === 'font' ? 'Font: ' + v : k === 'sz' ? v + ' pt' : k === 'color' ? 'Font color: #' + v : k === 'b' ? (v ? 'Bold' : 'Not Bold') : k === 'i' ? (v ? 'Italic' : 'Not Italic') : k === 'u' ? 'Underline: ' + v : k);
      if (pp.jc) parts.push({ left: 'Left', center: 'Centered', right: 'Right', both: 'Justified' }[pp.jc]);
      desc.textContent = parts.join(', ');
    }
    drawPv();
    const addTpl = ui.check('&Add to template', false, null);
    const auto = ui.check('Automatically &update', !!st.autoRedefine, (v) => { st.autoRedefine = v || undefined; });
    const fmtMenu = (anchor) => ui.openMenu([
      { label: '&Font...', run: () => G.font({ target: st.rPr, base: eff(), onOK: (r) => { Object.assign(st.rPr, r); for (const k in r) if (r[k] === undefined) delete st.rPr[k]; drawPv(); } }) },
      { label: '&Paragraph...', enabled: () => st.type === 'paragraph', run: () => G.paragraph({ target: st.pPr, base: st.basedOn ? D.styleProps(d, st.basedOn).pPr : {}, onOK: (p) => { st.pPr = D.mergePPr(st.pPr, p); drawPv(); } }) },
      { label: '&Tabs...', enabled: () => st.type === 'paragraph', run: () => G.tabs({ target: st.pPr, onOK: (p) => { st.pPr.tabs = p.tabs; } }) },
      { label: '&Border...', enabled: () => st.type === 'paragraph', run: () => G.borders(null, { target: st.pPr, onOK: (r) => { st.pPr.borders = r.borders; st.pPr.shd = r.shd; drawPv(); } }) },
      { label: '&Language...', run: () => G.language({ onOK: (lang) => { st.rPr.lang = lang; } }) },
      { label: '&Numbering...', enabled: () => st.type === 'paragraph', run: () => A().status('Apply numbering to text, then use Update to Match Selection to add it to the style.') },
    ], anchor.getBoundingClientRect());
    const fmtBtn = ui.button('F&ormat ▾', (e) => fmtMenu(e.currentTarget));
    ui.dialog({
      title: isNew ? 'New Style' : 'Modify Style', width: 520,
      body: h('div', { class: 'col' }, ui.group('Properties', G.f('&Name:', name, 'wide'), G.f('Style &type:', type, 'wide'), G.f('Style &based on:', based, 'wide'), nextRow), ui.group('Formatting', fmtRow, pv, desc), G.row(addTpl, auto)),
      buttons: [{ label: 'F&ormat ▾', onClick: () => { fmtMenu(L.$('.dlg:last-of-type .dlg-foot .btn') || fmtBtn); return false; } }, { label: 'OK', primary: true, onClick: () => {
        const nm = name.value.trim();
        if (!nm) { ui.msg('Please enter a name for the style.', { icon: 'warn' }); return false; }
        const exists = D.findStyleByName(d, nm);
        if (isNew && exists) { ui.msg(`A style named ${nm} already exists.`, { icon: 'warn' }); return false; }
        const sid = isNew ? D.uniqueStyleId(d, nm) : id;
        E.edit(isNew ? 'New Style' : 'Modify Style', () => {
          D.touchKey(d, 'styles');
          const out = Object.assign({}, st, { id: sid, name: isNew || !st.builtin ? nm : st.name });
          if (out.next === '__self') out.next = sid;
          if (isNew) { out.custom = true; out.q = true; delete out.builtin; }
          d.styles[sid] = out;
          D.stylesChanged();
          return E.sel;
        });
        if (isNew && st.type === 'paragraph' && E.sel && curP) A().applyStyle(sid);
      } }, { label: 'Cancel' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
  };

  /* ================= Language ================= */
  G.language = function (opts) {
    opts = opts || {};
    const cur = E.curRun().lang || 'en-US';
    const lb = G.listBox(A().LANGS.map(([v, l]) => ({ value: v, label: l })), cur, { height: 180 });
    const noProof = ui.check("&Do not check spelling or grammar", !!E.curRun().noProof, null);
    const detect = ui.check('Detect &language automatically', true, null);
    ui.dialog({
      title: 'Language', width: 320, body: h('div', { class: 'col' }, h('div', { text: 'Mark selected text as:' }), lb, h('div', { class: 'tp-note', text: 'The speller and other proofing tools use the dictionaries of the selected language.' }), noProof, detect),
      buttons: [{ label: '&Default...', onClick: async () => { const r = await ui.msg('Do you want to change the default language to ' + A().langName(lb.get()) + '?', { icon: 'question', buttons: ['&Yes', '&No'] }); if (r === 0) { const d = doc(); E.edit('Default Language', () => { D.touchKey(d, 'defaults'); d.defaults.rPr = Object.assign({}, d.defaults.rPr, { lang: lb.get() }); D.stylesChanged(); return E.sel; }); } return false; } }, { label: 'OK', primary: true, onClick: () => { if (opts.onOK) { opts.onOK(lb.get()); return; } E.formatRun({ lang: lb.get(), noProof: noProof.input.checked || undefined }, 'Language'); A().updateStatus(); } }, { label: 'Cancel' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
  };
})();
