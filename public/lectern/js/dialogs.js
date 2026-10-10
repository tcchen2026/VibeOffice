/* Lectern — dialog boxes. */
(function () {
  'use strict';
  const L = window.L;
  const { h } = L;
  const ui = L.ui;
  const D = (L.dlg = {});
  const E = () => L.ed;
  const M = L.model;
  const IN = (pt) => L.round(pt / 72, 2);
  const PT = (inch) => inch * 72;
  const design = () => M.design(L.pres, E().slide());
  const fillCSS = (f) => L.render.fillCSS(f, design());

  /* ---------- shared pickers ---------- */
  /** a dropdown button that edits a Fill; returns element with .get() */
  function fillPicker(fill, mode, onChange) {
    let cur = L.clone(fill || { t: 'none' });
    const sw = h('span', { class: 'fp-sw', style: 'display:inline-block;width:70px;height:14px;border:1px solid #555;vertical-align:middle' });
    const btn = h('button', { type: 'button', class: 'btn small', style: 'min-width:100px;display:inline-flex;align-items:center;gap:6px' }, sw, h('span', { class: 'dd-arrow' }));
    const paint = () => {
      if (!cur || cur.t === 'none') { sw.style.background = 'repeating-linear-gradient(45deg,#fff 0 4px,#eee 4px 8px)'; sw.title = mode === 'line' ? 'No Line' : 'No Fill'; }
      else sw.style.background = mode === 'line' ? L.model.resolveColor(cur.c, design()) : fillCSS(cur);
    };
    paint();
    btn.addEventListener('click', () => {
      ui.colorMenu(btn, { mode: mode === 'line' ? 'line' : mode === 'plain' ? 'plain' : 'fill', effects: mode !== 'plain' && mode !== 'line', grid: mode === 'plain' }, async (v) => {
        if (v && v.none) cur = mode === 'line' ? { t: 'none' } : { t: 'none' };
        else if (v && v.effects) { const f = await D.fillEffects(cur && cur.t !== 'none' ? cur : { t: 'solid', c: 'accent1', a: 1 }); if (!f) return; cur = f; }
        else if (typeof v === 'string') cur = mode === 'line' ? Object.assign({}, cur && cur.t !== 'none' ? cur : { w: 0.75, dash: 'solid' }, { c: v, t: undefined }) : { t: 'solid', c: v, a: cur && cur.a != null && cur.t === 'solid' ? cur.a : 1 };
        if (cur && cur.t === undefined) delete cur.t;
        paint();
        onChange && onChange(cur);
      });
    });
    btn.get = () => cur;
    btn.set = (f) => { cur = f; paint(); };
    return btn;
  }
  D.fillPicker = fillPicker;
  function colorPick(c, onChange, opts) {
    let cur = c;
    const sw = h('span', { style: 'display:inline-block;width:60px;height:14px;border:1px solid #555;vertical-align:middle' });
    const btn = h('button', { type: 'button', class: 'btn small', style: 'min-width:90px;display:inline-flex;align-items:center;gap:6px' }, sw, h('span', { class: 'dd-arrow' }));
    const paint = () => { sw.style.background = cur ? L.model.resolveColor(cur, design()) : 'repeating-linear-gradient(45deg,#fff 0 4px,#eee 4px 8px)'; };
    paint();
    btn.addEventListener('click', () => ui.colorMenu(btn, { mode: 'plain', grid: true, automatic: opts && opts.automatic }, (v) => { if (v && v.auto) cur = null; else if (typeof v === 'string') cur = v; else return; paint(); onChange && onChange(cur); }));
    btn.get = () => cur;
    return btn;
  }
  D.colorPick = colorPick;
  const LINE_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
  const DASHES = [['solid', 'Solid'], ['sysDot', 'Round Dot'], ['sysDash', 'Square Dot'], ['dash', 'Dash'], ['dashDot', 'Dash Dot'], ['lgDash', 'Long Dash'], ['lgDashDot', 'Long Dash Dot'], ['lgDashDotDot', 'Long Dash Dot Dot']];
  const ARROWS = [['none', 'None'], ['triangle', 'Arrow'], ['arrow', 'Open Arrow'], ['stealth', 'Stealth Arrow'], ['diamond', 'Diamond Arrow'], ['oval', 'Oval Arrow']];
  const SIZES = [['sm', 'Small'], ['med', 'Medium'], ['lg', 'Large']];
  D.LINE_WEIGHTS = LINE_WEIGHTS; D.DASHES = DASHES; D.ARROWS = ARROWS;

  /* ---------- Font ---------- */
  D.font = function () {
    const st = L.fmt.state();
    if (!st) { ui.msg('Select some text or a text object first.'); return; }
    const o = { font: st.font, sz: st.sz, b: st.b, i: st.i, u: st.u, shd: st.shd, emb: !!(st.r && st.r.emb), base: (st.r && st.r.base) || 0, color: st.r ? st.r.color : null };
    const fontList = h('select', { id: 'fd-font', size: 7, style: 'width:180px' });
    L.FONT_LIST.forEach((f) => { const op = h('option', { value: f, text: f }); if (f === o.font) op.selected = true; fontList.appendChild(op); });
    if (!L.FONT_LIST.includes(o.font)) fontList.insertBefore(h('option', { value: o.font, text: o.font, selected: true }), fontList.firstChild);
    const fontIn = h('input', { type: 'text', id: 'fd-font-in', value: o.font, style: 'width:180px' });
    const styleSel = h('select', { id: 'fd-style', size: 4, style: 'width:110px' }, ...['Regular', 'Italic', 'Bold', 'Bold Italic'].map((s) => h('option', { value: s, text: s })));
    styleSel.value = o.b && o.i ? 'Bold Italic' : o.b ? 'Bold' : o.i ? 'Italic' : 'Regular';
    const sizeIn = h('input', { type: 'text', id: 'fd-size', value: o.sz, style: 'width:60px' });
    const sizeSel = h('select', { size: 4, style: 'width:60px' }, ...L.SIZE_LIST.map((s) => h('option', { value: s, text: s })));
    sizeSel.value = String(o.sz);
    const preview = h('div', { class: 'preview-box font-preview', text: 'AaBbYyZz' });
    const upd = () => {
      const r = preview.style;
      r.fontFamily = L.fontStack(fontIn.value);
      r.fontWeight = /Bold/.test(styleSel.value) ? '700' : '400';
      r.fontStyle = /Italic/.test(styleSel.value) ? 'italic' : 'normal';
      r.textDecoration = u.input.checked ? 'underline' : 'none';
      r.textShadow = s.input.checked ? '2px 2px 0 rgba(0,0,0,.35)' : emb.input.checked ? '-1px -1px 0 #fff, 1px 1px 0 #666' : 'none';
      r.color = col.get() ? L.model.resolveColor(col.get(), design()) : '#000';
      r.fontSize = Math.min(40, Math.max(10, parseFloat(sizeIn.value) || 18)) + 'px';
    };
    fontList.addEventListener('change', () => { fontIn.value = fontList.value; upd(); });
    fontIn.addEventListener('input', upd);
    styleSel.addEventListener('change', upd);
    sizeSel.addEventListener('change', () => { sizeIn.value = sizeSel.value; upd(); });
    sizeIn.addEventListener('input', upd);
    const u = ui.check('&Underline', o.u, upd), s = ui.check('&Shadow', o.shd, upd), emb = ui.check('&Emboss', o.emb, upd);
    const sup = ui.check('Su&perscript', o.base > 0, () => { if (sup.input.checked) sub.input.checked = false; }), sub = ui.check('Su&bscript', o.base < 0, () => { if (sub.input.checked) sup.input.checked = false; });
    const off = ui.spin({ value: o.base ? Math.abs(o.base / 1000) : 30, min: 1, max: 100, step: 1, unit: '%', dec: 0 });
    const col = colorPick(o.color, upd, { automatic: true });
    const body = h('div', { class: 'col' },
      h('div', { class: 'row' },
        h('div', { class: 'col' }, h('label', { for: 'fd-font-in', html: '<u>F</u>ont:' }), fontIn, fontList),
        h('div', { class: 'col' }, h('label', { for: 'fd-style', html: 'F<u>o</u>nt style:' }), styleSel),
        h('div', { class: 'col' }, h('label', { for: 'fd-size', html: '<u>S</u>ize:' }), sizeIn, sizeSel)),
      h('div', { class: 'row' },
        ui.group('Effects', h('div', { class: 'row' }, h('div', { class: 'col' }, u, s, emb), h('div', { class: 'col' }, sup, sub, h('div', { class: 'field' }, h('label', { text: 'Offset:' }), off)))),
        h('div', { class: 'col' }, h('label', { html: '<u>C</u>olor:' }), col)),
      preview,
      h('div', { class: 'tp-note', text: 'This is a TrueType font. The same font will be used on both your printer and your screen.' }));
    upd();
    ui.dialog({
      title: 'Font', body, width: 470, buttons: [{
        label: 'OK', primary: true, onClick: () => {
          const props = {
            font: fontIn.value.trim() || o.font, sz: L.clamp(parseFloat(sizeIn.value) || o.sz, 1, 4000), b: /Bold/.test(styleSel.value), i: /Italic/.test(styleSel.value),
            u: u.input.checked ? 'sng' : undefined, shd: s.input.checked || undefined, emb: emb.input.checked || undefined,
            base: sup.input.checked ? off.get() * 1000 : sub.input.checked ? -off.get() * 1000 : undefined, color: col.get() || undefined,
          };
          L.fmt.run((r) => { const n = Object.assign({}, r, props); for (const k in props) if (props[k] === undefined) delete n[k]; if (props.color) delete n.fill; return n; }, 'Font', (wa) => { wa.font = props.font; wa.b = props.b; wa.i = props.i; });
        },
      }, { label: 'Cancel' }],
    });
  };

  /* ---------- Bullets and Numbering ---------- */
  const BULLET_SETS = [null, '•', '○', '■', '❖', '➢', '✓', '▪'];
  const NUM_SETS = [null, 'arabicPeriod', 'arabicParenR', 'romanUcPeriod', 'alphaUcPeriod', 'alphaLcParenR', 'alphaLcPeriod', 'romanLcPeriod'];
  D.bullets = function (tabIdx) {
    const st = L.fmt.state();
    if (!st || st.wa) { ui.msg('Select some text first.'); return; }
    const cur = st.ps.bu || { t: 'none' };
    let pick = cur.t === 'char' ? cur.ch : cur.t === 'num' ? cur.scheme : null;
    let kind = cur.t === 'num' ? 'num' : 'char';
    let customCh = cur.t === 'char' ? cur.ch : '•';
    const sizeSp = ui.spin({ value: Math.round((cur.sz || 1) * 100), min: 25, max: 400, step: 5, unit: '%', dec: 0 });
    const sizeSp2 = ui.spin({ value: Math.round((cur.sz || 1) * 100), min: 25, max: 400, step: 5, unit: '%', dec: 0 });
    const col = colorPick(cur.c || null, null, { automatic: true });
    const col2 = colorPick(cur.c || null, null, { automatic: true });
    const startSp = ui.spin({ value: cur.start || 1, min: 1, max: 9999, step: 1, dec: 0 });
    const bg = h('div', { class: 'bul-grid' });
    const ng = h('div', { class: 'num-grid' });
    const mkBtn = (grid, val, label, isNum) => {
      const b = h('button', { type: 'button' });
      if (val == null) b.appendChild(h('span', { text: 'None', style: 'font-size:13px;align-self:center' }));
      else for (let i = 1; i <= 3; i++) b.appendChild(h('div', null, h('span', { text: isNum ? L.render.numberText(val, i) : val, style: 'min-width:16px' }), h('i')));
      if (pick === val && ((isNum && kind === 'num') || (!isNum && kind === 'char') || val == null)) b.classList.add('on');
      b.addEventListener('click', () => { L.$$('button', bg).concat(L.$$('button', ng)).forEach((x) => x.classList.remove('on')); b.classList.add('on'); pick = val; kind = isNum ? 'num' : 'char'; if (!isNum && val) customCh = val; });
      b.addEventListener('dblclick', () => { okBtn.click(); });
      grid.appendChild(b);
      void label;
    };
    BULLET_SETS.forEach((v) => mkBtn(bg, v, '', false));
    NUM_SETS.forEach((v) => mkBtn(ng, v, '', true));
    const customize = ui.button('C&ustomize...', () => D.symbol((ch) => { customCh = ch; pick = ch; kind = 'char'; const b = L.$$('button', bg)[7]; L.clear(b); for (let i = 0; i < 3; i++) b.appendChild(h('div', null, h('span', { text: ch, style: 'min-width:16px' }), h('i'))); L.$$('button', bg).forEach((x) => x.classList.remove('on')); b.classList.add('on'); BULLET_SETS[7] = ch; }));
    const tabs = ui.tabs([
      { label: 'Bulle&ted', body: h('div', { class: 'col' }, bg, h('div', { class: 'row', style: 'align-items:center;margin-top:8px' }, h('label', { text: 'Size:' }), sizeSp, h('span', { text: '% of text' }), h('label', { text: 'Color:' }), col, customize)) },
      { label: '&Numbered', body: h('div', { class: 'col' }, ng, h('div', { class: 'row', style: 'align-items:center;margin-top:8px' }, h('label', { text: 'Size:' }), sizeSp2, h('span', { text: '% of text' }), h('label', { text: 'Color:' }), col2, h('label', { text: 'Start at:' }), startSp)) },
    ], tabIdx != null ? tabIdx : kind === 'num' ? 1 : 0);
    const d = ui.dialog({
      title: 'Bullets and Numbering', body: tabs, width: 420, buttons: [{
        label: 'OK', primary: true, onClick: () => {
          let bu;
          if (pick == null) bu = { t: 'none' };
          else if (kind === 'num') bu = { t: 'num', scheme: pick, start: startSp.get(), sz: sizeSp2.get() / 100, c: col2.get() || undefined };
          else bu = { t: 'char', ch: pick || customCh, sz: sizeSp.get() / 100, c: col.get() || undefined };
          if (bu.sz === 1) delete bu.sz;
          L.fmt.para((p, i, t) => {
            p.pp = Object.assign({}, p.pp, { bu });
            const cls = L.style.cls(t ? t.sh : null);
            if (bu.t !== 'none' && cls === 'other' && !(p.pp.marL > 0)) { p.pp.marL = 27 + (p.lvl || 0) * 36; p.pp.indent = -27; }
            if (bu.t === 'none' && cls === 'other') { delete p.pp.marL; delete p.pp.indent; }
          }, 'Bullets and Numbering');
        },
      }, { label: 'Cancel' }],
    });
    const okBtn = d.buttons[0];
  };

  /* ---------- Line spacing ---------- */
  D.lineSpacing = function () {
    const st = L.fmt.state();
    if (!st || st.wa) { ui.msg('Select some text first.'); return; }
    const ps = st.ps;
    const mk = (sp, def) => {
      const isPts = sp && sp.pts != null;
      const v = sp ? (isPts ? sp.pts : (sp.pct || 0) / 100) : def;
      const spin = ui.spin({ value: v, min: 0, max: 1584, step: isPts ? 1 : 0.1, dec: 2 });
      const unit = ui.select([['l', 'Lines'], ['p', 'Points']], isPts ? 'p' : 'l', (u2) => { spin.set(u2 === 'p' ? L.round(spin.get() * (st.sz || 18) * 1.2, 1) : L.round(spin.get() / ((st.sz || 18) * 1.2), 2)); });
      return { spin, unit, get: () => (unit.value === 'p' ? { pts: spin.get() } : { pct: L.round(spin.get() * 100, 1) }) };
    };
    const ls = mk(ps.lnSpc, 1), sb = mk(ps.spcBef, 0.2), sa = mk(ps.spcAft, 0);
    const body = h('div', { class: 'col' },
      ui.group('Line spacing', h('div', { class: 'row' }, ls.spin, ls.unit)),
      ui.group('Before paragraph', h('div', { class: 'row' }, sb.spin, sb.unit)),
      ui.group('After paragraph', h('div', { class: 'row' }, sa.spin, sa.unit)));
    ui.dialog({ title: 'Line Spacing', body, width: 280, buttons: [{ label: 'OK', primary: true, onClick: () => L.fmt.para((p) => { p.pp = Object.assign({}, p.pp, { lnSpc: ls.get(), spcBef: sb.get(), spcAft: sa.get() }); }, 'Line Spacing') }, { label: 'Cancel' }] });
  };

  /* ---------- Change Case ---------- */
  D.changeCase = function () {
    let mode = 'sentence';
    const opts = [['sentence', '&Sentence case.'], ['lower', '&lowercase'], ['upper', '&UPPERCASE'], ['title', '&Title Case'], ['toggle', 't&OGGLE cASE']];
    const body = h('div', { class: 'col' }, ...opts.map(([k, l], i) => ui.radio('cc', l, i === 0, () => { mode = k; })));
    ui.dialog({ title: 'Change Case', body, width: 240, buttons: [{ label: 'OK', primary: true, onClick: () => D.applyCase(mode) }, { label: 'Cancel' }] });
  };
  D.applyCase = function (mode) {
    const fn = { sentence: L.sentenceCase, lower: (s) => s.toLowerCase(), upper: (s) => s.toUpperCase(), title: L.titleCase, toggle: L.toggleCase }[mode];
    L.fmt.run((r) => (r.fld || typeof r.t !== 'string' ? r : Object.assign({}, r, { t: fn(r.t) })), 'Change Case', (wa) => { wa.text = fn(wa.text); });
  };

  /* ---------- Replace Fonts ---------- */
  D.replaceFonts = function () {
    const used = new Set();
    const scan = (tx, d) => { for (const p of tx.ps) for (const r of p.rs) used.add(L.style.font(r.font || '+mn', d)); };
    for (const s of L.pres.slides) { const d = M.design(L.pres, s); M.walk(s.shapes, (sh) => { if (sh.tx) scan(sh.tx, d); if (sh.tbl) sh.tbl.rows.forEach((r) => r.cells.forEach((c) => scan(c.tx, d))); }); used.add(d.fonts.major); used.add(d.fonts.minor); }
    const from = ui.select(Array.from(used), Array.from(used)[0]);
    const to = ui.select(L.FONT_LIST, 'Arial');
    const body = h('div', { class: 'col' }, ui.field('&Replace:', from), ui.field('&With:', to));
    ui.dialog({
      title: 'Replace Font', body, width: 330, buttons: [{
        label: 'Replace', primary: true, onClick: () => {
          const a = from.value, b = to.value;
          L.hist.push('Replace Fonts');
          for (const id in L.pres.designs) { const d = L.pres.designs[id]; if (d.fonts.major === a) d.fonts.major = b; if (d.fonts.minor === a) d.fonts.minor = b; }
          const fix = (tx, d) => { for (const p of tx.ps) { for (const r of p.rs) if (L.style.font(r.font || '+mn', d) === a && r.font && r.font[0] !== '+') r.font = b; if (p.end && p.end.font === a) p.end.font = b; } };
          for (const s of L.pres.slides) { const d = M.design(L.pres, s); M.walk(s.shapes, (sh) => { if (sh.tx) fix(sh.tx, d); if (sh.tbl) sh.tbl.rows.forEach((r) => r.cells.forEach((c) => fix(c.tx, d))); if (sh.wa && sh.wa.font === a) sh.wa.font = b; }); }
          E().render(); L.bus.emit('slides-changed');
        },
      }, { label: 'Close' }],
    });
  };

  /* ---------- Fill Effects ---------- */
  const GRAD_PRESETS = [
    ['Early Sunset', [[0, '#000082'], [0.3, '#66008F'], [0.65, '#BA0066'], [1, '#FF8200']]],
    ['Late Sunset', [[0, '#000000'], [0.5, '#000040'], [1, '#FF8C00']]],
    ['Nightfall', [[0, '#000000'], [0.6, '#0A128C'], [1, '#181CC7']]],
    ['Daybreak', [[0, '#5E9EFF'], [0.4, '#85C2FF'], [0.7, '#C4D6EB'], [1, '#FFEBFA']]],
    ['Horizon', [[0, '#DCEBF5'], [0.4, '#83A7C3'], [0.5, '#768FB9'], [0.6, '#83A7C3'], [1, '#FFFFFF']]],
    ['Desert', [[0, '#9C3500'], [0.3, '#E8A15C'], [0.6, '#FFDF9C'], [1, '#FFFFFF']]],
    ['Ocean', [[0, '#03D4A8'], [0.5, '#21D6E0'], [1, '#0087E6']]],
    ['Calm Water', [[0, '#CCCCFF'], [0.5, '#99CCFF'], [1, '#CCCCFF']]],
    ['Fire', [[0, '#FFF200'], [0.45, '#FF7A00'], [0.7, '#FF0300'], [1, '#4D0808']]],
    ['Fog', [[0, '#8488C4'], [0.5, '#D4DEFF'], [1, '#96AB94']]],
    ['Moss', [[0, '#DDEBCF'], [0.5, '#9CB86E'], [1, '#156B13']]],
    ['Peacock', [[0, '#3399FF'], [0.2, '#00CCCC'], [0.5, '#9999FF'], [0.8, '#2E6792'], [1, '#3333CC']]],
    ['Wheat', [[0, '#FBEAC7'], [0.2, '#FEE7F2'], [0.5, '#FAC77D'], [0.8, '#FBA97D'], [1, '#FBD49C']]],
    ['Parchment', [[0, '#FFEFD1'], [0.65, '#F0EBD5'], [1, '#D1C39F']]],
    ['Mahogany', [[0, '#D6B19C'], [0.3, '#D49E6C'], [0.7, '#A65528'], [1, '#663012']]],
    ['Rainbow', [[0, '#A603AB'], [0.2, '#0819FB'], [0.4, '#1A8D48'], [0.6, '#FFFF00'], [0.8, '#EE3F17'], [1, '#E81766']]],
    ['Gold', [[0, '#E6DCAC'], [0.15, '#E6D78A'], [0.4, '#C7AC4C'], [0.6, '#E6D78A'], [1, '#E6DCAC']]],
    ['Brass', [[0, '#825600'], [0.15, '#FFA800'], [0.35, '#825600'], [0.55, '#FFA800'], [0.8, '#825600'], [1, '#FFA800']]],
    ['Chrome', [[0, '#FFFFFF'], [0.2, '#000000'], [0.4, '#FFFFFF'], [0.6, '#000000'], [0.8, '#FFFFFF'], [1, '#000000']]],
    ['Silver', [[0, '#FFFFFF'], [0.3, '#E6E6E6'], [0.6, '#7D8496'], [1, '#FFFFFF']]],
    ['Sapphire', [[0, '#000082'], [0.5, '#0047FF'], [1, '#000082']]],
  ];
  /* procedural textures (generated tiles) */
  const TEXTURES = ['Canvas', 'Denim', 'Paper bag', 'Newsprint', 'Recycled paper', 'Parchment', 'Granite', 'White marble', 'Green marble', 'Sand', 'Cork', 'Oak', 'Walnut', 'Water droplets', 'Blue tissue paper', 'Woven mat', 'Stationery', 'Fabric'];
  const texCache = {};
  function texture(name) {
    if (texCache[name]) return texCache[name];
    const S = 128;
    const c = document.createElement('canvas');
    c.width = S; c.height = S;
    const g = c.getContext('2d');
    let seed = name.split('').reduce((a, ch) => a * 31 + ch.charCodeAt(0), 7) >>> 0;
    const rnd = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; };
    const noise = (base, amt, scale) => { const img = g.getImageData(0, 0, S, S); const [r, gg, b] = L.color.hexToRgb(base); for (let i = 0; i < S * S; i++) { const n = (rnd() - 0.5) * amt * (scale || 1); img.data[i * 4] = r + n; img.data[i * 4 + 1] = gg + n; img.data[i * 4 + 2] = b + n; img.data[i * 4 + 3] = 255; } g.putImageData(img, 0, 0); };
    const P = {
      'Canvas': () => { noise('#E9E1CC', 30); g.globalAlpha = 0.15; for (let i = 0; i < S; i += 2) { g.fillStyle = i % 4 ? '#fff' : '#a08a60'; g.fillRect(0, i, S, 1); g.fillRect(i, 0, 1, S); } },
      'Denim': () => { noise('#3E5F8A', 40); g.globalAlpha = 0.25; g.strokeStyle = '#cfe0f5'; for (let i = -S; i < S; i += 3) { g.beginPath(); g.moveTo(i, 0); g.lineTo(i + S, S); g.stroke(); } },
      'Paper bag': () => { noise('#B8915F', 26); g.globalAlpha = 0.2; for (let i = 0; i < 40; i++) { g.fillStyle = rnd() > 0.5 ? '#8a6a3e' : '#d6b88a'; g.fillRect(rnd() * S, rnd() * S, rnd() * 20, 1); } },
      'Newsprint': () => { noise('#E3E0D8', 20); g.globalAlpha = 0.2; g.fillStyle = '#555'; for (let i = 0; i < 200; i++) g.fillRect(rnd() * S, rnd() * S, 1, 1); },
      'Recycled paper': () => { noise('#DCD3C0', 22); g.globalAlpha = 0.5; for (let i = 0; i < 80; i++) { g.fillStyle = ['#7a6b55', '#a49884', '#5e5444'][i % 3]; g.fillRect(rnd() * S, rnd() * S, 1 + rnd() * 2, 1); } },
      'Parchment': () => { const gr = g.createRadialGradient(S / 2, S / 2, 10, S / 2, S / 2, S * 0.75); gr.addColorStop(0, '#FBF1D4'); gr.addColorStop(1, '#E2CC97'); g.fillStyle = gr; g.fillRect(0, 0, S, S); g.globalAlpha = 0.08; noise('#C9B27A', 40); },
      'Granite': () => { noise('#9A9A9A', 60); g.globalAlpha = 0.6; for (let i = 0; i < 400; i++) { g.fillStyle = ['#2b2b2b', '#e8e8e8', '#6b4f4f'][i % 3]; g.fillRect(rnd() * S, rnd() * S, 1 + rnd() * 2, 1 + rnd() * 2); } },
      'White marble': () => { noise('#F2F2F0', 12); g.globalAlpha = 0.35; g.strokeStyle = '#9a9a9a'; for (let k = 0; k < 6; k++) { g.beginPath(); let x = rnd() * S, y = 0; g.moveTo(x, y); while (y < S) { x += (rnd() - 0.5) * 14; y += 6; g.lineTo(x, y); } g.stroke(); } },
      'Green marble': () => { noise('#2E6B4F', 30); g.globalAlpha = 0.45; g.strokeStyle = '#cfe8d8'; for (let k = 0; k < 6; k++) { g.beginPath(); let x = rnd() * S, y = 0; g.moveTo(x, y); while (y < S) { x += (rnd() - 0.5) * 14; y += 6; g.lineTo(x, y); } g.stroke(); } },
      'Sand': () => { noise('#D9C29A', 40); },
      'Cork': () => { noise('#B07A45', 50); g.globalAlpha = 0.5; for (let i = 0; i < 300; i++) { g.fillStyle = rnd() > 0.5 ? '#6b4423' : '#d9a066'; g.fillRect(rnd() * S, rnd() * S, 2, 2); } },
      'Oak': () => { g.fillStyle = '#C49A6C'; g.fillRect(0, 0, S, S); g.globalAlpha = 0.35; for (let y = 0; y < S; y += 2) { g.strokeStyle = rnd() > 0.5 ? '#8a5f34' : '#e0bd8f'; g.beginPath(); g.moveTo(0, y); for (let x = 0; x <= S; x += 8) g.lineTo(x, y + Math.sin(x / 15 + y / 20) * 2); g.stroke(); } },
      'Walnut': () => { g.fillStyle = '#5C3A21'; g.fillRect(0, 0, S, S); g.globalAlpha = 0.35; for (let y = 0; y < S; y += 2) { g.strokeStyle = rnd() > 0.5 ? '#3a2312' : '#8a5a36'; g.beginPath(); g.moveTo(0, y); for (let x = 0; x <= S; x += 8) g.lineTo(x, y + Math.sin(x / 12 + y / 18) * 3); g.stroke(); } },
      'Water droplets': () => { noise('#9CC3D9', 14); for (let i = 0; i < 40; i++) { const x = rnd() * S, y = rnd() * S, r = 2 + rnd() * 6; const gr = g.createRadialGradient(x - r / 3, y - r / 3, 0, x, y, r); gr.addColorStop(0, 'rgba(255,255,255,.9)'); gr.addColorStop(1, 'rgba(60,110,150,.35)'); g.fillStyle = gr; g.beginPath(); g.arc(x, y, r, 0, Math.PI * 2); g.fill(); } },
      'Blue tissue paper': () => { noise('#BFD3EE', 18); g.globalAlpha = 0.15; for (let i = 0; i < 30; i++) { g.strokeStyle = '#fff'; g.beginPath(); g.moveTo(rnd() * S, rnd() * S); g.lineTo(rnd() * S, rnd() * S); g.stroke(); } },
      'Woven mat': () => { g.fillStyle = '#D8C08E'; g.fillRect(0, 0, S, S); for (let y = 0; y < S; y += 16) for (let x = 0; x < S; x += 16) { g.fillStyle = (x + y) % 32 ? '#c2a46b' : '#e8d3a5'; g.fillRect(x, y, 16, 8); g.fillStyle = (x + y) % 32 ? '#e8d3a5' : '#c2a46b'; g.fillRect(x, y + 8, 16, 8); } g.globalAlpha = 0.1; noise('#8a6a3e', 50); },
      'Stationery': () => { g.fillStyle = '#FFFFF4'; g.fillRect(0, 0, S, S); g.strokeStyle = '#B9D3EE'; for (let y = 8; y < S; y += 16) { g.beginPath(); g.moveTo(0, y + 0.5); g.lineTo(S, y + 0.5); g.stroke(); } },
      'Fabric': () => { noise('#B04A5A', 30); g.globalAlpha = 0.2; for (let i = 0; i < S; i += 4) { g.fillStyle = '#fff'; g.fillRect(i, 0, 1, S); g.fillRect(0, i, S, 1); } },
    };
    (P[name] || P.Canvas)();
    g.globalAlpha = 1;
    texCache[name] = c.toDataURL('image/png');
    return texCache[name];
  }
  async function textureMedia(name) {
    const url = texture(name);
    const blob = await (await fetch(url)).blob();
    return L.media.add(blob, name + '.png');
  }
  D.texture = texture;

  /** Fill Effects dialog → Promise<Fill|null> */
  D.fillEffects = function (initial) {
    return new Promise((resolve) => {
      let res = null;
      const d0 = design();
      const base = initial && initial.t === 'solid' ? L.model.resolveColor(initial.c, d0) : '#3366CC';
      let g = initial && initial.t === 'grad' ? L.clone(initial) : { t: 'grad', stops: [{ p: 0, c: base, a: 1 }, { p: 1, c: L.color.lighten(base, 0.75), a: 1 }], ang: 90, path: 'lin' };
      let mode = 'two', c1 = g.stops[0].c, c2 = g.stops[g.stops.length - 1].c, dark = 0.5, preset = 0, style = g.path !== 'lin' ? (g.path === 'circle' ? 'center' : 'corner') : g.ang === 0 ? 'vert' : g.ang === 45 ? 'dup' : g.ang === 135 ? 'ddn' : 'horz', variant = 0;
      let tr1 = 0, tr2 = 0;
      let current = initial && initial.t !== 'none' ? L.clone(initial) : g;
      const sample = h('div', { class: 'preview-box', style: 'width:110px;height:110px' }, h('div', { class: 'fx-swatch' }));
      const paintSample = () => { sample.firstChild.style.background = fillCSS(current); };
      const variants = h('div', { class: 'row', style: 'gap:6px' });
      const buildGrad = (v) => {
        let stops;
        const cc1 = L.model.resolveColor(c1, d0), cc2 = mode === 'one' ? (dark < 0.5 ? L.color.darken(cc1, 1 - dark * 2) : L.color.lighten(cc1, (dark - 0.5) * 2)) : L.model.resolveColor(c2, d0);
        if (mode === 'preset') stops = GRAD_PRESETS[preset][1].map(([p, c]) => ({ p, c, a: 1 }));
        else stops = [{ p: 0, c: cc1, a: 1 - tr1 }, { p: 1, c: cc2, a: 1 - tr2 }];
        if (v === 1 || v === 3) stops = stops.map((s) => ({ p: 1 - s.p, c: s.c, a: s.a })).reverse();
        if (v >= 2) stops = stops.map((s) => ({ p: s.p / 2, c: s.c, a: s.a })).concat(stops.slice(0, -1).reverse().map((s) => ({ p: 1 - s.p / 2, c: s.c, a: s.a })));
        const ang = { horz: 90, vert: 0, dup: 315, ddn: 45, corner: 0, center: 0 }[style];
        return { t: 'grad', stops, ang, path: style === 'center' ? 'circle' : style === 'corner' ? 'rect' : 'lin' };
      };
      const refreshVariants = () => {
        L.clear(variants);
        for (let v = 0; v < 4; v++) {
          const gv = buildGrad(v);
          const b = h('button', { type: 'button', style: `width:46px;height:46px;border:1px solid #888;padding:0;background:${fillCSS(gv)};${v === variant ? 'outline:2px solid #0b3fa8;outline-offset:1px' : ''}` });
          b.addEventListener('click', () => { variant = v; current = buildGrad(v); paintSample(); refreshVariants(); });
          variants.appendChild(b);
        }
        current = buildGrad(variant); paintSample();
      };
      const p1 = colorPick(c1, (c) => { c1 = c; refreshVariants(); }), p2 = colorPick(c2, (c) => { c2 = c; refreshVariants(); });
      const darkR = h('input', { type: 'range', min: 0, max: 100, value: 50, id: 'fe-dark' });
      darkR.addEventListener('input', () => { dark = darkR.value / 100; refreshVariants(); });
      const presetSel = ui.select(GRAD_PRESETS.map((p, i) => [i, p[0]]), 0, (v) => { preset = +v; refreshVariants(); });
      const t1 = ui.spin({ value: 0, min: 0, max: 100, step: 5, unit: ' %', dec: 0, onChange: (v) => { tr1 = v / 100; refreshVariants(); } });
      const t2 = ui.spin({ value: 0, min: 0, max: 100, step: 5, unit: ' %', dec: 0, onChange: (v) => { tr2 = v / 100; refreshVariants(); } });
      const colorsArea = h('div', { class: 'col' });
      const setMode = (m) => {
        mode = m;
        L.clear(colorsArea);
        if (m === 'one') colorsArea.append(ui.field('Color &1:', p1), h('div', { class: 'row', style: 'align-items:center' }, h('span', { text: 'Dark' }), darkR, h('span', { text: 'Light' })));
        if (m === 'two') colorsArea.append(ui.field('Color &1:', p1), ui.field('Color &2:', p2));
        if (m === 'preset') colorsArea.append(ui.field('Preset colors:', presetSel));
        refreshVariants();
      };
      const shadingRadios = h('div', { class: 'col' }, ...[['horz', '&Horizontal'], ['vert', '&Vertical'], ['dup', 'Diagonal &up'], ['ddn', 'Dia&gonal down'], ['corner', '&From corner'], ['center', 'Fro&m center']].map(([k, l]) => ui.radio('fe-shade', l, style === k, () => { style = k; variant = 0; refreshVariants(); })));
      const gradTab = h('div', { class: 'col' },
        h('div', { class: 'row' },
          ui.group('Colors', h('div', { class: 'col' }, ui.radio('fe-mode', '&One color', false, () => setMode('one')), ui.radio('fe-mode', '&Two colors', true, () => setMode('two')), ui.radio('fe-mode', 'Pr&eset', false, () => setMode('preset')))),
          colorsArea),
        ui.group('Transparency', h('div', { class: 'col' }, ui.field('F&rom:', t1), ui.field('T&o:', t2))),
        h('div', { class: 'row' }, ui.group('Shading styles', shadingRadios), ui.group('Variants', variants)));
      /* texture */
      const texGrid = h('div', { class: 'tex-grid' });
      let texSel = null;
      TEXTURES.forEach((n) => {
        const b = h('button', { type: 'button', 'data-tip': n, 'aria-label': n, style: `background-image:url(${texture(n)});background-size:44px 44px` });
        b.addEventListener('click', async () => { L.$$('button', texGrid).forEach((x) => x.classList.remove('on')); b.classList.add('on'); texSel = n; texName.textContent = n; const id = await textureMedia(n); current = { t: 'img', media: id, tile: true }; paintSample(); });
        texGrid.appendChild(b);
      });
      const texName = h('div', { class: 'tp-note', text: ' ' });
      const otherTex = ui.button('&Other Texture...', async () => { const f = await L.pickFiles('image/*'); if (f[0]) { const id = L.media.add(f[0], f[0].name); current = { t: 'img', media: id, tile: true }; paintSample(); texName.textContent = f[0].name; } });
      const texTab = h('div', { class: 'col' }, h('label', { text: 'Texture:' }), texGrid, texName, otherTex);
      /* pattern */
      let pfg = '#000000', pbg = '#FFFFFF', pprst = initial && initial.t === 'patt' ? initial.prst : 'pct50';
      const pattGrid = h('div', { class: 'patt-grid' });
      const pattName = h('div', { class: 'tp-note', text: pprst });
      const drawPatt = () => {
        L.clear(pattGrid);
        Object.keys(L.render.PATTERNS).forEach((k) => {
          const b = h('button', { type: 'button', 'aria-label': k, 'data-tip': k, class: k === pprst ? 'on' : '', style: `background:${L.render.patternURL(k, L.model.resolveColor(pfg, d0), L.model.resolveColor(pbg, d0))}` });
          b.addEventListener('click', () => { pprst = k; pattName.textContent = k; current = { t: 'patt', prst: k, fg: pfg, bg: pbg }; paintSample(); drawPatt(); });
          pattGrid.appendChild(b);
        });
      };
      drawPatt();
      const fgP = colorPick(pfg, (c) => { pfg = c; current = { t: 'patt', prst: pprst, fg: pfg, bg: pbg }; paintSample(); drawPatt(); });
      const bgP = colorPick(pbg, (c) => { pbg = c; current = { t: 'patt', prst: pprst, fg: pfg, bg: pbg }; paintSample(); drawPatt(); });
      const pattTab = h('div', { class: 'col' }, h('label', { text: 'Pattern:' }), pattGrid, pattName, h('div', { class: 'row' }, ui.field('&Foreground:', fgP), ui.field('&Background:', bgP)));
      /* picture */
      const picPrev = h('div', { class: 'preview-box', style: 'width:220px;height:150px' });
      const picTab = h('div', { class: 'col' }, ui.button('Select &Picture...', async () => { const f = await L.pickFiles('image/*'); if (!f[0]) return; const id = L.media.add(f[0], f[0].name); current = { t: 'img', media: id }; picPrev.style.background = `url(${L.media.url(id)}) center/contain no-repeat`; paintSample(); }), picPrev);
      const tabs = ui.tabs([{ label: '&Gradient', body: gradTab }, { label: 'Te&xture', body: texTab }, { label: '&Pattern', body: pattTab }, { label: 'P&icture', body: picTab }], initial && initial.t === 'patt' ? 2 : initial && initial.t === 'img' ? (initial.tile ? 1 : 3) : 0);
      setMode('two');
      if (initial && initial.t !== 'grad') { current = L.clone(initial); paintSample(); }
      const body = h('div', { class: 'row', style: 'flex-wrap:nowrap' }, tabs, h('div', { class: 'col', style: 'flex:none;padding-top:26px' }, h('label', { text: 'Sample:' }), sample));
      ui.dialog({ title: 'Fill Effects', body, width: 560, buttons: [{ label: 'OK', primary: true, onClick: () => { res = current; } }, { label: 'Cancel' }] }).done.then(() => resolve(res));
      void texSel;
    });
  };

  /* ---------- Format AutoShape / Picture / Text Box ---------- */
  D.formatShape = function (tabName) {
    const shapes = E().selected();
    if (!shapes.length) return;
    const s0 = shapes[0];
    const isPic = s0.type === 'image', isLine = s0.type === 'line', isWA = s0.type === 'wordart', isTable = s0.type === 'table';
    const title = isPic ? 'Format Picture' : isLine ? 'Format AutoShape' : isWA ? 'Format LettersArt' : s0.ph ? 'Format Placeholder' : s0.type === 'text' ? 'Format Text Box' : isTable ? 'Format Table' : 'Format AutoShape';
    const ch = {}; /* changes */
    const srcFill = isWA ? s0.wa.fill : s0.fill;
    const srcLine = isWA ? s0.wa.line : s0.line;
    /* colors & lines */
    const fp = fillPicker(srcFill, 'fill', (f) => { ch.fill = f; trans.set(f && f.a != null ? Math.round((1 - f.a) * 100) : 0); });
    const trans = ui.spin({ value: srcFill && srcFill.a != null && srcFill.t === 'solid' ? Math.round((1 - srcFill.a) * 100) : 0, min: 0, max: 100, step: 5, unit: ' %', dec: 0, onChange: (v) => { ch.alpha = 1 - v / 100; } });
    const lp = fillPicker(srcLine && srcLine.t !== 'none' ? srcLine : { t: 'none' }, 'line', (l) => { ch.lineColor = l; });
    const lnW = ui.spin({ value: srcLine && srcLine.w != null ? srcLine.w : 0.75, min: 0, max: 1584, step: 0.25, unit: ' pt', dec: 2, onChange: (v) => { ch.lineW = v; } });
    const dash = ui.select(DASHES, (srcLine && srcLine.dash) || 'solid', (v) => { ch.dash = v; });
    const cmpd = ui.select([['sng', 'Single'], ['dbl', 'Double'], ['thickThin', 'Thick Thin'], ['thinThick', 'Thin Thick'], ['tri', 'Triple']], (srcLine && srcLine.cmpd) || 'sng', (v) => { ch.cmpd = v; });
    const aB = ui.select(ARROWS, (srcLine && srcLine.head && srcLine.head.type) || 'none', (v) => { ch.head = v; });
    const aE = ui.select(ARROWS, (srcLine && srcLine.tail && srcLine.tail.type) || 'none', (v) => { ch.tail = v; });
    const sB = ui.select(SIZES, (srcLine && srcLine.head && srcLine.head.w) || 'med', (v) => { ch.headSz = v; });
    const sE = ui.select(SIZES, (srcLine && srcLine.tail && srcLine.tail.w) || 'med', (v) => { ch.tailSz = v; });
    const colorsTab = h('div', { class: 'col' },
      !isLine && !isPic ? ui.group('Fill', ui.field('C&olor:', fp), ui.field('&Transparency:', trans)) : null,
      isPic ? ui.group('Fill', ui.field('C&olor:', fp)) : null,
      ui.group('Line', h('div', { class: 'row' }, h('div', { class: 'col' }, ui.field('Co&lor:', lp), ui.field('&Dashed:', dash)), h('div', { class: 'col' }, ui.field('&Style:', cmpd), ui.field('&Weight:', lnW)))),
      isLine || (s0.type === 'shape' && L.geom.get(s0.geom).line) || (s0.geom === 'custom') ? ui.group('Arrows', h('div', { class: 'row' }, h('div', { class: 'col' }, ui.field('&Begin style:', aB), ui.field('Begin si&ze:', sB)), h('div', { class: 'col' }, ui.field('&End style:', aE), ui.field('E&nd size:', sE)))) : null);
    /* size */
    const hS = ui.spin({ value: IN(s0.h), min: 0, max: 56, step: 0.01, unit: '"', onChange: (v) => { ch.h = PT(v); if (lock.input.checked && s0.h) { ch.w = s0.w * (PT(v) / s0.h); wS.set(IN(ch.w)); } } });
    const wS = ui.spin({ value: IN(s0.w), min: 0, max: 56, step: 0.01, unit: '"', onChange: (v) => { ch.w = PT(v); if (lock.input.checked && s0.w) { ch.h = s0.h * (PT(v) / s0.w); hS.set(IN(ch.h)); } } });
    const rS = ui.spin({ value: s0.rot || 0, min: -360, max: 360, step: 1, unit: '°', dec: 1, onChange: (v) => { ch.rot = ((v % 360) + 360) % 360; } });
    rS.setDisabled(!shapes.every(M.canRotate));
    const lock = ui.check('Lock &aspect ratio', !!s0.lockAspect, (v) => { ch.lockAspect = v; });
    let natW = 0, natH = 0;
    const origLbl = h('div', { class: 'tp-note' });
    if (isPic && L.media.has(s0.media)) L.loadImage(L.media.url(s0.media)).then((im) => { natW = im.naturalWidth * 0.75; natH = im.naturalHeight * 0.75; origLbl.textContent = `Original size: Height ${IN(natH)}"  Width ${IN(natW)}"`; }).catch(() => {});
    const sizeTab = h('div', { class: 'col' },
      ui.group('Size and rotate', h('div', { class: 'row' }, ui.field('H&eight:', hS), ui.field('Wi&dth:', wS)), ui.field('Rota&tion:', rS)),
      ui.group('Scale', lock),
      isPic ? h('div', null, origLbl, ui.button('&Reset', () => { if (natW) { ch.w = natW; ch.h = natH; wS.set(IN(natW)); hS.set(IN(natH)); ch.crop = { l: 0, t: 0, r: 0, b: 0 }; ch.img = {}; } })) : null);
    /* position */
    const xS = ui.spin({ value: IN(s0.x), min: -56, max: 56, step: 0.01, unit: '"', onChange: (v) => { ch.x = PT(v); } });
    const yS = ui.spin({ value: IN(s0.y), min: -56, max: 56, step: 0.01, unit: '"', onChange: (v) => { ch.y = PT(v); } });
    const posTab = h('div', { class: 'col' }, ui.group('Position on slide', ui.field('&Horizontal:', xS), ui.field('&Vertical:', yS), h('div', { class: 'tp-note', text: 'From: Top Left Corner' })));
    /* picture */
    let picTab = null;
    if (isPic) {
      const c = s0.crop || { l: 0, t: 0, r: 0, b: 0 };
      const fullW = s0.w / (1 - c.l - c.r), fullH = s0.h / (1 - c.t - c.b);
      const cr = (k, full) => ui.spin({ value: IN(c[k] * full), min: -20, max: 20, step: 0.01, unit: '"', onChange: (v) => { ch.crop = Object.assign({}, ch.crop || c); ch.crop[k] = PT(v) / full; } });
      const im = s0.img || {};
      const mode = ui.select([['', 'Automatic'], ['gray', 'Grayscale'], ['bw', 'Black & White'], ['wash', 'Washout']], im.mode || '', (v) => { ch.imgMode = v; });
      const br = h('input', { type: 'range', min: -100, max: 100, value: Math.round((im.bright || 0) * 100), id: 'fp-bright' });
      const co = h('input', { type: 'range', min: -100, max: 100, value: Math.round((im.contrast || 0) * 100), id: 'fp-contrast' });
      br.addEventListener('input', () => { ch.bright = br.value / 100; });
      co.addEventListener('input', () => { ch.contrast = co.value / 100; });
      picTab = h('div', { class: 'col' },
        ui.group('Crop from', h('div', { class: 'row' }, h('div', { class: 'col' }, ui.field('&Left:', cr('l', fullW)), ui.field('&Right:', cr('r', fullW))), h('div', { class: 'col' }, ui.field('&Top:', cr('t', fullH)), ui.field('&Bottom:', cr('b', fullH))))),
        ui.group('Image control', ui.field('&Color:', mode), ui.field('Brig&htness:', br), ui.field('Co&ntrast:', co)));
    }
    /* text box */
    let tbTab = null;
    if (s0.tx) {
      const tx = s0.tx;
      const ins = tx.ins || [7.2, 3.6, 7.2, 3.6];
      const insSp = (i) => ui.spin({ value: IN(ins[i]), min: 0, max: 20, step: 0.01, unit: '"', onChange: (v) => { ch.ins = (ch.ins || ins.slice()); ch.ins[i] = PT(v); } });
      const anchor = ui.select([['t', 'Top'], ['ctr', 'Middle'], ['b', 'Bottom'], ['t|c', 'Top Centered'], ['ctr|c', 'Middle Centered'], ['b|c', 'Bottom Centered']], tx.anchor + (tx.anchorCtr ? '|c' : ''), (v) => { ch.anchor = v; });
      tbTab = h('div', { class: 'col' },
        ui.field('Text &anchor point:', anchor, { cls: 'wide' }),
        ui.group('Internal margin', h('div', { class: 'row' }, h('div', { class: 'col' }, ui.field('&Left:', insSp(0)), ui.field('&Right:', insSp(2))), h('div', { class: 'col' }, ui.field('&Top:', insSp(1)), ui.field('&Bottom:', insSp(3))))),
        ui.check('&Word wrap text in AutoShape', tx.wrap !== false, (v) => { ch.wrap = v; }),
        ui.check('Resize AutoShape to &fit text', tx.autofit === 'shape', (v) => { ch.autofit = v ? 'shape' : 'none'; }),
        ui.check('R&otate text within AutoShape by 90°', !!tx.vert, (v) => { ch.vert = v; }));
    }
    const alt = h('textarea', { id: 'fs-alt', rows: 4, style: 'width:100%' });
    alt.value = s0.alt || '';
    alt.addEventListener('input', () => { ch.alt = alt.value; });
    const webTab = h('div', { class: 'col' }, h('label', { for: 'fs-alt', text: 'Alternative text:' }), alt, h('div', { class: 'tp-note', text: 'Screen readers and web pages use alternative text to describe the object.' }));
    const tabs = [{ label: 'Colors and Lines', body: colorsTab }, { label: 'Size', body: sizeTab }, { label: 'Position', body: posTab }];
    if (picTab) tabs.push({ label: 'Picture', body: picTab });
    if (tbTab) tabs.push({ label: 'Text Box', body: tbTab });
    tabs.push({ label: 'Web', body: webTab });
    const ti = Math.max(0, tabs.findIndex((t) => t.label === tabName));
    const tabsEl = ui.tabs(tabs, ti);
    ui.dialog({
      title, body: tabsEl, width: 470, buttons: [{
        label: 'OK', primary: true, onClick: () => {
          E().commit(title, () => {
            for (const s of shapes) {
              const tgt = s.type === 'wordart' ? s.wa : s;
              if (ch.fill !== undefined) tgt.fill = L.clone(ch.fill);
              if (ch.alpha != null && tgt.fill && tgt.fill.t === 'solid') tgt.fill.a = ch.alpha;
              if (ch.alpha != null && tgt.fill && tgt.fill.t === 'grad') tgt.fill.stops.forEach((st) => { st.a = ch.alpha; });
              if (ch.lineColor !== undefined) tgt.line = ch.lineColor && ch.lineColor.t === 'none' ? { t: 'none', w: (tgt.line && tgt.line.w) || 0.75 } : Object.assign({ w: 0.75, dash: 'solid' }, tgt.line && tgt.line.t !== 'none' ? tgt.line : {}, { c: ch.lineColor.c, a: ch.lineColor.a });
              if (tgt.line && tgt.line.t !== 'none') {
                if (ch.lineW != null) tgt.line.w = ch.lineW;
                if (ch.dash) tgt.line.dash = ch.dash;
                if (ch.cmpd) tgt.line.cmpd = ch.cmpd;
                if (ch.head) tgt.line.head = ch.head === 'none' ? undefined : { type: ch.head, w: ch.headSz || 'med', len: ch.headSz || 'med' };
                if (ch.tail) tgt.line.tail = ch.tail === 'none' ? undefined : { type: ch.tail, w: ch.tailSz || 'med', len: ch.tailSz || 'med' };
                if (!ch.head && ch.headSz && tgt.line.head) tgt.line.head.w = tgt.line.head.len = ch.headSz;
                if (!ch.tail && ch.tailSz && tgt.line.tail) tgt.line.tail.w = tgt.line.tail.len = ch.tailSz;
              } else if (ch.lineW != null && (!tgt.line || tgt.line.t === 'none') && ch.lineColor === undefined && s.type !== 'image') { tgt.line = { c: 'tx1', w: ch.lineW, dash: ch.dash || 'solid' }; }
              if (shapes.length === 1) {
                const ob = { x: s.x, y: s.y, w: s.w, h: s.h };
                if (ch.w != null) s.w = Math.max(0, ch.w);
                if (ch.h != null) s.h = Math.max(0, ch.h);
                if (ch.x != null) s.x = ch.x;
                if (ch.y != null) s.y = ch.y;
                if (s.type === 'group') { if (ch.x != null || ch.y != null) M.translate(s, 0, 0); M.scaleGroup(s, ob, s); }
                if (s.type === 'table' && (ch.w != null || ch.h != null)) E().scaleTable(s, Object.assign(L.clone(s), ob));
              } else { if (ch.w != null) s.w = ch.w; if (ch.h != null) s.h = ch.h; }
              if (ch.rot != null && M.canRotate(s)) s.rot = ch.rot;
              if (ch.lockAspect != null) s.lockAspect = ch.lockAspect;
              if (s.type === 'image') {
                if (ch.crop) {
                  const c = s.crop || { l: 0, t: 0, r: 0, b: 0 };
                  const fullW = s.w / (1 - c.l - c.r), fullH = s.h / (1 - c.t - c.b);
                  s.crop = ch.crop;
                  s.w = fullW * (1 - ch.crop.l - ch.crop.r); s.h = fullH * (1 - ch.crop.t - ch.crop.b);
                }
                s.img = Object.assign({}, s.img || {}, ch.img || {});
                if (ch.imgMode != null) { if (ch.imgMode) s.img.mode = ch.imgMode; else delete s.img.mode; }
                if (ch.bright != null) s.img.bright = ch.bright;
                if (ch.contrast != null) s.img.contrast = ch.contrast;
              }
              if (s.tx) {
                if (ch.ins) s.tx.ins = ch.ins.map((v) => L.round(v, 2));
                if (ch.anchor) { const [a, c] = ch.anchor.split('|'); s.tx.anchor = a; s.tx.anchorCtr = !!c; }
                if (ch.wrap != null) s.tx.wrap = ch.wrap;
                if (ch.autofit) s.tx.autofit = ch.autofit;
                if (ch.vert != null) { if (ch.vert) s.tx.vert = 'vert'; else delete s.tx.vert; }
              }
              if (ch.alt != null) s.alt = ch.alt;
              const f = E().find(s.id);
              if (f && f.parent) E().fitGroup(f.parent);
            }
          });
        },
      }, { label: 'Cancel' }],
    });
  };

  /* ---------- Background ---------- */
  D.background = function () {
    const slide = E().slide();
    if (!slide) return;
    const d = design();
    let fill = L.clone(slide.bg || d.bg);
    let omit = !!slide.hideMaster;
    const prev = h('div', { class: 'preview-box', style: 'width:220px;height:150px;position:relative' });
    const paint = () => {
      L.clear(prev);
      const tmp = Object.assign({}, slide, { bg: fill, hideMaster: omit });
      prev.appendChild(L.render.thumb(L.pres, tmp, 218));
    };
    const fp = fillPicker(fill, 'fill', (f) => { fill = f && f.t === 'none' ? { t: 'solid', c: 'bg1', a: 1 } : f; paint(); });
    const om = ui.check('Omit background &graphics from master', omit, (v) => { omit = v; paint(); });
    paint();
    const apply = (all) => {
      L.hist.push('Background');
      const targets = all ? L.pres.slides : (L.app.selectedSlides && L.app.selectedSlides().length > 1 ? L.app.selectedSlides() : [slide]);
      if (E().view === 'master') { const ms = E().masterSlide; ms.bg = fill; E().syncMaster(); }
      else if (all) { for (const s of L.pres.slides) { delete s.bg; s.hideMaster = omit || undefined; } for (const id in L.pres.designs) { if (L.pres.slides.some((s) => s.design === id)) L.pres.designs[id].bg = L.clone(fill); } }
      else for (const s of targets) { s.bg = L.clone(fill); s.hideMaster = omit || undefined; }
      E().render(); L.bus.emit('slides-changed');
    };
    ui.dialog({ title: 'Background', body: h('div', { class: 'row' }, h('div', { class: 'col' }, h('label', { text: 'Background fill' }), prev, fp, om)), width: 380, buttons: [{ label: 'Apply to &All', onClick: () => apply(true) }, { label: '&Apply', primary: true, onClick: () => apply(false) }, { label: 'Cancel' }] });
  };

  /* ---------- Page Setup ---------- */
  D.pageSetup = function () {
    const p = L.pres;
    let cur = M.SLIDE_SIZES.find((s) => Math.abs(s.w - p.W) < 0.6 && Math.abs(s.h - p.H) < 0.6 && s.key !== 'letter' && s.key !== 'overhead') || M.SLIDE_SIZES[M.SLIDE_SIZES.length - 1];
    const w = ui.spin({ value: IN(p.W), min: 1, max: 56, step: 0.1, unit: '"' }), hh = ui.spin({ value: IN(p.H), min: 1, max: 56, step: 0.1, unit: '"' });
    const sel = ui.select(M.SLIDE_SIZES.map((s) => [s.key, s.name]), cur.key, (k) => { const s = M.SLIDE_SIZES.find((x) => x.key === k); if (k !== 'custom') { w.set(IN(s.w)); hh.set(IN(s.h)); } });
    const first = ui.spin({ value: p.firstNum || 1, min: 0, max: 9999, step: 1, dec: 0 });
    let portrait = p.H > p.W;
    const dflt = L.app.opts.slideSize;
    const asDefault = ui.check('&Use for new presentations', !!(dflt && Math.abs(dflt.w - p.W) < 0.6 && Math.abs(dflt.h - p.H) < 0.6));
    const body = h('div', { class: 'row' },
      h('div', { class: 'col' }, ui.field('&Slides sized for:', sel, { cls: 'wide' }), ui.field('&Width:', w, { cls: 'wide' }), ui.field('H&eight:', hh, { cls: 'wide' }), ui.field('&Number slides from:', first, { cls: 'wide' }), asDefault),
      ui.group('Orientation', h('div', { class: 'col' }, h('b', { text: 'Slides' }), ui.radio('ps-or', '&Portrait', portrait, () => { portrait = true; if (w.get() > hh.get()) { const t = w.get(); w.set(hh.get()); hh.set(t); } }), ui.radio('ps-or', '&Landscape', !portrait, () => { portrait = false; if (hh.get() > w.get()) { const t = w.get(); w.set(hh.get()); hh.set(t); } }))));
    ui.dialog({ title: 'Page Setup', body, width: 480, buttons: [{ label: 'OK', primary: true, onClick: () => {
      const preset = M.SLIDE_SIZES.find((x) => x.key === sel.value);
      /* use exact preset points (inch spinners round to 0.01") */
      const W = preset && preset.key !== 'custom' && Math.abs(IN(preset.w) - w.get()) < 0.011 && Math.abs(IN(preset.h) - hh.get()) < 0.011 ? preset.w : PT(w.get());
      const H = preset && preset.key !== 'custom' && Math.abs(IN(preset.w) - w.get()) < 0.011 && Math.abs(IN(preset.h) - hh.get()) < 0.011 ? preset.h : PT(hh.get());
      const W2 = portrait ? Math.min(W, H) : Math.max(W, H), H2 = portrait ? Math.max(W, H) : Math.min(W, H);
      L.app.resizeSlides(W2, H2);
      L.pres.firstNum = first.get();
      if (asDefault.input.checked) L.app.opts.slideSize = { w: W2, h: H2 };
      else if (dflt && Math.abs(dflt.w - W2) < 0.6 && Math.abs(dflt.h - H2) < 0.6) delete L.app.opts.slideSize;
      L.app.saveOpts();
      E().render(); L.bus.emit('slides-changed');
    } }, { label: 'Cancel' }] });
  };

  /* ---------- Header and Footer ---------- */
  D.headerFooter = function () {
    const slide = E().slide();
    const hf = Object.assign({}, L.pres.hf, slide && slide.hf);
    const dt = ui.check('&Date and time', !!hf.dt, () => sync());
    const auto = ui.radio('hf-dt', 'Update &automatically', hf.dtAuto !== false, () => sync());
    const fmt = ui.select(L.DATE_FORMATS.map((f) => [f, L.fmtDate(new Date(), f)]), hf.dtFmt || 'datetime1');
    const fixed = ui.radio('hf-dt', 'Fi&xed', hf.dtAuto === false, () => sync());
    const fixedIn = h('input', { type: 'text', id: 'hf-fixed', value: hf.dtText || L.fmtDate(new Date(), 'datetime1'), style: 'width:220px' });
    const num = ui.check('Slide &number', !!hf.num);
    const ftr = ui.check('&Footer', !!hf.ftr, () => sync());
    const ftrIn = h('input', { type: 'text', id: 'hf-ftr', value: hf.ftrText || '', style: 'width:260px' });
    const nt = ui.check('Don\'t show on title &slide', !!hf.notOnTitle);
    const sync = () => { const on = dt.input.checked; auto.input.disabled = fixed.input.disabled = !on; fmt.disabled = !on || !auto.input.checked; fixedIn.disabled = !on || !fixed.input.checked; ftrIn.disabled = !ftr.input.checked; };
    sync();
    const read = () => ({ dt: dt.input.checked, dtAuto: auto.input.checked, dtFmt: fmt.value, dtText: fixedIn.value, num: num.input.checked, ftr: ftr.input.checked, ftrText: ftrIn.value, notOnTitle: nt.input.checked });
    const body = ui.tabs([{ label: 'Slide', body: h('div', { class: 'col' }, h('b', { text: 'Include on slide' }), dt, h('div', { class: 'col', style: 'margin-left:20px' }, auto, fmt, fixed, fixedIn), num, ftr, h('div', { style: 'margin-left:20px' }, ftrIn), nt) }]);
    ui.dialog({
      title: 'Header and Footer', body, width: 420, buttons: [
        { label: 'Apply to &All', primary: true, onClick: () => { L.hist.push('Header and Footer'); L.pres.hf = read(); for (const s of L.pres.slides) delete s.hf; E().render(); L.bus.emit('slides-changed'); } },
        { label: '&Apply', onClick: () => { L.hist.push('Header and Footer'); if (slide) slide.hf = read(); E().render(); L.bus.emit('slides-changed'); } },
        { label: 'Cancel' }],
    });
  };

  /* ---------- Find / Replace ---------- */
  const findState = { what: '', with: '', mc: false, ww: false, pos: null };
  function findNext(fs) {
    const slides = L.pres.slides;
    if (!slides.length || !fs.what) return false;
    const flags = fs.mc ? 'g' : 'gi';
    const esc = fs.what.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(fs.ww ? `\\b${esc}\\b` : esc, flags);
    const startSlide = E().idx;
    let after = fs.pos;
    if (L.te.active()) { const s = L.te.getSel(); const t = L.te.target(); if (s && t) after = { slide: E().idx, sid: t.sh.id, cell: L.te.state.cell, p: Math.max(s.p0, s.p1), o: Math.max(s.o1, s.o0) }; }
    for (let k = 0; k <= slides.length; k++) {
      const si = (startSlide + k) % slides.length;
      const sl = slides[si];
      const targets = [];
      M.walk(sl.shapes, (sh) => {
        if (sh.tx) targets.push({ sh, tx: sh.tx, cell: null });
        if (sh.tbl) sh.tbl.rows.forEach((r, ri) => r.cells.forEach((c, ci) => targets.push({ sh, tx: c.tx, cell: [ri, ci] })));
      });
      let passed = !after || after.slide !== si || k > 0 && k === slides.length;
      for (const t of targets) {
        if (!passed) { if (t.sh.id === after.sid && String(t.cell) === String(after.cell)) passed = 'here'; else continue; }
        for (let pi = 0; pi < t.tx.ps.length; pi++) {
          const text = L.txt.paraText(t.tx.ps[pi]);
          re.lastIndex = 0;
          let m;
          while ((m = re.exec(text))) {
            if (passed === 'here' && (pi < after.p || (pi === after.p && m.index < after.o))) { if (!m[0].length) re.lastIndex++; continue; }
            fs.pos = { slide: si, sid: t.sh.id, cell: t.cell, p: pi, o: m.index + m[0].length };
            if (L.app.view !== 'normal') L.app.setView('normal');
            if (E().idx !== si) E().goto(si);
            L.te.begin(t.sh.id, { cell: t.cell || undefined, sel: { p0: pi, o0: m.index, p1: pi, o1: m.index + m[0].length } });
            return true;
          }
        }
        passed = true;
      }
      after = null;
    }
    return false;
  }
  D.find = function (replace) {
    const what = h('input', { type: 'text', id: 'fd-what', value: (L.te.active() && L.te.selectedText()) || findState.what, style: 'width:260px' });
    const repl = h('input', { type: 'text', id: 'fd-with', value: findState.with, style: 'width:260px' });
    const mc = ui.check('Matc&h case', findState.mc), ww = ui.check('Find &whole words only', findState.ww);
    const read = () => { findState.what = what.value; findState.with = repl.value; findState.mc = mc.input.checked; findState.ww = ww.input.checked; };
    const body = h('div', { class: 'col' }, ui.field('Fi&nd what:', what, { cls: 'wide' }), replace ? ui.field('Re&place with:', repl, { cls: 'wide' }) : null, mc, ww);
    const notFound = () => ui.msg(`${L.APP} has finished searching the presentation. The search item wasn't found.`, { icon: 'info' });
    const btns = [{ label: '&Find Next', primary: true, onClick: () => { read(); if (!findNext(findState)) notFound(); return false; } }];
    if (replace) {
      btns.push({ label: '&Replace', onClick: () => { read(); doReplaceOne(); return false; } });
      btns.push({ label: 'Replace &All', onClick: () => { read(); const n = replaceAll(); ui.msg(n ? `${L.APP} has finished searching the presentation and made ${n} replacement${n === 1 ? '' : 's'}.` : 'The search item wasn\'t found.', { icon: 'info' }); return false; } });
    } else btns.push({ label: 'Re&place...', onClick: (close) => { read(); close(null); setTimeout(() => D.find(true), 0); return false; } });
    btns.push({ label: 'Close' });
    ui.dialog({ title: replace ? 'Replace' : 'Find', body, width: 440, buttons: btns });
    function doReplaceOne() {
      if (L.te.active()) {
        const sel = L.te.selectedText();
        const match = findState.mc ? sel === findState.what : sel.toLowerCase() === findState.what.toLowerCase();
        if (match && sel) { L.hist.push('Replace'); L.te.insertText(findState.with); }
      }
      if (!findNext(findState)) notFound();
    }
    function replaceAll() {
      if (!findState.what) return 0;
      if (L.te.active()) L.te.end();
      const esc = findState.what.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp(findState.ww ? `\\b${esc}\\b` : esc, findState.mc ? 'g' : 'gi');
      let n = 0;
      L.hist.push('Replace All');
      const fix = (tx) => { for (const p of tx.ps) for (const r of p.rs) { if (r.fld) continue; r.t = r.t.replace(re, () => { n++; return findState.with; }); } };
      for (const s of L.pres.slides) M.walk(s.shapes, (sh) => { if (sh.tx) fix(sh.tx); if (sh.tbl) sh.tbl.rows.forEach((r) => r.cells.forEach((c) => fix(c.tx))); if (sh.wa) sh.wa.text = sh.wa.text.replace(re, () => { n++; return findState.with; }); });
      if (!n) L.hist.undo.pop();
      E().render(); L.bus.emit('slides-changed');
      return n;
    }
  };

  /* ---------- Insert Table ---------- */
  D.insertTable = function (cb) {
    const c = ui.spin({ value: 3, min: 1, max: 75, step: 1, dec: 0 }), r = ui.spin({ value: 2, min: 1, max: 75, step: 1, dec: 0 });
    ui.dialog({ title: 'Insert Table', body: h('div', { class: 'col' }, ui.field('Number of &columns:', c, { cls: 'wide' }), ui.field('Number of &rows:', r, { cls: 'wide' })), width: 280, buttons: [{ label: 'OK', primary: true, onClick: () => cb(r.get(), c.get()) }, { label: 'Cancel' }] });
  };

  /* ---------- Hyperlink ---------- */
  D.hyperlink = function () {
    const ed = E();
    const editing = L.te.active();
    const sh = ed.primary();
    if (!sh && !editing) { ui.msg('Select text or an object to link.'); return; }
    const runProps = editing ? L.te.runPropsAtSel() : {};
    const existing = editing ? (L.te.target() && (() => { const t = L.te.target(); const s = L.te.currentSel(); const p = t.tx.ps[s.p0]; let acc = 0; for (const r of p.rs) { acc += r.t.length; if (acc >= s.o0 && r.link) return r.link; } return null; })()) : sh.link;
    void runProps;
    let kind = existing && existing.slide ? 'place' : existing && existing.url && /^mailto:/.test(existing.url) ? 'mail' : existing && existing.action ? 'place' : 'web';
    const disp = h('input', { type: 'text', id: 'hl-disp', value: editing ? L.te.selectedText() : '<<Selection in Document>>', style: 'width:100%' });
    if (!editing) disp.disabled = true;
    const addr = h('input', { type: 'text', id: 'hl-addr', value: existing && existing.url && !/^mailto:/.test(existing.url) ? existing.url : 'https://', style: 'width:100%' });
    const tip = h('input', { type: 'text', id: 'hl-tip', value: (existing && existing.tip) || '', style: 'width:100%' });
    const mail = h('input', { type: 'text', id: 'hl-mail', value: existing && existing.url && /^mailto:/.test(existing.url) ? existing.url.replace(/^mailto:/, '').split('?')[0] : '', style: 'width:100%' });
    const subj = h('input', { type: 'text', id: 'hl-subj', value: '', style: 'width:100%' });
    const places = [['a:first', 'First Slide'], ['a:last', 'Last Slide'], ['a:next', 'Next Slide'], ['a:prev', 'Previous Slide']].concat(L.pres.slides.map((s, i) => ['s:' + s.id, `${i + 1}. ${M.slideTitle(s) || 'Slide ' + (i + 1)}`]));
    const placeSel = h('select', { id: 'hl-place', size: 10, style: 'width:100%' }, ...places.map(([v, l]) => h('option', { value: v, text: l })));
    if (existing && existing.slide) placeSel.value = 's:' + existing.slide; else if (existing && existing.action) placeSel.value = 'a:' + existing.action; else placeSel.selectedIndex = 2;
    const panes = {
      web: h('div', { class: 'col' }, h('label', { for: 'hl-addr', html: 'Addr<u>e</u>ss:' }), addr, h('label', { for: 'hl-tip', text: 'ScreenTip:' }), tip),
      place: h('div', { class: 'col' }, h('label', { for: 'hl-place', html: 'Select a place in this d<u>o</u>cument:' }), placeSel),
      mail: h('div', { class: 'col' }, h('label', { for: 'hl-mail', html: '<u>E</u>-mail address:' }), mail, h('label', { for: 'hl-subj', html: 'Su<u>b</u>ject:' }), subj),
    };
    const host = h('div', { style: 'flex:1;min-width:0' });
    const side = h('div', { class: 'list-box', style: 'width:130px;min-height:180px' });
    const show = (k) => { kind = k; L.clear(host).appendChild(panes[k]); L.$$('.li', side).forEach((x) => x.classList.toggle('on', x.dataset.k === k)); };
    [['web', 'Existing File or Web Page'], ['place', 'Place in This Document'], ['mail', 'E-mail Address']].forEach(([k, l]) => { const it = h('div', { class: 'li', 'data-k': k, text: l, style: 'white-space:normal;padding:8px 6px' }); it.addEventListener('click', () => show(k)); side.appendChild(it); });
    show(kind);
    const body = h('div', { class: 'col' }, h('div', { class: 'field' }, h('label', { for: 'hl-disp', html: '<u>T</u>ext to display:' }), disp), h('div', { class: 'row', style: 'flex-wrap:nowrap' }, h('div', { class: 'col' }, h('label', { text: 'Link to:' }), side), host));
    const apply = (link) => {
      if (editing) {
        const t = L.te.target();
        const s = L.te.currentSel();
        if (s && L.txt.collapsed(s) && disp.value) { L.hist.push('Hyperlink'); L.te.insertText(disp.value); const s2 = L.te.currentSel(); L.te.applyRun((r) => { const n = Object.assign({}, r); if (link) n.link = link; else delete n.link; return n; }, 'Hyperlink', { p0: s2.p0, o0: s2.o0 - disp.value.length, p1: s2.p1, o1: s2.o1 }); }
        else L.te.applyRun((r) => { const n = Object.assign({}, r); if (link) n.link = link; else delete n.link; return n; }, 'Hyperlink');
        void t;
      } else ed.commit('Hyperlink', () => { for (const x of ed.selected()) { if (link) x.link = link; else delete x.link; } });
    };
    const buttons = [{
      label: 'OK', primary: true, onClick: () => {
        let link = null;
        if (kind === 'web') { const u = addr.value.trim(); if (u && u !== 'https://') link = { url: /^[a-z]+:/i.test(u) ? u : 'https://' + u }; if (link && tip.value) link.tip = tip.value; }
        if (kind === 'mail' && mail.value.trim()) link = { url: 'mailto:' + mail.value.trim() + (subj.value ? '?subject=' + encodeURIComponent(subj.value) : '') };
        if (kind === 'place') { const v = placeSel.value || 'a:next'; link = v.startsWith('s:') ? { slide: v.slice(2) } : { action: v.slice(2) }; }
        if (link) apply(link);
      },
    }, { label: 'Cancel' }];
    if (existing) buttons.unshift({ label: '&Remove Link', onClick: () => apply(null) });
    ui.dialog({ title: existing ? 'Edit Hyperlink' : 'Insert Hyperlink', body, width: 560, buttons });
  };

  /* ---------- Action Settings ---------- */
  D.actionSettings = function () {
    const shapes = E().selected();
    if (!shapes.length) { ui.msg('Select an object first.'); return; }
    const cur = shapes[0].link;
    let none = !cur;
    const opts = [['next', 'Next Slide'], ['prev', 'Previous Slide'], ['first', 'First Slide'], ['last', 'Last Slide'], ['lastViewed', 'Last Slide Viewed'], ['end', 'End Show'], ['slide', 'Slide...'], ['url', 'URL...']];
    const sel = ui.select(opts, cur ? (cur.slide ? 'slide' : cur.url ? 'url' : cur.action) : 'next', () => sync());
    const slideSel = ui.select(L.pres.slides.map((s, i) => [s.id, `${i + 1}. ${M.slideTitle(s) || 'Slide ' + (i + 1)}`]), cur && cur.slide);
    const url = h('input', { type: 'text', id: 'as-url', value: (cur && cur.url) || 'https://', style: 'width:240px' });
    const rn = ui.radio('as-k', '&None', none, () => { none = true; sync(); }), rl = ui.radio('as-k', '&Hyperlink to:', !none, () => { none = false; sync(); });
    const sync = () => { sel.disabled = none; slideSel.hidden = none || sel.value !== 'slide'; url.hidden = none || sel.value !== 'url'; };
    sync();
    const body = ui.tabs([{ label: 'Mouse Click', body: h('div', { class: 'col' }, h('b', { text: 'Action on click' }), rn, rl, h('div', { class: 'col', style: 'margin-left:20px' }, sel, slideSel, url)) }]);
    ui.dialog({
      title: 'Action Settings', body, width: 360, buttons: [{
        label: 'OK', primary: true, onClick: () => {
          let link = null;
          if (!none) link = sel.value === 'slide' ? { slide: slideSel.value } : sel.value === 'url' ? { url: url.value } : { action: sel.value };
          E().commit('Action Settings', () => { for (const s of shapes) { if (link) s.link = link; else delete s.link; } });
        },
      }, { label: 'Cancel' }],
    });
  };

  /* ---------- Symbol ---------- */
  const SYM_SETS = [
    ['Latin-1 Supplement', 0x00a1, 0x00ff], ['Latin Extended-A', 0x0100, 0x017f], ['Greek and Coptic', 0x0391, 0x03c9], ['General Punctuation', 0x2010, 0x2044],
    ['Currency Symbols', 0x20a0, 0x20bf], ['Letterlike Symbols', 0x2100, 0x214f], ['Number Forms', 0x2150, 0x218b], ['Arrows', 0x2190, 0x21ff], ['Mathematical Operators', 0x2200, 0x22ff],
    ['Box Drawing', 0x2500, 0x257f], ['Geometric Shapes', 0x25a0, 0x25ff], ['Miscellaneous Symbols', 0x2600, 0x26ff], ['Dingbats', 0x2700, 0x27bf],
  ];
  const recentSyms = ['€', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β'];
  D.symbol = function (cb) {
    let pick = null;
    const grid = h('div', { class: 'sym-grid' });
    const font = ui.select(['(normal text)', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI Symbol'], '(normal text)', () => fill(sub.value));
    const sub = ui.select(SYM_SETS.map((s, i) => [i, s[0]]), 0, (v) => fill(v));
    const code = h('input', { type: 'text', id: 'sym-code', style: 'width:70px', readonly: true });
    const recent = h('div', { class: 'row', style: 'gap:2px' });
    const fill = (si) => {
      const [, a, b] = SYM_SETS[+si];
      L.clear(grid);
      grid.style.fontFamily = font.value === '(normal text)' ? '' : L.fontStack(font.value);
      for (let c = a; c <= b; c++) {
        const ch = String.fromCodePoint(c);
        const btn = h('button', { type: 'button', text: ch, 'aria-label': 'U+' + c.toString(16).toUpperCase() });
        btn.addEventListener('click', () => { L.$$('button.on', grid).forEach((x) => x.classList.remove('on')); btn.classList.add('on'); pick = ch; code.value = c.toString(16).toUpperCase().padStart(4, '0'); });
        btn.addEventListener('dblclick', () => { pick = ch; insertBtn.click(); });
        grid.appendChild(btn);
      }
    };
    fill(0);
    for (const ch of recentSyms) { const b = h('button', { type: 'button', class: 'btn small', text: ch, style: 'min-width:24px;padding:0' }); b.addEventListener('click', () => { pick = ch; code.value = ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); }); recent.appendChild(b); }
    const body = h('div', { class: 'col' }, h('div', { class: 'row' }, ui.field('&Font:', font), ui.field('Su&bset:', sub)), grid, h('label', { text: 'Recently used symbols:' }), recent, ui.field('&Character code:', code));
    const d = ui.dialog({
      title: 'Symbol', body, width: 470, buttons: [{
        label: '&Insert', primary: true, onClick: () => {
          if (!pick) return false;
          const i = recentSyms.indexOf(pick); if (i >= 0) recentSyms.splice(i, 1); recentSyms.unshift(pick); recentSyms.length = 16;
          if (cb) { cb(pick); return; }
          if (L.te.active()) { L.hist.push('Insert Symbol'); L.te.insertText(pick); }
        },
      }, { label: 'Cancel' }],
    });
    const insertBtn = d.buttons[0];
  };

  /* ---------- Date and Time ---------- */
  D.dateTime = function () {
    const list = h('select', { id: 'dt-list', size: 12, style: 'width:260px' }, ...L.DATE_FORMATS.map((f) => h('option', { value: f, text: L.fmtDate(new Date(), f) })));
    list.selectedIndex = 0;
    const upd = ui.check('&Update automatically', false);
    ui.dialog({
      title: 'Date and Time', body: h('div', { class: 'col' }, h('label', { for: 'dt-list', html: '<u>A</u>vailable formats:' }), list, upd), width: 300, buttons: [{
        label: 'OK', primary: true, onClick: () => {
          const f = list.value || 'datetime1';
          if (!L.te.active()) { L.app.ensureTextTarget(); }
          if (!L.te.active()) return;
          L.hist.push('Insert Date and Time');
          if (upd.input.checked) L.te.insertField(f); else L.te.insertText(L.fmtDate(new Date(), f));
        },
      }, { label: 'Cancel' }],
    });
  };

  /* ---------- Set Up Show ---------- */
  D.setupShow = function () {
    const s = L.pres.show;
    const n = L.pres.slides.length;
    const speaker = ui.radio('su-type', '&Presented by a speaker (full screen)', !s.kiosk), kiosk = ui.radio('su-type', 'Browsed at a &kiosk (full screen)', !!s.kiosk);
    const loop = ui.check('&Loop continuously until \'Esc\'', !!s.loop), noAnim = ui.check('Show without &animation', !!s.noAnim);
    const pen = colorPick(s.penColor || '#FF0000');
    const all = ui.radio('su-rng', '&All', !(s.from && s.to)), rng = ui.radio('su-rng', '&From:', !!(s.from && s.to));
    const from = ui.spin({ value: s.from || 1, min: 1, max: Math.max(1, n), step: 1, dec: 0 }), to = ui.spin({ value: s.to || n, min: 1, max: Math.max(1, n), step: 1, dec: 0 });
    const man = ui.radio('su-adv', '&Manually', s.useTimings === false), tim = ui.radio('su-adv', '&Using timings, if present', s.useTimings !== false);
    const body = h('div', { class: 'row' },
      h('div', { class: 'col' }, ui.group('Show type', speaker, kiosk), ui.group('Show options', loop, noAnim, ui.field('P&en color:', pen))),
      h('div', { class: 'col' }, ui.group('Show slides', all, h('div', { class: 'row', style: 'align-items:center' }, rng, from, h('span', { text: 'To:' }), to)), ui.group('Advance slides', man, tim)));
    ui.dialog({ title: 'Set Up Show', body, width: 560, buttons: [{ label: 'OK', primary: true, onClick: () => { L.hist.push('Set Up Show'); Object.assign(s, { kiosk: kiosk.input.checked, loop: loop.input.checked || kiosk.input.checked, noAnim: noAnim.input.checked, penColor: pen.get() && pen.get()[0] === '#' ? pen.get() : L.model.resolveColor(pen.get(), design()), from: rng.input.checked ? from.get() : 0, to: rng.input.checked ? to.get() : 0, useTimings: tim.input.checked }); } }, { label: 'Cancel' }] });
  };

  /* ---------- Options ---------- */
  D.options = function () {
    const o = L.app.opts;
    const c = (label, key) => ui.check(label, !!o[key], (v) => { o[key] = v; });
    const undo = ui.spin({ value: L.hist.limit, min: 3, max: 150, step: 1, dec: 0 });
    const user = h('input', { type: 'text', id: 'op-user', value: L.pres.props.author || '', style: 'width:220px' });
    const pwOpen = h('input', { type: 'password', id: 'op-pw', value: L.pres.password || '', autocomplete: 'new-password', style: 'width:180px' });
    const look = ui.lookField();
    const tabs = ui.tabs([
      { label: 'View', body: h('div', { class: 'col' }, ui.group('Show', c('Startup &Task Pane', 'startupPane'), c('&Slide Layout task pane when inserting new slides', 'layoutPaneOnNew'), c('Stat&us bar', 'statusBar'), c('&Vertical ruler', 'ruler')), ui.group('Slide show', c('&Popup menu on right mouse click', 'showPopup'), c('Show popup menu &button', 'showPopbar'), c('&End with black slide', 'endBlack'))) },
      { label: 'General', body: h('div', { class: 'col' }, look && look.el, ui.group('User information', ui.field('&Name:', user)), c('Recently used file &list', 'recent')) },
      { label: 'Edit', body: h('div', { class: 'col' }, ui.group('Text', c('Replace straight &quotes with smart quotes', 'smartQuotes'), c('&AutoCorrect as you type', 'autocorrect'), c('Capitalize first letter of &sentences', 'capSentence')), ui.group('Undo', ui.field('&Maximum number of undos:', undo, { cls: 'wide' }))) },
      { label: 'Spelling and Style', body: h('div', { class: 'col' }, ui.group('Spelling', c('Check spelling as you &type', 'spell')), h('div', { class: 'tp-note', text: 'Spelling suggestions come from your browser. Hold Shift while right-clicking a word to see them.' })) },
      { label: 'Security', body: h('div', { class: 'col' }, ui.group('File encryption settings for this document', ui.field('Password to &open:', pwOpen, { cls: 'wide' })), h('div', { class: 'tp-note', text: 'The presentation is encrypted (AES-256) the next time you save it as a PowerPoint file, the way PowerPoint 2007 and later do. Leave the box empty to save it without a password. If you lose or forget the password, it cannot be recovered.' })) },
    ]);
    const applyPw = () => {
      const pw = pwOpen.value || '';
      if (pw === (L.pres.password || '')) return true;
      if (!pw) { delete L.pres.password; L.hist.dirty = true; return true; }
      /* confirm a new password, as PowerPoint does */
      const p2 = h('input', { type: 'password', id: 'pw-confirm', autocomplete: 'new-password', style: 'width:100%' });
      ui.dialog({ title: 'Confirm Password', width: 320, body: h('div', { class: 'col' }, h('label', { for: 'pw-confirm', text: 'Reenter password to open:' }), p2, h('div', { class: 'tp-note', text: 'Caution: If you lose or forget the password, it cannot be recovered.' })), buttons: [{ label: 'OK', primary: true, onClick: () => { if (p2.value !== pw) { ui.msg('Confirmation password is not identical.', { icon: 'warn' }); return false; } L.pres.password = pw; L.hist.dirty = true; } }, { label: 'Cancel' }] });
      setTimeout(() => p2.focus(), 0);
      return true;
    };
    ui.dialog({ title: 'Options', body: tabs, width: 460, buttons: [{ label: 'OK', primary: true, onClick: () => { L.hist.limit = undo.get(); L.pres.props.author = user.value; L.app.saveOpts(); L.app.applyOpts(); if (look) look.apply(); applyPw(); } }, { label: 'Cancel', onClick: () => { L.app.loadOpts(); } }] });
  };

  /* ---------- AutoCorrect ---------- */
  D.autocorrect = function () {
    const o = L.app.opts;
    const AC = L.te.AUTOCORRECT;
    const list = h('div', { class: 'list-box', style: 'height:150px' });
    const rep = h('input', { type: 'text', id: 'ac-rep', style: 'width:120px' }), wit = h('input', { type: 'text', id: 'ac-with', style: 'width:200px' });
    const draw = () => { L.clear(list); Object.keys(AC).sort().forEach((k) => { const li = h('div', { class: 'li' }, h('span', { text: k, style: 'width:110px;display:inline-block' }), h('span', { text: AC[k] })); li.addEventListener('click', () => { L.$$('.li', list).forEach((x) => x.classList.remove('on')); li.classList.add('on'); rep.value = k; wit.value = AC[k]; }); list.appendChild(li); }); };
    draw();
    const body = ui.tabs([{
      label: 'AutoCorrect', body: h('div', { class: 'col' },
        ui.check('Correct TWo INitial CApitals', !!o.autocorrect, (v) => { o.autocorrect = v; }),
        ui.check('Capitalize first letter of &sentences', !!o.capSentence, (v) => { o.capSentence = v; }),
        ui.check('Replace straight quotes with smart quotes', !!o.smartQuotes, (v) => { o.smartQuotes = v; }),
        h('div', { class: 'row' }, ui.field('&Replace:', rep), ui.field('&With:', wit)), list,
        h('div', { class: 'row' }, ui.button('&Add', () => { if (rep.value) { AC[rep.value] = wit.value; draw(); } }), ui.button('&Delete', () => { delete AC[rep.value]; draw(); }))),
    }]);
    ui.dialog({ title: 'AutoCorrect: English (U.S.)', body, width: 420, buttons: [{ label: 'OK', primary: true, onClick: () => L.app.saveOpts() }, { label: 'Cancel' }] });
  };

  /* ---------- Properties ---------- */
  D.properties = function () {
    const p = L.pres.props;
    const f = (k, label) => { const i = h('input', { type: 'text', id: 'pr-' + k, value: p[k] || '', style: 'width:260px' }); i.addEventListener('input', () => { p[k] = i.value; if (k === 'title') L.pres.title = i.value; }); return ui.field(label, i, { cls: 'wide' }); };
    const cm = h('textarea', { id: 'pr-cm', rows: 3, style: 'width:260px' }); cm.value = p.comments || ''; cm.addEventListener('input', () => { p.comments = cm.value; });
    const words = L.pres.slides.reduce((a, s) => a + (M.allText(s).match(/\S+/g) || []).length, 0);
    const tabs = ui.tabs([
      { label: 'Summary', body: h('div', { class: 'col' }, f('title', '&Title:'), f('subject', '&Subject:'), f('author', '&Author:'), f('company', 'C&ompany:'), f('category', 'Ca&tegory:'), f('keywords', '&Keywords:'), ui.field('&Comments:', cm, { cls: 'wide' })) },
      { label: 'Statistics', body: h('table', { class: 'kbd-table' }, ...[['Created:', p.created ? new Date(p.created).toLocaleString() : '—'], ['Slides:', L.pres.slides.length], ['Hidden slides:', L.pres.slides.filter((s) => s.hidden).length], ['Notes:', L.pres.slides.filter((s) => s.notes).length], ['Words:', words], ['Pictures:', L.media.all().size]].map(([a, b]) => h('tr', null, h('td', { text: a }), h('td', { text: String(b) })))) },
    ]);
    ui.dialog({ title: (L.app.fileName || 'Presentation') + ' Properties', body: tabs, width: 430, buttons: [{ label: 'OK', primary: true }] });
  };

  /* ---------- Save As ---------- */
  D.saveAs = function (cb) {
    const base = (L.app.fileName || 'Presentation1').replace(/\.[^.]+$/, '');
    const name = h('input', { type: 'text', id: 'sa-name', value: base, style: 'width:300px' });
    const types = [['pptx', 'PowerPoint Presentation (*.pptx)'], ['pptm', 'Macro-Enabled Presentation (*.pptm)'], ['ppsx', 'PowerPoint Show (*.ppsx)'], ['ppsm', 'Macro-Enabled Show (*.ppsm)'], ['potx', 'Design Template (*.potx)'], ['potm', 'Macro-Enabled Template (*.potm)'], ['pdf', 'PDF, one slide per page (*.pdf)'], ['html', 'Single File Web Page (*.html)'], ['png', 'PNG Graphics Format, current slide (*.png)'], ['txt', 'Outline (*.txt)']];
    const cur = types.some(([k]) => k === L.app.fileType) ? L.app.fileType : 'pptx';
    const type = ui.select(types, cur);
    return ui.dialog({ title: 'Save As', body: h('div', { class: 'col' }, ui.field('File &name:', name, { cls: 'wide' }), ui.field('Save as &type:', type, { cls: 'wide' }), h('div', { class: 'tp-note', text: 'Your browser decides where downloaded files go.' })), width: 470, buttons: [{ label: '&Save', primary: true, onClick: () => cb((name.value.trim() || base).replace(/\.(pptx|pptm|ppsx|ppsm|potx|potm)$/i, ''), type.value) }, { label: 'Cancel' }] });
  };

  /* ---------- Zoom ---------- */
  D.zoom = function () {
    const presets = [['fit', 'Fit'], [4, '400%'], [2, '200%'], [1, '100%'], [0.66, '66%'], [0.5, '50%'], [0.33, '33%']];
    let pick = E().zoom;
    const pct = ui.spin({ value: E().zoomPct(), min: 10, max: 400, step: 1, unit: '%', dec: 0, onChange: (v) => { pick = v / 100; } });
    const body = h('div', { class: 'row' }, ui.group('Zoom to', h('div', { class: 'col' }, ...presets.map(([v, l]) => ui.radio('zm', l, String(v) === String(E().zoom), () => { pick = v; if (v !== 'fit') pct.set(Math.round(v * 100)); })))), h('div', { class: 'col' }, ui.field('&Percent:', pct)));
    ui.dialog({ title: 'Zoom', body, width: 300, buttons: [{ label: 'OK', primary: true, onClick: () => E().setZoom(pick) }, { label: 'Cancel' }] });
  };

  /* ---------- Grid and Guides ---------- */
  D.grid = function () {
    const g = E().grid;
    const snap = ui.check('&Snap objects to grid', g.snap);
    const sp = ui.select([[4.5, '1/16"'], [6, '0.083"'], [7.2, '0.1"'], [9, '0.125"'], [12, '0.167"'], [14.4, '0.2"'], [18, '0.25"'], [24, '0.333"'], [36, '0.5"'], [72, '1"']], g.size);
    const disp = ui.check('&Display grid on screen', g.show), gd = ui.check('Display drawing &guides on screen', g.guides);
    ui.dialog({ title: 'Grid and Guides', body: h('div', { class: 'col' }, ui.group('Snap to', snap), ui.group('Grid settings', ui.field('Spac&ing:', sp), disp), ui.group('Guide settings', gd)), width: 320, buttons: [{ label: 'OK', primary: true, onClick: () => { g.snap = snap.input.checked; g.size = +sp.value; g.show = disp.input.checked; g.guides = gd.input.checked; E().layout(); L.ui.refresh(); L.app.saveOpts(); } }, { label: 'Cancel' }] });
  };

  /* ---------- WordArt ---------- */
  const WA_STYLES = (() => {
    const s = [];
    const grad = (a, b, ang) => ({ t: 'grad', stops: [{ p: 0, c: a, a: 1 }, { p: 1, c: b, a: 1 }], ang: ang == null ? 90 : ang, path: 'lin' });
    const solid = (c) => ({ t: 'solid', c, a: 1 });
    const ln = (c, w) => ({ c, w: w || 0.75, dash: 'solid' });
    const shd = { c: '#000000', a: 0.35, dx: 3, dy: 3, blur: 0 };
    const rows = [
      [solid('#000000'), null, null, 'textPlain'], [solid('#FFFFFF'), ln('#000000'), null, 'textPlain'], [grad('#3366FF', '#99CCFF'), null, null, 'textPlain'], [solid('#C0C0C0'), ln('#404040'), shd, 'textPlain'], [grad('#FF6600', '#FFCC00'), ln('#993300'), null, 'textArchUp'], [solid('#3366CC'), null, shd, 'textWave1'],
      [grad('#FFFFFF', '#808080', 90), ln('#000000'), shd, 'textPlain'], [solid('#CC0000'), ln('#660000'), null, 'textSlantUp'], [grad('#00CC99', '#006666'), null, shd, 'textPlain'], [solid('#FFFF00'), ln('#FF6600', 1.5), shd, 'textPlain'], [grad('#CC99FF', '#660099'), null, null, 'textArchDown'], [solid('#006600'), ln('#003300'), null, 'textTriangle'],
      [grad('#FFCC00', '#CC6600', 0), ln('#663300'), shd, 'textPlain'], [solid('#000080'), null, shd, 'textChevron'], [grad('#E6E6E6', '#333333'), ln('#000000'), null, 'textPlain'], [solid('#FF99CC'), ln('#CC0066'), shd, 'textWave1'], [grad('#66CCFF', '#003399', 0), null, shd, 'textCanUp'], [solid('#996633'), ln('#663300'), null, 'textSlantDown'],
      [grad('#FFFF99', '#FF9900'), ln('#996600'), shd, 'textArchUp'], [solid('#00CCFF'), ln('#006699'), null, 'textPlain'], [grad('#FF0000', '#660000'), null, shd, 'textPlain'], [solid('#808080'), null, { c: '#000000', a: 0.6, dx: 4, dy: 4, blur: 2 }, 'textPlain'], [grad('#99FF99', '#009900'), ln('#006600'), null, 'textCircle'], [solid('#FF6600'), null, shd, 'textTriangle'],
      [grad('#000000', '#808080'), null, shd, 'textPlain'], [solid('#FFFFFF'), ln('#3366CC', 1.5), shd, 'textArchUp'], [grad('#CC66FF', '#3399FF', 0), null, null, 'textWave1'], [solid('#003366'), ln('#99CCFF'), shd, 'textSlantUp'], [grad('#FFCCCC', '#FF3366'), ln('#990033'), null, 'textChevron'], [solid('#669900'), null, shd, 'textCanUp'],
    ];
    rows.forEach(([fill, line, shadow, warp], i) => s.push({ fill, line: line || { t: 'none' }, shadow, warp, b: i % 3 !== 1, font: ['Arial Black', 'Impact', 'Arial', 'Times New Roman', 'Georgia', 'Trebuchet MS'][i % 6] }));
    return s;
  })();
  D.WA_STYLES = WA_STYLES;
  D.wordartGallery = function (cb, current) {
    let pick = current != null ? current : 0;
    const grid = h('div', { class: 'wa-grid' });
    WA_STYLES.forEach((st, i) => {
      const b = h('button', { type: 'button', 'aria-label': 'LettersArt style ' + (i + 1), class: i === pick ? 'on' : '' });
      const sh = { w: 66, h: 46, wa: Object.assign({}, st, { text: 'LettersArt', vert: false }) };
      b.appendChild(L.render.wordart(sh, design(), {}));
      b.addEventListener('click', () => { L.$$('button', grid).forEach((x) => x.classList.remove('on')); b.classList.add('on'); pick = i; });
      b.addEventListener('dblclick', () => { okb.click(); });
      grid.appendChild(b);
    });
    const d = ui.dialog({ title: 'LettersArt Gallery', body: h('div', { class: 'col' }, h('label', { text: 'Select a LettersArt style:' }), grid), width: 470, buttons: [{ label: 'OK', primary: true, onClick: () => cb(pick) }, { label: 'Cancel' }] });
    const okb = d.buttons[0];
  };
  D.wordartText = function (init, cb) {
    const font = ui.select(L.FONT_LIST, init.font || 'Arial Black');
    const size = ui.select(L.SIZE_LIST, 36);
    const b = ui.check('&B', !!init.b), i = ui.check('&I', !!init.i);
    const ta = h('textarea', { id: 'wa-text', rows: 5, style: 'width:100%;font-size:18px' });
    ta.value = init.text || 'Your Text Here';
    const sync = () => { ta.style.fontFamily = L.fontStack(font.value); ta.style.fontWeight = b.input.checked ? '700' : '400'; ta.style.fontStyle = i.input.checked ? 'italic' : 'normal'; };
    font.addEventListener('change', sync); b.input.addEventListener('change', sync); i.input.addEventListener('change', sync);
    sync();
    ui.dialog({ title: 'Edit LettersArt Text', body: h('div', { class: 'col' }, h('div', { class: 'row' }, ui.field('&Font:', font), ui.field('&Size:', size), b, i), h('label', { for: 'wa-text', html: '<u>T</u>ext:' }), ta), width: 470, focus: '#wa-text', buttons: [{ label: 'OK', primary: true, onClick: () => cb({ text: ta.value, font: font.value, b: b.input.checked, i: i.input.checked, size: +size.value }) }, { label: 'Cancel' }] });
  };

  /* ---------- Photo Album ---------- */
  D.photoAlbum = function () {
    const pics = [];
    const list = h('div', { class: 'list-box', style: 'height:150px;width:200px' });
    const prev = h('div', { class: 'preview-box', style: 'width:180px;height:135px' });
    let sel = -1;
    const draw = () => {
      L.clear(list);
      pics.forEach((p, i) => { const li = h('div', { class: 'li' + (i === sel ? ' on' : ''), text: `${i + 1}  ${p.name}` }); li.addEventListener('click', () => { sel = i; draw(); prev.style.background = `url(${p.url}) center/contain no-repeat`; }); list.appendChild(li); });
    };
    const layout = ui.select([['fit', 'Fit to slide'], ['1', '1 picture'], ['2', '2 pictures'], ['4', '4 pictures'], ['1t', '1 picture with title'], ['2t', '2 pictures with title'], ['4t', '4 pictures with title']], '1');
    const frame = ui.select([['rect', 'Rectangle'], ['roundRect', 'Rounded Rectangle'], ['simple', 'Simple Frame, White'], ['black', 'Simple Frame, Black'], ['shadow', 'Soft Shadow']], 'rect');
    const caps = ui.check('Captions &below ALL pictures', false);
    const add = ui.button('&File/Disk...', async () => { const fs = await L.pickFiles('image/*', true); for (const f of fs) pics.push({ file: f, name: f.name, url: URL.createObjectURL(f) }); draw(); });
    const rm = ui.button('Re&move', () => { if (sel >= 0) { pics.splice(sel, 1); sel = -1; draw(); prev.style.background = ''; } });
    const up = ui.button('↑', () => { if (sel > 0) { [pics[sel - 1], pics[sel]] = [pics[sel], pics[sel - 1]]; sel--; draw(); } });
    const dn = ui.button('↓', () => { if (sel >= 0 && sel < pics.length - 1) { [pics[sel + 1], pics[sel]] = [pics[sel], pics[sel + 1]]; sel++; draw(); } });
    const body = h('div', { class: 'col' },
      ui.group('Album Content', h('div', { class: 'row' }, h('div', { class: 'col' }, h('label', { text: 'Insert picture from:' }), add), h('div', { class: 'col' }, h('label', { text: 'Pictures in album:' }), list, h('div', { class: 'row' }, up, dn, rm)), h('div', { class: 'col' }, h('label', { text: 'Preview:' }), prev))),
      ui.group('Album Layout', ui.field('&Picture layout:', layout, { cls: 'wide' }), ui.field('Frame s&hape:', frame, { cls: 'wide' }), caps));
    ui.dialog({ title: 'Photo Album', body, width: 600, buttons: [{ label: '&Create', primary: true, onClick: () => { if (!pics.length) { ui.msg('Add at least one picture.'); return false; } L.app.createPhotoAlbum(pics, layout.value, frame.value, caps.input.checked); } }, { label: 'Cancel' }] });
  };

  /* ---------- Slide Finder (Insert ▸ Slides from Files) ---------- */
  D.slideFinder = async function () {
    const files = await L.pickFiles('.pptx,.ppsx,.potx,.pptm');
    if (!files[0]) return;
    ui.busy(true, 'Opening ' + files[0].name + '...');
    let src;
    try { src = await L.pptx.read(await L.readAsArrayBuffer(files[0])); } catch (e) { ui.busy(false); ui.msg(`${files[0].name} could not be opened: ${e.message}`, { icon: 'error' }); return; }
    ui.busy(false);
    const picked = new Set();
    const grid = h('div', { style: 'display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:10px;max-height:320px;overflow:auto;padding:4px;background:#fff;border:1px solid #888' });
    src.slides.forEach((s, i) => {
      const cell = h('label', { style: 'display:flex;flex-direction:column;gap:3px;align-items:flex-start;cursor:default' });
      const ck = h('input', { type: 'checkbox' });
      ck.addEventListener('change', () => { if (ck.checked) picked.add(i); else picked.delete(i); });
      const th = L.render.thumb(src, s, 124);
      th.style.border = '1px solid #999';
      cell.append(th, h('span', { style: 'display:flex;gap:4px;align-items:center' }, ck, h('span', { text: `${i + 1}. ${(M.slideTitle(s) || 'Slide').slice(0, 24)}` })));
      grid.appendChild(cell);
    });
    const keep = ui.check('&Keep source formatting', false);
    const doInsert = (all) => {
      const idxs = all ? src.slides.map((_, i) => i) : Array.from(picked).sort((a, b) => a - b);
      if (!idxs.length) { ui.msg('Select at least one slide.'); return false; }
      L.app.insertSlidesFrom(src, idxs, keep.input.checked);
    };
    ui.dialog({ title: 'Slide Finder', body: h('div', { class: 'col' }, h('div', { text: 'Presentation: ' + files[0].name }), grid, keep), width: 640, buttons: [{ label: '&Insert', primary: true, onClick: () => doInsert(false) }, { label: 'Insert &All', onClick: () => doInsert(true) }, { label: 'Close' }] });
  };

  /* ---------- Custom Animation effect options ---------- */
  D.effectOptions = function (a, onDone) {
    const info = L.anim.info(a.cls, a.eff) || {};
    const dir = info.dirs ? ui.select(info.dirs.map((d) => [d[0], d[1]]), a.dir || info.dirs[0][0]) : null;
    const amt = a.cls === 'emph' && (a.eff === 'growShrink' || a.eff === 'spin') ? ui.spin({ value: a.amount || (a.eff === 'spin' ? 360 : 150), min: 10, max: 1440, step: 10, unit: a.eff === 'spin' ? '°' : '%', dec: 0 }) : null;
    const start = ui.select([['click', 'On Click'], ['with', 'With Previous'], ['after', 'After Previous']], a.start);
    const delay = ui.spin({ value: (a.delay || 0) / 1000, min: 0, max: 60, step: 0.5, unit: ' seconds', dec: 1 });
    const speed = ui.select(L.anim.SPEEDS.map(([n, ms]) => [ms, `${n} (${ms / 1000} seconds)`]), L.anim.SPEEDS.reduce((best, s) => (Math.abs(s[1] - (a.dur || 500)) < Math.abs(best - (a.dur || 500)) ? s[1] : best), 500));
    const repeat = ui.select([[1, '(none)'], [2, '2'], [3, '3'], [4, '4'], [5, '5'], [10, '10']], a.repeat || 1);
    const sh = M.shapeById(E().slide(), a.sid);
    const grp = ui.select([['all', 'As One Object'], ['para', 'By 1st Level Paragraphs']], a.by || 'all');
    if (!sh || !sh.tx || sh.tx.ps.length < 2) grp.disabled = true;
    const tabs = ui.tabs([
      { label: 'Effect', body: h('div', { class: 'col' }, ui.group('Settings', dir ? ui.field('&Direction:', dir) : h('div', { class: 'tp-note', text: 'This effect has no direction settings.' }), amt ? ui.field(a.eff === 'spin' ? '&Amount:' : '&Size:', amt) : null)) },
      { label: 'Timing', body: h('div', { class: 'col' }, ui.field('&Start:', start), ui.field('&Delay:', delay), ui.field('Sp&eed:', speed), ui.field('&Repeat:', repeat)) },
      { label: 'Text Animation', body: h('div', { class: 'col' }, ui.field('&Group text:', grp, { cls: 'wide' })) },
    ]);
    ui.dialog({ title: L.anim.name(a), body: tabs, width: 380, buttons: [{ label: 'OK', primary: true, onClick: () => onDone({ dir: dir ? dir.value : a.dir, amount: amt ? amt.get() : a.amount, start: start.value, delay: delay.get() * 1000, dur: +speed.value, repeat: +repeat.value, by: grp.value }) }, { label: 'Cancel' }] });
  };

  /* ---------- Print ---------- */
  D.print = function () {
    const n = L.pres.slides.length;
    const what = ui.select([['slides', 'Slides'], ['handouts', 'Handouts'], ['notes', 'Notes Pages'], ['outline', 'Outline View']], L.app.printOpts.what);
    const per = ui.select([[1, '1'], [2, '2'], [3, '3'], [4, '4'], [6, '6'], [9, '9']], L.app.printOpts.per);
    const color = ui.select([['color', 'Color'], ['gray', 'Grayscale'], ['bw', 'Pure Black and White']], L.app.printOpts.color);
    const frame = ui.check('Fra&me slides', L.app.printOpts.frame);
    const hid = ui.check('Print &hidden slides', L.app.printOpts.hidden);
    const all = ui.radio('pr-r', '&All', true), cur = ui.radio('pr-r', 'Curr&ent slide', false), rng = ui.radio('pr-r', '&Slides:', false);
    const rIn = h('input', { type: 'text', id: 'pr-range', value: `1-${n}`, style: 'width:120px' });
    const body = h('div', { class: 'col' }, ui.group('Print range', all, cur, h('div', { class: 'row', style: 'align-items:center' }, rng, rIn)), h('div', { class: 'row' }, ui.field('Print &what:', what), ui.field('Slides per page:', per)), ui.field('Color/gra&yscale:', color), frame, hid);
    ui.dialog({
      title: 'Print', body, width: 420, buttons: [{
        label: 'Preview', primary: true, onClick: () => {
          const range = all.input.checked ? null : cur.input.checked ? [E().idx] : parseRange(rIn.value, n);
          Object.assign(L.app.printOpts, { what: what.value, per: +per.value, color: color.value, frame: frame.input.checked, hidden: hid.input.checked, range });
          L.app.setView('preview');
        },
      }, { label: 'Cancel' }],
    });
  };
  function parseRange(s, n) {
    const out = [];
    for (const part of String(s).split(/[,;]/)) {
      const m = /^\s*(\d+)\s*(?:-\s*(\d+))?\s*$/.exec(part);
      if (!m) continue;
      const a = +m[1], b = m[2] ? +m[2] : a;
      for (let i = Math.min(a, b); i <= Math.max(a, b); i++) if (i >= 1 && i <= n) out.push(i - 1);
    }
    return out.length ? Array.from(new Set(out)) : null;
  }

  /* ---------- About ---------- */
  D.about = () => ui.about({ name: 'Lectern', office: 'PowerPoint', text: [
    'A presentation editor in the style of Microsoft Office PowerPoint 2003 that runs entirely in your browser. Presentations are read and written as Office Open XML (.pptx, .pptm); nothing is uploaded.',
    'Slide shows play in the browser, and presentations export to PDF, PNG pictures and web pages.'] });
})();
