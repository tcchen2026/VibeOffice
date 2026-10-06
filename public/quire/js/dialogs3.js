/* Quire — dialogs, part 3: Format Object (pictures, AutoShapes, text boxes, WordArt), Fill Effects,
 * Tools (Word Count, AutoSummarize, Hyphenation, Envelopes and Labels, Letter Wizard, Macros, Templates,
 * AutoCorrect, Customize, Options) and File (Properties, Print, Save As, New from template, Zoom,
 * Paste Special, About).
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const { h } = L;
  const ui = L.ui;
  const G = L.dlg;
  const doc = () => D.doc;
  const A = () => L.app;

  /* ================= Format Object ================= */
  G.formatObject = function (tab) {
    const it0 = L.drawing.selectedItem();
    if (!it0) return;
    const it = L.clone(it0);
    const isImg = it.t === 'img', isShape = it.t === 'shape', isTb = !!it.tb, isWA = !!it.wordart, isLine = isShape && (it.geom === 'line' || L.geom.isLine(it.geom));
    const title = isImg ? 'Format Picture' : isTb ? 'Format Text Box' : isWA ? 'Format WordArt' : it.t === 'chart' ? 'Format Object' : 'Format AutoShape';
    /* Colors and Lines */
    let fill = it.fill ? L.clone(it.fill) : { t: 'none' };
    const fillBtn = G.color(fill.t === 'solid' ? String(fill.c || '#FFFFFF').replace('#', '') : null, { id: 'fo-fill', mode: 'fill', automatic: false, noneText: fill.t === 'none' ? 'No Fill' : fill.t === 'grad' ? 'Gradient' : fill.t === 'patt' ? 'Pattern' : fill.t === 'img' ? 'Picture' : 'No Fill', onChange: (v) => { fill = v ? { t: 'solid', c: '#' + v, a: 1 - transp.get() / 100 } : { t: 'none' }; } });
    const fx = ui.button('Fill &Effects...', async () => { const f = await G.fillEffects(fill); if (f) { fill = f; fillBtn.set(null); fillBtn.querySelector('.cpk-l').textContent = { grad: 'Gradient', patt: 'Pattern', img: 'Picture', solid: 'Color' }[f.t] || 'Fill'; } }, { class: 'btn small' });
    const transp = ui.spin({ value: Math.round((1 - (fill.a == null ? 1 : fill.a)) * 100), min: 0, max: 100, step: 5, dec: 0, unit: ' %', id: 'fo-tr', onChange: (v) => { if (fill.t === 'solid') fill.a = 1 - v / 100; } });
    const ln = isImg ? (it.border ? { c: '#' + (it.border.color || '000000'), w: it.border.sz || 0.75, dash: 'solid' } : { t: 'none' }) : it.line ? L.clone(it.line) : { t: 'none' };
    const lineBtn = G.color(ln.t === 'none' || !ln.c ? null : String(ln.c).replace('#', ''), { id: 'fo-line', mode: 'line', automatic: false, noneText: 'No Line' });
    const dash = ui.select(Object.keys(L.drawing.DASH).map((k) => [k, k.replace(/([A-Z])/g, ' $1').toLowerCase()]), ln.dash || 'solid', null, { id: 'fo-dash' });
    const weight = ui.spin({ value: ln.w || 0.75, min: 0, max: 1584, step: 0.25, dec: 2, unit: ' pt', id: 'fo-wt' });
    const arrows = ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'].map((x) => [x, x === 'none' ? '(none)' : x]);
    const head = ui.select(arrows, (ln.head && ln.head.type) || 'none', null, { id: 'fo-ah' });
    const tail = ui.select(arrows, (ln.tail && ln.tail.type) || 'none', null, { id: 'fo-at' });
    /* Size */
    const hgt = G.len(it.h || 72, { id: 'fo-h', min: 1, onChange: (v) => { if (lock.input.checked && it.h) { wid.set((it.w * v) / it.h); scW.set(Math.round((wid.get() / (it.natW || it.w)) * 100)); } scH.set(Math.round((v / (it.natH || it.h)) * 100)); } });
    const wid = G.len(it.w || 72, { id: 'fo-w', min: 1, onChange: (v) => { if (lock.input.checked && it.w) { hgt.set((it.h * v) / it.w); scH.set(Math.round((hgt.get() / (it.natH || it.h)) * 100)); } scW.set(Math.round((v / (it.natW || it.w)) * 100)); } });
    const rot = ui.spin({ value: it.rot || 0, min: -3600, max: 3600, step: 1, dec: 1, unit: '°', id: 'fo-rot' });
    const scH = ui.spin({ value: 100, min: 1, max: 5000, step: 1, dec: 0, unit: ' %', id: 'fo-sh', onChange: (v) => { hgt.set(((it.natH || it.h) * v) / 100); if (lock.input.checked) { wid.set(((it.natW || it.w) * v) / 100); scW.set(v); } } });
    const scW = ui.spin({ value: 100, min: 1, max: 5000, step: 1, dec: 0, unit: ' %', id: 'fo-sw', onChange: (v) => { wid.set(((it.natW || it.w) * v) / 100); if (lock.input.checked) { hgt.set(((it.natH || it.h) * v) / 100); scH.set(v); } } });
    if (it.natW) { scW.set(Math.round((it.w / it.natW) * 100)); scH.set(Math.round((it.h / it.natH) * 100)); }
    const lock = ui.check('Lock &aspect ratio', isImg || it.t === 'chart' || !!it.lockAspect, null);
    const reset = ui.button('Re&set', () => { if (it.natW) { wid.set(it.natW); hgt.set(it.natH); scW.set(100); scH.set(100); } rot.set(0); }, { class: 'btn small' });
    /* Layout */
    let wrap = it.float ? it.float.wrap || 'none' : 'inline';
    const wrapBox = h('div', { class: 'preset-grid' });
    const wrapSvg = { inline: '<svg width="40" height="36"><path d="M2 6h36M2 12h8M30 12h8M2 18h36M2 24h36M2 30h36" stroke="#999"/><rect x="12" y="8" width="16" height="10" fill="#9cf" stroke="#333"/></svg>', square: '<svg width="40" height="36"><path d="M2 6h36M2 12h10M28 12h10M2 18h10M28 18h10M2 24h10M28 24h10M2 30h36" stroke="#999"/><rect x="13" y="10" width="14" height="16" fill="#9cf" stroke="#333"/></svg>', tight: '<svg width="40" height="36"><path d="M2 6h36M2 12h12M26 12h12M2 18h9M29 18h9M2 24h12M26 24h12M2 30h36" stroke="#999"/><circle cx="20" cy="18" r="8" fill="#9cf" stroke="#333"/></svg>', behind: '<svg width="40" height="36"><rect x="10" y="8" width="20" height="20" fill="#9cf" stroke="#333"/><path d="M2 6h36M2 12h36M2 18h36M2 24h36M2 30h36" stroke="#555"/></svg>', none: '<svg width="40" height="36"><path d="M2 6h36M2 12h36M2 18h36M2 24h36M2 30h36" stroke="#999"/><rect x="10" y="8" width="20" height="20" fill="#9cf" stroke="#333"/></svg>' };
    for (const [k, l] of [['inline', 'In l<u>i</u>ne with text'], ['square', '<u>S</u>quare'], ['tight', '<u>T</u>ight'], ['behind', '<u>B</u>ehind text'], ['none', 'I<u>n</u> front of text']]) {
      const b = h('button', { type: 'button', class: wrap === k ? 'on' : '' }, h('span', { html: wrapSvg[k] }), h('span', { html: l }));
      b.addEventListener('click', () => { wrap = k; L.$$('button', wrapBox).forEach((x) => x.classList.toggle('on', x === b)); });
      wrapBox.appendChild(b);
    }
    const hAlign = G.radios('fo-ha', [['left', '&Left'], ['center', '&Center'], ['right', '&Right'], ['other', '&Other']], it.float && it.float.posH && it.float.posH.align ? it.float.posH.align : 'other');
    const adv = ui.button('Ad&vanced...', () => advanced(), { class: 'btn small' });
    let posH = it.float ? L.clone(it.float.posH) : { rel: 'column', off: 0 }, posV = it.float ? L.clone(it.float.posV) : { rel: 'paragraph', off: 0 };
    let dist = it.float ? L.clone(it.float.dist || { t: 0, b: 0, l: 9, r: 9 }) : { t: 0, b: 0, l: 9, r: 9 };
    let moveWith = it.float ? it.float.posV.rel === 'paragraph' : true, overlap = it.float ? it.float.allowOverlap !== false : true;
    function advanced() {
      const rels = [['column', 'Column'], ['margin', 'Margin'], ['page', 'Page'], ['character', 'Character']];
      const vrels = [['paragraph', 'Paragraph'], ['margin', 'Margin'], ['page', 'Page'], ['line', 'Line']];
      const hMode = G.radios('fa-hm', [['align', 'Ali&gnment'], ['abs', 'Absolute &position']], posH.align ? 'align' : 'abs');
      const hAl = ui.select([['left', 'Left'], ['center', 'Centered'], ['right', 'Right']], posH.align || 'left', null, { id: 'fa-ha' });
      const hRel = ui.select(rels, posH.rel || 'column', null, { id: 'fa-hr' });
      const hOff = G.len(posH.off || 0, { id: 'fa-ho' });
      const vMode = G.radios('fa-vm', [['align', 'Alig&nment'], ['abs', 'Absolute po&sition']], posV.align ? 'align' : 'abs');
      const vAl = ui.select([['top', 'Top'], ['center', 'Centered'], ['bottom', 'Bottom']], posV.align || 'top', null, { id: 'fa-va' });
      const vRel = ui.select(vrels, posV.rel || 'paragraph', null, { id: 'fa-vr' });
      const vOff = G.len(posV.off || 0, { id: 'fa-vo' });
      const mv = ui.check('&Move object with text', moveWith, null), ov = ui.check('Allo&w overlap', overlap, null);
      const dT = G.len(dist.t || 0, { id: 'fa-dt', min: 0 }), dB = G.len(dist.b || 0, { id: 'fa-db', min: 0 }), dL = G.len(dist.l || 0, { id: 'fa-dl', min: 0 }), dR = G.len(dist.r || 0, { id: 'fa-dr', min: 0 });
      ui.dialog({ title: 'Advanced Layout', width: 460, body: ui.tabs([{ label: 'Picture &Position', body: h('div', { class: 'col' }, ui.group('Horizontal', hMode, G.row(hAl, h('span', { text: 'relative to' }), hRel), G.row(hOff, h('span', { text: 'to the right of' }))), ui.group('Vertical', vMode, G.row(vAl, h('span', { text: 'relative to' }), vRel), G.row(vOff, h('span', { text: 'below' }))), ui.group('Options', mv, ov)) }, { label: '&Text Wrapping', body: ui.group('Distance from text', G.row(G.f('To&p:', dT, 'narrow'), G.f('&Left:', dL, 'narrow')), G.row(G.f('B&ottom:', dB, 'narrow'), G.f('&Right:', dR, 'narrow'))) }]), buttons: [{ label: 'OK', primary: true, onClick: () => {
        posH = hMode.get() === 'align' ? { rel: hRel.value, align: hAl.value } : { rel: hRel.value, off: hOff.get() };
        posV = vMode.get() === 'align' ? { rel: vRel.value === 'paragraph' ? 'margin' : vRel.value, align: vAl.value } : { rel: vRel.value, off: vOff.get() };
        moveWith = mv.input.checked; overlap = ov.input.checked;
        dist = { t: dT.get(), b: dB.get(), l: dL.get(), r: dR.get() };
        hAlign.set(posH.align || 'other');
        if (wrap === 'inline') wrap = 'square';
      } }, { label: 'Cancel' }] });
    }
    /* Picture */
    const crop = Object.assign({ l: 0, t: 0, r: 0, b: 0 }, it.crop || {});
    const cropL = G.len((crop.l || 0) * (it.natW || it.w), { id: 'fo-cl' }), cropR = G.len((crop.r || 0) * (it.natW || it.w), { id: 'fo-cr' }), cropT = G.len((crop.t || 0) * (it.natH || it.h), { id: 'fo-ct' }), cropB = G.len((crop.b || 0) * (it.natH || it.h), { id: 'fo-cb' });
    const color = ui.select([['', 'Automatic'], ['gray', 'Grayscale'], ['bw', 'Black & White'], ['washout', 'Washout']], it.gray ? 'gray' : it.bw ? 'bw' : it.washout ? 'washout' : '', null, { id: 'fo-col' });
    const bright = ui.spin({ value: Math.round(((it.bright || 0) + 1) * 50), min: 0, max: 100, step: 3, dec: 0, unit: ' %', id: 'fo-br' });
    const contrast = ui.spin({ value: Math.round(((it.contrast || 0) + 1) * 50), min: 0, max: 100, step: 3, dec: 0, unit: ' %', id: 'fo-co' });
    /* Text Box */
    const ins = Object.assign({ l: 7.2, t: 3.6, r: 7.2, b: 3.6 }, it.ins || {});
    const iL = G.len(ins.l, { id: 'fo-il', min: 0 }), iR = G.len(ins.r, { id: 'fo-ir', min: 0 }), iT = G.len(ins.t, { id: 'fo-it', min: 0 }), iB = G.len(ins.b, { id: 'fo-ib', min: 0 });
    const anchor = ui.select([['t', 'Top'], ['ctr', 'Center'], ['b', 'Bottom']], it.anchor || 't', null, { id: 'fo-anc' });
    const autofit = ui.check('Resize AutoShape to &fit text', !!it.autofit, null);
    /* Web */
    const alt = h('textarea', { rows: 4, id: 'fo-alt', style: 'width:100%' });
    alt.value = it.alt || it.descr || '';
    const tabsDef = [
      { label: 'Colors and Lines', body: h('div', { class: 'col' }, isImg ? null : ui.group('Fill', G.row(G.f('C&olor:', fillBtn), fx), G.f('&Transparency:', transp)), ui.group('Line', G.row(G.f('Co&lor:', lineBtn), G.f('&Dashed:', dash)), G.f('&Weight:', weight)), isLine || (isShape && !isTb && !isWA) ? ui.group('Arrows', G.row(G.f('&Begin style:', head), G.f('&End style:', tail))) : null) },
      { label: 'Size', body: h('div', { class: 'col' }, ui.group('Size and rotate', G.row(G.f('H&eight:', hgt, 'narrow'), G.f('Wi&dth:', wid, 'narrow')), G.f('Ro&tation:', rot)), ui.group('Scale', G.row(G.f('&Height:', scH, 'narrow'), G.f('&Width:', scW, 'narrow')), lock), it.natW ? h('div', { class: 'tp-note', text: `Original size — Height: ${D.fmtLen(it.natH)}, Width: ${D.fmtLen(it.natW)}` }) : null, h('div', { class: 'tp-btns' }, reset)) },
      { label: 'Layout', body: h('div', { class: 'col' }, ui.group('Wrapping style', wrapBox), ui.group('Horizontal alignment', h('div', { class: 'row' }, ...hAlign.querySelectorAll('.ck'))), h('div', { class: 'tp-btns' }, adv)) },
    ];
    if (isImg) tabsDef.push({ label: 'Picture', body: h('div', { class: 'col' }, ui.group('Crop from', G.row(G.f('&Left:', cropL, 'narrow'), G.f('T&op:', cropT, 'narrow')), G.row(G.f('Ri&ght:', cropR, 'narrow'), G.f('&Bottom:', cropB, 'narrow'))), ui.group('Image control', G.f('&Color:', color), G.f('Brig&htness:', bright), G.f('Co&ntrast:', contrast))) });
    if (isTb) tabsDef.push({ label: 'Text Box', body: h('div', { class: 'col' }, ui.group('Internal margin', G.row(G.f('&Left:', iL, 'narrow'), G.f('&Top:', iT, 'narrow')), G.row(G.f('&Right:', iR, 'narrow'), G.f('&Bottom:', iB, 'narrow'))), G.f('&Vertical alignment:', anchor), autofit) });
    tabsDef.push({ label: 'Web', body: h('div', { class: 'col' }, h('label', { for: 'fo-alt', html: '<u>A</u>lternative text:' }), alt, h('div', { class: 'tp-note', text: 'Web browsers display alternative text while pictures are loading or if they are missing. Screen readers read it aloud.' })) });
    const tabIdx = tab === 'picture' ? tabsDef.findIndex((t) => t.label === 'Picture') : tab === 'layout' ? 2 : tab === 'size' ? 1 : 0;
    ui.dialog({
      title, width: 470, body: ui.tabs(tabsDef, Math.max(0, tabIdx)),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        L.drawing.modifySelected(title, (x) => {
          if (!isImg) { if (fill.t === 'solid') fill.a = 1 - transp.get() / 100; x.fill = fill; }
          const lc = lineBtn.get();
          if (isImg) x.border = lc ? { val: dash.value === 'solid' ? 'single' : 'dashed', sz: weight.get(), color: lc } : undefined;
          else x.line = lc ? Object.assign({}, x.line && x.line.t !== 'none' ? x.line : {}, { c: '#' + lc, w: weight.get(), dash: dash.value, head: head.value === 'none' ? undefined : { type: head.value, w: 'med', len: 'med' }, tail: tail.value === 'none' ? undefined : { type: tail.value, w: 'med', len: 'med' } }) : { t: 'none' };
          if (x.line && x.line.t !== 'none') delete x.line.t;
          if (x.t === 'group' && (wid.get() !== x.w || hgt.get() !== x.h)) { const kx = wid.get() / x.w, ky = hgt.get() / x.h; const sc = (g) => { for (const k of g.kids || []) { k.x *= kx; k.y *= ky; k.w *= kx; k.h *= ky; sc(k); } }; sc(x); }
          x.w = L.round(wid.get(), 2); x.h = L.round(hgt.get(), 2);
          x.rot = rot.get() ? ((rot.get() % 360) + 360) % 360 : undefined;
          if (lock.input.checked && !isImg) x.lockAspect = true;
          L.drawing.setWrap(x, wrap);
          if (x.float) {
            x.float.posH = hAlign.get() !== 'other' ? { rel: posH.rel === 'character' ? 'column' : posH.rel || 'column', align: hAlign.get() } : posH;
            x.float.posV = posV;
            x.float.dist = dist;
            x.float.allowOverlap = overlap;
          }
          if (isImg) {
            const nw = x.natW || it.w, nh = x.natH || it.h;
            const c = { l: (cropL.get() || 0) / nw, r: (cropR.get() || 0) / nw, t: (cropT.get() || 0) / nh, b: (cropB.get() || 0) / nh };
            x.crop = c.l || c.r || c.t || c.b ? c : undefined;
            x.gray = color.value === 'gray' || undefined; x.bw = color.value === 'bw' || undefined; x.washout = color.value === 'washout' || undefined;
            x.bright = L.round(bright.get() / 50 - 1, 2) || undefined; x.contrast = L.round(contrast.get() / 50 - 1, 2) || undefined;
          }
          if (isTb) { x.ins = { l: iL.get(), r: iR.get(), t: iT.get(), b: iB.get() }; x.anchor = anchor.value === 't' ? undefined : anchor.value; x.autofit = autofit.input.checked || undefined; }
          x.alt = alt.value || undefined;
        });
      } }, { label: 'Cancel' }],
    });
  };
  /** Fill Effects: resolves with a fill object or null */
  G.fillEffects = function (cur) {
    return new Promise((resolve) => {
      cur = cur || {};
      let res = null;
      const c1 = G.color(cur.t === 'grad' && cur.stops ? String(cur.stops[0].c).replace('#', '') : '3366FF', { id: 'fe-c1', automatic: false });
      const c2 = G.color(cur.t === 'grad' && cur.stops ? String(cur.stops[cur.stops.length - 1].c).replace('#', '') : 'FFFFFF', { id: 'fe-c2', automatic: false });
      const colors = G.radios('fe-cols', [['one', 'O&ne color'], ['two', '&Two colors'], ['preset', 'Pr&eset']], 'two');
      const PRESETS = { 'Early Sunset': ['#000082', '#66008F', '#BA0066', '#FF0000', '#FF8200'], Horizon: ['#DCEBF5', '#83A7C3', '#768FB9', '#83A7C3', '#FFFFFF'], Ocean: ['#03D4A8', '#21D6E0', '#0087E6', '#005CBF'], Fire: ['#FFF200', '#FF7A00', '#FF0300', '#4D0808'], Silver: ['#FFFFFF', '#E6E6E6', '#7D8496', '#E6E6E6'], Gold: ['#E6DCAC', '#E6D78A', '#C7AC4C', '#E6D78A', '#E6DCAC'], Rainbow: ['#A603AB', '#0819FB', '#1A8D48', '#FFFF00', '#EE3F17', '#E81766'] };
      const preset = ui.select(Object.keys(PRESETS).map((k) => [k, k]), 'Early Sunset', null, { id: 'fe-pre' });
      const style = G.radios('fe-st', [['h', 'Hori&zontal'], ['v', '&Vertical'], ['du', 'Diagonal &up'], ['dd', 'Dia&gonal down'], ['corner', '&From corner'], ['center', 'From ce&nter']], 'h');
      const pattGrid = h('div', { class: 'shape-grid', style: 'grid-template-columns:repeat(8,26px)' });
      let patt = cur.t === 'patt' ? cur.prst : 'pct50';
      for (const k of Object.keys(L.drawing.PATTERNS)) { const b = h('button', { type: 'button', class: k === patt ? 'on' : '', 'data-tip': k, 'aria-label': k, html: L.drawing.patternSVG(k, '#000', '#fff', 22) }); b.addEventListener('click', () => { patt = k; L.$$('button', pattGrid).forEach((x) => x.classList.toggle('on', x === b)); }); pattGrid.appendChild(b); }
      const fg = G.color('000000', { id: 'fe-fg', automatic: false }), bg = G.color('FFFFFF', { id: 'fe-bg', automatic: false });
      let pic = cur.t === 'img' ? cur.media : null;
      const picName = h('div', { class: 'tp-note', text: pic ? 'Picture selected' : 'No picture selected' });
      let which = cur.t === 'patt' ? 2 : cur.t === 'img' ? 3 : 0;
      const tabs = ui.tabs([
        { label: '&Gradient', body: h('div', { class: 'col' }, ui.group('Colors', colors, G.row(G.f('Color &1:', c1), G.f('Color &2:', c2)), G.f('Preset colors:', preset)), ui.group('Shading styles', style)), onShow: () => { which = 0; } },
        { label: 'Te&xture', body: h('div', { class: 'tp-note', text: 'Textures are drawn as patterns; choose a pattern or a picture.' }), onShow: () => { which = 1; } },
        { label: '&Pattern', body: h('div', { class: 'col' }, pattGrid, G.row(G.f('&Foreground:', fg), G.f('&Background:', bg))), onShow: () => { which = 2; } },
        { label: 'P&icture', body: h('div', { class: 'col' }, ui.button('Select &Picture...', async () => { const f = (await L.pickFiles('image/*'))[0]; if (!f) return; pic = L.media.add(new Blob([await L.readAsArrayBuffer(f)], { type: f.type }), f.name); picName.textContent = f.name; }), picName), onShow: () => { which = 3; } },
      ], which);
      ui.dialog({ title: 'Fill Effects', width: 420, body: tabs, buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (which === 0) {
          const ang = { h: 90, v: 0, du: 45, dd: 135, corner: 45, center: 0 }[style.get()];
          let stops;
          if (colors.get() === 'preset') { const cs = PRESETS[preset.value]; stops = cs.map((c, i) => ({ p: i / (cs.length - 1), c, a: 1 })); }
          else if (colors.get() === 'one') stops = [{ p: 0, c: '#' + c1.get(), a: 1 }, { p: 1, c: '#FFFFFF', a: 1 }];
          else stops = [{ p: 0, c: '#' + c1.get(), a: 1 }, { p: 1, c: '#' + c2.get(), a: 1 }];
          res = { t: 'grad', ang, stops, path: style.get() === 'center' ? 'circle' : style.get() === 'corner' ? 'rect' : undefined };
        } else if (which === 2 || which === 1) res = { t: 'patt', prst: patt, fg: '#' + fg.get(), bg: '#' + bg.get() };
        else if (which === 3 && pic) res = { t: 'img', media: pic };
      } }, { label: 'Cancel' }], onClose: () => resolve(res) });
    });
  };
  G.drawingGrid = function () {
    const snap = ui.check('&Snap objects to grid', !!A().opts.snapGrid, null);
    const hs = G.len(A().opts.gridH || 9, { id: 'dg-h', min: 1 }), vs = G.len(A().opts.gridV || 9, { id: 'dg-v', min: 1 });
    ui.dialog({ title: 'Drawing Grid', width: 300, body: h('div', { class: 'col' }, snap, ui.group('Grid settings', G.f('&Horizontal spacing:', hs, 'wide'), G.f('&Vertical spacing:', vs, 'wide'))), buttons: [{ label: 'OK', primary: true, onClick: () => { A().opts.snapGrid = snap.input.checked; A().opts.gridH = hs.get(); A().opts.gridV = vs.get(); A().saveOpts(); } }, { label: 'Cancel' }] });
  };


  /* ================= Format ▸ AutoFormat ================= */
  G.autoFormatDoc = function () {
    const mode = G.radios('af-m', [['now', '&AutoFormat now'], ['review', 'AutoFormat and &review each change']], 'now');
    const kind = ui.select([['general', 'General document'], ['letter', 'Letter'], ['email', 'E-mail']], 'general', null, { id: 'af-k' });
    const o = A().opts;
    const opts = { headings: ui.check('Built-in Heading &styles', true, null), lists: ui.check('&List styles', true, null), quotes: ui.check('"Straight &quotes" with “smart quotes”', o.smartQuotes !== false, null), dashes: ui.check('&Hyphens (--) with dash (—)', o.dashes !== false, null), ordinals: ui.check('&Ordinals (1st) with superscript', o.ordinals !== false, null), fractions: ui.check('&Fractions (1/2) with fraction character', o.fractions !== false, null), urls: ui.check('&Internet paths with hyperlinks', o.autoUrl !== false, null), empty: ui.check('Remove &extra empty paragraphs', true, null) };
    ui.dialog({ title: 'AutoFormat', width: 400, body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Quire will automatically format the document.' }), mode, G.f('Please select a document type to help improve the formatting process:', kind, 'wide'), ui.group('Apply / replace', ...Object.values(opts))), buttons: [{ label: 'OK', primary: true, onClick: () => {
      const cfg = {}; for (const k in opts) cfg[k] = opts[k].input.checked;
      if (kind.value === 'letter') cfg.headings = false;
      const n = autoFormat(cfg);
      if (mode.get() === 'review') ui.dialog({ title: 'AutoFormat', width: 380, body: h('div', { class: 'tp-note', text: `Formatting completed: ${n} change${n === 1 ? '' : 's'}. You can accept all of them, or reject all and return the document to its previous state.` }), buttons: [{ label: '&Accept All', primary: true }, { label: '&Reject All', onClick: () => { D.undo(); LY.update(); E.restoreDom(true); } }] });
      else A().status(`AutoFormat made ${n} change${n === 1 ? '' : 's'}.`, 6000);
    } }, { label: 'Cancel' }] });
  };
  function autoFormat(cfg) {
    const d = doc();
    let n = 0;
    E.edit('AutoFormat', () => {
      D.reindex(d);
      const paras = d.main.paras.slice();
      const textReplace = (p, re, fn) => {
        const t = D.ptext(p);
        const hits = [];
        let m;
        re.lastIndex = 0;
        while ((m = re.exec(t))) { hits.push([m.index, m[0], fn(m, t)]); if (!re.global) break; }
        for (let i = hits.length - 1; i >= 0; i--) {
          const [at, src, rep] = hits[i];
          if (rep == null || rep === src) continue;
          const rp = O.rPrAt(D.pos(p, at + 1));
          O.deleteRange(D.pos(p, at), D.pos(p, at + src.length));
          if (typeof rep === 'string') O.insertText(D.pos(p, at), rep, rp);
          else { O.insertText(D.pos(p, at), rep.text, Object.assign({}, rp, rep.rPr)); }
          n++;
        }
      };
      for (const p of paras) {
        if (cfg.quotes) {
          textReplace(p, /"/g, (m, t) => { const pr = t[m.index - 1]; return !pr || /[\s([{—–\-]/.test(pr) ? '“' : '”'; });
          textReplace(p, /'/g, (m, t) => { const pr = t[m.index - 1]; return !pr || /[\s([{—–\-“]/.test(pr) ? '‘' : '’'; });
        }
        if (cfg.dashes) { textReplace(p, /(\S) ?-- ?(\S)/g, (m) => m[1] + '—' + m[2]); textReplace(p, /(\w) - (\w)/g, (m) => m[1] + ' – ' + m[2]); }
        if (cfg.fractions) textReplace(p, /\b(1\/2|1\/4|3\/4)\b/g, (m) => ({ '1/2': '½', '1/4': '¼', '3/4': '¾' })[m[1]]);
        if (cfg.ordinals) {
          const t = D.ptext(p);
          const re = /\b(\d+)(st|nd|rd|th)\b/g;
          let m;
          const hits = [];
          while ((m = re.exec(t))) hits.push([m.index + m[1].length, m[2].length]);
          for (const [a, l] of hits) { O.setRunProps(D.pos(p, a), D.pos(p, a + l), { vert: 'superscript' }); n++; }
        }
        if (cfg.urls) {
          const t = D.ptext(p);
          const re = /\b((?:https?:\/\/|www\.)[^\s<>"“”]+[^\s<>"“”.,;:!?)\]])/g;
          let m;
          const hits = [];
          while ((m = re.exec(t))) hits.push([m.index, m[1]]);
          for (const [a, u] of hits) { const r = O.rPrAt(D.pos(p, a + 1)); if (r.link) continue; O.setRunProps(D.pos(p, a), D.pos(p, a + u.length), { link: { url: /^www\./.test(u) ? 'http://' + u : u }, style: 'Hyperlink' }); n++; }
        }
      }
      D.reindex(d);
      const blocks = d.main.blocks;
      const plain = (p) => D.ptext(p).replace(/[-]/g, '');
      if (cfg.lists) {
        let bulletNum = null, numNum = null, lastKind = null, touchedNum = false;
        for (let i = 0; i < blocks.length; i++) {
          const p = blocks[i];
          if (p.t !== 'p' || (p.pPr.num && p.pPr.num.id)) { lastKind = null; continue; }
          const t = plain(p);
          const mb = /^\s*([*•\-–>])\s+\S/.exec(t), mn = /^\s*(\d+|[a-zA-Z])[.)]\s+\S/.exec(t);
          if (!mb && !mn) { if (t.trim()) lastKind = null; continue; }
          const kind = mb ? 'b' : 'n';
          const cut = (mb || mn)[0].length - 1;
          O.deleteRange(D.pos(p, 0), D.pos(p, cut));
          D.touch(p);
          if (!touchedNum) { D.touchKey(d, 'numbering'); touchedNum = true; }
          if (kind === 'b') { if (!bulletNum || lastKind !== 'b') bulletNum = D.addNum(d, D.makeBulletAbs('•', 'Symbol', '')); p.pPr.num = { id: bulletNum, lvl: 0 }; }
          else { if (!numNum || lastKind !== 'n') numNum = D.addNum(d, D.makeNumberAbs(/[a-z]/.test(mn[1]) ? 'lowerLetter' : /[A-Z]/.test(mn[1]) ? 'upperLetter' : 'decimal', '%1.')); p.pPr.num = { id: numNum, lvl: 0 }; }
          if (!p.pPr.style || p.pPr.style === 'Normal') p.pPr.style = 'ListParagraph' in d.styles ? 'ListParagraph' : p.pPr.style;
          lastKind = kind;
          n++;
        }
      }
      if (cfg.headings) {
        for (let i = 0; i < blocks.length; i++) {
          const p = blocks[i];
          if (p.t !== 'p' || D.headingLevel(d, p) || (p.pPr.num && p.pPr.num.id) || (p.pPr.style && p.pPr.style !== 'Normal')) continue;
          const t = plain(p).trim();
          const next = blocks[i + 1];
          if (!t || t.length > 70 || /[.,;:!?]$/.test(t) || t.split(/\s+/).length > 10) continue;
          if (!next || next.t !== 'p' || !plain(next).trim() || plain(next).trim().length < t.length) continue;
          const prev = blocks[i - 1];
          if (prev && prev.t === 'p' && plain(prev).trim() && !D.headingLevel(d, prev)) continue;
          D.touch(p);
          p.pPr.style = /^[A-Z0-9 \-:&]+$/.test(t) || i < 3 ? 'Heading1' : 'Heading2';
          if (!d.styles[p.pPr.style]) { D.touchKey(d, 'styles'); d.styles[p.pPr.style] = D.builtinStyles()[p.pPr.style]; D.stylesChanged(); }
          n++;
        }
      }
      if (cfg.empty) {
        const cont = d.main;
        let removed = false;
        for (let i = blocks.length - 2; i > 0; i--) {
          const a = blocks[i], b = blocks[i - 1];
          if (a.t === 'p' && b.t === 'p' && !a.sect && !b.sect && !plain(a).trim() && !plain(b).trim() && a.runs.every((r) => r.t === 'text') && b.runs.every((r) => r.t === 'text')) { if (!removed) { D.touchList(cont); removed = true; } blocks.splice(i, 1); n++; }
        }
        if (removed) d._idxDirty = true;
      }
      D.reindex(d);
      return E.sel && D.byId(d, E.sel.f.p.id) ? E.sel : D.pos(D.firstPara(d.main.blocks), 0);
    });
    return n;
  }
  /* File ▸ New ▸ From existing document: open a copy that is not tied to the original file name */
  L.app.newFromExisting = async function () {
    const f = (await L.pickFiles('.docx,.dotx,.docm,.dotm,.htm,.html,.txt,.rtf'))[0];
    if (!f) return;
    await A().openFile(f);
    const nm = 'Document' + (++A().untitled);
    A().fileName = nm;
    A().fileType = 'docx';
    doc().dirty = true;
    A().updateTitle();
  };

  /* ================= Word Count ================= */
  G.wordCount = function () {
    const d = doc();
    const selOnly = E.sel && !E.collapsed();
    const incl = ui.check('Include &footnotes and endnotes', !!L.store.get('wcNotes', false), () => { L.store.set('wcNotes', incl.input.checked); draw(); });
    const tbl = h('table', { class: 'wc-table' });
    const draw = () => {
      L.clear(tbl);
      let paras;
      if (selOnly) { const [a, b] = E.range(); paras = D.parasBetween(d, a.p, b.p).map((p) => (p === a.p || p === b.p ? D.para(D.sliceRuns(p, p === a.p ? a.o : 0, p === b.p ? b.o : D.plen(p))) : p)); }
      else { paras = D.allParas(d); if (incl.input.checked) for (const k of ['fn', 'en']) for (const id in d[k]) D.walk(d[k][id], (b) => { if (b.t === 'p') paras.push(b); }); }
      const st = D.stats(paras);
      let lines = 0;
      if (LY.pages.length) for (const pg of LY.pages) for (const f of pg.el.querySelectorAll('.pg-body .p')) { if (f.closest('.balloons')) continue; const lh = parseFloat(getComputedStyle(f.firstChild).lineHeight) || 16; lines += Math.max(1, Math.round(f.firstChild.offsetHeight / lh)); }
      const rows = [['Pages', selOnly ? '' : LY.pages.length || 1], ['Words', st.words], ['Characters (no spaces)', st.charsNoSp], ['Characters (with spaces)', st.chars], ['Paragraphs', st.paras], ['Lines', selOnly ? '' : lines || st.paras]];
      for (const [k, v] of rows) if (v !== '') tbl.appendChild(h('tr', null, h('td', { text: k }), h('td', { text: Number(v).toLocaleString() })));
    };
    draw();
    ui.dialog({ title: 'Word Count', width: 300, body: h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Statistics:' + (selOnly ? ' (selection)' : '') }), tbl, incl), buttons: [{ label: 'Close', primary: true }] });
  };

  /* ================= AutoSummarize ================= */
  G.autoSummarize = function () {
    const kind = G.radios('as-k', [['highlight', '&Highlight key points'], ['top', '&Insert an executive summary or abstract at the top of the document'], ['new', '&Create a new document and put the summary there'], ['hide', 'Hide &everything but the summary without leaving the original document']], 'highlight');
    const pct = ui.select([[10, '10 sentences'], [20, '20 sentences'], [100, '100 words or less'], [500, '500 words or less'], ['10%', '10%'], ['25%', '25%'], ['50%', '50%'], ['75%', '75%']], '25%', null, { id: 'as-pct' });
    const upd = ui.check('&Update document statistics', true, null);
    ui.dialog({ title: 'AutoSummarize', width: 460, body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Quire has examined the document and picked the sentences most relevant to the main theme.' }), ui.group('Type of summary', kind), ui.group('Length of summary', G.f('&Percent of original:', pct, 'wide')), upd), buttons: [{ label: 'OK', primary: true, onClick: () => summarize(kind.get(), pct.value, upd.input.checked) }, { label: 'Cancel' }] });
  };
  function summarize(kind, len, updStats) {
    const d = doc();
    D.reindex(d);
    const sents = [];
    const freq = new Map();
    const STOP = new Set('a an and are as at be but by for from has have he her his i in is it its of on or our she that the their them there they this to was we were what when which who will with you your not no so if than then also can may'.split(' '));
    for (const p of d.main.paras) {
      const t = D.ptext(p);
      const re = /[^.!?]+[.!?]*\s*/g;
      let m;
      while ((m = re.exec(t))) {
        const s = m[0];
        if (!s.trim()) continue;
        const words = (s.toLowerCase().match(/[\p{L}']+/gu) || []).filter((w) => !STOP.has(w) && w.length > 2);
        for (const w of words) freq.set(w, (freq.get(w) || 0) + 1);
        sents.push({ p, a: m.index, b: m.index + s.trimEnd().length, text: s.trim(), words, idx: sents.length });
      }
    }
    if (!sents.length) return;
    for (const s of sents) s.score = s.words.reduce((a, w) => a + (freq.get(w) || 0), 0) / Math.max(4, s.words.length) + (D.headingLevel(d, s.p) ? 2 : 0) + (s.idx === 0 ? 1 : 0);
    let pick;
    const totalWords = sents.reduce((a, s) => a + s.text.split(/\s+/).length, 0);
    const ranked = sents.slice().sort((a, b) => b.score - a.score);
    if (/%$/.test(len)) { const target = (totalWords * parseFloat(len)) / 100; pick = []; let w = 0; for (const s of ranked) { if (w >= target && pick.length) break; pick.push(s); w += s.text.split(/\s+/).length; } }
    else if (+len >= 100) { pick = []; let w = 0; for (const s of ranked) { const n = s.text.split(/\s+/).length; if (w + n > +len && pick.length) break; pick.push(s); w += n; } }
    else pick = ranked.slice(0, +len);
    pick.sort((a, b) => a.idx - b.idx);
    if (kind === 'highlight' || kind === 'hide') {
      E.edit('AutoSummarize', () => {
        if (kind === 'highlight') for (const s of pick) O.setRunProps(D.pos(s.p, s.a), D.pos(s.p, s.b), { hl: 'yellow' });
        else { const set = new Set(pick); for (const s of sents) if (!set.has(s)) O.setRunProps(D.pos(s.p, s.a), D.pos(s.p, s.b), { hidden: true }); }
        return E.sel;
      });
      A().status(`${pick.length} key sentence${pick.length === 1 ? '' : 's'} ${kind === 'highlight' ? 'highlighted' : 'kept visible; the rest is hidden text'}. Undo to restore.`, 8000);
    } else if (kind === 'top') {
      E.edit('AutoSummarize', () => {
        const cont = D.touchList(d.main);
        const title = D.para([D.text('Summary', { b: true })], {});
        const body = D.para([D.text(pick.map((s) => s.text).join(' '))], {});
        cont.blocks.unshift(title, body, D.para());
        d._idxDirty = true;
        return D.pos(title, 0);
      });
    } else {
      const text = pick.map((s) => s.text).join(' ');
      A().newDocument({ build: (nd) => { nd.main.blocks = [D.para([D.text('Summary of ' + A().fileName, { b: true, sz: 14 })]), D.para([D.text(text)])]; } });
    }
    if (updStats) { d.props.description = (pick.map((s) => s.text).join(' ')).slice(0, 255); }
  }

  /* ================= Hyphenation ================= */
  G.hyphenation = function () {
    const d = doc();
    const auto = ui.check('&Automatically hyphenate document', !!d.settings.autoHyphen, null);
    const caps = ui.check('Hyphenate words in &CAPS', d.settings.hyphenCaps !== false, null);
    const zone = G.len(d.settings.hyphenZone || 18, { id: 'hy-z', min: 0 });
    const limit = G.num(d.settings.consecHyphen || null, { id: 'hy-l', min: 0, autoLabel: 'No limit' });
    ui.dialog({ title: 'Hyphenation', width: 320, body: h('div', { class: 'col' }, auto, caps, G.f('Hyphenation &zone:', zone, 'wide'), G.f('&Limit consecutive hyphens to:', limit, 'wide')), buttons: [{ label: 'OK', primary: true, onClick: () => { E.edit('Hyphenation', () => { D.touchKey(d, 'settings'); d.settings.autoHyphen = auto.input.checked; d.settings.hyphenCaps = caps.input.checked; d.settings.hyphenZone = zone.get(); d.settings.consecHyphen = limit.get() || undefined; return E.sel; }); LY.root.classList.toggle('hyphens', auto.input.checked); } }, { label: '&Manual...', onClick: () => { ui.msg('Manual hyphenation: place the insertion point inside a long word and press Ctrl+- to insert an optional hyphen.', { icon: 'info' }); return false; } }, { label: 'Cancel' }] });
  };

  /* ================= Envelopes and Labels ================= */
  const LABELS = [['5160', 'Avery 5160 – Address', 3, 10, 189, 72, 13.5, 36, 9], ['5161', 'Avery 5161 – Address', 2, 10, 288, 72, 11.25, 36, 13.5], ['5162', 'Avery 5162 – Address', 2, 7, 288, 96, 11.25, 63, 13.5], ['5163', 'Avery 5163 – Shipping', 2, 5, 288, 144, 11.25, 36, 13.5], ['5164', 'Avery 5164 – Shipping', 2, 3, 288, 240, 11.25, 36, 13.5], ['5167', 'Avery 5167 – Return Address', 4, 20, 126, 36, 21.6, 36, 21.6], ['5395', 'Avery 5395 – Name Badge', 2, 4, 243, 168, 49.5, 41.4, 27], ['L7160', 'Avery L7160 (A4) – Address', 3, 7, 179.15, 107.7, 20.7, 43.9, 7.2]];
  G.envelopes = function (tab) {
    const d = doc();
    let sel = E.sel && !E.collapsed() ? O.textRange(...E.range(), '\n') : '';
    const addr = h('textarea', { rows: 5, id: 'el-addr', style: 'width:100%' });
    addr.value = sel;
    const ret = h('textarea', { rows: 4, id: 'el-ret', style: 'width:100%' });
    ret.value = L.store.get('userAddress', '') || (L.store.get('userName', '') + '\n');
    const omit = ui.check('&Omit', false, null);
    const size = ui.select([['10', 'Size 10 (4 1/8 x 9 1/2 in)'], ['DL', 'DL (110 x 220 mm)'], ['C5', 'C5 (162 x 229 mm)'], ['Monarch', 'Monarch (3 7/8 x 7 1/2 in)']], '10', null, { id: 'el-size' });
    const laddr = h('textarea', { rows: 5, id: 'el-laddr', style: 'width:100%' });
    laddr.value = sel;
    const lprod = ui.select(LABELS.map((x) => [x[0], x[1]]), '5160', null, { id: 'el-prod' });
    const full = G.radios('el-full', [['full', '&Full page of the same label'], ['single', 'Single &label']], 'full');
    const row = G.num(1, { min: 1, max: 30, id: 'el-row' }), colN = G.num(1, { min: 1, max: 4, id: 'el-col' });
    let cur = tab === 'labels' ? 1 : 0;
    const tabs = ui.tabs([
      { label: '&Envelopes', body: h('div', { class: 'col' }, h('label', { for: 'el-addr', html: '<u>D</u>elivery address:' }), addr, G.row(h('label', { for: 'el-ret', html: '<u>R</u>eturn address:' }), omit), ret, G.f('Envelope si&ze:', size)), onShow: () => { cur = 0; } },
      { label: '&Labels', body: h('div', { class: 'col' }, h('label', { for: 'el-laddr', html: '<u>A</u>ddress:' }), laddr, G.f('Label &product:', lprod), ui.group('Print', full, G.row(G.f('Ro&w:', row, 'narrow'), G.f('Colu&mn:', colN, 'narrow')))), onShow: () => { cur = 1; } },
    ], cur);
    ui.dialog({ title: 'Envelopes and Labels', width: 420, body: tabs, buttons: [{ label: '&Add to Document', primary: true, onClick: () => { if (cur === 0) addEnvelope(); else newLabels(); } }, { label: cur === 0 ? '&Print (PDF)...' : 'New &Document', onClick: () => { if (cur === 0) { addEnvelope(); setTimeout(() => A().printPDF(), 100); } else newLabels(); } }, { label: 'Cancel' }] });
    function addEnvelope() {
      L.store.set('userAddress', ret.value);
      const sz = { 10: [684, 297], DL: [623.6, 311.8], C5: [649.15, 459.2], Monarch: [540, 279] }[size.value];
      E.edit('Add Envelope', () => {
        const cont = D.touchList(d.main);
        if (!d.styles.EnvelopeAddress) { D.touchKey(d, 'styles'); d.styles.EnvelopeAddress = D.builtinStyles().EnvelopeAddress; d.styles.EnvelopeReturn = D.builtinStyles().EnvelopeReturn; D.stylesChanged(); }
        const blocks = [];
        if (!omit.input.checked) for (const line of ret.value.split('\n')) blocks.push(D.para(line ? [D.text(line)] : [], { style: 'EnvelopeReturn' }));
        const lines = addr.value.split('\n');
        lines.forEach((line, i) => blocks.push(D.para(line ? [D.text(line)] : [], { style: 'EnvelopeAddress', ind: { l: sz[0] * 0.42 }, sp: i === 0 ? { b: Math.max(0, sz[1] * 0.42 - (omit.input.checked ? 0 : 60)) } : undefined })));
        const last = blocks[blocks.length - 1];
        last.sect = Object.assign(D.defaultSect(), { pgW: sz[0], pgH: sz[1], orient: 'landscape', mt: 18, mb: 18, ml: 36, mr: 36, type: 'nextPage' });
        cont.blocks.unshift(...blocks);
        d._idxDirty = true;
        return D.pos(blocks[0], 0);
      });
    }
    function newLabels() {
      const spec = LABELS.find((x) => x[0] === lprod.value);
      const [, , nc, nr, lw, lh, ml, mt, gap] = spec;
      const text = laddr.value;
      A().newDocument({ build: (nd) => {
        nd.sect.ml = ml; nd.sect.mr = 0; nd.sect.mt = mt; nd.sect.mb = 0; nd.sect.hdr = 0; nd.sect.ftr = 0;
        if (/A4/.test(spec[1])) { nd.sect.pgW = 595.3; nd.sect.pgH = 841.9; }
        const cols = [];
        for (let c = 0; c < nc; c++) { cols.push(lw); if (c < nc - 1 && gap) cols.push(gap); }
        const t = D.table([], cols, { style: 'TableNormal', w: { type: 'dxa', v: cols.reduce((a, b) => a + b, 0) }, layout: 'fixed', cellMar: { l: 5.4, r: 5.4, t: 0, b: 0 }, look: {} });
        for (let r = 0; r < nr; r++) {
          const cells = [];
          for (let c = 0; c < cols.length; c++) {
            const isGap = gap && c % 2 === 1;
            const lr = Math.floor(c / (gap ? 2 : 1));
            const fill = !isGap && (full.get() === 'full' || (r === row.get() - 1 && lr === colN.get() - 1));
            const blocks = fill ? text.split('\n').map((ln) => D.para(ln ? [D.text(ln)] : [], { ind: { l: 5 }, sp: { b: 0, a: 0 } })) : [D.para()];
            cells.push(D.cell(blocks, { w: cols[c], vAlign: 'center' }));
          }
          t.rows.push(D.row(cells, { h: lh, hRule: 'exact', cantSplit: true }));
        }
        nd.main.blocks = [t, D.para()];
      } });
    }
  };

  /* ================= Letter Wizard ================= */
  G.letterWizard = function () {
    const date = ui.check('&Date line:', true, null);
    const datePic = ui.select(L.fields.DATE_PICTURES.slice(0, 6).map((p) => [p, L.fields.formatDate(new Date(), p)]), 'MMMM d, yyyy', null, { id: 'lw-date' });
    const style = ui.select([['full', 'Full block'], ['modified', 'Modified block'], ['semi', 'Semi-block']], 'full', null, { id: 'lw-style' });
    const rname = h('input', { type: 'text', id: 'lw-rn', style: 'width:100%' });
    const raddr = h('textarea', { rows: 4, id: 'lw-ra', style: 'width:100%' });
    const sal = h('input', { type: 'text', id: 'lw-sal', value: 'Dear ', list: 'lw-sals', style: 'width:100%' });
    const sals = h('datalist', { id: 'lw-sals' }, ...['Dear ', 'Dear Sir or Madam:', 'Ladies and Gentlemen:', 'To Whom It May Concern:'].map((x) => h('option', { value: x })));
    const refLine = h('input', { type: 'text', id: 'lw-ref', style: 'width:100%', placeholder: 'In reply to:' });
    const subj = h('input', { type: 'text', id: 'lw-sub', style: 'width:100%' });
    const cc = h('input', { type: 'text', id: 'lw-cc', style: 'width:100%' });
    const sname = h('input', { type: 'text', id: 'lw-sn', value: L.store.get('userName', ''), style: 'width:100%' });
    const saddr = h('textarea', { rows: 3, id: 'lw-sa', style: 'width:100%' });
    saddr.value = L.store.get('userAddress', '');
    const closing = h('input', { type: 'text', id: 'lw-cl', value: 'Sincerely,', list: 'lw-cls', style: 'width:100%' });
    const cls = h('datalist', { id: 'lw-cls' }, ...L.autocorrect.AUTOTEXT.Closing.map((x) => h('option', { value: x })));
    const title = h('input', { type: 'text', id: 'lw-ti', style: 'width:100%' });
    const encl = G.num(0, { id: 'lw-enc', min: 0 });
    const tabs = ui.tabs([
      { label: 'Letter Format', body: h('div', { class: 'col' }, G.row(date, datePic), G.f('Page design / Letter style:', style, 'wide')) },
      { label: 'Recipient Info', body: h('div', { class: 'col' }, G.f('Recipient\'s name:', rname, 'wide'), h('label', { for: 'lw-ra', text: 'Delivery address:' }), raddr, sals, G.f('Salutation:', sal, 'wide')) },
      { label: 'Other Elements', body: h('div', { class: 'col' }, G.f('Reference line:', refLine, 'wide'), G.f('Subject:', subj, 'wide'), G.f('Courtesy copies (cc):', cc, 'wide')) },
      { label: 'Sender Info', body: h('div', { class: 'col' }, G.f('Sender\'s name:', sname, 'wide'), h('label', { for: 'lw-sa', text: 'Return address:' }), saddr, cls, G.f('Complimentary closing:', closing, 'wide'), G.f('Job title:', title, 'wide'), G.f('Enclosures:', encl, 'wide')) },
    ]);
    ui.dialog({ title: 'Letter Wizard', width: 460, body: tabs, buttons: [{ label: 'OK', primary: true, onClick: () => {
      L.store.set('userAddress', saddr.value);
      const indent = style.value === 'full' ? 0 : 252;
      const P = (t, o) => D.para(t ? [D.text(t)] : [], o || {});
      const bl = [];
      for (const line of saddr.value.split('\n').filter(Boolean)) bl.push(P(line, { ind: { l: indent } }));
      if (bl.length) bl.push(P(''));
      if (date.input.checked) bl.push(P(L.fields.formatDate(new Date(), datePic.value), { ind: { l: indent } }), P(''));
      if (refLine.value) bl.push(P(refLine.value), P(''));
      if (rname.value) bl.push(P(rname.value));
      for (const line of raddr.value.split('\n').filter(Boolean)) bl.push(P(line));
      bl.push(P(''));
      if (subj.value) bl.push(D.para([D.text('Subject: ', { b: true }), D.text(subj.value)]), P(''));
      bl.push(P(sal.value.trim() === 'Dear' || sal.value === 'Dear ' ? 'Dear ' + (rname.value || '') + ':' : sal.value), P(''));
      bl.push(P('Type your letter here.', { ind: { fl: style.value === 'semi' ? 36 : 0 } }), P(''));
      bl.push(P(closing.value, { ind: { l: indent } }), P('', { ind: { l: indent } }), P('', { ind: { l: indent } }), P(sname.value, { ind: { l: indent } }));
      if (title.value) bl.push(P(title.value, { ind: { l: indent } }));
      if (encl.get()) bl.push(P(''), P(`Enclosure${encl.get() > 1 ? 's' : ''} (${encl.get()})`));
      if (cc.value) bl.push(P(''), P('cc: ' + cc.value));
      E.pasteBlocks(bl.map((b) => (delete b.partial, b)));
    } }, { label: 'Cancel' }] });
  };

  /* ================= Macros, Templates & Organizer ================= */
  G.macros = function () {
    const lb = G.listBox([], null, { height: 140 });
    ui.dialog({ title: 'Macros', width: 360, body: h('div', { class: 'col' }, h('label', { text: 'Macro name:' }), h('input', { type: 'text', style: 'width:100%', disabled: true }), lb, h('div', { class: 'tp-note', text: 'Macros are not run in Quire. VBA projects in .docm files are kept when the document is saved as .docm by Word, but Quire saves the document text only.' })), buttons: [{ label: 'Close', primary: true }] });
  };
  G.templates = function () {
    const d = doc();
    const auto = ui.check('Automatically &update document styles', !!d.settings.attachedTemplateUpdate, null);
    const name = h('input', { type: 'text', value: d.props.template || 'Normal', style: 'width:100%', readonly: true });
    ui.dialog({ title: 'Templates and Add-ins', width: 400, body: h('div', { class: 'col' }, ui.group('Document template', name, auto, h('div', { class: 'tp-btns' }, ui.button('&Attach...', async () => { const f = (await L.pickFiles('.dotx,.dotm,.docx'))[0]; if (!f) return; const res = await L.docx.read(await L.readAsArrayBuffer(f)); d.props.template = f.name; name.value = f.name; if (auto.input.checked) A().mergeStylesFrom(res.doc, true); A().status('Template attached: ' + f.name); }, { class: 'btn small' }), ui.button('&Organizer...', () => G.organizer(), { class: 'btn small' }))), h('div', { class: 'tp-note', text: 'Attaching a template with "Automatically update document styles" copies its styles into this document.' })), buttons: [{ label: 'OK', primary: true, onClick: () => { d.settings.attachedTemplateUpdate = auto.input.checked || undefined; } }, { label: 'Cancel' }] });
  };
  G.organizer = function () {
    const d = doc();
    let other = null, otherName = '';
    const mine = h('select', { size: 12, multiple: true, style: 'width:100%' });
    const theirs = h('select', { size: 12, multiple: true, style: 'width:100%' });
    const fillSel = (sel, dd) => { L.clear(sel); if (!dd) return; for (const k of Object.keys(dd.styles).sort()) sel.appendChild(h('option', { value: k, text: D.styleDisplayName(dd.styles[k]) })); };
    fillSel(mine, d);
    const lblOther = h('div', { class: 'cd-cap', text: 'In (no file):' });
    ui.dialog({ title: 'Organizer', width: 560, body: G.row(h('div', { class: 'col', style: 'flex:1' }, h('div', { class: 'cd-cap', text: 'In ' + A().fileName + ':' }), mine), h('div', { class: 'col', style: 'width:110px;justify-content:center' }, ui.button('<- &Copy', () => { if (!other) return; for (const o of theirs.selectedOptions) { E.edit('Copy Style', () => { D.touchKey(d, 'styles'); d.styles[o.value] = L.clone(other.styles[o.value]); D.stylesChanged(); return E.sel; }); } fillSel(mine, d); }), ui.button('&Delete', () => { E.edit('Delete Styles', () => { D.touchKey(d, 'styles'); for (const o of mine.selectedOptions) if (!d.styles[o.value].builtin && o.value !== 'Normal') delete d.styles[o.value]; D.stylesChanged(); return E.sel; }); fillSel(mine, d); }), ui.button('&Rename...', async () => { const o = mine.selectedOptions[0]; if (!o) return; const n = await ui.prompt('New name:', D.styleDisplayName(d.styles[o.value]), 'Rename'); if (!n) return; E.edit('Rename Style', () => { D.touchKey(d, 'styles'); d.styles[o.value].name = n; return E.sel; }); fillSel(mine, d); })), h('div', { class: 'col', style: 'flex:1' }, lblOther, theirs, ui.button('Open &File...', async () => { const f = (await L.pickFiles('.docx,.dotx,.docm,.dotm'))[0]; if (!f) return; other = (await L.docx.read(await L.readAsArrayBuffer(f))).doc; otherName = f.name; lblOther.textContent = 'In ' + otherName + ':'; fillSel(theirs, other); }))), buttons: [{ label: 'Close', primary: true }] });
  };

  /* ================= AutoCorrect ================= */
  G.autocorrect = function (tab) {
    const o = A().opts;
    const ac = L.autocorrect;
    const chk = (k, l, def) => ui.check(l, o[k] !== undefined ? o[k] !== false : def !== false, (v) => { o[k] = v; A().saveOpts(); });
    const from = h('input', { type: 'text', id: 'ac-from', style: 'width:110px' }), to = h('input', { type: 'text', id: 'ac-to', style: 'flex:1;min-width:0' });
    const list = G.listBox([], null, { height: 150, onChange: (v) => { from.value = v; to.value = ac.list().get(v) || ''; } });
    const fill = () => { const m = ac.list(); list.setItems(Array.from(m.keys()).sort((a, b) => a.localeCompare(b)).map((k) => ({ value: k, label: `${k}  →  ${m.get(k)}` }))); };
    fill();
    from.addEventListener('input', () => { const m = ac.list(); if (m.has(from.value)) to.value = m.get(from.value); });
    const exc = () => { const ta = h('textarea', { rows: 10, style: 'width:100%' }); ta.value = ac.exceptions().join('\n'); ui.dialog({ title: 'AutoCorrect Exceptions', width: 300, body: h('div', { class: 'col' }, h('div', { text: "Don't capitalize after:" }), ta), buttons: [{ label: 'OK', primary: true, onClick: () => ac.setExceptions(ta.value.split(/\s+/).filter(Boolean)) }, { label: 'Cancel' }] }); };
    const atName = h('input', { type: 'text', id: 'at-name', style: 'width:100%' });
    const atList = G.listBox(L.store.get('autoText', []).map((x) => x.name), null, { height: 120, onChange: (v) => { const e = L.store.get('autoText', []).find((x) => x.name === v); atName.value = v; atPv.textContent = e ? e.text : ''; } });
    const atPv = h('div', { class: 'tp-box', style: 'min-height:40px;white-space:pre-wrap' });
    const tabs = ui.tabs([
      { label: '&AutoCorrect', body: h('div', { class: 'col' }, chk('acButtons', 'Show Auto&Correct Options buttons'), chk('twoCaps', 'Correct TWo INitial &CApitals'), chk('capSentence', 'Capitalize first letter of &sentences'), chk('capCells', 'Capitalize first letter of table &cells'), chk('capDays', 'Capitalize &names of days'), chk('capsLock', 'Correct accidental usage of cAPS &LOCK key'), ui.button('&Exceptions...', exc, { class: 'btn small' }), chk('autocorrect', 'Replace &text as you type'), G.row(G.f('&Replace:', from), G.f('&With:', to)), list, h('div', { class: 'tp-btns' }, ui.button('&Add', () => { if (!from.value) return; ac.setEntry(from.value, to.value); fill(); }, { class: 'btn small' }), ui.button('&Delete', () => { if (!from.value) return; ac.setEntry(from.value, null); fill(); from.value = ''; to.value = ''; }, { class: 'btn small' }))) },
      { label: 'AutoFormat As You T&ype', body: h('div', { class: 'col' }, ui.group('Replace as you type', chk('smartQuotes', '"&Straight quotes" with “smart quotes”'), chk('ordinals', '&Ordinals (1st) with superscript'), chk('fractions', '&Fractions (1/2) with fraction character (½)'), chk('dashes', '&Hyphens (--) with dash (—)'), chk('autoFormat', '*&Bold* and _italic_ with real formatting'), chk('autoUrl', '&Internet and network paths with hyperlinks')), ui.group('Apply as you type', chk('autoBullets', 'Automatic b&ulleted lists'), chk('autoBorders', 'Border &lines'), chk('autoTables', '&Tables'), chk('autoNumbers', 'Automatic &numbered lists'), chk('autoHeadings', 'Built-in Heading st&yles', false))) },
      { label: 'Auto&Text', body: h('div', { class: 'col' }, chk('autoComplete', 'Show &AutoComplete suggestions'), h('label', { for: 'at-name', text: 'Enter AutoText entries here:' }), atName, atList, h('div', { class: 'cd-cap', text: 'Preview' }), atPv, h('div', { class: 'tp-btns' }, ui.button('&Add', () => { const t = E.sel && !E.collapsed() ? O.textRange(...E.range(), '\n') : ''; if (!t) { ui.msg('Select the text you want to store as AutoText first.', { icon: 'info' }); return; } const n = atName.value.trim() || t.slice(0, 32); ac.addAutoText(n, t); atList.setItems(L.store.get('autoText', []).map((x) => x.name)); }, { class: 'btn small' }), ui.button('&Delete', () => { ac.deleteAutoText(atName.value); atList.setItems(L.store.get('autoText', []).map((x) => x.name)); atPv.textContent = ''; }, { class: 'btn small' }), ui.button('&Insert', () => { const e = L.store.get('autoText', []).find((x) => x.name === atName.value); if (e) E.insertTextAt(e.text, 'AutoText'); }, { class: 'btn small' }))) },
    ], tab === 'autotext' ? 2 : tab === 'format' ? 1 : 0);
    ui.dialog({ title: 'AutoCorrect: English (U.S.)', width: 460, body: tabs, buttons: [{ label: 'OK', primary: true }, { label: 'Cancel' }] });
  };

  /* ================= Customize ================= */
  G.customize = function () {
    const tbs = ['standard', 'formatting', 'drawing', 'tables', 'reviewing', 'picture', 'outlining', 'headerFooter', 'mailMerge', 'wordart'].map((k) => ui.check(ui.cmds['tb_' + k].label, A().tbShown(k), () => A().toggleToolbar(k)));
    const cmds = Object.values(ui.cmds).filter((c) => c.label && !c.id.startsWith('tb_')).sort((a, b) => ui.stripAmp(a.label).localeCompare(ui.stripAmp(b.label)));
    const cmdList = G.listBox(cmds.map((c) => ({ value: c.id, label: ui.stripAmp(c.menuLabel || c.label) + (c.key ? '   ' + c.key : '') })), null, { height: 220 });
    const keys = h('table', { class: 'kbd-table' });
    for (const [k, id] of Object.entries(A().SHORTCUTS).sort((a, b) => a[1].localeCompare(b[1]))) { const c = ui.cmds[id]; if (c) keys.appendChild(h('tr', null, h('td', { text: k }), h('td', { text: ui.stripAmp(c.menuLabel || c.label) }))); }
    const o = A().opts;
    ui.dialog({ title: 'Customize', width: 460, body: ui.tabs([
      { label: 'Tool&bars', body: h('div', { class: 'col' }, ...tbs) },
      { label: '&Commands', body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'All Quire commands. Double-click one to run it.' }), cmdList), onShow: () => {} },
      { label: '&Options', body: h('div', { class: 'col' }, ui.check('Show Standard and Formatting toolbars on two rows', true, null), ui.check('Always show full menus', true, null), ui.check('Show Scree&nTips on toolbars', o.screenTips !== false, (v) => { o.screenTips = v; A().saveOpts(); document.body.classList.toggle('notips', !v); }), ui.check('Show shortcut &keys in ScreenTips', true, null)) },
      { label: '&Keyboard', body: h('div', { style: 'max-height:300px;overflow:auto' }, keys) },
    ]), buttons: [{ label: 'Close', primary: true }] });
    cmdList.querySelector('.listbox').addEventListener('dblclick', () => { const id = cmdList.get(); if (id) ui.exec(id); });
  };

  /* ================= Options ================= */
  G.options = function (tab) {
    const o = A().opts;
    const d = doc();
    const pwInp = h('input', { type: 'password', id: 'op-pw', value: d.openPassword || '', style: 'width:220px', autocomplete: 'new-password' });
    const chk = (label, get, set) => ui.check(label, get(), set);
    const opt = (k, l, def) => chk(l, () => (o[k] !== undefined ? !!o[k] : !!def), (v) => { o[k] = v; });
    const lyo = (k, l) => chk(l, () => !!LY.opts[k], (v) => { LY.opts[k] = v; rerender = true; });
    let rerender = false;
    const fieldShade = ui.select([['never', 'Never'], ['always', 'Always'], ['selected', 'When selected']], o.fieldShading || 'selected', (v) => { o.fieldShading = v; }, { id: 'op-fs' });
    const units = ui.select([['in', 'Inches'], ['cm', 'Centimeters'], ['mm', 'Millimeters'], ['pt', 'Points'], ['pi', 'Picas']], o.unit || 'in', (v) => { o.unit = v; }, { id: 'op-unit' });
    const recentN = G.num(4, { min: 0, max: 9, id: 'op-rec' });
    const uname = h('input', { type: 'text', id: 'op-un', value: o.userName || '', style: 'width:100%' });
    const uinit = h('input', { type: 'text', id: 'op-ui', value: o.userInitials || '', style: 'width:80px' });
    const uaddr = h('textarea', { rows: 3, id: 'op-ua', style: 'width:100%' });
    uaddr.value = L.store.get('userAddress', '');
    const insColor = ui.select([['author', 'By author'], ['red', 'Red'], ['blue', 'Blue'], ['green', 'Green']], o.insColor || 'author', (v) => { o.insColor = v; }, { id: 'op-ic' });
    const balloons = ui.select([['always', 'Always'], ['never', 'Never'], ['comments', 'Only for comments/formatting']], LY.opts.balloons === false ? 'never' : 'comments', (v) => { LY.opts.balloons = v !== 'never'; rerender = true; }, { id: 'op-bal' });
    const compat = ui.select([['11', 'Microsoft Word 2003'], ['12', 'Microsoft Word 2007'], ['14', 'Microsoft Word 2010'], ['15', 'Microsoft Word 2013 and later']], String(d.settings.compat || 11), (v) => { d.settings.compat = +v; d.dirty = true; rerender = true; }, { id: 'op-compat' });
    const autoRec = G.num(o.autoSaveMin || 10, { min: 1, max: 120, id: 'op-asr' });
    const tabs = ui.tabs([
      { label: 'View', body: h('div', { class: 'col' }, ui.group('Show', h('div', { class: 'grid2' }, opt('startupPane', 'Startup Task Pane', true), opt('highlightShow', 'Highlight', true), lyo('showBookmarks', 'Bookmarks'), opt('statusBar', 'Status bar', true), opt('screenTips', 'ScreenTips', true), lyo('fieldCodes', 'Field codes'), opt('smartTags', 'Smart tags', false), opt('animText', 'Animated text', false)), G.f('Field shading:', fieldShade)), ui.group('Formatting marks', h('div', { class: 'grid2' }, lyo('showMarks', 'All'), lyo('showHidden', 'Hidden text'))), ui.group('Print and Web Layout options', h('div', { class: 'grid2' }, lyo('textBoundaries', 'Text boundaries'), opt('vruler', 'Vertical ruler (Print view only)', true), opt('whiteSpace', 'White space between pages (Print view only)', true)))) },
      { label: 'General', body: h('div', { class: 'col' }, opt('bgRepag', 'Background repagination', true), opt('blueBg', 'Blue background, white text', false), opt('sound', 'Provide feedback with sound', false), G.row(opt('recentOn', 'Recently used file list:', true), recentN, h('span', { text: 'entries' })), G.f('Measurement units:', units, 'wide')) },
      { label: 'Edit', body: h('div', { class: 'col' }, ui.group('Editing options', opt('typingReplaces', 'Typing replaces selection', true), opt('dragDrop', 'Drag-and-drop text editing', true), opt('ctrlClick', 'Use CTRL + Click to follow hyperlink', true), opt('smartPara', 'Use smart paragraph selection', true), chk('Overtype mode', () => E.overtype, (v) => { E.overtype = v; }), opt('smartCut', 'Smart cut and paste', true)), ui.group('Cut and paste options', opt('pasteOptions', 'Show Paste Options buttons', true))) },
      { label: 'Print', body: h('div', { class: 'col' }, ui.group('Printing options', opt('updateFieldsOnPrint', 'Update fields', false), opt('printBg', 'Background colors and images', true)), ui.group('Include with document', opt('printProps', 'Document properties', false), opt('printCodes', 'Field codes', false), opt('printHidden', 'Hidden text', false), opt('printDrawings', 'Drawing objects', true))) },
      { label: 'Save', body: h('div', { class: 'col' }, opt('promptProps', 'Prompt for document properties', false), G.row(opt('autoRecover', 'Save AutoRecover info every:', true), autoRec, h('span', { text: 'minutes' })), h('div', { class: 'tp-note', text: 'AutoRecover keeps a copy in this browser only; it is offered when Quire opens after an unexpected close.' }), G.f('Save Word files as:', ui.select([['docx', 'Word Document (*.docx)'], ['dotx', 'Word Template (*.dotx)']], 'docx', null, { id: 'op-fmt' }), 'wide')) },
      { label: 'Security', body: h('div', { class: 'col' }, ui.group('Privacy options', opt('removePersonal', 'Remove personal information from file properties on save', false), opt('warnTracked', 'Warn before printing, saving or sending a file that contains tracked changes or comments', false)), ui.group('File encryption options for this document', h('div', { class: 'col' }, h('label', { for: 'op-pw', text: 'Password to open:' }), pwInp, h('div', { class: 'tp-note', text: 'Saved .docx files are encrypted with AES-256, as Word 2010 and later do. Passwords are case-sensitive; if you forget it, the document cannot be recovered.' }))), h('div', { class: 'tp-note', text: 'To restrict editing instead, use Tools ▸ Protect Document.' })) },
      { label: 'Spelling & Grammar', body: h('div', { class: 'col' }, ui.group('Spelling', opt('spell', 'Check spelling as you type', true), chk('Hide spelling and grammar errors in this document', () => !!(L.spell && L.spell.hidden), (v) => { if (L.spell) L.spell.hidden = v; }), opt('ignoreUpper', 'Ignore words in UPPERCASE', true), opt('ignoreNum', 'Ignore words with numbers', true), opt('ignoreUrl', 'Ignore Internet and file addresses', true), ui.button('Custom Dictionaries...', () => { const ta = h('textarea', { rows: 12, style: 'width:100%' }); ta.value = L.store.get('customDict', []).join('\n'); ui.dialog({ title: 'CUSTOM.DIC', width: 300, body: ta, buttons: [{ label: 'OK', primary: true, onClick: () => { L.store.set('customDict', ta.value.split(/\s+/).filter(Boolean).map((x) => x.toLowerCase())); if (L.spell) L.spell.reloadCustom(); } }, { label: 'Cancel' }] }); }, { class: 'btn small' })), ui.group('Grammar', opt('grammarAsType', 'Check grammar as you type', true), opt('grammar', 'Check grammar with spelling', true), ui.button('Readability Statistics...', () => { const r = L.spell.readability(); ui.msg(`Words: ${r.words}\nCharacters: ${r.chars}\nParagraphs: ${r.paras}\nSentences: ${r.sentences}\n\nSentences per Paragraph: ${r.sPerPara}\nWords per Sentence: ${r.wPerS}\nCharacters per Word: ${r.cPerW}\n\nPassive Sentences: ${r.passive}%\nFlesch Reading Ease: ${r.ease}\nFlesch-Kincaid Grade Level: ${r.grade}`, { icon: 'info', title: 'Readability Statistics' }); }, { class: 'btn small' })), ui.button('Recheck Document', () => { if (L.spell) { L.spell.ignoredOnce.clear(); L.spell.refreshSoon(0); } }, { class: 'btn small' })) },
      { label: 'Track Changes', body: h('div', { class: 'col' }, ui.group('Markup', G.f('Insertions color:', insColor, 'wide')), ui.group('Balloons', G.f('Use Balloons (Print and Web Layout):', balloons, 'wide'))) },
      { label: 'User Information', body: h('div', { class: 'col' }, G.f('Name:', uname, 'wide'), G.f('Initials:', uinit, 'wide'), h('label', { for: 'op-ua', text: 'Mailing address:' }), uaddr) },
      { label: 'Compatibility', body: h('div', { class: 'col' }, G.f('Lay out this document as if created in:', compat, 'wide'), h('div', { class: 'tp-note', text: 'Word 2003 layout positions tables by their cell margins and keeps legacy spacing rules; later versions align tables with the text.' })) },
    ], { view: 0, general: 1, edit: 2, print: 3, save: 4, security: 5, spelling: 6, track: 7, user: 8, compat: 9 }[tab] || 0);
    tabs.classList.add('options-tabs');
    ui.dialog({ title: 'Options', width: 560, body: tabs, buttons: [{ label: 'OK', primary: true, onClick: () => {
      const newPw = pwInp.value;
      if (newPw !== (d.openPassword || '')) {
        if (!newPw) { d.openPassword = null; A().status('The password to open was removed; the next save is not encrypted.'); }
        else {
          /* Word asks for the password a second time before using it */
          const again = h('input', { type: 'password', id: 'op-pw2', style: 'width:100%' });
          let val = null;
          ui.dialog({ title: 'Confirm Password', width: 340, body: h('div', { class: 'col' }, h('label', { for: 'op-pw2', text: 'Reenter password to open:' }), again, h('div', { class: 'tp-note', text: 'Caution: If you lose or forget the password, it cannot be recovered.' })), buttons: [{ label: 'OK', primary: true, onClick: () => { val = again.value; } }, { label: 'Cancel' }] }).done.then(() => {
            if (val == null) return;
            if (val !== newPw) { ui.msg('The password confirmation does not match.', { icon: 'warn' }); return; }
            d.openPassword = newPw;
            A().dirty && A().dirty();
            A().status('The document will be encrypted with this password when you save it.');
          });
          setTimeout(() => again.focus(), 0);
        }
      }
      o.userName = uname.value.trim() || 'Quire User';
      o.userInitials = uinit.value.trim();
      o.autoSaveMin = autoRec.get();
      L.store.set('userAddress', uaddr.value);
      if (o.blueBg) document.body.classList.add('bluebg'); else document.body.classList.remove('bluebg');
      A().saveOpts();
      A().applyOpts();
      if (rerender) { LY.root.classList.toggle('marks', !!LY.opts.showMarks); LY.render(); E.restoreDom(true); }
    } }, { label: 'Cancel' }] });
  };

  /* ================= Properties ================= */
  G.properties = function () {
    const d = doc();
    const p = d.props;
    const st = D.stats(D.allParas(d));
    const fld = (k, l) => { const inp = h('input', { type: 'text', id: 'pr-' + k, value: p[k] || '', style: 'width:100%' }); return { k, inp, el: G.f(l, inp, 'wide') }; };
    const fields = [fld('title', '&Title:'), fld('subject', '&Subject:'), fld('creator', '&Author:'), fld('manager', '&Manager:'), fld('company', 'C&ompany:'), fld('category', 'Cat&egory:'), fld('keywords', '&Keywords:')];
    const comments = h('textarea', { rows: 3, id: 'pr-desc', style: 'width:100%' });
    comments.value = p.description || '';
    const hb = fld('hyperlinkBase', '&Hyperlink base:');
    const fmt = (s) => (s ? new Date(s).toLocaleString() : '');
    const statT = h('table', { class: 'wc-table' });
    for (const [k, v] of [['Pages:', LY.pages.length || 1], ['Paragraphs:', st.paras], ['Lines:', A().docStats().lines], ['Words:', st.words], ['Characters:', st.charsNoSp], ['Characters (with spaces):', st.chars]]) statT.appendChild(h('tr', null, h('td', { text: k }), h('td', { text: Number(v).toLocaleString() })));
    const heads = h('div', { class: 'tp-box', style: 'max-height:200px;overflow:auto' });
    D.reindex(d);
    for (const q of d.main.paras) { const lv = D.headingLevel(d, q); if (lv) heads.appendChild(h('div', { text: D.plainText(q).trim(), style: `padding-left:${(lv - 1) * 12}px` })); }
    if (!heads.childNodes.length) heads.textContent = '(no headings)';
    /* custom */
    const cName = h('input', { type: 'text', id: 'pr-cn', list: 'pr-cnl', style: 'width:100%' });
    const cnl = h('datalist', { id: 'pr-cnl' }, ...['Checked by', 'Client', 'Date completed', 'Department', 'Destination', 'Disposition', 'Division', 'Document number', 'Editor', 'Forward to', 'Group', 'Language', 'Mailstop', 'Matter', 'Office', 'Owner', 'Project', 'Publisher', 'Purpose', 'Received from', 'Recorded by', 'Recorded date', 'Reference', 'Source', 'Status', 'Telephone number', 'Typist'].map((x) => h('option', { value: x })));
    const cType = ui.select([['text', 'Text'], ['date', 'Date'], ['number', 'Number'], ['bool', 'Yes or no']], 'text', null, { id: 'pr-ct' });
    const cVal = h('input', { type: 'text', id: 'pr-cv', style: 'width:100%' });
    const custom = Object.assign({}, d.custom || {});
    const cList = G.listBox([], null, { height: 110, onChange: (v) => { const c = custom[v]; cName.value = v; cVal.value = c && c.v != null ? c.v : c; cType.value = (c && c.type) || 'text'; } });
    const fillC = () => cList.setItems(Object.keys(custom).map((k) => ({ value: k, label: `${k}: ${custom[k] && custom[k].v != null ? custom[k].v : custom[k]}` })));
    fillC();
    const tabs = ui.tabs([
      { label: 'General', body: h('div', { class: 'col' }, h('div', { class: 'row', style: 'align-items:center' }, h('span', { html: L.icons.get('new') }), h('b', { text: A().fileName + '.docx' })), h('table', { class: 'wc-table' }, ...[['Type:', 'Microsoft Word Document'], ['Created:', fmt(p.created)], ['Modified:', fmt(p.modified)], ['Last printed:', fmt(p.printed)]].map(([a, b]) => h('tr', null, h('td', { text: a }), h('td', { text: b }))))) },
      { label: 'Summary', body: h('div', { class: 'col' }, ...fields.map((f) => f.el), h('label', { for: 'pr-desc', html: 'Co<u>m</u>ments:' }), comments, hb.el, h('div', { class: 'tp-note', text: 'Template: ' + (p.template || 'Normal') })) },
      { label: 'Statistics', body: h('div', { class: 'col' }, h('table', { class: 'wc-table' }, ...[['Created:', fmt(p.created)], ['Modified:', fmt(p.modified)], ['Last saved by:', p.lastModifiedBy || ''], ['Revision number:', String(p.revision || 1)], ['Total editing time:', (p.totalTime || 0) + ' Minutes']].map(([a, b]) => h('tr', null, h('td', { text: a }), h('td', { text: b })))), h('div', { class: 'cd-cap', text: 'Statistics:' }), statT) },
      { label: 'Contents', body: h('div', { class: 'col' }, h('div', { text: 'Document contents:' }), heads) },
      { label: 'Custom', body: h('div', { class: 'col' }, cnl, G.f('&Name:', cName, 'wide'), G.f('&Type:', cType, 'wide'), G.f('&Value:', cVal, 'wide'), h('div', { class: 'tp-btns' }, ui.button('&Add', () => { if (!cName.value.trim()) return; custom[cName.value.trim()] = { type: cType.value, v: cType.value === 'number' ? +cVal.value : cType.value === 'bool' ? /^(y|yes|true|1)$/i.test(cVal.value) : cVal.value }; fillC(); }, { class: 'btn small' }), ui.button('&Delete', () => { delete custom[cName.value.trim()]; fillC(); }, { class: 'btn small' })), h('label', { text: 'Properties:' }), cList) },
    ], 1);
    ui.dialog({ title: A().fileName + ' Properties', width: 420, body: tabs, buttons: [{ label: 'OK', primary: true, onClick: () => {
      E.edit('Properties', () => { D.touchKey(d, 'props'); D.touchKey(d, 'custom'); for (const f of fields) d.props[f.k] = f.inp.value; d.props.description = comments.value; d.props.hyperlinkBase = hb.inp.value || undefined; d.custom = custom; return E.sel; });
      A().updateTitle();
    } }, { label: 'Cancel' }] });
  };

  /* ================= Zoom ================= */
  G.zoom = function () {
    const cur = Math.round(LY.zoom * 100);
    const r = G.radios('zm', [['200', '&200%'], ['100', '&100%'], ['75', '&75%'], ['pw', '&Page width'], ['tw', '&Text width'], ['wp', '&Whole page'], ['many', '&Many pages']], [200, 100, 75].includes(cur) ? String(cur) : 'pct', (v) => { const n = { 200: 200, 100: 100, 75: 75 }[v]; if (n) pct.set(n); });
    const pct = ui.spin({ value: cur, min: 10, max: 500, step: 10, dec: 0, unit: '%', id: 'zm-pct' });
    const pv = h('div', { class: 'preview-box', style: 'width:150px;height:110px', text: 'AaBbCcDdEeXxYyZz' });
    ui.dialog({ title: 'Zoom', width: 360, body: G.row(ui.group('Zoom to', r), h('div', { class: 'col' }, G.f('Percen&t:', pct), h('div', { class: 'cd-cap', text: 'Preview' }), pv)), buttons: [{ label: 'OK', primary: true, onClick: () => { const v = r.get(); if (v === 'pw') A().zoomTo('Page Width'); else if (v === 'tw') A().zoomTo('Text Width'); else if (v === 'wp') A().zoomTo('Whole Page'); else if (v === 'many') A().zoomTo('Two Pages'); else A().zoomTo(pct.get() + '%'); } }, { label: 'Cancel' }] });
  };

  /* ================= Print (to PDF) ================= */
  G.print = function () {
    const n = LY.pages.length || 1;
    const range = G.radios('pr-r', [['all', '&All'], ['current', 'Curr&ent page'], ['selection', '&Selection'], ['pages', 'Pa&ges:']], 'all');
    const pagesIn = h('input', { type: 'text', id: 'pr-pages', style: 'width:150px' });
    pagesIn.addEventListener('focus', () => range.set('pages'));
    const what = ui.select([['doc', 'Document'], ['markup', 'Document showing markup'], ['props', 'Document properties'], ['list', 'List of markup'], ['styles', 'Styles'], ['autotext', 'AutoText entries'], ['keys', 'Key assignments']], 'doc', null, { id: 'pr-what' });
    const which = ui.select([['all', 'All pages in range'], ['odd', 'Odd pages'], ['even', 'Even pages']], 'all', null, { id: 'pr-which' });
    const perSheet = ui.select([[1, '1 page'], [2, '2 pages'], [4, '4 pages'], [6, '6 pages'], [8, '8 pages'], [16, '16 pages']], 1, null, { id: 'pr-pps' });
    const copies = G.num(1, { min: 1, max: 99, id: 'pr-copies' });
    ui.dialog({
      title: 'Print', width: 470,
      body: h('div', { class: 'col' }, ui.group('Printer', G.f('&Name:', ui.select([['pdf', 'Save as PDF (Quire)']], 'pdf', null, { id: 'pr-printer' }), 'wide'), h('div', { class: 'tp-note', text: 'Status: Ready — the pages are saved as a PDF that you can print from any viewer.' })), G.row(ui.group('Page range', range, pagesIn, h('div', { class: 'tp-note', text: 'Enter page numbers and/or page ranges separated by commas. For example, 1,3,5–12' })), ui.group('Copies', G.f('Number of &copies:', copies))), G.row(G.f('Print &what:', what, 'wide'), G.f('P&rint:', which)), ui.group('Zoom', G.f('Pages per s&heet:', perSheet, 'wide'))),
      buttons: [{ label: '&Options...', onClick: () => { G.options('print'); return false; } }, { label: 'OK', primary: true, onClick: () => { doPrint(); } }, { label: 'Cancel' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
    async function doPrint() {
      const w = what.value;
      if (w !== 'doc' && w !== 'markup') { printList(w); return; }
      let pages = null;
      const rv = range.get();
      if (rv === 'current') pages = [Math.max(0, LY.pageOf(E.sel.f)) + 1];
      else if (rv === 'selection' && E.sel && !E.collapsed()) { const [a, b] = E.range(); const p0 = LY.pageOf(a), p1 = LY.pageOf(b); pages = []; for (let i = p0; i <= p1; i++) pages.push(i + 1); }
      else if (rv === 'pages') { pages = []; for (const part of pagesIn.value.split(',')) { const m = /^\s*(\d+)\s*(?:[-–]\s*(\d+))?\s*$/.exec(part); if (m) { const a = +m[1], b = m[2] ? +m[2] : a; for (let i = a; i <= b; i++) if (i >= 1 && i <= n) pages.push(i); } } }
      if (which.value !== 'all') pages = (pages || Array.from({ length: n }, (_, i) => i + 1)).filter((i) => (which.value === 'odd' ? i % 2 === 1 : i % 2 === 0));
      const prevMarkup = LY.opts.markup;
      if (w === 'doc' && prevMarkup !== 'final') { LY.opts.markup = 'final'; LY.render(); }
      try {
        const blob = await A().buildPDF(pages);
        L.ui.busy(false);
        doc().props.printed = new Date().toISOString();
        await L.saveFile(`${A().fileName}.pdf`, blob);
      } catch (e) { L.ui.busy(false); ui.msg('The PDF could not be created: ' + (e.message || e), { icon: 'error' }); }
      finally { if (LY.opts.markup !== prevMarkup) { LY.opts.markup = prevMarkup; LY.render(); E.restoreDom(true); } }
    }
    function printList(kind) {
      const d = doc();
      const lines = [];
      if (kind === 'props') { for (const [k, v] of Object.entries(d.props)) if (v) lines.push(`${k}: ${v}`); }
      else if (kind === 'styles') { for (const k of Object.keys(d.styles).sort()) { const s = d.styles[k]; lines.push(`${D.styleDisplayName(s)} (${s.type})${s.basedOn ? ' — based on ' + s.basedOn : ''}`); } }
      else if (kind === 'autotext') { for (const e of L.autocorrect.autoTextEntries()) lines.push(`${e.cat}: ${e.name}`); }
      else if (kind === 'keys') { for (const [k, id] of Object.entries(A().SHORTCUTS)) if (ui.cmds[id]) lines.push(`${k}\t${ui.stripAmp(ui.cmds[id].label)}`); }
      else if (kind === 'list') { for (const r of L.review.list(d.main)) lines.push(`${r.type} — ${r.stamp.author || ''}: ${r.text}`); for (const k in d.comments) lines.push(`Comment ${d.comments[k].author}: ${d.comments[k].blocks.map((b) => (b.t === 'p' ? D.plainText(b) : '')).join(' ')}`); }
      const nd = D.newDoc();
      nd.main.blocks = [D.para([D.text(A().fileName + ' — ' + what.selectedOptions[0].text, { b: true, sz: 14 })])].concat(lines.map((l) => D.para([D.text(l)])));
      const keep = D.doc;
      D.doc = nd;
      LY.render();
      A().buildPDF().then(async (blob) => { L.ui.busy(false); D.doc = keep; LY.render(); E.restoreDom(true); await L.saveFile(`${A().fileName} (${kind}).pdf`, blob); }).catch(() => { D.doc = keep; LY.render(); L.ui.busy(false); });
    }
  };

  /* ================= Save As ================= */
  G.saveAs = function (cb) {
    const name = h('input', { type: 'text', id: 'sa-name', value: A().fileName, style: 'width:100%' });
    const type = ui.select([['docx', 'Word Document (*.docx)'], ['dotx', 'Word Template (*.dotx)'], ['html', 'Single File Web Page (*.htm)'], ['rtf', 'Rich Text Format (*.rtf)'], ['txt', 'Plain Text (*.txt)'], ['pdf', 'PDF (*.pdf)']], A().fileType === 'dotx' ? 'dotx' : 'docx', null, { id: 'sa-type' });
    const note = h('div', { class: 'tp-note', text: 'Your browser asks where to save the file, or saves it to the Downloads folder.' });
    ui.dialog({ title: 'Save As', width: 430, focus: name, body: h('div', { class: 'col' }, G.f('File &name:', name, 'wide'), G.f('Save as &type:', type, 'wide'), note), buttons: [{ label: '&Save', primary: true, onClick: () => { const n = name.value.trim().replace(/\.(docx|dotx|htm|html|rtf|txt|pdf)$/i, '') || 'Document'; cb(n, type.value); } }, { label: 'Cancel' }] });
  };
  G.pasteSpecial = function () {
    const lb = G.listBox([{ value: 'rich', label: 'Formatted Text (RTF)' }, { value: 'html', label: 'HTML Format' }, { value: 'text', label: 'Unformatted Text' }, { value: 'utext', label: 'Unformatted Unicode Text' }], 'rich', { height: 110 });
    const clip = E.clip;
    ui.dialog({ title: 'Paste Special', width: 380, body: h('div', { class: 'col' }, h('div', { text: clip ? 'Source: Quire document' : 'Source: (the Quire clipboard is empty — copy something first, or press Ctrl+V)' }), G.radios('ps-m', [['paste', '&Paste:'], ['link', 'Paste &link:']], 'paste'), lb, h('div', { class: 'tp-note', text: 'Formatted text keeps fonts and formatting; unformatted text takes the formatting of the destination.' })), buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (!clip) return;
      const v = lb.get();
      if (v === 'text' || v === 'utext') E.pasteText(clip.text);
      else E.pasteBlocks(clip.blocks, { keepFormat: v === 'rich' || v === 'html' });
    } }, { label: 'Cancel' }] });
  };
  G.about = function () {
    ui.dialog({ title: 'About Quire', width: 400, body: h('div', { class: 'col', style: 'align-items:center;text-align:center;gap:8px' }, h('span', { html: L.icons.app ? L.icons.app(48) : '' }), h('b', { text: 'Quire 2003 Web Edition', style: 'font-size:15px;color:#0b2e7c' }), h('div', { text: 'A word processor in the style of Microsoft Office Word 2003 that reads and writes Office Open XML (.docx) documents.' }), h('div', { class: 'tp-note', text: 'Microsoft, Word and Office are trademarks of Microsoft Corporation. Quire is an independent program and is not affiliated with Microsoft.' }), h('div', { class: 'tp-note', text: `Version ${L.VERSION || '1.0'}` })), buttons: [{ label: 'OK', primary: true }] });
  };
  G.newFromTemplate = function () {
    const cats = L.templates.categories();
    let pick = 'blank';
    const pv = h('div', { class: 'preview-box', style: 'width:150px;height:190px' });
    const drawPv = (id) => { L.clear(pv); const page = h('div', { class: 'ps-page', style: 'width:110px;height:140px' }); for (let i = 0; i < (id === 'blank' ? 0 : 14); i++) page.appendChild(h('i', { style: `top:${14 + i * 8}px;${i % 5 === 4 ? 'right:40px' : ''}` })); pv.appendChild(page); };
    const tabs = ui.tabs(cats.map((c) => {
      const grid = h('div', { class: 'preset-grid' });
      const items = c === 'General' ? [{ id: 'blank', name: 'Blank Document' }, { id: 'web', name: 'Web Page' }, { id: 'mail', name: 'E-mail Message' }] : L.templates.list().filter((t) => t.cat === c);
      for (const t of items) { const b = h('button', { type: 'button', class: t.id === pick ? 'on' : '' }, h('span', { html: L.icons.get('new', 32) || L.icons.get('new') }), h('span', { text: t.name })); b.addEventListener('click', () => { pick = t.id; L.$$('.preset-grid button', tabs).forEach((x) => x.classList.toggle('on', x === b)); drawPv(t.id); }); b.addEventListener('dblclick', () => ok.click()); grid.appendChild(b); }
      return { label: c, body: grid };
    }));
    drawPv('blank');
    const dlg = ui.dialog({ title: 'Templates', width: 520, body: G.row(h('div', { style: 'flex:1;min-width:240px' }, tabs), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), pv)), buttons: [{ label: 'OK', primary: true, onClick: () => { if (pick === 'blank') A().newDocument(); else if (pick === 'web') { A().newDocument().then(() => A().setView('web')); } else if (pick === 'mail') A().newDocument({ build: (d) => { d.main.blocks = [D.para([D.text('To: ')]), D.para([D.text('Subject: ')]), D.para()]; } }); else L.templates.create(pick); } }, { label: 'Cancel' }] });
    const ok = dlg.buttons[0];
  };
})();
