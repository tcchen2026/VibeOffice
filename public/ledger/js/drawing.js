/* Ledger — drawing objects on worksheets: pictures, charts, AutoShapes and text boxes.
 * Rendering into the grid's object layer, selection, move / resize, text editing, the drawing tools, Format Object. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, O = L.ops, LY = L.layout;
  const { h } = L;
  const ui = L.ui;
  const DR = (L.drawing = {});
  const G = () => L.grid;
  const wb = () => G().wb;
  const sh = () => G().sheet();
  const SVGNS = 'http://www.w3.org/2000/svg';

  /* ------------------------------------------------------------ media URLs */
  const urls = new Map(); /* media entry → object URL */
  function mediaURL(id, cb) {
    const m = wb().media && wb().media.get(id);
    if (!m) return null;
    if (urls.has(m)) return urls.get(m);
    const ext = String(m.ext || '').toLowerCase();
    if ((ext === 'emf' || ext === 'wmf') && L.metafile) {
      urls.set(m, '');
      L.metafile.toPNG(m.bytes, ext).then((b) => { if (b) { urls.set(m, URL.createObjectURL(b)); if (cb) cb(); } });
      return '';
    }
    const u = URL.createObjectURL(new Blob([m.bytes], { type: m.type || L.extToMime(ext) || 'image/png' }));
    urls.set(m, u);
    return u;
  }

  /* ------------------------------------------------------------ shape properties */
  /** fill / line / geometry / text of a shape, from our own properties or its DrawingML */
  function shapeProps(d) {
    if (d._props && d._props.src === d.xml) return d._props;
    const p = { src: d.xml, geom: d.geom || 'rect', fill: d.fill === undefined ? '#FFFFFF' : d.fill, line: d.line === undefined ? '#000000' : d.line, lineW: d.lineW || 0.75, text: d.text || '', font: d.font || {}, align: d.align || 'l', vAlign: d.vAlign || 't', paras: null };
    if (d.xml && !d.dirty && L.dml) {
      try {
        const el = L.xml.parse(d.xml);
        const X = L.dml;
        const tctx = { theme: { colors: themeMap() } };
        const spPr = X.kid(el, 'spPr') || X.desc(el, 'spPr');
        if (spPr) {
          const g = X.geometry(spPr); p.geom = g.geom; p.adj = g.adj; p.path = g.path;
          const f = X.fillIn(spPr, tctx);
          /* shapes without their own fill take the theme style's (fillRef) */
          const style = X.kid(el, 'style');
          if (f) p.fill = f.t === 'none' ? null : f.t === 'solid' ? f.c : f.t === 'grad' ? f.stops[0].c : '#FFFFFF';
          else if (style) { const fr = X.kid(style, 'fillRef'); const c = fr ? X.colorOf(fr, tctx) : null; p.fill = c ? c.c : '#4F81BD'; }
          const ln = X.kid(spPr, 'ln');
          if (ln) { const lf = X.fillIn(ln, tctx); p.line = lf ? (lf.t === 'none' ? null : lf.c || '#000000') : p.line; p.lineW = X.num(ln, 'w', 9525) / 12700; }
          else if (style) { const lr = X.kid(style, 'lnRef'); const c = lr ? X.colorOf(lr, tctx) : null; p.line = c ? c.c : '#385D8A'; }
          const xf = X.kid(spPr, 'xfrm'); if (xf) { p.rot = X.num(xf, 'rot', 0) / 60000; p.flipH = X.at(xf, 'flipH') === '1'; p.flipV = X.at(xf, 'flipV') === '1'; }
        }
        const tb = X.desc(el, 'txBody');
        if (tb) {
          const bp = X.kid(tb, 'bodyPr');
          if (bp) { const an = X.at(bp, 'anchor'); p.vAlign = an === 'ctr' ? 'ctr' : an === 'b' ? 'b' : 't'; }
          p.paras = X.kids(tb, 'p').map((pe) => {
            const ppr = X.kid(pe, 'pPr');
            const runs = X.kids(pe).filter((r) => r.localName === 'r' || r.localName === 'fld' || r.localName === 'br').map((r) => {
              if (r.localName === 'br') return { t: '\n' };
              const rp = X.kid(r, 'rPr');
              const run = { t: (X.kid(r, 't') || { textContent: '' }).textContent };
              if (rp) {
                if (X.at(rp, 'sz')) run.sz = +X.at(rp, 'sz') / 100;
                if (X.at(rp, 'b') === '1') run.b = true;
                if (X.at(rp, 'i') === '1') run.i = true;
                if (X.at(rp, 'u') && X.at(rp, 'u') !== 'none') run.u = true;
                const c = X.fillIn(rp, tctx); if (c && c.t === 'solid') run.color = c.c;
                const lat = X.kid(rp, 'latin'); if (lat && X.at(lat, 'typeface') && !/^\+/.test(X.at(lat, 'typeface'))) run.font = X.at(lat, 'typeface');
              }
              return run;
            });
            return { algn: (ppr && X.at(ppr, 'algn')) || 'l', runs };
          });
          if (!d.text) p.text = p.paras.map((q) => q.runs.map((r) => r.t).join('')).join('\n');
        }
      } catch (e) { /* draw as a plain rectangle */ }
    }
    d._props = p;
    return p;
  }
  const themeMap = () => { const t = wb().theme; const names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']; const o = {}; names.forEach((n, i) => { o[n] = (t && t.colors && t.colors[i]) || M.DEFAULT_THEME.colors[i]; }); return o; };

  /* ------------------------------------------------------------ element */
  DR.element = function (d, s) {
    const el = h('div', { class: 'gobj ' + d.kind, role: 'img', 'aria-label': d.name || d.kind });
    el._d = d;
    const content = h('div', { class: 'gobj-c' });
    el.appendChild(content);
    el.addEventListener('pointerdown', (e) => onDown(e, el));
    el.addEventListener('dblclick', (e) => { e.stopPropagation(); DR.activate(el._d); });
    el.addEventListener('contextmenu', (e) => { e.preventDefault(); e.stopPropagation(); if (G().selectedObject() !== el._d) G().selectObject(el._d); DR.contextMenu(e, el._d); });
    void s;
    return el;
  };
  DR.update = function (el, d, s, rc, selected) {
    el._d = d;
    const content = el.firstChild;
    const z = (s.view.zoom || 100) / 100;
    const w = Math.max(1, Math.round(rc.w)), hh = Math.max(1, Math.round(rc.h));
    if (d.kind === 'image') {
      const crop = d.crop;
      const key = 'img' + d.media + (crop ? JSON.stringify(crop) : '');
      if (el._key !== key) {
        el._key = key;
        content.textContent = '';
        const url = mediaURL(d.media, () => { el._key = null; G().syncObjects(); });
        if (url) {
          const img = h('img', { src: url, alt: d.descr || d.name || '', draggable: 'false' });
          if (crop && (crop.l || crop.t || crop.r || crop.b)) {
            const sx = 1 / Math.max(0.01, 1 - crop.l / 100 - crop.r / 100), sy = 1 / Math.max(0.01, 1 - crop.t / 100 - crop.b / 100);
            Object.assign(img.style, { position: 'absolute', width: sx * 100 + '%', height: sy * 100 + '%', left: -crop.l / 100 * sx * 100 + '%', top: -crop.t / 100 * sy * 100 + '%', maxWidth: 'none' });
          }
          content.appendChild(img);
        } else content.appendChild(h('div', { class: 'gobj-ph', text: (d.name || 'Picture') }));
      }
    } else if (d.kind === 'chart') {
      const key = `c${w}x${hh}@${z}#${wb().version}:${d.chart ? JSON.stringify([d.chart.kind, d.chart.title, d.chart.legend, d.chart.labels, d.chart.gridY, d.chart.series && d.chart.series.length]) : ''}`;
      if (el._key !== key) {
        el._key = key;
        content.textContent = '';
        if (d.chart) { try { content.appendChild(L.xchart.render(d.chart, wb(), s, w, hh, z)); } catch (e) { console.error(e); content.appendChild(h('div', { class: 'gobj-ph', text: d.name || 'Chart' })); } }
        else content.appendChild(h('div', { class: 'gobj-ph', text: (d.name || 'Chart') + '\n(this chart type is not displayed)' }));
      }
    } else {
      const p = shapeProps(d);
      const key = `s${w}x${hh}@${z}:${d.xml ? d.xml.length : ''}:${JSON.stringify([p.geom, p.fill, p.line, p.lineW, d.text, p.text, d.font, d.align])}`;
      if (el._key !== key && !el.classList.contains('editing')) {
        el._key = key;
        content.textContent = '';
        content.appendChild(shapeSVG(p, w, hh, z));
        const tx = textBox(d, p, z);
        if (tx) content.appendChild(tx);
      }
    }
    /* sizing handles */
    let hs = el.querySelector('.gobj-hs');
    if (selected && !hs) {
      hs = h('div', { class: 'gobj-hs' });
      for (const k of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) { const x = h('div', { class: 'gobj-h ' + k }); x.dataset.h = k; hs.appendChild(x); }
      el.appendChild(hs);
    } else if (!selected && hs) hs.remove();
  };
  function shapeSVG(p, w, hh, z) {
    const svg = document.createElementNS(SVGNS, 'svg');
    svg.setAttribute('width', w); svg.setAttribute('height', hh);
    svg.setAttribute('viewBox', `0 0 ${w} ${hh}`);
    svg.style.overflow = 'visible';
    const lw = p.line ? Math.max(1, (p.lineW || 0.75) * (96 / 72) * z) : 0;
    const geomSh = { geom: p.geom, adj: p.adj, path: p.path, w, h: hh, flipH: p.flipH, flipV: p.flipV };
    if (L.geom && L.geom.isLine(p.geom)) {
      const path = document.createElementNS(SVGNS, 'path');
      path.setAttribute('d', L.geom.linePath(geomSh));
      path.setAttribute('stroke', p.line || '#000'); path.setAttribute('stroke-width', lw || 1); path.setAttribute('fill', 'none');
      svg.appendChild(path);
      return svg;
    }
    const parts = L.geom ? L.geom.parts(geomSh, w, hh) : [{ d: `M0,0 H${w} V${hh} H0 Z`, fill: 'norm', stroke: true }];
    for (const part of parts) {
      const path = document.createElementNS(SVGNS, 'path');
      path.setAttribute('d', part.d);
      let fill = p.fill || 'none';
      if (part.fill === 'none') fill = 'none';
      else if (p.fill && /^darken/.test(part.fill || '')) fill = L.color.applyMods ? L.color.applyMods(p.fill, [{ name: 'shade', val: 60000 }]) : p.fill;
      else if (p.fill && /^lighten/.test(part.fill || '')) fill = L.color.applyMods ? L.color.applyMods(p.fill, [{ name: 'tint', val: 60000 }]) : p.fill;
      path.setAttribute('fill', fill);
      if (part.rule) path.setAttribute('fill-rule', part.rule);
      if (part.stroke !== false && p.line) { path.setAttribute('stroke', p.line); path.setAttribute('stroke-width', lw); } else path.setAttribute('stroke', 'none');
      svg.appendChild(path);
    }
    if (p.flipH || p.flipV) svg.style.transform = `scale(${p.flipH ? -1 : 1},${p.flipV ? -1 : 1})`;
    return svg;
  }
  function textBox(d, p, z) {
    const text = d.dirty || !p.paras ? (d.text != null ? d.text : p.text) : null;
    if (!text && !(p.paras && p.paras.some((q) => q.runs.length))) return null;
    const box = h('div', { class: 'gobj-t v-' + (p.vAlign || 't') });
    const f = d.font || {};
    box.style.fontSize = ((f.sz || 10) * z) + 'pt';
    box.style.fontFamily = LY.fontStack(f.name || 'Arial');
    if (f.b) box.style.fontWeight = 'bold';
    if (f.i) box.style.fontStyle = 'italic';
    if (f.color) box.style.color = f.color;
    if (text != null) {
      box.style.textAlign = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }[d.align || 'l'] || 'left';
      box.textContent = text;
      return box;
    }
    for (const para of p.paras) {
      const pe = h('div', { style: 'text-align:' + ({ l: 'left', ctr: 'center', r: 'right', just: 'justify' }[para.algn] || 'left') });
      if (!para.runs.length) pe.appendChild(h('br'));
      for (const r of para.runs) {
        if (r.t === '\n') { pe.appendChild(h('br')); continue; }
        const sp = h('span', { text: r.t });
        if (r.sz) sp.style.fontSize = r.sz * z + 'pt';
        if (r.b) sp.style.fontWeight = 'bold';
        if (r.i) sp.style.fontStyle = 'italic';
        if (r.u) sp.style.textDecoration = 'underline';
        if (r.color) sp.style.color = r.color;
        if (r.font) sp.style.fontFamily = LY.fontStack(r.font);
        pe.appendChild(sp);
      }
      box.appendChild(pe);
    }
    return box;
  }

  /* ------------------------------------------------------------ interaction */
  const guard = () => G().guardProtect && G().guardProtect('objects');
  function replace(d, patch, label) {
    const s = sh();
    const nd = Object.assign({}, d, patch);
    delete nd._props;
    O.tx(wb(), label || 'Move Object', () => O.setDrawings(s, s.drawings.map((x) => (x === d ? nd : x))));
    G().selectObject(nd);
    G().syncObjects(true);
    return nd;
  }
  DR.replace = replace;
  /** canvas rect (CSS px in the grid layer) → anchor of the same kind as d's */
  function anchorFor(d, x, y, w, hh) {
    const s = sh();
    const a = d.anchor || { type: 'two' };
    const z = (s.view.zoom || 100) / 100;
    if (a.type === 'abs') return Object.assign({}, a, { x: Math.max(0, (x - G().frame().hdrW) / z * 0.75), y: Math.max(0, (y - G().frame().hdrH) / z * 0.75), w: w / z * 0.75, h: hh / z * 0.75 });
    const two = G().rectToAnchor(s, x, y, w, hh, { editAs: a.editAs });
    if (a.type === 'one') return { type: 'one', from: two.from, w: w / z * 0.75, h: hh / z * 0.75 };
    return two;
  }
  function onDown(e, el) {
    if (e.button === 2) return;
    e.stopPropagation();
    /* a shape whose text is being edited: the press places the caret and a drag selects text */
    if (el.classList.contains('editing')) return;
    e.preventDefault();
    if (L.editor && L.editor.active && !L.editor.commit()) return;
    const d = el._d;
    if (G().selectedObject() !== d) G().selectObject(d);
    G().focus();
    const handle = e.target.closest('.gobj-h');
    if ((d.locked || (sh().protection && !d.unlocked)) && guard()) return;
    const x0 = e.clientX, y0 = e.clientY;
    const r0 = { x: el.offsetLeft, y: el.offsetTop, w: el.offsetWidth, h: el.offsetHeight };
    const ratio = r0.w / Math.max(1, r0.h);
    let moved = false, cur = Object.assign({}, r0);
    const ghost = h('div', { class: 'gobj-ghost' });
    const mv = (ev) => {
      const dx = ev.clientX - x0, dy = ev.clientY - y0;
      if (!moved && Math.abs(dx) + Math.abs(dy) < 3) return;
      if (!moved) { moved = true; el.parentNode.appendChild(ghost); }
      cur = Object.assign({}, r0);
      if (!handle) { cur.x += dx; cur.y += dy; if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) cur.y = r0.y; else cur.x = r0.x; } }
      else {
        const k = handle.dataset.h;
        if (/e/.test(k)) cur.w = Math.max(4, r0.w + dx);
        if (/s/.test(k)) cur.h = Math.max(4, r0.h + dy);
        if (/w/.test(k)) { cur.w = Math.max(4, r0.w - dx); cur.x = r0.x + r0.w - cur.w; }
        if (/n/.test(k)) { cur.h = Math.max(4, r0.h - dy); cur.y = r0.y + r0.h - cur.h; }
        if ((ev.shiftKey || d.kind === 'image') && k.length === 2) { /* corners keep the aspect ratio (pictures always) */
          const wv = cur.w, hv = cur.h;
          if (wv / hv > ratio) cur.w = hv * ratio; else cur.h = wv / ratio;
          if (/w/.test(k)) cur.x = r0.x + r0.w - cur.w;
          if (/n/.test(k)) cur.y = r0.y + r0.h - cur.h;
        }
      }
      Object.assign(ghost.style, { left: cur.x + 'px', top: cur.y + 'px', width: cur.w + 'px', height: cur.h + 'px' });
    };
    const up = () => {
      document.removeEventListener('pointermove', mv);
      document.removeEventListener('pointerup', up);
      ghost.remove();
      if (!moved) return;
      const fr = G().frame();
      if (d.anchor && d.anchor.type !== 'abs') { cur.x = Math.max(fr.hdrW, cur.x); cur.y = Math.max(fr.hdrH, cur.y); }
      replace(d, { anchor: anchorFor(d, cur.x, cur.y, cur.w, cur.h) }, handle ? 'Resize Object' : 'Move Object');
    };
    document.addEventListener('pointermove', mv);
    document.addEventListener('pointerup', up);
  }
  /** keys while an object is selected; true when handled */
  DR.key = function (e, d) {
    const k = e.key, ctrl = e.ctrlKey || e.metaKey;
    if (k === 'Escape') { G().deselectObject(); return true; }
    if (k === 'Delete' || k === 'Backspace') { if (guard()) return true; L.app.deleteObject(d); return true; }
    if (ctrl && /^[cxv]$/i.test(k)) {
      if (k.toLowerCase() === 'v') { if (DR.clipboard) DR.pasteObject(); else { G().deselectObject(); return false; } return true; }
      DR.copy(d, k.toLowerCase() === 'x');
      return true;
    }
    if (/^Arrow/.test(k)) {
      if (guard()) return true;
      const el = Array.from(L.$('#glayer').children).find((x) => x._d === d);
      if (!el) return true;
      const step = ctrl ? 1 : 8;
      const dx = k === 'ArrowLeft' ? -step : k === 'ArrowRight' ? step : 0, dy = k === 'ArrowUp' ? -step : k === 'ArrowDown' ? step : 0;
      replace(d, { anchor: anchorFor(d, el.offsetLeft + dx, el.offsetTop + dy, el.offsetWidth, el.offsetHeight) });
      return true;
    }
    if (k === 'Enter' || k === 'F2') { DR.activate(d); return true; }
    if (k === 'Tab') { const list = sh().drawings.filter((x) => !x.hidden); const i = list.indexOf(d); G().selectObject(list[(i + (e.shiftKey ? -1 : 1) + list.length) % list.length]); return true; }
    if (k.length === 1 && !ctrl && !e.altKey && d.kind === 'shape') { DR.editText(d, k); return true; }
    if (ctrl && /^[zy]$/i.test(k)) return false;
    return !ctrl; /* swallow typing so it does not go into cells */
  };
  /** double-click / Enter: edit a chart's options or a shape's text */
  DR.activate = function (d) {
    if (d.kind === 'chart') { if (d.chart) L.chartWizard.open({ edit: d }); return; }
    if (d.kind === 'shape') { DR.editText(d); return; }
    if (d.kind === 'image') DR.format(d);
  };
  DR.editText = function (d, firstKey) {
    if (guard()) return;
    const el = Array.from(L.$('#glayer').children).find((x) => x._d === d);
    if (!el) return;
    const p = shapeProps(d);
    el.classList.add('editing');
    let box = el.querySelector('.gobj-t');
    if (!box) { box = h('div', { class: 'gobj-t v-' + (p.vAlign || 't') }); el.firstChild.appendChild(box); }
    const z = (sh().view.zoom || 100) / 100;
    box.style.fontSize = (((d.font && d.font.sz) || 10) * z) + 'pt';
    box.textContent = d.text != null && (d.dirty || !p.paras) ? d.text : p.text || '';
    if (firstKey) box.textContent = firstKey;
    box.contentEditable = 'true';
    box.spellcheck = false;
    box.focus();
    const range = document.createRange(); range.selectNodeContents(box); range.collapse(false); const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
    const finish = () => {
      box.removeEventListener('blur', finish);
      box.contentEditable = 'false';
      el.classList.remove('editing');
      const text = box.innerText.replace(/\n$/, '');
      if (text !== (p.text || '')) replace(d, { text, dirty: true, geom: p.geom, fill: p.fill, line: p.line, lineW: p.lineW, font: d.font || { name: 'Arial', sz: 10 } }, 'Typing');
      else { el._key = null; G().syncObjects(); }
      G().focus();
    };
    box.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') { e.preventDefault(); box.blur(); } });
    box.addEventListener('blur', finish);
  };

  /* ------------------------------------------------------------ clipboard */
  DR.clipboard = null;
  DR.copy = function (d, cut) {
    DR.clipboard = { d: JSON.parse(JSON.stringify(Object.assign({}, d, { _props: undefined }), (k, v) => (v instanceof Uint8Array ? undefined : v))), media: d.media && wb().media ? wb().media.get(d.media) : null, sheet: sh() };
    if (L.clip && L.clip.clearMarquee) L.clip.clearMarquee();
    if (cut) L.app.deleteObject(d);
  };
  DR.pasteObject = function () {
    const c = DR.clipboard;
    if (!c) return;
    const s = sh();
    const d = L.sheetObjects.copy(c.d);
    if (c.media && d.media) {
      let id = d.media;
      if (!wb().media.get(id) || wb().media.get(id) !== c.media) { id = 'img' + (wb().media.size + 1) + '_' + Date.now().toString(36); wb().media.set(id, c.media); }
      d.media = id;
      if (d.objectKeep?.values.media === c.d.media) d.objectKeep.values.media = id;
    }
    d.id = Math.max(d.objectKeep ? 1024 : 1, ...s.drawings.map((x) => x.id || 0)) + 1;
    const sel = G().sel();
    const a = d.anchor;
    if (a && a.type !== 'abs') {
      const dr = sel.r - a.from.r, dc = sel.c - a.from.c;
      const same = c.sheet === s && G().selectedObject();
      a.from = Object.assign({}, a.from, same ? { r: a.from.r + 1, c: a.from.c + 1 } : { r: sel.r, c: sel.c });
      if (a.to) a.to = Object.assign({}, a.to, same ? { r: a.to.r + 1, c: a.to.c + 1 } : { r: a.to.r + dr, c: a.to.c + dc });
    }
    if (d.chart) d.chart.dirty = true;
    O.tx(wb(), 'Paste', () => O.setDrawings(s, s.drawings.concat([d])));
    G().selectObject(d);
    G().syncObjects(true);
  };

  /* ------------------------------------------------------------ order, menus */
  DR.order = function (where) {
    const d = G().selectedObject();
    if (!d) return;
    const s = sh();
    const rest = s.drawings.filter((x) => x !== d);
    const list = where === 'front' ? rest.concat([d]) : where === 'back' ? [d].concat(rest) : (() => { const i = s.drawings.indexOf(d); const l = s.drawings.slice(); const j = where === 'forward' ? Math.min(l.length - 1, i + 1) : Math.max(0, i - 1); l.splice(i, 1); l.splice(j, 0, d); return l; })();
    O.tx(wb(), 'Order', () => O.setDrawings(s, list));
    G().syncObjects(true);
  };
  DR.contextMenu = function (e, d) {
    const items = ['cut', 'copy', 'paste', '-'];
    if (d.kind === 'chart') items.push({ label: 'Chart &Type...', run: () => L.chartWizard.typeMenu({ x: e.clientX, y: e.clientY }) }, { label: '&Chart Options...', run: () => L.chartWizard.open({ edit: d }) }, '-');
    if (d.kind === 'shape') items.push({ label: 'Edit Te&xt', run: () => DR.editText(d) }, '-');
    items.push({ label: 'O&rder', sub: [{ label: 'Bring to Fr&ont', run: () => DR.order('front') }, { label: 'Send to Bac&k', run: () => DR.order('back') }, { label: 'Bring &Forward', run: () => DR.order('forward') }, { label: 'Send &Backward', run: () => DR.order('backward') }] }, { label: 'Assign Macro...', disabled: true }, '-',
      { label: d.kind === 'image' ? 'F&ormat Picture...' : d.kind === 'chart' ? 'F&ormat Chart Area...' : d.kind === 'shape' && shapeProps(d).geom === 'rect' && d.txBox ? 'F&ormat Text Box...' : 'F&ormat AutoShape...', run: () => DR.format(d) },
      { label: '&Delete', run: () => L.app.deleteObject(d) });
    ui.openMenu(items, { x: e.clientX, y: e.clientY });
  };

  /* ------------------------------------------------------------ Format Object */
  DR.format = function (d) {
    if (!d) return;
    const s = sh();
    const p = d.kind === 'shape' ? shapeProps(d) : null;
    const el = Array.from(L.$('#glayer').children).find((x) => x._d === d);
    const z = (s.view.zoom || 100) / 100;
    const curW = el ? el.offsetWidth / z * 0.75 : 72, curH = el ? el.offsetHeight / z * 0.75 : 72;
    const st = { fill: p ? p.fill : d.kind === 'chart' && d.chart && d.chart.bg ? d.chart.bg.color : null, line: p ? p.line : null, lineW: p ? p.lineW : 0.75, w: curW, h: curH, editAs: (d.anchor && d.anchor.editAs) || 'twoCell', alt: d.descr || '', locked: !d.unlocked, print: !d.noPrint };
    const tabs = [];
    const swatch = (key, mode) => {
      const b = h('button', { type: 'button', class: 'swatch-btn', style: `background:${st[key] || 'transparent'}` }, h('span', { class: 'dd-arrow' }));
      b.addEventListener('click', () => ui.colorMenu(b, { mode, grid: true }, (c) => { if (c.none) st[key] = null; else if (c.auto) st[key] = mode === 'line' ? '#000000' : '#FFFFFF'; else if (typeof c === 'string') st[key] = c; b.style.background = st[key] || 'transparent'; }));
      return b;
    };
    if (d.kind !== 'image') {
      const weight = ui.spin({ value: st.lineW, min: 0.25, max: 20, step: 0.25, unit: ' pt', onChange: (v) => { st.lineW = v; } });
      tabs.push({ label: 'Colors and Lines', body: h('div', { class: 'col' }, ui.group('Fill', ui.field('&Color:', swatch('fill', 'fill'))), d.kind === 'shape' ? ui.group('Line', ui.field('C&olor:', swatch('line', 'line')), ui.field('&Weight:', weight)) : h('span')) });
    }
    const wS = ui.spin({ value: st.w / 72, min: 0.01, max: 50, step: 0.1, unit: '"', onChange: (v) => { const r = st.h / st.w; st.w = v * 72; if (lock.input.checked) { st.h = st.w * r; hS.set(st.h / 72); } } });
    const hS = ui.spin({ value: st.h / 72, min: 0.01, max: 50, step: 0.1, unit: '"', onChange: (v) => { const r = st.w / st.h; st.h = v * 72; if (lock.input.checked) { st.w = st.h * r; wS.set(st.w / 72); } } });
    const lock = ui.check('Lock &aspect ratio', d.kind === 'image', null);
    tabs.push({ label: 'Size', body: h('div', { class: 'col' }, ui.group('Size and rotate', ui.field('H&eight:', hS), ui.field('Wi&dth:', wS)), lock) });
    const props = h('div', { class: 'col' }, ui.group('Object positioning',
      ui.radio('fo-pos', '&Move and size with cells', st.editAs === 'twoCell', () => { st.editAs = 'twoCell'; }),
      ui.radio('fo-pos', 'Move &but don\'t size with cells', st.editAs === 'oneCell', () => { st.editAs = 'oneCell'; }),
      ui.radio('fo-pos', '&Don\'t move or size with cells', st.editAs === 'absolute', () => { st.editAs = 'absolute'; })),
    ui.check('&Print object', st.print, (v) => { st.print = v; }), ui.check('&Locked', st.locked, (v) => { st.locked = v; }));
    if (d.anchor && d.anchor.type !== 'abs') tabs.push({ label: 'Properties', body: props });
    const alt = h('textarea', { rows: '3', style: 'width:100%' });
    alt.value = st.alt;
    tabs.push({ label: 'Web', body: h('div', { class: 'col' }, h('label', { text: 'Alternative text:' }), alt) });
    if (d.kind === 'chart' && d.chart) tabs.unshift({ label: 'Chart', body: h('div', { class: 'col' }, ui.button('Chart &Options...', () => { dlg.close(null); L.chartWizard.open({ edit: d }); }), ui.button('Chart &Type...', (e) => { dlg.close(null); L.chartWizard.typeMenu({ x: e.clientX, y: e.clientY }); })) });
    const title = d.kind === 'image' ? 'Format Picture' : d.kind === 'chart' ? 'Format Chart Area' : d.txBox ? 'Format Text Box' : 'Format AutoShape';
    const dlg = ui.dialog({ title, body: ui.tabs(tabs), width: 420, buttons: [{ label: 'OK', primary: true, onClick: () => {
      const patch = { descr: alt.value || undefined, unlocked: st.locked ? undefined : true, noPrint: st.print ? undefined : true };
      if (d.kind === 'shape') Object.assign(patch, { fill: st.fill, line: st.line, lineW: st.lineW, geom: p.geom, text: d.text != null ? d.text : p.text, dirty: true, font: d.font || { name: 'Arial', sz: 10 } });
      if (d.kind === 'chart' && d.chart && st.fill !== (d.chart.bg && d.chart.bg.color)) patch.chart = Object.assign({}, d.chart, { bg: st.fill ? { t: 'solid', color: st.fill } : { t: 'none' }, dirty: true });
      const a = JSON.parse(JSON.stringify(d.anchor));
      if (Math.abs(st.w - curW) > 0.5 || Math.abs(st.h - curH) > 0.5) {
        if (a.type === 'two' && el) { const nw = st.w / 0.75 * z, nh = st.h / 0.75 * z; Object.assign(a, anchorFor(d, el.offsetLeft, el.offsetTop, nw, nh)); }
        else { a.w = st.w; a.h = st.h; }
      }
      if (a.type === 'two') a.editAs = st.editAs;
      patch.anchor = a;
      replace(d, patch, title.replace('Format ', 'Format '));
    } }, { label: 'Cancel' }] });
  };

  /* ------------------------------------------------------------ drawing tools */
  let tool = null;
  DR.startTool = function (kind) {
    if (guard()) return;
    DR.cancelTool();
    const wrap = L.$('#gridwrap');
    const ov = h('div', { class: 'draw-overlay', 'aria-label': 'Drag to draw' });
    wrap.appendChild(ov);
    tool = { kind, ov };
    L.app.status('Drag in the worksheet to insert the ' + (kind === 'textbox' ? 'text box' : 'AutoShape') + '. Press Esc to cancel.');
    const esc = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); DR.cancelTool(); } };
    document.addEventListener('keydown', esc, true);
    tool.esc = esc;
    ov.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      const r = wrap.getBoundingClientRect();
      const x0 = e.clientX - r.left, y0 = e.clientY - r.top;
      const band = h('div', { class: 'gobj-ghost' });
      ov.appendChild(band);
      let cur = { x: x0, y: y0, w: 0, h: 0 };
      const mv = (ev) => {
        const x1 = ev.clientX - r.left, y1 = ev.clientY - r.top;
        cur = { x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0) };
        if (ev.shiftKey) { const m = Math.max(cur.w, cur.h); cur.w = cur.h = m; }
        Object.assign(band.style, { left: cur.x + 'px', top: cur.y + 'px', width: cur.w + 'px', height: cur.h + 'px' });
      };
      const up = () => {
        document.removeEventListener('pointermove', mv);
        document.removeEventListener('pointerup', up);
        const k = tool.kind;
        DR.cancelTool();
        const z = (sh().view.zoom || 100) / 100;
        if (cur.w < 6 && cur.h < 6) { cur.w = (k === 'textbox' ? 96 : 72) * z * (96 / 72); cur.h = (k === 'textbox' ? 48 : 72) * z * (96 / 72); }
        DR.insertShape(k, cur);
      };
      document.addEventListener('pointermove', mv);
      document.addEventListener('pointerup', up);
    });
  };
  DR.cancelTool = function () {
    if (!tool) return;
    tool.ov.remove();
    document.removeEventListener('keydown', tool.esc, true);
    tool = null;
    L.app.status('');
    G().focus();
  };
  DR.insertShape = function (kind, rect) {
    const s = sh();
    const fr = G().frame();
    rect.x = Math.max(fr.hdrW, rect.x); rect.y = Math.max(fr.hdrH, rect.y);
    const id = Math.max(1, ...s.drawings.map((x) => x.id || 0)) + 1;
    const isText = kind === 'textbox';
    const d = { kind: 'shape', id, name: (isText ? 'Text Box ' : (kind === 'ellipse' ? 'Oval ' : kind === 'line' || kind === 'arrow' ? 'Line ' : 'Rectangle ')) + id, geom: isText ? 'rect' : kind === 'arrow' ? 'straightConnector1' : kind === 'line' ? 'line' : kind, txBox: isText, fill: kind === 'line' || kind === 'arrow' ? null : '#FFFFFF', line: '#000000', lineW: 0.75, text: '', font: { name: 'Arial', sz: 10 }, dirty: true };
    d.anchor = G().rectToAnchor(s, rect.x, rect.y, rect.w, rect.h, { editAs: 'twoCell' });
    O.tx(wb(), 'Insert ' + (isText ? 'Text Box' : 'AutoShape'), () => O.setDrawings(s, s.drawings.concat([d])));
    G().selectObject(d);
    G().syncObjects(true);
    if (isText) setTimeout(() => DR.editText(d), 0);
    return d;
  };
  /** AutoShapes menu (Drawing toolbar) */
  DR.autoShapesMenu = function (at) {
    const groups = L.geom ? L.geom.MENU : { 'Basic Shapes': ['rect', 'ellipse'] };
    ui.openMenu([{ label: '&Lines', sub: [{ label: 'Line', run: () => DR.startTool('line') }, { label: 'Arrow', run: () => DR.startTool('arrow') }] }].concat(Object.keys(groups).map((g) => ({ label: g, sub: [{ custom: (close) => {
      const grid = h('div', { class: 'shape-grid' });
      for (const k of groups[g]) { const b = h('button', { type: 'button', html: L.geom ? L.geom.preview(k, 20) : k, 'data-tip': k, 'aria-label': k }); b.addEventListener('click', () => { close(); DR.startTool(k); }); grid.appendChild(b); }
      return grid;
    } }] }))), at);
  };
})(typeof window !== 'undefined' ? window : globalThis);
