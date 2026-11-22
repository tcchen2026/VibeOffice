/* VibeOffice — core utilities, shared by Quire, Ledger and Lectern.
 * Shared namespace, DOM helpers, units, colors, events, file I/O, media store, PDF writer.
 * Every other script attaches itself to window.L. Each page names its app on its <html> (no inline script:
 * the pages' Content-Security-Policy):
 *   <html lang="en" data-app="quire" data-app-name="Quire">
 * APP_ID keeps each app's saved settings apart ('quire.…', L.store); APP_NAME and APP appear in messages.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const named = root.document?.documentElement?.dataset || {};
  L.VERSION = root.VO?.VERSION || 'dev';   // one version for the suite: common/suite.js
  L.APP_ID = L.APP_ID || named.app || 'vibeoffice';
  L.APP_NAME = L.APP_NAME || named.appName || 'VibeOffice';
  L.APP = L.APP || L.APP_NAME;

  /* ---------- DOM ---------- */
  /* the text h() puts in the UI is shown in the UI's language (common/i18n.js): text, these attributes and string
     children whose whole text has a translation */
  const SHOWN = { title: 1, placeholder: 1, 'aria-label': 1, 'data-tip': 1 };
  const tr = (v, ...args) => (typeof v !== 'string' ? v : L.t ? L.t(v, ...args) : v.replace(/(\S)\|[A-Za-z][A-Za-z -]*$/, '$1').replace(/\{(\d+)\}/g, (m, i) => (args[i] != null ? args[i] : m)));
  L.h = function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    /* translate: 'no' marks data (a font's name, a slide's title): shown as it is */
    const data = attrs && attrs.translate === 'no';
    if (attrs) {
      for (const k in attrs) {
        let v = attrs[k];
        if (v == null || v === false) continue;
        if (SHOWN[k] && !data) v = tr(v);
        if (k === 'class') el.className = v;
        else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
        else if (k === 'style') el.style.cssText = v;
        else if (k === 'html') el.innerHTML = v;
        else if (k === 'text') el.textContent = data ? v : tr(v);
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
        else if (k === 'dataset') Object.assign(el.dataset, v);
        else if (v === true) el.setAttribute(k, '');
        else el.setAttribute(k, v);
      }
    }
    appendKids(el, kids, data);
    return el;
  };
  function appendKids(el, kids, data) {
    for (const k of kids) {
      if (k == null || k === false) continue;
      if (Array.isArray(k)) appendKids(el, k, data);
      else if (k instanceof Node) el.appendChild(k);
      else el.appendChild(document.createTextNode(typeof k === 'string' && !data ? tr(k) : String(k)));
    }
  }
  L.$ = (s, r) => (r || document).querySelector(s);
  L.$$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  L.svgNS = 'http://www.w3.org/2000/svg';
  L.s = function s(tag, attrs, ...kids) {
    const el = document.createElementNS(L.svgNS, tag);
    if (attrs) for (const k in attrs) {
      const v = attrs[k];
      if (v == null || v === false) continue;
      if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else el.setAttribute(k, v);
    }
    for (const k of kids.flat()) if (k != null) el.appendChild(k instanceof Node ? k : document.createTextNode(String(k)));
    return el;
  };
  L.clear = (el) => { while (el.firstChild) el.removeChild(el.firstChild); return el; };

  /* ---------- misc ---------- */
  let idc = 0;
  L.uid = (p) => (p || 'i') + Date.now().toString(36).slice(-4) + (idc++).toString(36) + Math.random().toString(36).slice(2, 6);
  L.clone = (o) => (o == null ? o : JSON.parse(JSON.stringify(o)));
  L.clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  L.round = (v, d) => { const m = Math.pow(10, d == null ? 2 : d); return Math.round(v * m) / m; };
  L.debounce = (fn, ms) => { let t; return function (...a) { clearTimeout(t); t = setTimeout(() => fn.apply(this, a), ms); }; };
  L.rafThrottle = (fn) => { let q = false, args; return function (...a) { args = a; if (q) return; q = true; requestAnimationFrame(() => { q = false; fn.apply(this, args); }); }; };
  L.esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  /* text for XML content or an attribute; tabs and line breaks as character references, which an attribute would
     otherwise turn into spaces (a form field's help text, a picture's alt text) */
  L.xesc = (s) => String(s == null ? '' : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
    .replace(/[&<>"\t\n\r]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\t': '&#9;', '\n': '&#10;', '\r': '&#13;' }[c]));
  L.isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
  /** iOS or Android (an iPad says it is a Mac, but has a touch screen) */
  L.isMobile = typeof navigator !== 'undefined' && (/Android|iPhone|iPad|iPod/.test(navigator.userAgent || '') || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1));
  L.deepMerge = function deepMerge(a, b) {
    if (b == null) return a;
    if (a == null || typeof a !== 'object' || typeof b !== 'object' || Array.isArray(b)) return L.clone(b);
    const out = Object.assign({}, a);
    for (const k in b) {
      if (b[k] === undefined) continue;
      const bv = b[k];
      if (bv && typeof bv === 'object' && !Array.isArray(bv) && out[k] && typeof out[k] === 'object' && !Array.isArray(out[k])) out[k] = deepMerge(out[k], bv);
      else out[k] = bv && typeof bv === 'object' ? L.clone(bv) : bv;
    }
    return out;
  };
  L.equal = (a, b) => JSON.stringify(a) === JSON.stringify(b);

  /* ---------- units ----------
   * The model stores geometry and font sizes in points.
   * 1 pt = 12700 EMU, 72 pt = 1 inch, 96 CSS px = 1 inch.
   */
  L.EMU = 12700;
  L.pt2emu = (pt) => Math.round((pt || 0) * 12700);
  L.emu2pt = (e) => (Number(e) || 0) / 12700;
  L.PX_PER_PT = 96 / 72;
  L.fmtIn = (pt) => (L.round(pt / 72, 2)) + '"';

  /* ---------- event bus ---------- */
  const handlers = {};
  L.bus = {
    on(ev, fn) { (handlers[ev] = handlers[ev] || []).push(fn); return () => this.off(ev, fn); },
    off(ev, fn) { const a = handlers[ev]; if (a) { const i = a.indexOf(fn); if (i >= 0) a.splice(i, 1); } },
    emit(ev, ...args) { const a = handlers[ev]; if (a) for (const f of a.slice()) { try { f(...args); } catch (e) { console.error(e); } } },
  };

  /* ---------- colors ---------- */
  const C = (L.color = {});
  C.hexToRgb = (hex) => {
    let h = String(hex || '#000').replace('#', '').trim();
    if (h.length === 3) h = h.split('').map((c) => c + c).join('');
    const n = parseInt(h.slice(0, 6), 16) || 0;
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  };
  C.rgbToHex = (r, g, b) => '#' + [r, g, b].map((v) => L.clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  /** RGB (0-255) to HSL (0-1): lightness the middle of the largest and smallest channel, saturation their difference
      over their sum (over 2 − their sum above half lightness), hue from the largest channel's sextant (Foley and
      van Dam, "RGB to HLS") */
  C.rgbToHsl = (r, g, b) => {
    const R = r / 255, G = g / 255, B = b / 255;
    const top = Math.max(R, G, B), bottom = Math.min(R, G, B), l = (top + bottom) / 2;
    if (top === bottom) return [0, 0, l];
    const span = top - bottom;
    const s = l > 0.5 ? span / (2 - top - bottom) : span / (top + bottom);
    let h;
    if (top === R) h = (G - B) / span + (G < B ? 6 : 0);
    else if (top === G) h = 2 + (B - R) / span;
    else h = 4 + (R - G) / span;
    return [h / 6, s, l];
  };
  /** HSL (0-1) to RGB (0-255): the algorithm of CSS Color Level 3, 4.2.4 ("hsl.to.rgb" and "hue.to.rgb") */
  C.hslToRgb = (h, s, l) => {
    if (s === 0) return [l * 255, l * 255, l * 255];
    const m2 = l < 0.5 ? l * (s + 1) : l + s - l * s, m1 = l * 2 - m2;
    const hueToRgb = (x) => {
      if (x < 0) x += 1;
      if (x > 1) x -= 1;
      if (x < 1 / 6) return m1 + (m2 - m1) * 6 * x;
      if (x < 1 / 2) return m2;
      if (x < 2 / 3) return m1 + (m2 - m1) * (2 / 3 - x) * 6;
      return m1;
    };
    return [hueToRgb(h + 1 / 3) * 255, hueToRgb(h) * 255, hueToRgb(h - 1 / 3) * 255];
  };
  C.mix = (a, b, t) => { const x = C.hexToRgb(a), y = C.hexToRgb(b); return C.rgbToHex(x[0] + (y[0] - x[0]) * t, x[1] + (y[1] - x[1]) * t, x[2] + (y[2] - x[2]) * t); };
  C.lighten = (hex, t) => C.mix(hex, '#FFFFFF', t);
  C.darken = (hex, t) => C.mix(hex, '#000000', t);
  C.luma = (hex) => { const [r, g, b] = C.hexToRgb(hex); return (0.299 * r + 0.587 * g + 0.114 * b) / 255; };
  C.rgba = (hex, a) => { const [r, g, b] = C.hexToRgb(hex); return a == null || a >= 1 ? C.rgbToHex(r, g, b) : `rgba(${r},${g},${b},${L.round(a, 3)})`; };
  C.gray = (hex) => { const v = Math.round(C.luma(hex) * 255); return C.rgbToHex(v, v, v); };
  /** a gradient's stops ({ p: 0–1, c: '#RRGGBB', a }, in order) as Office paints them. A gradient of two colours, two
      opaque stops at 0 and 100 % (or the first again at 100 % around one at 50 %), it blends in linear light along an
      S-curve, a normal distribution's (sigma a quarter of the way; erf as Abramowitz and Stegun 7.1.26): FF0000 to
      00B050 is #BC8038 halfway and #EE461C a quarter of the way, where a browser's straight sRGB blend has #805828 and
      #BF2C14 (PowerPoint 2016's PDF of LibreOffice's tdf105739 samples its curve: #BA813A and #EE4920). So stops are
      put in between (n to a gap) for a browser's straight runs to follow it. Any other gradient (more stops, stops
      inside, transparency) Office blends in sRGB, straight from stop to stop, as a browser does */
  const SIGMA = 0.25, erf = (x) => { const t = 1 / (1 + 0.3275911 * Math.abs(x)), y = 1 - ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x); return x < 0 ? -y : y; };
  const sigmoid = (t) => { const e0 = erf(-0.5 / SIGMA / Math.SQRT2); return (erf((t - 0.5) / SIGMA / Math.SQRT2) - e0) / (-2 * e0); };
  C.lightStops = (stops, n = 16) => {
    const opaque = stops.every((st) => st.a == null || st.a >= 1), [s0, s1, s2] = stops;
    const two = stops.length === 2 && s0.p === 0 && s1.p === 1, reflected = stops.length === 3 && s0.p === 0 && s1.p === 0.5 && s2.p === 1 && s0.c === s2.c;
    if (!opaque || !(two || reflected)) return stops;
    const out = [];
    stops.forEach((st, i) => {
      const prev = stops[i - 1];
      if (prev && prev.c !== st.c) {
        const x = C.hexToRgb(prev.c).map(toLin), y = C.hexToRgb(st.c).map(toLin);
        for (let k = 1; k < n; k++) { const f = sigmoid(k / n); out.push({ p: prev.p + ((st.p - prev.p) * k) / n, c: C.rgbToHex(...x.map((v, j) => toS(v + (y[j] - v) * f))) }); }
      }
      out.push(st);
    });
    return out;
  };
  /* OOXML colour modifiers (values in 1/1000 %) */
  const toLin = (c) => { c /= 255; return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
  const toS = (l) => { l = Math.min(1, Math.max(0, l)); return 255 * (l <= 0.0031308 ? l * 12.92 : 1.055 * Math.pow(l, 1 / 2.4) - 0.055); };
  C.applyMods = (hex, mods) => {
    if (!mods || !mods.length) return hex;
    let [r, g, b] = C.hexToRgb(hex);
    for (const m of mods) {
      const v = m.val / 100000;
      let [h, s, l] = C.rgbToHsl(r, g, b);
      switch (m.name) {
        case 'lumMod': l = l * v; [r, g, b] = C.hslToRgb(h, s, L.clamp(l, 0, 1)); break;
        case 'lumOff': l = l + v; [r, g, b] = C.hslToRgb(h, s, L.clamp(l, 0, 1)); break;
        case 'lum': [r, g, b] = C.hslToRgb(h, s, L.clamp(v, 0, 1)); break;
        case 'satMod': [r, g, b] = C.hslToRgb(h, L.clamp(s * v, 0, 1), l); break;
        case 'satOff': [r, g, b] = C.hslToRgb(h, L.clamp(s + v, 0, 1), l); break;
        case 'sat': [r, g, b] = C.hslToRgb(h, L.clamp(v, 0, 1), l); break;
        case 'hueMod': [r, g, b] = C.hslToRgb((h * v) % 1, s, l); break;
        case 'hueOff': [r, g, b] = C.hslToRgb((h + m.val / 21600000 + 1) % 1, s, l); break;
        case 'hue': [r, g, b] = C.hslToRgb((m.val / 21600000) % 1, s, l); break;
        /* DrawingML tint/shade work in linear light (scRGB), not on gamma-encoded values */
        case 'tint': [r, g, b] = [r, g, b].map((c) => toS(toLin(c) * v + (1 - v))); break;
        case 'shade': [r, g, b] = [r, g, b].map((c) => toS(toLin(c) * v)); break;
        case 'comp': [r, g, b] = C.hslToRgb((h + 0.5) % 1, s, l); break;
        case 'inv': r = 255 - r; g = 255 - g; b = 255 - b; break;
        case 'gray': { const y = 0.299 * r + 0.587 * g + 0.114 * b; r = g = b = y; break; }
        default: break;
      }
    }
    return C.rgbToHex(r, g, b);
  };
  C.PRESET = {
    black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', cyan: '#00FFFF', magenta: '#FF00FF',
    gray: '#808080', grey: '#808080', silver: '#C0C0C0', maroon: '#800000', navy: '#000080', olive: '#808000', purple: '#800080', teal: '#008080', lime: '#00FF00',
    orange: '#FFA500', darkBlue: '#00008B', darkRed: '#8B0000', darkGreen: '#006400', ltGray: '#D3D3D3', dkGray: '#A9A9A9', lightGray: '#D3D3D3', darkGray: '#A9A9A9',
    gold: '#FFD700', pink: '#FFC0CB', brown: '#A52A2A', indigo: '#4B0082', violet: '#EE82EE', tan: '#D2B48C', beige: '#F5F5DC', lavender: '#E6E6FA', coral: '#FF7F50',
  };
  /* Office 2003 standard colour grid (the 40-swatch palette) */
  C.STANDARD = [
    '#000000', '#993300', '#333300', '#003300', '#003366', '#000080', '#333399', '#333333',
    '#800000', '#FF6600', '#808000', '#008000', '#008080', '#0000FF', '#666699', '#808080',
    '#FF0000', '#FF9900', '#99CC00', '#339966', '#33CCCC', '#3366FF', '#800080', '#969696',
    '#FF00FF', '#FFCC00', '#FFFF00', '#00FF00', '#00FFFF', '#00CCFF', '#993366', '#C0C0C0',
    '#FF99CC', '#FFCC99', '#FFFF99', '#CCFFCC', '#CCFFFF', '#99CCFF', '#CC99FF', '#FFFFFF',
  ];

  /* ---------- the app's settings (never relied on): VO.settings (common/look.js), as JSON under '<APP_ID>.<k>' ---------- */
  const settings = () => root.VO?.settings;
  L.store = {
    get(k, d) { const v = settings()?.get(L.APP_ID + '.' + k); if (v == null) return d; try { return JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { return !!settings()?.set(L.APP_ID + '.' + k, JSON.stringify(v)); },
    del(k) { settings()?.remove(L.APP_ID + '.' + k); },
  };

  /* ---------- the look (common/looks.css) for chrome drawn on canvases ---------- */
  let lookCache = {};
  /** a look setting by name ('ruler-off' reads --ruler-off), or d where there is no page (Node) */
  L.lookColor = (name, d) => {
    if (!(name in lookCache)) lookCache[name] = typeof document === 'undefined' ? '' : getComputedStyle(document.documentElement).getPropertyValue('--' + name).trim();
    return lookCache[name] || d;
  };
  /** call fn after the look changes (VO.look.set, here or in another window), once the new values apply */
  L.onLook = (fn) => { if (typeof window !== 'undefined' && window.addEventListener) window.addEventListener('vo-look', () => { lookCache = {}; fn(); }); };
  if (typeof window !== 'undefined' && window.addEventListener) window.addEventListener('vo-look', () => { lookCache = {}; });

  /* ---------- file I/O ---------- */
  L.PDF_DPI = 300;
  const PDF_MAX_PIXELS = 40_000_000;
  /** Pixels per page unit, capped before allocating a canvas (including rounded pixel dimensions). */
  L.pdfScale = function (w, h, scale) {
    if (Math.round(w * scale) * Math.round(h * scale) <= PDF_MAX_PIXELS) return scale;
    const limit = Math.sqrt(PDF_MAX_PIXELS / (w * h));
    return Math.min(scale, Math.floor(w * limit) / w, Math.floor(h * limit) / h);
  };
  // The writer and text collectors are needed only when their export path is used.
  const pdfLoads = new Map();
  async function pdfScript(name, ready) {
    if (!ready()) {
      if (!pdfLoads.has(name)) {
        const script = document.createElement('script');
        script.src = (root.VO?.base || new URL('..', location.href).href) + 'common/' + name + '.js';
        const pending = new Promise((resolve, reject) => {
          script.onload = resolve;
          script.onerror = () => {
            pdfLoads.delete(name);
            script.remove();
            reject(new Error(L.t('The PDF exporter could not be loaded. Please try again.')));
          };
        });
        pdfLoads.set(name, pending);
        document.head.append(script);
        await pending;
      } else await pdfLoads.get(name);
    }
  }
  L.pdfModule = async function (text = false) {
    await pdfScript('pdf', () => L.pdf?.write);
    if (text) await pdfScript('pdftext', () => L.pdf?.collectDOM);
    return L.pdf;
  };
  L.buildPDF = async (pages, info) => (await L.pdfModule()).write(pages, info);
  /* ---------- saving to real files (the desktop app) ----------
     On the web a save is a download (L.saveFile). In the desktop app it is a real file (docs/desktop.md, Saving):
     Save writes the window's own file in place ('save'), Save As and exports what the user picks in the system's
     Save dialog (choose, then 'chosen'); the native side checks and writes (desktop/src/files.rs) and says what went
     wrong as a code, which is told here in plain words with the next step. One write at a time per window. */
  const native = () => (root.VO && root.VO.host.kind === 'desktop' ? root.VO.host : null);
  let pending = null;   // the Save dialog's pick, until written or given up
  let queue = Promise.resolve();
  const extOf = (name) => (/\.([^.]+)$/.exec(name || '') || ['', ''])[1].toLowerCase();
  const without = (name) => String(name || '').replace(/\.[^.]+$/, '');
  /* "{0}" is the file's name; every failure before the file is replaced says it hasn't changed */
  const N = (s) => s;   // UI text, translated where it is shown
  const SAYS = {
    changed: [N('“{0}” was changed by another program after you opened it. Replace it with this version, or save this version under another name?'), ['&Replace It', '&Save As...', 'Cancel'], ['replace', 'saveas']],
    missing: [N('“{0}” is no longer where it was: it was moved, renamed or deleted.'), ['&Save As...', 'Cancel'], ['saveas']],
    denied: [N('{1} can\'t save “{0}”: it is read-only, or you don\'t have permission to change it. Your file hasn\'t changed.'), ['&Save As...', 'Cancel'], ['saveas']],
    /* the Mac's sandbox: the app may need the user's leave to save in the file's folder */
    sandboxed: [N('{1} doesn\'t have permission to save “{0}” in its folder. Allow {1} into the folder, or save this version under another name? Your file hasn\'t changed.'), ['&Allow Access...', '&Save As...', 'Cancel'], ['allow', 'saveas']],
    busy: [N('“{0}” is open in another program, which is keeping it from being changed. Your file hasn\'t changed.'), ['&Try Again', '&Save As...', 'Cancel'], ['retry', 'saveas']],
    full: [N('There isn\'t enough space on the disk to save “{0}”. Your file hasn\'t changed.'), ['&Save As...', 'Cancel'], ['saveas']],
  };
  L.file = {
    /** whether saves are real files (the desktop app), not downloads */
    native: () => !!native(),
    /** the app's types for the Save dialog, in its order: [{ id, label, ext }] (each app sets them) */
    types: [],
    /** the type of a file name, by its extension */
    typeOf: (name) => (L.file.types.find((t) => t.ext.includes(extOf(name))) || {}).id || null,
    /** the desktop app: whether the window has a file of its own */
    bound: () => { const h = native(); return !!(h && h.file.bound); },
    /** whether File ▸ Versions... is there: the desktop app keeps the versions saves replace (not the lite edition) */
    versions: () => { const h = native(); return !!(h && !h.lite); },
    /** the document's name for the title bar: "(Read-Only)" after it when the window's file may not be changed */
    titled: (name) => { const h = native(); return h && h.file.bound && h.file.readOnly ? L.t('{0} (Read-Only)', name) : name; },
    /** the desktop app: whether Save writes the window's own file in place (it has one, and may change it) */
    inPlace: () => { const h = native(); return !!(h && h.file.bound && !h.file.readOnly); },
    /** the desktop app's Save dialog for a Save As ('saveAs') or an export ('export'): the app's types (or `ids`),
     *  `type` first, `name` (no extension) filled in. { name (no extension), file (with it), type }, or null */
    async choose(purpose, type, name, ids) {
      const types = L.file.types.filter((t) => !ids || ids.includes(t.id)).map((t) => ({ id: t.id, label: L.t(t.label) }));
      if (native().lite) return chooseByName(purpose, types, type, without(name) || 'Document');
      for (;;) {
        try {
          const c = await native().choose({ purpose, types, type, name: without(name) || 'Document', title: L.t('Save As') });
          pending = c;
          return c && { name: without(c.name), file: c.name, type: c.type };
        } catch (e) {
          if (e && e.code === 'open-elsewhere') { await L.ui.msg(L.t('“{0}” is open in another {1} window. Close it there, or choose another name.', e.file || name, L.APP), { icon: 'warn' }); continue; }
          await L.ui.msg(String(e && e.message || e), { icon: 'error' });
          return null;
        }
      }
    },
    /** the Save dialog's pick is not written after all (the compatibility check's Cancel) */
    cancel() { if (pending) { pending = null; native().cancelChoice().catch(() => {}); } },
    /** File ▸ Versions... (the desktop app): the versions of the document's file its saves replaced (kept 30 days),
     *  one of them opened in a window of its own, to be saved where the user likes */
    async showVersions(fileName) {
      const ui = L.ui;
      let list = [];
      try { list = await native().versions(); } catch (e) { ui.msg(String(e && e.message || e), { icon: 'error' }); return; }
      if (!list.length) { ui.msg(L.t('{0} hasn\'t kept any earlier versions of “{1}”. It keeps the version a save replaces, for 30 days.', L.APP, fileName), { icon: 'info' }); return; }
      const box = L.h('select', { size: 8, style: 'width:100%' }, ...list.map((v, i) => L.h('option', { value: v.id, selected: i === 0 ? '' : null, translate: 'no' },
        new Date(v.time).toLocaleString(root.VO && root.VO.i18n ? root.VO.i18n.dates : []) + ' · ' + L.t('{0} KB', Math.max(1, Math.round(v.size / 1024))))));
      const open = () => native().openVersion(box.value).catch((e) => ui.msg(String(e && e.message || e), { icon: 'error' }));
      box.addEventListener('dblclick', open);
      ui.dialog({ title: L.t('Versions'), width: 420, focus: box, body: L.h('div', { class: 'col' }, L.h('div', { text: L.t('Earlier versions of “{0}”, kept as saves replaced them:', fileName) }), box),
        buttons: [{ label: '&Open', primary: true, onClick: () => { open(); } }, { label: 'Close', cancel: true }] });
    },
    /** writes `blob` as `type`: how 'save' (the window's own file) or 'chosen' (what choose picked). A failure is
     *  told with its next step (Replace It, Try Again, Save As...). { result: 'saved' or 'cancelled', name (the
     *  file's, with its extension), bound (it is now the window's file) } */
    write(blob, how, { type, name }) {
      const run = queue.then(() => writeNow(blob, how, type, name));
      queue = run.catch(() => {});
      return run;
    },
  };
  /* the lite edition (desktop/src/lite.rs): no system Save dialog on its small screen. The name and type are asked
     here and the file goes in its Documents folder; one already there is replaced only once the user says so */
  async function chooseByName(purpose, types, type, name) {
    const ui = L.ui, title = L.t(purpose === 'export' ? 'Export' : 'Save As');
    for (;;) {
      const inp = L.h('input', { type: 'text', id: 'save-name', value: name, spellcheck: 'false', style: 'width:100%' });
      const sel = types.length > 1 ? ui.select(types.map((t) => [t.id, t.label]), type, null, { id: 'save-type', style: 'width:100%' }) : null;
      let ok = false;
      await ui.dialog({ title, width: 300, focus: inp,
        body: L.h('div', { class: 'col' }, ui.field('File &name:', inp), sel ? ui.field('Save as &type:', sel) : null, L.h('div', { class: 'hint', text: L.t('In your Documents folder.') })),
        buttons: [{ label: '&Save', primary: true, onClick: () => { ok = true; } }, { label: 'Cancel', cancel: true }] }).done;
      if (!ok) return null;
      /* a type's extension typed (report.xlsx) chooses it */
      const typed = types.find((t) => t.id === L.file.typeOf(inp.value.trim()));
      type = typed ? typed.id : sel ? sel.value : type;
      name = (typed ? without(inp.value.trim()) : inp.value.trim()) || name;
      let replace = false;
      for (;;) {
        try {
          const c = await native().choose({ purpose, types, type, name, title, replace });
          pending = c;
          return c && { name: without(c.name), file: c.name, type: c.type };
        } catch (e) {
          if (e && e.code === 'exists' && !replace) {
            if (await ui.msg(L.t('“{0}” already exists. Replace it?', e.file || name), { icon: 'warn', buttons: ['&Yes', '&No'] }) === 0) { replace = true; continue; }
          } else await ui.msg(String(e && e.message || e), { icon: 'error' });
          break;   // the name again
        }
      }
    }
  }
  async function writeNow(blob, how, type, name) {
    const host = native();
    const bytes = new Uint8Array(await blob.arrayBuffer());
    let replace = false;
    for (;;) {
      const shown = how === 'save' ? host.file.name : (pending && pending.name) || name;
      try {
        if (how === 'save') {
          await host.save(bytes, replace);
          return { result: 'saved', name: host.file.name, bound: true };
        }
        const w = await host.writeChosen(bytes);
        pending = null;
        if (w.bound) host.file = { bound: true, readOnly: false, name: w.name };
        return { result: 'saved', name: w.name, bound: w.bound };
      } catch (e) {
        const code = e && e.code, say = SAYS[code === 'denied' && host.sandboxed ? 'sandboxed' : code];
        if (!say) {
          await L.ui.msg(L.t('“{0}” couldn\'t be saved: {1} Your file hasn\'t changed.', shown, String(e && e.message || e)), { icon: 'error' });
          if (how === 'chosen') L.file.cancel();
          return { result: 'cancelled' };
        }
        const r = await L.ui.msg(L.t(say[0], shown, L.APP), { icon: 'warn', buttons: say[1] });
        const act = say[2][r];
        if (act === 'replace') { replace = true; continue; }
        if (act === 'retry') continue;
        if (act === 'allow') {
          /* the folder chooser, at the file's folder; then the same save again (asked again if still refused) */
          try { await host.allowAccess(how, L.t('Choose the folder of “{0}” and click Open to let {1} save there.', shown, L.APP)); }
          catch (err) { await L.ui.msg(String(err && err.message || err), { icon: 'warn' }); }
          continue;
        }
        if (act === 'saveas') {
          const c = await L.file.choose('saveAs', type, without(shown), [type]);
          if (!c) return { result: 'cancelled' };
          how = 'chosen';
          continue;
        }
        if (how === 'chosen') L.file.cancel();
        return { result: 'cancelled' };
      }
    }
  }
  /** Offer a generated file to the user as a download. Returns 'saved', and tells vo-saved { name } on window once
   *  the file is handed over (the assistant acts out Save and Save As then, common/scout/assistant.js). In the
   *  desktop app it is an export: the Save dialog asks where (unless the app's own Save As just did), then the file
   *  is written there; 'cancelled' if the user gave up. */
  L.saveFile = async function (name, blob) {
    if (native()) {
      const type = L.file.typeOf(name);
      if (!type) throw new Error(L.t('{0} doesn\'t save “{1}”.', L.APP, name));
      if (!pending || pending.type !== type) { const c = await L.file.choose('export', type, name, [type]); if (!c) return 'cancelled'; }
      const w = await L.file.write(blob, 'chosen', { type, name });
      if (w.result === 'saved') window.dispatchEvent(new CustomEvent('vo-saved', { detail: { name: w.name } }));
      return w.result;
    }
    const url = URL.createObjectURL(blob);
    const a = L.h('a', { href: url, download: name, style: 'display:none' });
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 5000);
    window.dispatchEvent(new CustomEvent('vo-saved', { detail: { name } }));
    return 'saved';
  };
  L.pickFiles = function (accept, multiple) {
    return new Promise((resolve) => {
      const inp = L.h('input', { type: 'file', accept: accept || '', style: 'position:fixed;left:-9999px;top:0' });
      if (multiple) inp.multiple = true;
      document.body.appendChild(inp);
      let done = false;
      inp.addEventListener('change', () => { done = true; resolve(Array.from(inp.files || [])); inp.remove(); });
      window.addEventListener('focus', function f() {
        window.removeEventListener('focus', f);
        setTimeout(() => { if (!done) { resolve([]); inp.remove(); } }, 1500);
      });
      inp.click();
    });
  };
  L.readAsArrayBuffer = (blob) => blob.arrayBuffer ? blob.arrayBuffer() : new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(blob); });
  L.readAsText = (blob) => blob.text ? blob.text() : new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsText(blob); });
  L.blobToDataURL = (blob) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(blob); });
  L.loadImage = (src) => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('image load failed')); i.src = src; });
  L.extToMime = (ext) => ({
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp',
    tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', ico: 'image/x-icon',
  }[String(ext || '').toLowerCase()] || 'application/octet-stream');
  L.mimeToExt = (m) => ({ 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/tiff': 'tiff', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf' }[m] || 'png');

  /** pixel size and resolution of a PNG / JPEG / GIF / BMP from its header: {w, h, dpiX, dpiY} or null */
  L.imageSize = function (b) {
    try {
      const u32 = (o) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
      const u16 = (o) => (b[o] << 8) | b[o + 1];
      const le16 = (o) => b[o] | (b[o + 1] << 8), le32 = (o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
      if (b[0] === 0x89 && b[1] === 0x50) {
        const r = { w: u32(16), h: u32(20), dpiX: 96, dpiY: 96 };
        for (let o = 8; o + 12 <= b.length && o < 4096;) {
          const len = u32(o), type = String.fromCharCode(b[o + 4], b[o + 5], b[o + 6], b[o + 7]);
          if (type === 'pHYs' && b[o + 16] === 1) { r.dpiX = Math.round(u32(o + 8) * 0.0254) || 96; r.dpiY = Math.round(u32(o + 12) * 0.0254) || 96; }
          if (type === 'IDAT') break;
          o += 12 + len;
        }
        return r;
      }
      if (b[0] === 0xff && b[1] === 0xd8) {
        let dpiX = 96, dpiY = 96;
        for (let o = 2; o + 9 < b.length;) {
          if (b[o] !== 0xff) { o++; continue; }
          const m = b[o + 1], len = u16(o + 2);
          if (m === 0xe0 && b[o + 4] === 0x4a && b[o + 11] >= 1) { const k = b[o + 11] === 2 ? 2.54 : 1; dpiX = Math.round(u16(o + 12) * k) || 96; dpiY = Math.round(u16(o + 14) * k) || 96; }
          if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { w: u16(o + 7), h: u16(o + 5), dpiX, dpiY };
          o += 2 + len;
        }
        return null;
      }
      if (b[0] === 0x47 && b[1] === 0x49) return { w: le16(6), h: le16(8), dpiX: 96, dpiY: 96 };
      if (b[0] === 0x42 && b[1] === 0x4d) { const px = le32(38), py = le32(42); return { w: le32(18), h: Math.abs(le32(22) | 0), dpiX: px ? Math.round(px * 0.0254) : 96, dpiY: py ? Math.round(py * 0.0254) : 96 }; }
    } catch (e) { /* unknown */ }
    return null;
  };

  /* ---------- media store (binary assets live outside undo snapshots) ---------- */
  const media = new Map();
  L.media = {
    /** blob: the original bytes (what gets saved); view: an optional browser-displayable rendition (e.g. PNG of a WMF) */
    add(blob, name, view) {
      const id = L.uid('m');
      media.set(id, { blob, view: view || blob, url: URL.createObjectURL(view || blob), type: blob.type || 'image/png', name: name || id });
      return id;
    },
    get: (id) => media.get(id),
    url: (id) => (media.get(id) ? media.get(id).url : ''),
    has: (id) => media.has(id),
    all: () => media,
    clear() { for (const m of media.values()) URL.revokeObjectURL(m.url); media.clear(); },
  };

  /** a copy of a picture with one colour made transparent (Set Transparent Color, a:clrChange); cached per colour */
  const clearCache = new Map();
  L.media.withTransparent = function (id, hex, tol) {
    const key = id + '|' + hex + '|' + (tol || 0);
    if (!clearCache.has(key)) clearCache.set(key, (async () => {
      try {
        const im = await L.loadImage(L.media.url(id));
        const c = document.createElement('canvas');
        c.width = im.naturalWidth || 1; c.height = im.naturalHeight || 1;
        const g = c.getContext('2d');
        g.drawImage(im, 0, 0);
        const data = g.getImageData(0, 0, c.width, c.height);
        const [tr, tg, tb] = L.color.hexToRgb(hex);
        const t = tol || 0, d = data.data;
        for (let i = 0; i < d.length; i += 4) if (Math.abs(d[i] - tr) <= t && Math.abs(d[i + 1] - tg) <= t && Math.abs(d[i + 2] - tb) <= t) d[i + 3] = 0;
        g.putImageData(data, 0, 0);
        const blob = await new Promise((res) => c.toBlob(res, 'image/png'));
        if (!blob) return null;
        const nid = L.media.add(blob, 'transparent.png');
        const o = L.media.get(id);
        if (o && o.size) L.media.get(nid).size = o.size;
        return nid;
      } catch (e) { return null; }
    })());
    return clearCache.get(key);
  };

  /* ---------- text helpers ---------- */
  L.titleCase = (s) => s.replace(/\w\S*/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase());
  L.sentenceCase = (s) => s.toLowerCase().replace(/(^\s*\w|[.!?]\s+\w)/g, (c) => c.toUpperCase());
  L.toggleCase = (s) => s.split('').map((c) => (c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase())).join('');
  L.romanize = (n) => {
    if (n <= 0 || n > 3999) return String(n);
    const t = [['M', 1000], ['CM', 900], ['D', 500], ['CD', 400], ['C', 100], ['XC', 90], ['L', 50], ['XL', 40], ['X', 10], ['IX', 9], ['V', 5], ['IV', 4], ['I', 1]];
    let s = ''; for (const [r, v] of t) while (n >= v) { s += r; n -= v; } return s;
  };
  L.alpha = (n) => { let s = ''; n = Math.max(1, n); while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };
  L.fmtDate = (d, fmt) => {
    const M = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const D = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const p = (n) => String(n).padStart(2, '0');
    const h12 = d.getHours() % 12 || 12, ap = d.getHours() < 12 ? 'AM' : 'PM';
    switch (fmt) {
      case 'datetime2': return `${D[d.getDay()]}, ${M[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
      case 'datetime3': return `${d.getDate()} ${M[d.getMonth()]} ${d.getFullYear()}`;
      case 'datetime4': return `${M[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
      case 'datetime5': return `${p(d.getDate())}-${M[d.getMonth()].slice(0, 3)}-${String(d.getFullYear()).slice(2)}`;
      case 'datetime6': return `${M[d.getMonth()]} ${String(d.getFullYear()).slice(2)}`;
      case 'datetime7': return `${M[d.getMonth()].slice(0, 3)}-${String(d.getFullYear()).slice(2)}`;
      case 'datetime8': return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()} ${h12}:${p(d.getMinutes())} ${ap}`;
      case 'datetime9': return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()} ${h12}:${p(d.getMinutes())}:${p(d.getSeconds())} ${ap}`;
      case 'datetime10': return `${d.getHours()}:${p(d.getMinutes())}`;
      case 'datetime11': return `${d.getHours()}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
      case 'datetime12': return `${h12}:${p(d.getMinutes())} ${ap}`;
      case 'datetime13': return `${h12}:${p(d.getMinutes())}:${p(d.getSeconds())} ${ap}`;
      default: return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
    }
  };
  L.DATE_FORMATS = ['datetime1', 'datetime2', 'datetime3', 'datetime4', 'datetime5', 'datetime6', 'datetime7', 'datetime8', 'datetime9', 'datetime10', 'datetime11', 'datetime12', 'datetime13'];

  /* ---------- fonts ---------- */
  const FONT_FALLBACK = {
    'arial': '"Arimo","Liberation Sans",Helvetica,sans-serif',
    'arial black': '"Arial Black","Arimo",sans-serif',
    'arial narrow': '"Arial Narrow","Arimo",sans-serif',
    'helvetica': '"Helvetica Neue","TeX Gyre Heros","Nimbus Sans",Arial,"Arimo",sans-serif',
    'times new roman': '"Tinos","Liberation Serif",Times,serif',
    'times': '"Times New Roman","Tinos","TeX Gyre Termes",serif',
    'courier new': '"Cousine","Liberation Mono",Courier,monospace',
    'courier': '"Courier New","Cousine",monospace',
    'calibri': '"Carlito",Candara,"Segoe UI",sans-serif',
    'calibri light': '"Carlito",sans-serif',
    'aptos': 'Calibri,"Carlito",sans-serif',
    'aptos display': '"Calibri Light",Calibri,"Carlito",sans-serif',
    'aptos narrow': '"Arial Narrow",Calibri,"Carlito",sans-serif',
    'liberation sans': 'Arial,"Arimo",sans-serif',
    'liberation serif': '"Times New Roman","Tinos",serif',
    'liberation mono': '"Courier New","Cousine",monospace',
    'cambria': '"Caladea",Georgia,serif',
    'georgia': '"Gelasio","DejaVu Serif","Times New Roman","Tinos",serif',
    'garamond': '"EB Garamond","Times New Roman","Tinos",serif',
    'book antiqua': '"Palatino Linotype",Palatino,"TeX Gyre Pagella","P052","URW Palladio L","Times New Roman","Tinos",serif',
    'palatino linotype': 'Palatino,"Book Antiqua","TeX Gyre Pagella","P052","URW Palladio L","Times New Roman","Tinos",serif',
    'palatino': '"Palatino Linotype","Book Antiqua","TeX Gyre Pagella","P052","Times New Roman","Tinos",serif',
    'century schoolbook': '"New Century Schoolbook","TeX Gyre Schola","C059","Times New Roman","Tinos",serif',
    'century': '"Century Schoolbook","TeX Gyre Schola","C059","Times New Roman","Tinos",serif',
    'consolas': '"Inconsolata","DejaVu Sans Mono","Courier New","Cousine",monospace',
    'bookman old style': '"Bookman","TeX Gyre Bonum","URW Bookman","URW Bookman L","Times New Roman","Tinos",serif',
    'century gothic': '"TeX Gyre Adventor","URW Gothic","URW Gothic L","Futura",Arial,"Arimo",sans-serif',
    'tahoma': '"DejaVu Sans Condensed",Verdana,Arial,"Arimo",sans-serif',
    'verdana': '"DejaVu Sans",Tahoma,Arial,"Arimo",sans-serif',
    'trebuchet ms': '"Segoe UI",Arial,"Arimo",sans-serif',
    'comic sans ms': '"Comic Neue","Chalkboard SE",cursive',
    'impact': 'Haettenschweiler,"Arial Black",sans-serif',
    'lucida console': 'Monaco,"DejaVu Sans Mono","Courier New","Cousine",monospace',
    'segoe ui': 'Arial,"Arimo",sans-serif',
    'franklin gothic medium': '"Libre Franklin",Arial,"Arimo",sans-serif',
    'wingdings': '"Segoe UI Symbol",sans-serif',
    'symbol': '"Segoe UI Symbol",serif',
  };
  /* East Asian fonts (docs/i18n.md, Fonts): Office's names in English and as Japanese, Korean and Chinese Office
     write them, and the common open-source Chinese families, each with the fonts a Mac, Windows or Linux computer has
     in their place. None is vendored (several megabytes each, in every offline install). [English name, stack, other names] */
  const CJK_STACKS = {
    jpSerif: '"Yu Mincho","YuMincho","Hiragino Mincho ProN","Hiragino Mincho Pro","Noto Serif CJK JP","Noto Serif JP","IPAexMincho","IPAMincho",serif',
    jpSans: '"Yu Gothic","YuGothic","Hiragino Sans","Hiragino Kaku Gothic ProN","Meiryo","Noto Sans CJK JP","Noto Sans JP","IPAexGothic","IPAGothic",sans-serif',
    krSerif: '"Batang","AppleMyungjo","Noto Serif CJK KR","Noto Serif KR","UnBatang",serif',
    krSans: '"Malgun Gothic","Apple SD Gothic Neo","Noto Sans CJK KR","Noto Sans KR","UnDotum",sans-serif',
    scSerif: '"SimSun","Songti SC","STSong","Noto Serif CJK SC","Noto Serif SC","Source Han Serif SC","AR PL UMing CN",serif',
    scSans: '"Microsoft YaHei","DengXian","PingFang SC","Hiragino Sans GB","Noto Sans CJK SC","Noto Sans SC","Source Han Sans SC","WenQuanYi Micro Hei",sans-serif',
    tcSerif: '"PMingLiU","Songti TC","LiSong Pro","Noto Serif CJK TC","Noto Serif TC","AR PL UMing TW",serif',
    tcSans: '"Microsoft JhengHei","PingFang TC","Noto Sans CJK TC","Noto Sans TC",sans-serif',
    /* Source Han and Noto CJK are the same designs: each stands in for the other first */
    scOpenSerif: '"Source Han Serif SC","Noto Serif CJK SC","Noto Serif SC","SimSun","Songti SC","STSong","AR PL UMing CN",serif',
    scOpenSans: '"Source Han Sans SC","Noto Sans CJK SC","Noto Sans SC","Microsoft YaHei","DengXian","PingFang SC","Hiragino Sans GB","WenQuanYi Micro Hei",sans-serif',
  };
  const CJK = [['MS Mincho', 'jpSerif', 'ＭＳ 明朝'], ['MS PMincho', 'jpSerif', 'ＭＳ Ｐ明朝'], ['MS Gothic', 'jpSans', 'ＭＳ ゴシック'], ['MS PGothic', 'jpSans', 'ＭＳ Ｐゴシック'],
    ['MS UI Gothic', 'jpSans'], ['Yu Mincho', 'jpSerif', '游明朝'], ['Yu Gothic', 'jpSans', '游ゴシック'], ['Yu Gothic UI', 'jpSans'], ['Yu Gothic Light', 'jpSans', '游ゴシック Light'],
    ['Meiryo', 'jpSans', 'メイリオ'], ['Meiryo UI', 'jpSans'], ['HG Mincho E', 'jpSerif', 'HG明朝E'], ['HG Gothic E', 'jpSans', 'HGゴシックE'],
    ['Malgun Gothic', 'krSans', '맑은 고딕'], ['Batang', 'krSerif', '바탕'], ['BatangChe', 'krSerif', '바탕체'], ['Gulim', 'krSans', '굴림'], ['GulimChe', 'krSans', '굴림체'],
    ['Dotum', 'krSans', '돋움'], ['DotumChe', 'krSans', '돋움체'], ['Gungsuh', 'krSerif', '궁서'], ['GungsuhChe', 'krSerif', '궁서체'],
    ['SimSun', 'scSerif', '宋体'], ['NSimSun', 'scSerif', '新宋体'], ['SimHei', 'scSans', '黑体'], ['DengXian', 'scSans', '等线'], ['DengXian Light', 'scSans', '等线 Light'],
    ['Microsoft YaHei', 'scSans', '微软雅黑'], ['KaiTi', 'scSerif', '楷体'], ['FangSong', 'scSerif', '仿宋'],
    ['WenQuanYi Micro Hei', 'scSans', '文泉驿微米黑', '文泉驛微米黑'],
    ['Source Han Sans SC', 'scOpenSans', '思源黑体'], ['Source Han Serif SC', 'scOpenSerif', '思源宋体'],
    ['Noto Sans CJK SC', 'scOpenSans'], ['Noto Sans SC', 'scOpenSans'], ['Noto Serif CJK SC', 'scOpenSerif'], ['Noto Serif SC', 'scOpenSerif'],
    ['PMingLiU', 'tcSerif', '新細明體'], ['MingLiU', 'tcSerif', '細明體'], ['Microsoft JhengHei', 'tcSans', '微軟正黑體'], ['DFKai-SB', 'tcSerif', '標楷體']];
  const CJK_BY = new Map(), ALIAS = new Map();
  for (const [en, stack, ...names] of CJK) { CJK_BY.set(en.toLowerCase(), { en, stack }); for (const n of names) ALIAS.set(n.toLowerCase(), en.toLowerCase()); }
  /** the key every font table uses for a font: its English name, lower case (ＭＳ 明朝 → ms mincho) */
  L.fontKey = (name) => { const n = String(name || '').replace(/"/g, '').trim().toLowerCase(); return ALIAS.get(n) || n; };
  /** whether a font is one of the East Asian fonts above */
  L.fontIsEA = (name) => CJK_BY.has(L.fontKey(name));
  /** whether text has East Asian characters (Han, kana, Hangul, full-width forms), which Word draws with a run's
      East Asian font */
  L.hasEAText = (s) => /[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\ufe30-\ufe4f\uff00-\uffef]/.test(s || '');
  /** a CSS font-family list for a font and what stands in for it; with ea, an East Asian font, its letters come
      from the first and its East Asian characters from the second (as Word draws a run with both fonts) */
  /** what a document's font table says about fonts it names (an app fills it when it opens one): font key ->
      { alt, family, pitch }, for a font we have no stand-in of our own for */
  L.fontHints = new Map();
  /* Word's stand-in for a missing font of a kind: a serif face for roman, a fixed-pitch one for modern */
  const BY_FAMILY = { roman: 'times new roman', modern: 'courier new', swiss: 'arial' };
  const hinted = (k, depth) => {
    const h = L.fontHints.get(k);
    if (!h || depth > 2) return null;
    const alt = h.alt && L.fontKey(h.alt);
    if (alt && alt !== k && (FONT_FALLBACK[alt] || CJK_BY.has(alt) || L.fontHints.has(alt))) return `"${h.alt}",` + (hinted(alt, depth + 1) || L.fontStack(h.alt).replace(/^"[^"]*",/, ''));
    const fam = h.pitch === 'fixed' ? 'courier new' : BY_FAMILY[h.family];
    return fam ? `"${{ 'times new roman': 'Times New Roman', 'courier new': 'Courier New', arial: 'Arial' }[fam]}",${FONT_FALLBACK[fam]}` : null;
  };
  L.fontStack = function (name, ea) {
    if (!name) return ea ? L.fontStack(ea) : 'Arial,"Arimo",sans-serif';
    const n = String(name).replace(/"/g, ''), k = L.fontKey(n), cjk = CJK_BY.get(k);
    const stack = cjk ? `"${n}",` + (cjk.en !== n ? `"${cjk.en}",` : '') + CJK_STACKS[cjk.stack] : `"${n}",${FONT_FALLBACK[k] || hinted(k, 0) || 'Arial,"Arimo",sans-serif'}`;
    if (!ea || L.fontKey(ea) === k || !L.fontIsEA(ea)) return stack;
    return stack.replace(/,(serif|sans-serif|monospace|cursive)$/, '') + ',' + L.fontStack(ea);
  };
  /* Which family in a stack the browser will draw with: an installed font (detected by measuring
     text against the generic fallbacks), or one of the suite's own stand-ins in public/fonts/. */
  const installed = new Map();
  let probe = null;
  L.fontInstalled = function (family) {
    if (installed.has(family)) return installed.get(family);
    if (!probe) probe = document.createElement('canvas').getContext('2d');
    const text = 'mmmmmmmmmmlli1WQ@ ', width = (font) => { probe.font = '72px ' + font; return probe.measureText(text).width; };
    const yes = ['monospace', 'serif', 'sans-serif'].some((g) => width(`"${family}",${g}`) !== width(g));
    installed.set(family, yes);
    return yes;
  };
  const standIns = () => new Set(Array.from(document.fonts || [], (f) => f.family.replace(/["']/g, '')));
  /** the family that draws `name` (an Office font name): itself when installed, else its stand-in */
  L.fontUsed = function (name) {
    const own = standIns();
    for (const f of L.fontStack(name).split(',').map((s) => s.trim().replace(/^"|"$/g, ''))) {
      if (['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy'].includes(f)) return f;
      if (own.has(f) || L.fontInstalled(f)) return f;
    }
    return 'sans-serif';
  };
  const GENERIC = ['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy'];
  /** what draws `name` on this computer when it isn't installed: a family, a generic one (serif…), or '' */
  L.fontStandIn = function (name) {
    if (L.localFonts().includes(name)) return '';   // listed by the browser as installed here
    const used = L.fontUsed(name);
    return used.toLowerCase() === String(name).toLowerCase() ? '' : used;
  };
  /** the Font boxes' grey note beside a font name, and its tooltip: what draws it on this computer, if not itself
      ({ text, tip }, or '' when it is installed) */
  L.fontNote = function (name) {
    const used = L.fontStandIn(name);
    if (!used) return '';
    const generic = GENERIC.includes(used), t = tr;   // t: UI text (tools/i18n.mjs)
    return {
      text: generic ? t('Not installed') : t('Not installed; uses {0}', used),
      tip: generic ? t('{0} is not installed on this computer, so its text is shown in a default font. The document keeps the font name {0} and shows in that font on computers that have it.', name)
        : t('{0} is not installed on this computer, so its text is shown in {1}. The document keeps the font name {0} and shows in that font on computers that have it.', name, used),
    };
  };
  /** a string as a quoted CSS string (a font name in HTML we write), with the quotes, backslash, angle brackets
   * and line breaks escaped: a document's font name can end neither the string, an attribute nor a <style>. */
  L.cssString = (s) => `'${String(s).replace(/[\\'"<>\n\r\f]/g, (c) => '\\' + c.charCodeAt(0).toString(16) + ' ')}'`;
  /** the Content-Security-Policy meta tag of a page we export (Save as Web Page). The file is opened from disk,
   * away from the suite's own policy, so it carries one: pictures and fonts are inside it (data:), styles
   * inline, nothing loads from anywhere, and no script runs but the page's own `script` (the text between its
   * <script> tags, allowed by its SHA-256; common/sha.js), if it has one. */
  L.exportCSP = function (script) {
    const hash = script == null ? null : btoa(String.fromCharCode(...L.sha.sha256(new TextEncoder().encode(script))));
    return `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; font-src data:; style-src 'unsafe-inline'; script-src ${hash ? `'sha256-${hash}'` : "'none'"}; base-uri 'none'; form-action 'none'">`;
  };
  /** @font-face rules for a standalone export (HTML or an SVG image): the stand-ins from public/fonts/ that
   * `markup` names, as data: URLs, with bold/italic faces only if the markup uses them. */
  L.embedFonts = async function (markup) {
    const base = (root.VO?.base || new URL('..', location.href).href) + 'fonts/';
    const css = await (await fetch(base + 'fonts.css')).text();
    const bold = /bold|font-weight[:=]\s*"?[6-9]00/i.test(markup), italic = /italic/i.test(markup);
    const files = new Map();   // one rule per file: a variable font's regular and bold share it
    for (const rule of css.match(/@font-face\s*\{[^}]*\}/g) || []) {
      const family = /font-family:\s*'([^']+)'/.exec(rule)[1], file = /url\('([^']+)'\)/.exec(rule)[1];
      const style = /font-style:\s*(\w+)/.exec(rule)[1], weight = +/font-weight:\s*(\d+)/.exec(rule)[1];
      if (!markup.includes(family) || (!bold && weight > 400) || (!italic && style === 'italic')) continue;
      if (!files.has(file)) files.set(file, { family, style, weights: [] });
      files.get(file).weights.push(weight);
    }
    const rules = [];
    for (const [file, f] of files) {
      const bytes = new Uint8Array(await (await fetch(base + file)).arrayBuffer());
      let bin = ''; for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      const weight = f.weights.length > 1 ? Math.min(...f.weights) + ' ' + Math.max(...f.weights) : f.weights[0];
      rules.push(`@font-face{font-family:'${f.family}';font-style:${f.style};font-weight:${weight};src:url(data:font/woff2;base64,${btoa(bin)}) format('woff2')}`);
    }
    return rules.join('\n');
  };
  /* the Font boxes: the fonts English Office users pick most, by Office's names (documents add their own) */
  /* with a Japanese, Korean or Chinese UI (common/lang.js), that Office's own fonts too, and common open-source
     Chinese families under their installed names (sources and English aliases: docs/i18n.md, Fonts) */
  const uiLang = typeof document !== 'undefined' ? (document.documentElement.getAttribute('lang') || 'en') : 'en';
  L.FONT_LIST = ['Aptos', 'Aptos Display', 'Aptos Narrow', 'Arial', 'Arial Black', 'Calibri', 'Calibri Light', 'Cambria', 'Century Gothic', 'Comic Sans MS', 'Consolas',
    'Courier New', 'Garamond', 'Georgia', 'Impact', 'Palatino Linotype', 'Segoe UI', 'Symbol', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Wingdings']
    .concat({ ja: ['游明朝', '游ゴシック', 'ＭＳ 明朝', 'ＭＳ ゴシック', 'ＭＳ Ｐ明朝', 'ＭＳ Ｐゴシック', 'メイリオ'], ko: ['맑은 고딕', '바탕', '바탕체', '굴림', '굴림체', '돋움', '돋움체', '궁서'],
      'zh-CN': ['宋体', '新宋体', '黑体', '楷体', '仿宋', '等线', '微软雅黑', '文泉驿微米黑', '思源黑体', '思源宋体',
        'Noto Sans CJK SC', 'Noto Sans SC', 'Noto Serif CJK SC', 'Noto Serif SC'], 'zh-Hant': ['新細明體', '細明體', '標楷體', '微軟正黑體'] }[uiLang] || []);
  /** the Font box list in Office's sections: Theme Fonts (the document's heading and body fonts),
   * Recently Used Fonts, All Fonts (L.FONT_LIST plus fonts the document uses), then the computer's
   * fonts once loaded (Chrome and Edge: the first item loads them). Items are strings,
   * { heading } rows or { value, label }. */
  L.fontSections = function ({ major, minor, extra = [] } = {}) {
    const out = [], recent = L.store.get('recentFonts', []).slice(0, 5), local = L.localFonts();
    // Chrome and Edge can list the computer's fonts (Local Font Access, after the user allows it)
    if (typeof root.queryLocalFonts === 'function') out.push({ label: local.length ? 'Reload fonts on your computer...' : 'Load all fonts on your computer...', run: L.loadLocalFonts }, { sep: true });
    if (major || minor) {
      out.push({ heading: 'Theme Fonts' });
      if (major) out.push({ value: major, label: tr('{0} (Headings)', major) });
      if (minor) out.push({ value: minor, label: tr('{0} (Body)', minor) });
    }
    if (recent.length) out.push({ heading: 'Recently Used Fonts' }, ...recent);
    const all = Array.from(new Set(L.FONT_LIST.concat(extra.filter(Boolean)))).sort((x, y) => x.localeCompare(y));
    out.push({ heading: 'All Fonts' }, ...all);
    const more = local.filter((f) => !all.includes(f));
    if (more.length) out.push({ heading: 'Fonts on This Computer' }, ...more);
    return out;
  };
  /* the families queryLocalFonts() found, kept for every app of the suite (VO.localFonts, common/suite.js;
     the Start Center's Settings can load or forget them) */
  L.localFonts = () => (root.VO ? root.VO.localFonts.list() : []);
  L.loadLocalFonts = async function () {
    try {
      return await root.VO.localFonts.load();
    } catch (e) {
      L.ui?.msg(e.name === 'NotAllowedError' || e.name === 'SecurityError' ? 'The fonts on this computer could not be listed: permission was not given. You can allow it in the site settings of your browser.' : L.t('The fonts on this computer could not be listed: {0}', e.message || e), { icon: 'warn' });
      return 0;
    }
  };
  /** remember a font picked in a Font box (for Recently Used Fonts) */
  L.addRecentFont = function (name) {
    if (name) L.store.set('recentFonts', [name].concat(L.store.get('recentFonts', []).filter((f) => f !== name)).slice(0, 10));
  };
  /* the Font Size box of Word and Excel; Lectern sets PowerPoint's list in its app.js */
  L.SIZE_LIST = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

  /* the Symbol font's characters at 0x20-0x7e and 0xa0-0xfe (Adobe's Symbol encoding; the arrow and bracket
     extenders as the nearest whole characters, the unassigned 0xf0 as \0) */
  const SYMBOL = [' !∀#∃%&∋()∗+,−./0123456789:;<=>?≅ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ[∴]⊥_‾αβχδεφγηιϕκλμνοπθρστυϖωξψζ{|}∼',
    '€ϒ′≤⁄∞ƒ♣♦♥♠↔←↑→↓°±″≥×∝∂•÷≠≡≈…|—↵ℵℑℜ℘⊗⊕∅∩∪⊃⊇⊄⊂⊆∈∉∠∇®©™∏√⋅¬∧∨⇔⇐⇑⇒⇓◊〈®©™∑⎛⎜⎝⎡⎢⎣⎧⎨⎩⎪\0〉∫⌠⎮⌡⎞⎟⎠⎤⎥⎦⎫⎬⎭'];
  /* Wingdings / Symbol bullet glyphs mapped to Unicode so imported documents look right; a character may come as
     the font's own code (v) or in the private-use area Word stores it in (U+F076) */
  L.mapSymbolChar = function (ch, font) {
    const f = String(font || '').toLowerCase();
    const code = ch ? ch.charCodeAt(0) & 0xff : 0, byte = ch && (ch.charCodeAt(0) < 0x100 || (ch.charCodeAt(0) >= 0xF000 && ch.charCodeAt(0) <= 0xF0FF));
    const own = byte ? String.fromCharCode(code) : ch;
    /* Wingdings 2 and 3 carry the round, square and triangular bullets of many Office themes */
    if (/^wingdings\s*2/.test(f)) {
      if (code >= 0x96 && code <= 0x9d) return '●';
      if (code >= 0x9e && code <= 0xa5) return '■';
      if (code >= 0xa6 && code <= 0xad) return '◆';
      return { 0x50: '✓', 0x4f: '✗', 0x51: '☒', 0x52: '☑', 0xe9: '★', 0xea: '★' }[code] || '•';
    }
    if (/^wingdings\s*3/.test(f)) return { 0x7d: '►', 0x84: '►', 0x75: '►', 0x71: '◄', 0x70: '▲', 0x72: '▼', 0xc6: '➔', 0xc7: '⬅' }[code] || '▸';
    if (f.startsWith('wingdings')) {
      const m = { 'l': '●', 'n': '■', 'q': '❑', 'u': '◆', 'v': '❖', 'Ø': '➢', 'ü': '✓', '§': '▪', 'o': '❍', 'p': '□', 'w': '⬥', 'Ü': '➢', 'à': '➔', 'è': '➜', 'Þ': '➝', 'ð': '⇨', 'ª': '✦', '¨': '◻' };
      return m[own] || '•';
    }
    if (f === 'symbol') {
      const u = !byte ? '' : code >= 0xa0 ? SYMBOL[1][code - 0xa0] : code >= 0x20 && code < 0x7f ? SYMBOL[0][code - 0x20] : '';
      return u && u !== '\0' ? u : ch;
    }
    if (ch === '' || ch === '') return '•';
    if (ch && ch.charCodeAt(0) >= 0xF000 && ch.charCodeAt(0) <= 0xF0FF) return '•';
    return ch;
  };

  /* A picture's brightness and contrast as Office draws them, measured in Word 2016 on a grey ramp under 13 VML and
     DrawingML settings: a slope s around mid-grey, with half the brightness added before it and half after,
     v' = s·(v − ½ + b/2) + ½ + b/2, clamped; s = 1/(1 − c) for c ≥ 0 and 1 + c below, where b and c are DrawingML's
     lum bright and contrast (−1…1; VML's gain is s and its blacklevel b/2). CSS filters can't add an offset, so this
     is an SVG filter: { css, defs }, where defs is the element defining it, to be kept beside the picture so that a
     copy of the picture (a page or slide rasterised for export) carries it. Null when there is nothing to do. */
  L.lumFilter = function (bright, contrast) {
    bright = +bright || 0; contrast = +contrast || 0;
    if (!bright && !contrast) return null;
    const s = contrast >= 0 ? 1 / Math.max(1 - contrast, 0.001) : 1 + contrast;
    const intercept = 0.5 - s / 2 + (bright * (1 + s)) / 2;
    const num = (v) => String(L.round(v, 4)).replace('-', 'm').replace('.', '_');
    const id = `vo-lum-${num(s)}-${num(intercept)}`;
    const NS = 'http://www.w3.org/2000/svg', svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('width', '0');
    svg.setAttribute('height', '0');
    svg.setAttribute('aria-hidden', 'true');
    svg.style.cssText = 'position:absolute;width:0;height:0;overflow:hidden';
    const filter = document.createElementNS(NS, 'filter'), transfer = document.createElementNS(NS, 'feComponentTransfer');
    filter.id = id;
    filter.setAttribute('color-interpolation-filters', 'sRGB');
    for (const c of 'RGB') {
      const fn = document.createElementNS(NS, 'feFunc' + c);
      fn.setAttribute('type', 'linear');
      fn.setAttribute('slope', String(L.round(s, 4)));
      fn.setAttribute('intercept', String(L.round(intercept, 4)));
      transfer.appendChild(fn);
    }
    filter.appendChild(transfer);
    svg.appendChild(filter);
    return { css: `url(#${id})`, defs: svg };
  };

  /* ---------- geometry helpers ---------- */
  L.rotPt = (x, y, cx, cy, deg) => {
    const a = (deg * Math.PI) / 180, c = Math.cos(a), s = Math.sin(a);
    const dx = x - cx, dy = y - cy;
    return [cx + dx * c - dy * s, cy + dx * s + dy * c];
  };
  /* Axis-aligned bounds of a rotated rectangle */
  L.rotBounds = (sh) => {
    if (!sh.rot) return { x: sh.x, y: sh.y, w: sh.w, h: sh.h };
    const cx = sh.x + sh.w / 2, cy = sh.y + sh.h / 2;
    const pts = [[sh.x, sh.y], [sh.x + sh.w, sh.y], [sh.x + sh.w, sh.y + sh.h], [sh.x, sh.y + sh.h]].map(([x, y]) => L.rotPt(x, y, cx, cy, sh.rot));
    const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
    const x0 = Math.min(...xs), y0 = Math.min(...ys);
    return { x: x0, y: y0, w: Math.max(...xs) - x0, h: Math.max(...ys) - y0 };
  };
  L.unionBounds = (list) => {
    if (!list.length) return { x: 0, y: 0, w: 0, h: 0 };
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const b of list) { x0 = Math.min(x0, b.x); y0 = Math.min(y0, b.y); x1 = Math.max(x1, b.x + b.w); y1 = Math.max(y1, b.y + b.h); }
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  };
})(typeof window !== 'undefined' ? window : globalThis);
