/* VibeOffice — core utilities, shared by Quire, Ledger and Lectern.
 * Shared namespace, DOM helpers, units, colors, events, file I/O, media store, PDF writer.
 * Every other script attaches itself to window.L. Each page names its app before loading this file:
 *   <script>window.L = { APP_ID: 'quire', APP_NAME: 'Quire', APP: 'Quire 2003' };</script>
 * APP_ID keeps each app's saved settings apart (localStorage 'quire.…'); APP_NAME and APP appear in messages.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  L.VERSION = root.VO?.VERSION || 'dev';   // one version for the suite: common/suite.js
  L.APP_ID = L.APP_ID || 'vibeoffice';
  L.APP_NAME = L.APP_NAME || 'VibeOffice';
  L.APP = L.APP || L.APP_NAME;

  /* ---------- DOM ---------- */
  L.h = function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    if (attrs) {
      for (const k in attrs) {
        const v = attrs[k];
        if (v == null || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
        else if (k === 'style') el.style.cssText = v;
        else if (k === 'html') el.innerHTML = v;
        else if (k === 'text') el.textContent = v;
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
        else if (k === 'dataset') Object.assign(el.dataset, v);
        else if (v === true) el.setAttribute(k, '');
        else el.setAttribute(k, v);
      }
    }
    appendKids(el, kids);
    return el;
  };
  function appendKids(el, kids) {
    for (const k of kids) {
      if (k == null || k === false) continue;
      if (Array.isArray(k)) appendKids(el, k);
      else if (k instanceof Node) el.appendChild(k);
      else el.appendChild(document.createTextNode(String(k)));
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
  L.xesc = (s) => String(s == null ? '' : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
    .replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  L.isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
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
  C.rgbToHsl = (r, g, b) => {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0; const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h /= 6;
    }
    return [h, s, l];
  };
  C.hslToRgb = (h, s, l) => {
    if (s === 0) return [l * 255, l * 255, l * 255];
    const hue = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
  };
  C.mix = (a, b, t) => { const x = C.hexToRgb(a), y = C.hexToRgb(b); return C.rgbToHex(x[0] + (y[0] - x[0]) * t, x[1] + (y[1] - x[1]) * t, x[2] + (y[2] - x[2]) * t); };
  C.lighten = (hex, t) => C.mix(hex, '#FFFFFF', t);
  C.darken = (hex, t) => C.mix(hex, '#000000', t);
  C.luma = (hex) => { const [r, g, b] = C.hexToRgb(hex); return (0.299 * r + 0.587 * g + 0.114 * b) / 255; };
  C.rgba = (hex, a) => { const [r, g, b] = C.hexToRgb(hex); return a == null || a >= 1 ? C.rgbToHex(r, g, b) : `rgba(${r},${g},${b},${L.round(a, 3)})`; };
  C.gray = (hex) => { const v = Math.round(C.luma(hex) * 255); return C.rgbToHex(v, v, v); };
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

  /* ---------- storage (never relied on) ---------- */
  L.store = {
    get(k, d) { try { const v = localStorage.getItem(L.APP_ID + '.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(L.APP_ID + '.' + k, JSON.stringify(v)); return true; } catch (e) { return false; } },
    del(k) { try { localStorage.removeItem(L.APP_ID + '.' + k); } catch (e) { /* ignore */ } },
  };

  /* ---------- file I/O ---------- */
  /**
   * Minimal PDF writer: one full-page JPEG per page.
   * pages: [{ jpeg: Uint8Array, iw, ih (pixels), pw, ph (points) }]
   */
  L.buildPDF = function (pages, info) {
    const enc = new TextEncoder();
    const chunks = [];
    let len = 0;
    const offs = [];
    const push = (x) => { const b = typeof x === 'string' ? enc.encode(x) : x; chunks.push(b); len += b.length; };
    const pdfStr = (t) => { let hex = 'FEFF'; for (const ch of String(t || '')) { const c = ch.codePointAt(0); if (c > 0xffff) { const v = c - 0x10000; hex += (0xd800 + (v >> 10)).toString(16).padStart(4, '0') + (0xdc00 + (v & 0x3ff)).toString(16).padStart(4, '0'); } else hex += c.toString(16).padStart(4, '0'); } return '<' + hex.toUpperCase() + '>'; };
    const num = (v) => String(Math.round(v * 100) / 100);
    push('%PDF-1.4\n');
    push(new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));
    const obj = (id, body) => { offs[id] = len; push(id + ' 0 obj\n'); if (typeof body === 'string') push(body); else body(); push('\nendobj\n'); };
    const n = pages.length;
    const infoId = 3 + n * 3;
    obj(1, '<< /Type /Catalog /Pages 2 0 R >>');
    obj(2, `<< /Type /Pages /Kids [${pages.map((_, i) => (3 + i * 3) + ' 0 R').join(' ')}] /Count ${n} >>`);
    pages.forEach((p, i) => {
      const pid = 3 + i * 3, cid = pid + 1, iid = pid + 2;
      obj(pid, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${num(p.pw)} ${num(p.ph)}] /Resources << /XObject << /Im${i} ${iid} 0 R >> >> /Contents ${cid} 0 R >>`);
      const cs = `q ${num(p.pw)} 0 0 ${num(p.ph)} 0 0 cm /Im${i} Do Q`;
      obj(cid, `<< /Length ${cs.length} >>\nstream\n${cs}\nendstream`);
      obj(iid, () => { push(`<< /Type /XObject /Subtype /Image /Width ${p.iw} /Height ${p.ih} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>\nstream\n`); push(p.jpeg); push('\nendstream'); });
    });
    const d = new Date();
    const pad = (v) => String(v).padStart(2, '0');
    obj(infoId, `<< /Title ${pdfStr(info && info.title)} /Author ${pdfStr(info && info.author)} /Producer ${pdfStr(L.APP)} /CreationDate (D:${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}) >>`);
    const xref = len;
    const total = infoId + 1;
    push(`xref\n0 ${total}\n0000000000 65535 f \n`);
    for (let i = 1; i < total; i++) push(String(offs[i]).padStart(10, '0') + ' 00000 n \n');
    push(`trailer\n<< /Size ${total} /Root 1 0 R /Info ${infoId} 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
    return new Blob(chunks, { type: 'application/pdf' });
  };
  /** Offer a generated file to the user as a download. Returns 'saved'. */
  L.saveFile = async function (name, blob) {
    const url = URL.createObjectURL(blob);
    const a = L.h('a', { href: url, download: name, style: 'display:none' });
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 5000);
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
    'helvetica': '"Helvetica Neue","TeX Gyre Heros","Nimbus Sans","Arimo",sans-serif',
    'times new roman': '"Tinos","Liberation Serif",Times,serif',
    'times': '"Tinos","TeX Gyre Termes",serif',
    'courier new': '"Cousine","Liberation Mono",Courier,monospace',
    'courier': '"Cousine",monospace',
    'calibri': '"Carlito",Candara,"Segoe UI",sans-serif',
    'calibri light': '"Carlito",sans-serif',
    'cambria': '"Caladea",Georgia,serif',
    'georgia': '"Gelasio","DejaVu Serif","Tinos",serif',
    'garamond': '"EB Garamond","Tinos",serif',
    'book antiqua': '"Palatino Linotype",Palatino,"TeX Gyre Pagella","P052","URW Palladio L","Tinos",serif',
    'palatino linotype': 'Palatino,"Book Antiqua","TeX Gyre Pagella","P052","URW Palladio L","Tinos",serif',
    'palatino': '"Palatino Linotype","Book Antiqua","TeX Gyre Pagella","P052","Tinos",serif',
    'century schoolbook': '"New Century Schoolbook","TeX Gyre Schola","C059","Tinos",serif',
    'century': '"Century Schoolbook","TeX Gyre Schola","C059","Tinos",serif',
    'consolas': '"Inconsolata","DejaVu Sans Mono","Cousine",monospace',
    'bookman old style': '"Bookman","TeX Gyre Bonum","URW Bookman","URW Bookman L","Tinos",serif',
    'century gothic': '"TeX Gyre Adventor","URW Gothic","URW Gothic L","Futura","Arimo",sans-serif',
    'tahoma': '"DejaVu Sans Condensed",Verdana,"Arimo",sans-serif',
    'verdana': '"DejaVu Sans",Tahoma,"Arimo",sans-serif',
    'trebuchet ms': '"Segoe UI","Arimo",sans-serif',
    'comic sans ms': '"Comic Neue","Chalkboard SE",cursive',
    'impact': 'Haettenschweiler,"Arial Black",sans-serif',
    'lucida console': 'Monaco,"DejaVu Sans Mono","Cousine",monospace',
    'segoe ui': '"Arimo",sans-serif',
    'franklin gothic medium': '"Libre Franklin","Arimo",sans-serif',
    'wingdings': '"Segoe UI Symbol",sans-serif',
    'symbol': '"Segoe UI Symbol",serif',
  };
  L.fontStack = function (name) {
    if (!name) return 'Arial,"Arimo",sans-serif';
    const n = String(name).replace(/"/g, '');
    const fb = FONT_FALLBACK[n.toLowerCase()] || '"Arimo",sans-serif';
    return `"${n}",${fb}`;
  };
  L.FONT_LIST = ['Arial', 'Arial Black', 'Arial Narrow', 'Book Antiqua', 'Bookman Old Style', 'Calibri', 'Calibri Light', 'Cambria', 'Century Gothic', 'Comic Sans MS', 'Courier New',
    'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Impact', 'Lucida Console', 'Palatino Linotype', 'Segoe UI', 'Symbol', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Wingdings'];
  /* the Font Size box of Word and Excel; Lectern sets PowerPoint's list in its app.js */
  L.SIZE_LIST = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

  /* Wingdings / Symbol bullet glyphs mapped to Unicode so imported documents look right */
  L.mapSymbolChar = function (ch, font) {
    const f = String(font || '').toLowerCase();
    const code = ch ? ch.charCodeAt(0) & 0xff : 0;
    /* Wingdings 2 and 3 carry the round, square and triangular bullets of many Office themes */
    if (/^wingdings\s*2/.test(f)) {
      if (code >= 0x96 && code <= 0x9d) return '●';
      if (code >= 0x9e && code <= 0xa5) return '■';
      if (code >= 0xa6 && code <= 0xad) return '◆';
      return { 0x50: '✓', 0x4f: '✗', 0x51: '☒', 0x52: '☑', 0xe9: '★', 0xea: '★' }[code] || '•';
    }
    if (/^wingdings\s*3/.test(f)) return { 0x7d: '►', 0x84: '►', 0x75: '►', 0x71: '◄', 0x70: '▲', 0x72: '▼', 0xc6: '➔', 0xc7: '⬅' }[code] || '▸';
    if (f.startsWith('wingdings')) {
      const m = { 'l': '●', 'n': '■', 'q': '❑', 'u': '◆', 'v': '❖', 'Ø': '➢', 'ü': '✓', '§': '▪', 'o': '❍', 'p': '□', 'w': '⬥', 'Ü': '➢', 'à': '➔', 'è': '➜', 'Þ': '➝', 'ð': '⇨', 'ª': '✦', '¨': '◻', '\uF06E': '■', '\uF06C': '●', '\uF075': '◆', '\uF0FC': '✓', '\uF0D8': '➢', '\uF0A7': '▪' };
      return m[ch] || '•';
    }
    if (f === 'symbol') { const m = { '·': '•', '\uF0B7': '•', 'Þ': '⇒', '®': '→', '¾': '—' }; return m[ch] || ch; }
    if (ch === '' || ch === '') return '•';
    if (ch && ch.charCodeAt(0) >= 0xF000 && ch.charCodeAt(0) <= 0xF0FF) return '•';
    return ch;
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
