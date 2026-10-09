/* Ledger — .xlsx / .xlsm reader (SpreadsheetML, transitional and strict).
 * L.xlsxRead.read(bytes, {password}) → Promise<Workbook>
 */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, F = L.formula, NF = L.numfmt, XML = L.xml, K = L.opc;
  const R = (L.xlsxRead = {});

  /* ------------------------------------------------------------ element helpers */
  const kids = (el, name) => (el ? el.children.filter((c) => !name || c.localName === name) : []);
  const kid = (el, name) => { if (!el) return null; for (const c of el.children) if (c.localName === name) return c; return null; };
  const path = (el, ...names) => { let e = el; for (const n of names) { e = kid(e, n); if (!e) return null; } return e; };
  function at(el, n) {
    if (!el || !el.attrs) return null;
    const v = el.attrs[n];
    if (v !== undefined) return v;
    for (const k in el.attrs) { const i = k.indexOf(':'); if (i >= 0 && k.slice(i + 1) === n && k.slice(0, i) !== 'xmlns') return el.attrs[k]; }
    return null;
  }
  const rid = (el, n) => { if (!el || !el.attrs) return null; n = n || 'id'; for (const k in el.attrs) { const i = k.indexOf(':'); if (i > 0 && k.slice(i + 1) === n && k.slice(0, i) !== 'xmlns') return el.attrs[k]; } return null; };
  const num = (el, n, d) => { const v = at(el, n); if (v == null || v === '') return d; const x = parseFloat(v); return isNaN(x) ? d : x; };
  const bool = (el, n, d) => { const v = at(el, n); if (v == null) return d; return v === '1' || v === 'true' || v === 'on'; };
  const valOf = (el, d) => (el ? (at(el, 'val') != null ? at(el, 'val') : d) : undefined);
  /** <b/>, <b val="1"/>, <b val="0"/> */
  const flag = (el) => (el ? at(el, 'val') == null || at(el, 'val') === '1' || at(el, 'val') === 'true' : undefined);
  R.helpers = { kids, kid, path, at, rid, num, bool };

  /* ------------------------------------------------------------ package */
  class Pkg {
    constructor(zip) { this.zip = zip; this.cache = new Map(); }
    has(p) { return this.zip.has(p) || this.find(p) != null; }
    find(p) {
      if (this.zip.has(p)) return p;
      const low = p.toLowerCase();
      for (const k of this.zip.keys()) if (k.toLowerCase() === low) return k;
      return null;
    }
    async bytes(p) { const k = this.find(p); if (!k) return null; return this.zip.get(k).bytes(); }
    async text(p) { const b = await this.bytes(p); return b ? XML.decode(b) : null; }
    async xml(p) {
      const t = await this.text(p);
      if (t == null) return null;
      try { return XML.parse(t); } catch (e) {
        const clean = t.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
        try { return XML.parse(clean); } catch (e2) { return null; }
      }
    }
    async rels(p) {
      const i = p.lastIndexOf('/');
      const rp = (i < 0 ? '' : p.slice(0, i + 1)) + '_rels/' + p.slice(i + 1) + '.rels';
      if (this.cache.has(rp)) return this.cache.get(rp);
      const out = new Map();
      const el = await this.xml(rp);
      for (const r of kids(el, 'Relationship')) {
        const mode = at(r, 'TargetMode');
        const target = at(r, 'Target') || '';
        out.set(at(r, 'Id'), { type: (at(r, 'Type') || '').replace(/^.*\//, ''), target: mode === 'External' ? target : resolve(p, target), external: mode === 'External', raw: target });
      }
      this.cache.set(rp, out);
      return out;
    }
  }
  function resolve(base, target) {
    target = target.replace(/\\/g, '/');
    if (target[0] === '/') return target.slice(1);
    const parts = base.split('/');
    parts.pop();
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.' && seg !== '') parts.push(decodeURIComponent(seg));
    }
    return parts.join('/');
  }
  R.resolve = resolve;
  const relOfType = (rels, type) => { for (const r of rels.values()) if (r.type === type) return r; return null; };
  const relsOfType = (rels, type) => Array.from(rels.values()).filter((r) => r.type === type);

  /* ------------------------------------------------------------ colours, fonts, fills, borders */
  function color(el) {
    if (!el) return undefined;
    if (at(el, 'auto') === '1' || at(el, 'auto') === 'true') return { auto: true };
    const tint = num(el, 'tint', 0);
    let c = null;
    /* with both theme and rgb, Excel goes by the theme slot (LibreOffice tdf#113271) */
    if (at(el, 'theme') != null && at(el, 'theme') !== '') c = { theme: +at(el, 'theme') };
    else if (at(el, 'rgb') != null) { let v = at(el, 'rgb').replace('#', ''); if (v.length === 6) v = 'FF' + v; c = { rgb: v.toUpperCase() }; }
    else if (at(el, 'indexed') != null) c = { indexed: +at(el, 'indexed') };
    else return undefined;
    if (tint) c.tint = tint;
    return c;
  }
  R.color = color;
  function font(el) {
    if (!el) return null;
    const f = {};
    for (const c of el.children) {
      switch (c.localName) {
        case 'name': case 'rFont': f.name = at(c, 'val'); break;
        case 'sz': f.sz = num(c, 'val', 11); break;
        case 'b': f.b = flag(c); break;
        case 'i': f.i = flag(c); break;
        case 'strike': f.strike = flag(c); break;
        case 'u': { const v = at(c, 'val'); f.u = v == null ? 'single' : v === 'none' ? undefined : v; break; }
        case 'color': f.color = color(c); break;
        case 'vertAlign': { const v = at(c, 'val'); if (v && v !== 'baseline') f.vert = v; break; }
        case 'family': f.family = num(c, 'val', undefined); break;
        case 'scheme': { const v = at(c, 'val'); if (v && v !== 'none') f.scheme = v; break; }
        case 'charset': f.charset = num(c, 'val', undefined); break;
        case 'outline': f.outline = flag(c); break;
        case 'shadow': f.shadow = flag(c); break;
        case 'condense': case 'extend': break;
        default: break;
      }
    }
    for (const k of Object.keys(f)) if (f[k] === undefined || f[k] === false) delete f[k];
    return f;
  }
  R.font = font;
  function fill(el, dxf) {
    if (!el) return null;
    const pf = kid(el, 'patternFill');
    if (pf) {
      let pattern = at(pf, 'patternType');
      const fg = color(kid(pf, 'fgColor')), bg = color(kid(pf, 'bgColor'));
      if (dxf) {
        /* differential fills: a solid fill's colour is in bgColor */
        if (!pattern && !fg && !bg) return null;
        const o = { pattern: pattern || 'solid' };
        if (fg) o.fg = fg;
        if (bg) o.bg = bg;
        if (o.pattern === 'solid' && !fg && bg) o.fg = bg;
        return o;
      }
      if (!pattern || pattern === 'none') return null;
      const o = { pattern };
      if (fg) o.fg = fg;
      if (bg) o.bg = bg;
      return o;
    }
    const gf = kid(el, 'gradientFill');
    if (gf) {
      const g = { type: at(gf, 'type') || 'linear', degree: num(gf, 'degree', 0), stops: kids(gf, 'stop').map((s) => ({ pos: num(s, 'position', 0), color: color(kid(s, 'color')) || { rgb: 'FF000000' } })) };
      for (const k of ['left', 'right', 'top', 'bottom']) if (at(gf, k) != null) g[k] = num(gf, k, 0);
      return { gradient: g };
    }
    return null;
  }
  R.fill = fill;
  const SIDES = { left: 'l', right: 'r', top: 't', bottom: 'b', diagonal: 'd', start: 'l', end: 'r', vertical: 'v', horizontal: 'h' };
  function border(el) {
    if (!el) return null;
    const b = {};
    for (const c of el.children) {
      const k = SIDES[c.localName];
      if (!k) continue;
      const st = at(c, 'style');
      if (!st || st === 'none') continue;
      b[k] = { style: st };
      const col = color(kid(c, 'color'));
      if (col) b[k].color = col;
    }
    if (b.d) { b.du = bool(el, 'diagonalUp', false); b.dd = bool(el, 'diagonalDown', false); if (!b.du && !b.dd) delete b.d; if (!b.du) delete b.du; if (!b.dd) delete b.dd; }
    return Object.keys(b).length ? b : null;
  }
  R.border = border;
  function align(el) {
    if (!el) return null;
    const a = {};
    const h = at(el, 'horizontal'); if (h && h !== 'general') a.h = h;
    const v = at(el, 'vertical'); if (v && v !== 'bottom') a.v = v;
    if (bool(el, 'wrapText', false)) a.wrap = true;
    const ind = num(el, 'indent', 0); if (ind) a.indent = ind;
    const rot = num(el, 'textRotation', 0); if (rot) a.rot = rot;
    if (bool(el, 'shrinkToFit', false)) a.shrink = true;
    const ro = num(el, 'readingOrder', 0); if (ro) a.rtl = ro === 2 ? true : false;
    if (a.rtl === false) a.ltr = true, delete a.rtl;
    if (bool(el, 'justifyLastLine', false)) a.justLast = true;
    return Object.keys(a).length ? a : null;
  }
  function prot(el) {
    if (!el) return null;
    const p = {};
    if (at(el, 'locked') != null && !bool(el, 'locked', true)) p.locked = false;
    if (bool(el, 'hidden', false)) p.hidden = true;
    return Object.keys(p).length ? p : null;
  }

  /* ------------------------------------------------------------ theme */
  function readTheme(el) {
    const th = { colors: M.DEFAULT_THEME.colors.slice(), major: M.DEFAULT_THEME.major, minor: M.DEFAULT_THEME.minor, name: 'Office' };
    if (!el) return th;
    th.name = at(el, 'name') || 'Office';
    const elems = kid(el, 'themeElements');
    const cs = kid(elems, 'clrScheme');
    if (cs) {
      const get = (n) => {
        const c = kid(cs, n);
        const e = c && c.children[0];
        if (!e) return null;
        if (e.localName === 'srgbClr') return '#' + (at(e, 'val') || '000000').toUpperCase();
        if (e.localName === 'sysClr') return '#' + (at(e, 'lastClr') || (at(e, 'val') === 'window' ? 'FFFFFF' : '000000')).toUpperCase();
        return null;
      };
      const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      order.forEach((n, i) => { const v = get(n); if (v) th.colors[i] = v; });
      th.scheme = {};
      for (const n of order) th.scheme[n] = th.colors[order.indexOf(n)];
    }
    const fs = kid(elems, 'fontScheme');
    if (fs) {
      const mj = path(fs, 'majorFont', 'latin'), mn = path(fs, 'minorFont', 'latin');
      if (mj && at(mj, 'typeface')) th.major = at(mj, 'typeface');
      if (mn && at(mn, 'typeface')) th.minor = at(mn, 'typeface');
    }
    return th;
  }

  /* ------------------------------------------------------------ styles */
  function readStyles(el, wb) {
    const out = { xf: [0], dxfs: [], numFmts: {}, cellStyles: [], namedStyles: [], tableStyles: null };
    if (!el) return out;
    const nfs = kid(el, 'numFmts');
    for (const n of kids(nfs, 'numFmt')) out.numFmts[+at(n, 'numFmtId')] = at(n, 'formatCode') || 'General';
    const fonts = kids(kid(el, 'fonts'), 'font').map(font);
    const fills = kids(kid(el, 'fills'), 'fill').map((f) => fill(f));
    const borders = kids(kid(el, 'borders'), 'border').map(border);
    const colors = kid(el, 'colors');
    const previewColors = el.namespaceURI === 'http://schemas.microsoft.com/office/excel/2006/2' && kid(colors, 'themeColors');
    if (previewColors) {
      wb.extra.previewThemeColors = kids(previewColors, 'rgbColor').map(e => at(e, 'rgb'));
      const names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      wb.theme.scheme ||= {};
      wb.extra.previewThemeColors.forEach((color, i) => { if (color && names[i]) wb.theme.scheme[names[i]] = wb.theme.colors[i] = '#' + color.slice(-6).toUpperCase(); });
    }
    const idx = kid(colors, 'indexedColors');
    if (idx) {
      const pal = M.INDEXED.slice();
      kids(idx, 'rgbColor').forEach((c, i) => { const v = at(c, 'rgb'); if (v) pal[i] = '#' + v.slice(-6).toUpperCase(); });
      wb.palette = pal;
    }
    const mru = kids(kid(colors, 'mruColors'), 'color').map(color).filter(Boolean);
    if (mru.length) wb.mruColors = mru;
    const nfCode = (id) => (out.numFmts[id] != null ? out.numFmts[id] : NF.BUILTIN[id] != null ? NF.BUILTIN[id] : 'General');
    /* named cell styles */
    const csXfs = kids(kid(el, 'cellStyleXfs'), 'xf');
    const styleOf = (x, parent) => {
      const st = {
        font: Object.assign({}, fonts[num(x, 'fontId', 0)] || fonts[0] || { name: 'Calibri', sz: 11 }),
        fill: fills[num(x, 'fillId', 0)] || null,
        border: borders[num(x, 'borderId', 0)] || null,
        align: align(kid(x, 'alignment')),
        nf: nfCode(num(x, 'numFmtId', 0)),
        prot: prot(kid(x, 'protection')),
      };
      /* In cellXfs, applyX="0" does not mean "take X from the parent style": the xf's own fontId, fillId, … are what
         Excel shows; the flag only says whether a later change to the parent style should flow into this format
         (MS-OI29500 §18.8.45). Files from ClosedXML, EPPlus and others write applyFill="0" next to a real fill. */
      void parent;
      if (st.font) for (const k of Object.keys(st.font)) if (st.font[k] === undefined) delete st.font[k];
      return st;
    };
    const named = csXfs.map((x) => styleOf(x, null));
    const csNames = {};
    for (const cs of kids(kid(el, 'cellStyles'), 'cellStyle')) {
      const xid = num(cs, 'xfId', 0);
      const o = { name: at(cs, 'name') || 'Style ' + xid, xfId: xid };
      if (at(cs, 'builtinId') != null) o.builtinId = num(cs, 'builtinId', 0);
      if (bool(cs, 'customBuiltin', false)) o.customBuiltin = true;
      if (bool(cs, 'hidden', false)) o.hidden = true;
      out.cellStyles.push(o);
      if (csNames[xid] == null) csNames[xid] = o.name;
    }
    /* cell formats */
    const xfs = kids(kid(el, 'cellXfs'), 'xf');
    /* a stylesheet without cell formats, or a default font without a name or size, means Excel's own default
       (the theme's body font at 11 pt), not Ledger's Arial 10 for new workbooks */
    const base = xfs.length ? styleOf(xfs[0], named[num(xfs[0], 'xfId', 0)]) : styleOf({ children: [], attrs: {} }, null);
    if (base && base.font && !base.font.name) base.font.name = (wb.theme && wb.theme.minor) || 'Calibri';
    if (base && base.font && base.font.sz == null) base.font.sz = 11;
    if (base) {
      /* the workbook's default font (xf 0) */
      wb.styles = new M.StyleTable(base);
    }
    out.xf = xfs.map((x, i) => {
      const xfId = num(x, 'xfId', 0);
      const st = styleOf(x, named[xfId]);
      if (xfId && csNames[xfId] && csNames[xfId] !== 'Normal') st.xs = csNames[xfId];
      return i === 0 && base ? 0 : wb.styles.add(st);
    });
    out.namedStyles = out.cellStyles.map((cs) => ({ name: cs.name, builtinId: cs.builtinId, hidden: cs.hidden, customBuiltin: cs.customBuiltin, style: named[cs.xfId] ? wb.styles.add(named[cs.xfId]) : 0 }));
    /* differential formats (conditional formatting, tables) */
    for (const d of kids(kid(el, 'dxfs'), 'dxf')) {
      const o = {};
      const f = kid(d, 'font'); if (f) o.font = font(f);
      const fl = kid(d, 'fill'); if (fl) { const x = fill(fl, true); if (x) o.fill = x; }
      const b = kid(d, 'border'); if (b) { const x = border(b); if (x) o.border = x; }
      const n = kid(d, 'numFmt'); if (n) o.nf = at(n, 'formatCode');
      const a = kid(d, 'alignment'); if (a) o.align = align(a);
      out.dxfs.push(o);
    }
    const ts = kid(el, 'tableStyles');
    if (ts) {
      out.tableStyles = { def: at(ts, 'defaultTableStyle'), defPivot: at(ts, 'defaultPivotStyle'), custom: kids(ts, 'tableStyle').map((t) => ({ name: at(t, 'name'), pivot: at(t, 'pivot') !== '0', table: at(t, 'table') !== '0', elements: kids(t, 'tableStyleElement').map((e) => ({ type: at(e, 'type'), dxf: num(e, 'dxfId', -1), size: num(e, 'size', 1) })) })) };
    }
    return out;
  }

  /* ------------------------------------------------------------ shared strings */
  const RPR_KEYS = new Set(['rFont', 'name', 'sz', 'b', 'i', 'strike', 'u', 'color', 'vertAlign', 'family', 'scheme', 'charset', 'outline', 'shadow']);
  /** parse <si> / <is> content scanned from text [from, to) → string or {text, runs} */
  function readSST(text) {
    const out = [];
    let runs = null, cur = null, inT = false, inRPh = 0, buf = '', rpr = null, rprStack = null;
    let siText = '';
    XML.scan(text, 0, text.length, (name, kind, attrs, before) => {
      if (inT && before) { const t = XML.unx(XML.chars(before)); if (!inRPh) { if (cur) cur.t += t; siText += t; } }
      switch (name) {
        case 'si': case 'sstItem':
          if (kind === 1) { runs = null; cur = null; siText = ''; }
          else if (kind === 2) out.push(runs && runs.some((r) => r.font) ? { text: siText, runs } : siText);
          else out.push('');
          break;
        case 'r': if (kind === 1 && !inRPh) { if (!runs) runs = []; cur = { t: '' }; runs.push(cur); } else if (kind === 2) cur = null; break;
        case 'rPr': if (kind === 1) { rpr = {}; rprStack = []; } else if (kind === 2) { if (cur) cur.font = finishRpr(rpr); rpr = null; } break;
        case 't':
          if (kind === 1) {
            inT = true;
            if (!cur && !inRPh) { if (!runs) runs = []; cur = { t: '' }; runs.push(cur); cur._plain = true; }
          } else if (kind === 2) { inT = false; if (cur && cur._plain) { delete cur._plain; cur = null; } }
          break;
        case 'rPh': if (kind === 1) inRPh++; else if (kind === 2) inRPh--; break;
        default:
          if (rpr && RPR_KEYS.has(name) && kind !== 2) rpr[name] = XML.attrs(attrs);
          else if (rpr && kind === 1) rprStack.push(name);
          break;
      }
      void buf;
    });
    return out;
  }
  function finishRpr(o) {
    const f = {};
    const v = (k) => (o[k] && o[k].val);
    const fl = (k) => (o[k] ? o[k].val == null || o[k].val === '1' || o[k].val === 'true' : undefined);
    if (o.rFont || o.name) f.name = v('rFont') || v('name');
    if (o.sz) f.sz = parseFloat(v('sz'));
    if (fl('b')) f.b = true;
    if (fl('i')) f.i = true;
    if (fl('strike')) f.strike = true;
    if (o.u) { const u = o.u.val == null ? 'single' : o.u.val; if (u !== 'none') f.u = u; }
    if (o.color) { const c = colorAttrs(o.color); if (c) f.color = c; }
    if (o.vertAlign && v('vertAlign') !== 'baseline') f.vert = v('vertAlign');
    if (o.family) f.family = +v('family');
    if (o.scheme && v('scheme') !== 'none') f.scheme = v('scheme');
    if (fl('outline')) f.outline = true;
    if (fl('shadow')) f.shadow = true;
    return Object.keys(f).length ? f : undefined;
  }
  function colorAttrs(a) {
    if (a.auto === '1') return { auto: true };
    let c = null;
    if (a.theme != null && a.theme !== '') c = { theme: +a.theme };
    else if (a.rgb != null) { let v = a.rgb.replace('#', ''); if (v.length === 6) v = 'FF' + v; c = { rgb: v.toUpperCase() }; }
    else if (a.indexed != null) c = { indexed: +a.indexed };
    if (c && a.tint && +a.tint) c.tint = +a.tint;
    return c;
  }
  R.readSST = readSST;

  /* ------------------------------------------------------------ cell data */
  function cellRef(s) {
    let i = 0, c = 0;
    const n = s.length;
    while (i < n) { const ch = s.charCodeAt(i); if (ch >= 65 && ch <= 90) c = c * 26 + (ch - 64); else if (ch >= 97 && ch <= 122) c = c * 26 + (ch - 96); else break; i++; }
    let r = 0;
    for (; i < n; i++) { const ch = s.charCodeAt(i); if (ch === 36) continue; r = r * 10 + (ch - 48); }
    return { r: r - 1, c: c - 1 };
  }
  const reAttrFast = /([\w:]+)="([^"]*)"/g;
  function attrsFast(s) {
    const o = {};
    if (!s) return o;
    reAttrFast.lastIndex = 0;
    let m;
    while ((m = reAttrFast.exec(s))) o[m[1]] = m[2].indexOf('&') < 0 ? m[2] : XML.unescape(m[2]);
    if (s.indexOf("'") >= 0) { const o2 = XML.attrs(s); for (const k in o2) if (o[k] === undefined) o[k] = o2[k]; }
    return o;
  }
  function isoToSerial(s, date1904) {
    const m = /^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d(?:\.\d+)?))?)?/.exec(s);
    if (!m) return null;
    let v = NF.dateToSerial(+m[1], +m[2], +m[3], date1904);
    if (m[4]) v += (+m[4] * 3600 + +m[5] * 60 + (m[6] ? +m[6] : 0)) / 86400;
    return v;
  }
  /** parse <sheetData> into the sheet */
  function readSheetData(text, from, to, sh, ctx, rowUnit = 1) {
    const { sst, xf } = ctx;
    const shared = ctx.sharedF = new Map();
    let r = -1, c = -1, row = null;
    let cell = null, cAttr = null, fAttr = null, vText = null, fText = null, isText = null, isRuns = null, inIs = false, inT = false, inRPh = 0, rpr = null, curRun = null;
    let maxR = -1, maxC = -1;
    const arrays = ctx.arrays = [];
    const commit = () => {
      const t = cAttr.t;
      let v = null;
      if (t === 's') {
        const k = vText == null ? -1 : parseInt(vText, 10);
        const sv = sst[k];
        if (sv != null && ctx.wb.pkg) cell.keep = { source: ctx.wb.pkg.id, sst: k };
        if (sv == null) v = '';
        else if (typeof sv === 'string') v = sv;
        else { v = sv.text; cell.rt = sv.runs; }
      } else if (t === 'str') v = vText == null ? '' : XML.unx(XML.chars(vText));
      else if (t === 'inlineStr') { v = isText != null ? isText : vText != null ? XML.chars(vText) : ''; if (isRuns && isRuns.some((x) => x.font)) cell.rt = isRuns; }
      else if (t === 'b') v = vText != null && (vText.trim() === '1' || vText.trim() === 'true');
      else if (t === 'e') v = vText != null ? M.err(vText.trim()) : null;
      else if (t === 'd') { v = vText != null ? isoToSerial(vText.trim(), ctx.wb.date1904) : null; if (v == null && vText != null) v = vText; }
      else if (vText != null && vText !== '') { v = +vText; if (isNaN(v)) { const tv = vText.trim(); v = tv === '' ? null : isNaN(+tv) ? XML.unescape(tv) : +tv; } }
      cell.v = v;
      if (fAttr) {
        const ft = fAttr.t;
        let ftext = fText != null ? XML.chars(fText) : '';
        if (ftext && ftext[0] === '=') ftext = ftext.slice(1);
        if (ft === 'shared') {
          const si = fAttr.si;
          if (ftext.trim() && fAttr.ref) {
            shared.set(si, { r, c, text: ftext, ast: undefined });
            cell.f = ftext;
          } else if (ftext.trim()) { cell.f = ftext; if (!shared.has(si)) shared.set(si, { r, c, text: ftext, ast: undefined }); }
          else {
            const m = shared.get(si);
            if (m) {
              if (m.ast === undefined) { try { m.ast = F.parse(m.text); } catch (e) { m.ast = null; } }
              if (m.ast) {
                const a = F.shift(m.ast, r - m.r, c - m.c);
                cell.f = F.toText(a);
                cell.ast = a; cell._fsrc = cell.f;
              } else cell.f = m.text;
            }
          }
        } else if (ft === 'array') {
          cell.f = ftext;
          const rg = fAttr.ref ? F.parseRange(fAttr.ref) : { r1: r, c1: c, r2: r, c2: c };
          cell.af = rg || { r1: r, c1: c, r2: r, c2: c };
          if (cAttr.cm) cell.dyn = true; /* dynamic array (spills) */
          arrays.push({ r, c, cell });
        } else if (ft === 'dataTable') {
          cell.dt = { ref: fAttr.ref, dt2D: fAttr.dt2D === '1', dtr: fAttr.dtr === '1', r1: fAttr.r1, r2: fAttr.r2, del1: fAttr.del1 === '1', del2: fAttr.del2 === '1', ca: fAttr.ca === '1' };
        } else if (ftext.trim()) {
          cell.f = ftext;
        }
        if (cell.f != null && (vText == null && t !== 'inlineStr')) cell.dirty = true; /* never calculated */
        if (fAttr.ca === '1' && cell.f != null) cell.ca = true;
      }
      const s = cAttr.s ? +cAttr.s : 0;
      if (s) { const id = xf[s]; if (id) cell.s = id; }
      if (cAttr.s != null && xf[s] != null && ctx.wb.pkg) cell.keep = { ...cell.keep, source: ctx.wb.pkg.id, style: s, modelStyle: cell.s || 0 };
      if (cell.v == null && cell.f == null && !cell.s && cell.keep?.style == null && !cell.dt) { /* nothing to keep */ }
      else {
        if (!row) row = sh.rowObj(r);
        row.cells[c] = cell;
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
      }
    };
    XML.scan(text, from, to, (name, kind, attrs, before) => {
      if (inT && before) { const t = XML.unx(XML.chars(before)); if (!inRPh) { isText = (isText || '') + t; if (curRun) curRun.t += t; } }
      switch (name) {
        case 'row': {
          if (kind === 2) { row = null; break; }
          const a = attrsFast(attrs);
          r = a.r ? +a.r - 1 : r + 1;
          c = -1;
          row = null;
          const ht = a.ht != null ? parseFloat(a.ht) / rowUnit : null;
          const hidden = a.hidden === '1' || a.hidden === 'true';
          const lvl = a.outlineLevel ? +a.outlineLevel : 0;
          const coll = a.collapsed === '1' || a.collapsed === 'true';
          const custFmt = (a.customFormat === '1' || a.customFormat === 'true') && a.s;
          if (ht != null || hidden || lvl || coll || custFmt || a.thickBot || a.thickTop) {
            row = sh.rowObj(r);
            if (ht != null && !isNaN(ht)) { row.ht = ht; if (a.customHeight === '1' || a.customHeight === 'true') row.customHeight = true; }
            if (hidden) row.hidden = true;
            if (lvl) row.level = lvl;
            if (coll) row.collapsed = true;
            if (custFmt) { const id = xf[+a.s]; if (id) row.s = id; }
            if (custFmt && xf[+a.s] != null && ctx.wb.pkg) row.keep = { source: ctx.wb.pkg.id, style: +a.s, modelStyle: row.s || 0 };
            if (a.thickBot === '1') row.thickBot = true;
            if (a.thickTop === '1') row.thickTop = true;
            if (r > maxR && (row.s || row.ht != null)) { /* styled empty rows extend the used range only through cells */ }
          }
          break;
        }
        case 'c': {
          if (kind === 2) { if (cell) commit(); cell = null; break; }
          cAttr = attrsFast(attrs);
          if (cAttr.r) { const p = cellRef(cAttr.r); if (p.r >= 0 && p.r !== r) { r = p.r; row = null; } c = p.c; } else c = c + 1;
          cell = { v: null };
          fAttr = null; vText = null; fText = null; isText = null; isRuns = null;
          if (kind === 3) { commit(); cell = null; }
          break;
        }
        case 'v': if (kind === 2) vText = before; else if (kind === 3) vText = ''; break;
        case 'f':
          if (kind === 1) fAttr = attrsFast(attrs);
          else if (kind === 2) fText = before;
          else { fAttr = attrsFast(attrs); fText = ''; }
          break;
        case 'is': inIs = kind === 1; if (kind === 1) { isText = ''; isRuns = []; } break;
        case 't': if (inIs) { if (kind === 1) { inT = true; if (!curRun && !inRPh) { curRun = { t: '' }; isRuns.push(curRun); curRun._plain = true; } } else if (kind === 2) { inT = false; if (curRun && curRun._plain) { delete curRun._plain; curRun = null; } } } break;
        case 'r': if (inIs) { if (kind === 1) { curRun = { t: '' }; isRuns.push(curRun); } else if (kind === 2) curRun = null; } break;
        case 'rPr': if (inIs) { if (kind === 1) rpr = {}; else if (kind === 2) { if (curRun) curRun.font = finishRpr(rpr); rpr = null; } } break;
        case 'rPh': if (kind === 1) inRPh++; else if (kind === 2) inRPh--; break;
        default:
          if (rpr && RPR_KEYS.has(name) && kind !== 2) rpr[name] = XML.attrs(attrs);
          break;
      }
    });
    if (maxR > sh.maxR) sh.maxR = maxR;
    if (maxC > sh.maxC) sh.maxC = maxC;
    /* array formula members point at their master */
    for (const a of arrays) {
      const rg = a.cell.af;
      if ((rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1) > 1e6) continue;
      for (let rr = rg.r1; rr <= rg.r2; rr++) for (let cc = rg.c1; cc <= rg.c2; cc++) {
        if (rr === a.r && cc === a.c) continue;
        let m = sh.get(rr, cc);
        if (!m) { m = { v: null }; sh.put(rr, cc, m); }
        m.am = { r: a.r, c: a.c };
        delete m.f;
      }
    }
  }

  /* ------------------------------------------------------------ worksheet */
  function sliceSheetData(text) {
    const re = /<((?:[\w.-]+:)?)sheetData\b[^>]*?(\/?)>/g;
    const m = re.exec(text);
    if (!m) return { rest: text, from: 0, to: 0 };
    if (m[2] === '/') return { rest: text.slice(0, m.index) + text.slice(re.lastIndex), from: 0, to: 0 };
    const endTag = '</' + m[1] + 'sheetData>';
    const e = text.indexOf(endTag, re.lastIndex);
    const end = e < 0 ? text.length : e;
    return { rest: text.slice(0, m.index) + text.slice(e < 0 ? text.length : e + endTag.length), from: re.lastIndex, to: end };
  }

  function readPageSetup(ps, p) {
    if (!ps) return;
    if (at(ps, 'orientation')) p.orientation = at(ps, 'orientation') === 'landscape' ? 'landscape' : 'portrait';
    if (at(ps, 'paperSize')) p.paper = num(ps, 'paperSize', 1);
    if (at(ps, 'scale')) p.scale = num(ps, 'scale', 100);
    if (at(ps, 'fitToWidth') != null) p.fitW = num(ps, 'fitToWidth', 1); else if (p.fit) p.fitW = 1;
    if (at(ps, 'fitToHeight') != null) p.fitH = num(ps, 'fitToHeight', 1); else if (p.fit) p.fitH = 1;
    if (bool(ps, 'useFirstPageNumber', false) && at(ps, 'firstPageNumber') != null) p.firstPage = num(ps, 'firstPageNumber', 1);
    if (at(ps, 'pageOrder') === 'overThenDown') p.pageOrder = 'overThenDown';
    if (bool(ps, 'blackAndWhite', false)) p.bw = true;
    if (bool(ps, 'draft', false)) p.draft = true;
    if (at(ps, 'cellComments')) p.comments = at(ps, 'cellComments');
    if (at(ps, 'errors')) p.errors = at(ps, 'errors');
    if (at(ps, 'horizontalDpi')) p.dpi = num(ps, 'horizontalDpi', 600);
  }
  async function readWorksheet(pkg, part, sh, ctx) {
    const text = await pkg.text(part);
    if (text == null) return;
    const sd = sliceSheetData(text);
    let el;
    try { el = XML.parse(sd.rest); } catch (e) { el = XML.parse(sd.rest.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')); }
    const preview2005 = el.namespaceURI === 'http://schemas.microsoft.com/office/excel/2005/8/worksheet';
    const rowUnit = preview2005 ? 20 : 1;
    L.sheetObjects.capture(sh, part, sd.rest);
    L.opc.captureAC(el, sd.rest);
    resolveAC(el);
    const rels = await pkg.rels(part);
    /* sheet properties */
    const pr = kid(el, 'sheetPr');
    if (pr) {
      const tc = color(kid(pr, 'tabColor')); if (tc) sh.tabColor = tc;
      const op = kid(pr, 'outlinePr');
      if (op) { sh.outline.below = bool(op, 'summaryBelow', true); sh.outline.right = bool(op, 'summaryRight', true); if (bool(op, 'applyStyles', false)) sh.outline.applyStyles = true; }
      const ps = kid(pr, 'pageSetUpPr');
      if (ps && bool(ps, 'fitToPage', false)) sh.print.fit = true;
      if (at(pr, 'codeName')) sh.codeName = at(pr, 'codeName');
      if (bool(pr, 'filterMode', false)) sh.filterMode = true;
      if (bool(pr, 'transitionEvaluation', false)) sh.extra.transitionEvaluation = true;
    }
    /* views */
    const sv = path(el, 'sheetViews', 'sheetView');
    if (sv) {
      const v = sh.view;
      if (!bool(sv, 'showGridLines', true)) v.grid = false;
      if (!bool(sv, 'showRowColHeaders', true)) v.headings = false;
      if (!bool(sv, 'showZeros', true)) v.zeros = false;
      if (bool(sv, 'rightToLeft', false)) v.rtl = true;
      if (bool(sv, 'showFormulas', false)) v.formulas = true;
      if (bool(sv, 'tabSelected', false)) sh.selected = true;
      if (at(sv, 'zoomScale')) v.zoom = num(sv, 'zoomScale', 100) || 100;
      if (at(sv, 'view') === 'pageBreakPreview') { v.pageBreakPreview = true; if (at(sv, 'zoomScaleSheetLayoutView')) v.zoom = num(sv, 'zoomScaleSheetLayoutView', 60) || 60; else if (!at(sv, 'zoomScale')) v.zoom = 60; }
      if (at(sv, 'view') === 'pageLayout') v.pageLayout = true;
      if (!bool(sv, 'showOutlineSymbols', true)) v.outlineSymbols = false;
      if (at(sv, 'colorId') && !bool(sv, 'defaultGridColor', true)) v.gridColor = { indexed: num(sv, 'colorId', 64) };
      if (at(sv, 'topLeftCell')) { const p = F.parseCell(at(sv, 'topLeftCell')); if (p) v.top = { r: p.r, c: p.c }; }
      const pane = kid(sv, 'pane');
      if (pane) {
        const xs = num(pane, 'xSplit', 0), ys = num(pane, 'ySplit', 0);
        const state = at(pane, 'state') || 'split';
        const tl = at(pane, 'topLeftCell') ? F.parseCell(at(pane, 'topLeftCell')) : null;
        if (state === 'frozen' || state === 'frozenSplit') v.freeze = { r: Math.round(ys), c: Math.round(xs), top: tl ? { r: tl.r, c: tl.c } : null };
        else if (xs || ys) v.split = { x: xs, y: ys, top: tl ? { r: tl.r, c: tl.c } : null }; /* twips */
      }
      const sels = kids(sv, 'selection');
      const ap = pane ? at(pane, 'activePane') : null;
      const sel = sels.find((s) => at(s, 'pane') === ap) || sels[sels.length - 1];
      if (sel) {
        const ac = at(sel, 'activeCell') ? F.parseCell(at(sel, 'activeCell')) : null;
        const rg = M.parseSqref(at(sel, 'sqref') || at(sel, 'activeCell') || 'A1');
        if (ac || rg.length) v.sel = { r: ac ? ac.r : rg[0].r1, c: ac ? ac.c : rg[0].c1, ranges: rg.length ? rg : [{ r1: ac.r, c1: ac.c, r2: ac.r, c2: ac.c }] };
      }
    }
    /* format */
    const fp = kid(el, 'sheetFormatPr');
    if (fp) {
      if (at(fp, 'defaultColWidth') != null) sh.defColW = num(fp, 'defaultColWidth', null);
      if (at(fp, 'baseColWidth') != null) sh.baseColW = num(fp, 'baseColWidth', 8);
      if (at(fp, 'defaultRowHeight') != null && (bool(fp, 'customHeight', false) || bool(fp, 'zeroHeight', false))) sh.defRowH = num(fp, 'defaultRowHeight', 0) / rowUnit;
      if (at(fp, 'defaultRowHeight') != null) sh.fileDefRowH = num(fp, 'defaultRowHeight', 0) / rowUnit;
      if (bool(fp, 'zeroHeight', false)) sh.zeroHeight = true;
      if (at(fp, 'outlineLevelRow')) sh.outline.levelRow = num(fp, 'outlineLevelRow', 0);
      if (at(fp, 'outlineLevelCol')) sh.outline.levelCol = num(fp, 'outlineLevelCol', 0);
    }
    L.preserve.keepSheetFormat(sh, fp);
    for (const col of kids(kid(el, 'cols'), 'col')) {
      const offset = preview2005 ? 0 : 1;
      const mn = Math.max(0, num(col, 'min', offset) - offset), mx = Math.min(num(col, 'max', offset) - offset, 16383);
      const o = {};
      if (at(col, 'width') != null) o.w = num(col, 'width', 8.43);
      // The 2005 preview stores column widths in 1/256-character units and
      // numbers columns from zero; the 2006 preview already uses OOXML units.
      if (preview2005 && at(col, 'defaultWidth') != null) o.w = num(col, 'defaultWidth', 2158) / 256;
      if (bool(col, 'customWidth', false)) o.custom = true;
      if (bool(col, 'hidden', false)) o.hidden = true;
      if (bool(col, 'bestFit', false)) o.bestFit = true;
      const lvl = num(col, 'outlineLevel', 0); if (lvl) o.level = lvl;
      if (bool(col, 'collapsed', false)) o.collapsed = true;
      const s = num(col, 'style', 0); if (s && ctx.xf[s]) o.s = ctx.xf[s];
      if (at(col, 'style') != null && ctx.xf[s] != null && ctx.wb.pkg) o.keep = { source: ctx.wb.pkg.id, style: s, modelStyle: o.s || 0 };
      if (!Object.keys(o).length) continue;
      if (mx - mn > 2000 && !o.hidden && !o.level && (o.w == null || !o.s)) {
        /* a run to the last column: record it once as a tail default */
        for (let c = mn; c <= Math.min(mx, mn + 255); c++) sh.cols[c] = Object.assign({}, o);
        if (mx >= 16383 || mx - mn > 255) sh.colTail = { from: mn + 256, o };
        continue;
      }
      for (let c = mn; c <= mx; c++) sh.cols[c] = Object.assign({}, o);
    }
    /* cells */
    if (sd.to > sd.from) readSheetData(text, sd.from, sd.to, sh, ctx, rowUnit);
    /* protection */
    const sp = kid(el, 'sheetProtection');
    if (sp && bool(sp, 'sheet', false)) {
      const p = { attrs: Object.assign({}, sp.attrs) };
      sh.protection = p;
    }
    /* merges */
    for (const m of kids(kid(el, 'mergeCells'), 'mergeCell')) { const rg = F.parseRange(at(m, 'ref') || ''); if (rg && (rg.r1 !== rg.r2 || rg.c1 !== rg.c2)) sh.merges.push(rg); }
    /* autofilter */
    const af = kid(el, 'autoFilter');
    if (af) sh.autoFilter = readAutoFilter(af);
    const ss = kid(el, 'sortState');
    if (ss) sh.sortState = readSortState(ss);
    /* hidden rows inside an active filter are "filtered" (SUBTOTAL 1–11 skips them, manual hides it keeps) */
    const markFiltered = (f) => {
      if (!f || !f.ref || !f.cols.some((c) => c.values || c.custom || c.top10 || c.dynamic || c.color || c.icon || c.blank || c.dates)) return;
      for (let r = f.ref.r1 + 1; r <= f.ref.r2; r++) { const row = sh.rows[r]; if (row && row.hidden) row.filtered = true; }
    };
    markFiltered(sh.autoFilter);
    ctx.markFiltered = markFiltered;
    /* conditional formatting */
    for (const cf of kids(el, 'conditionalFormatting')) {
      const ranges = M.parseSqref(at(cf, 'sqref'));
      if (!ranges.length) continue;
      const rules = kids(cf, 'cfRule').map((r) => { const o = readCfRule(r, ctx); L.sheetExtensions.base(o, r, sh); return o; });
      if (rules.length) sh.cf.push({ ranges, rules, pivot: bool(cf, 'pivot', false) || undefined });
    }
    /* data validation */
    for (const d of kids(kid(el, 'dataValidations'), 'dataValidation')) { const o = readDV(d); if (o) sh.dv.push(o); }
    /* hyperlinks */
    for (const h of kids(kid(el, 'hyperlinks'), 'hyperlink')) {
      const rg = F.parseRange(at(h, 'ref') || '');
      if (!rg) continue;
      const o = { ref: rg };
      const id = rid(h, 'id');
      if (id && rels.get(id)) o.target = rels.get(id).raw;
      if (at(h, 'location') != null) o.location = at(h, 'location');
      if (at(h, 'display') != null) o.display = at(h, 'display');
      if (at(h, 'tooltip') != null) o.tooltip = at(h, 'tooltip');
      sh.links.push(o);
    }
    /* printing */
    const po = kid(el, 'printOptions');
    if (po) { sh.print.gridLines = bool(po, 'gridLines', false); sh.print.headings = bool(po, 'headings', false); sh.print.hCenter = bool(po, 'horizontalCentered', false); sh.print.vCenter = bool(po, 'verticalCentered', false); }
    const pm = kid(el, 'pageMargins');
    if (pm) sh.print.margins = { l: num(pm, 'left', 0.7), r: num(pm, 'right', 0.7), t: num(pm, 'top', 0.75), b: num(pm, 'bottom', 0.75), header: num(pm, 'header', 0.3), footer: num(pm, 'footer', 0.3) };
    const ps = kid(el, 'pageSetup');
    readPageSetup(ps, sh.print);
    L.preserve.keepPageSetup(sh, ps);
    const hf = kid(el, 'headerFooter');
    if (hf) {
      const p = sh.print;
      const t = (n) => { const e = kid(hf, n); return e ? XML.unx(e.textContent) : ''; };
      p.header = t('oddHeader'); p.footer = t('oddFooter');
      if (bool(hf, 'differentFirst', false)) { p.diffFirst = true; p.firstHeader = t('firstHeader'); p.firstFooter = t('firstFooter'); }
      if (bool(hf, 'differentOddEven', false)) { p.diffOddEven = true; p.evenHeader = t('evenHeader'); p.evenFooter = t('evenFooter'); }
      if (!bool(hf, 'scaleWithDoc', true)) p.hfScale = false;
      if (!bool(hf, 'alignWithMargins', true)) p.hfAlign = false;
    }
    for (const b of kids(kid(el, 'rowBreaks'), 'brk')) sh.print.rowBreaks.push(num(b, 'id', 0));
    for (const b of kids(kid(el, 'colBreaks'), 'brk')) sh.print.colBreaks.push(num(b, 'id', 0));
    /* comments (legacy) */
    const cm = relOfType(rels, 'comments');
    if (cm) await readComments(pkg, cm.target, sh);
    const vmlRel = (() => { const ld = kid(el, 'legacyDrawing'); return ld && rels.get(rid(ld, 'id')); })();
    if (vmlRel) await readVmlNotes(pkg, vmlRel.target, sh);
    /* Complete threads, including files without legacy note text. */
    const tcm = relOfType(rels, 'threadedComment');
    if (tcm) await readThreaded(pkg, tcm.target, sh, ctx);
    /* tables */
    for (const tp of kids(kid(el, 'tableParts'), 'tablePart')) {
      const r = rels.get(rid(tp, 'id'));
      if (r) await readTable(pkg, r.target, sh, ctx);
    }
    /* drawing (pictures, charts, shapes) */
    const dr = kid(el, 'drawing');
    if (dr) { const r = rels.get(rid(dr, 'id')); if (r) await readDrawing(pkg, r.target, sh, ctx); }
    /* Excel 2010+ extensions: x14 data validations, conditional formats, sparklines */
    const ext = kid(el, 'extLst');
    if (ext) readSheetExt(ext, sh, ctx);
    /* background picture */
    const pic = kid(el, 'picture');
    if (pic) { const r = rels.get(rid(pic, 'id')); if (r) { const b = await pkg.bytes(r.target); if (b) sh.background = ctx.media(r.target, b); } }
    /* parts we preserve verbatim */
    const keep = [];
    for (const n of ['customSheetViews', 'phoneticPr', 'scenarios', 'dataConsolidate', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'oleObjects', 'controls', 'webPublishItems', 'protectedRanges']) {
      const e = kid(el, n);
      if (e) keep.push({ name: n, xml: XML.serialize(e) });
    }
    if (keep.length) sh.extra.keep = keep;
    const scn = kid(el, 'scenarios');
    if (scn) sh.scenarios = kids(scn, 'scenario').map((x) => ({ name: at(x, 'name') || '', comment: at(x, 'comment') || '', user: at(x, 'user') || '', cells: kids(x, 'inputCells').map((ic) => { const p = F.parseCell(at(ic, 'r') || 'A1'); const v = at(ic, 'val'); return { r: p ? p.r : 0, c: p ? p.c : 0, v: v != null && v !== '' && !isNaN(+v) ? +v : v }; }) }));
    L.sheetObjects.read(sh);
  }

  function readAutoFilter(af) {
    const o = { ref: F.parseRange(at(af, 'ref') || ''), cols: [] };
    for (const fc of kids(af, 'filterColumn')) {
      const c = { col: num(fc, 'colId', 0) };
      if (at(fc, 'hiddenButton') === '1') c.hiddenButton = true;
      if (at(fc, 'showButton') === '0') c.hiddenButton = true;
      const fs = kid(fc, 'filters');
      if (fs) {
        c.values = kids(fs, 'filter').map((f) => at(f, 'val'));
        if (bool(fs, 'blank', false)) c.blank = true;
        const dg = kids(fs, 'dateGroupItem');
        if (dg.length) c.dates = dg.map((d) => ({ y: num(d, 'year', 0), m: num(d, 'month', 0), d: num(d, 'day', 0), H: num(d, 'hour', 0), M: num(d, 'minute', 0), S: num(d, 'second', 0), g: at(d, 'dateTimeGrouping') }));
        // Preserve source ordering when a filter mixes date groups with text values.
        if (c.values.length && c.dates?.length) c.filtersKeep = { xml: XML.serialize(fs), values: L.preserve.filterValues(c) };
      }
      const cf = kid(fc, 'customFilters');
      if (cf) c.custom = { and: bool(cf, 'and', false), list: kids(cf, 'customFilter').map((f) => ({ op: at(f, 'operator') || 'equal', val: at(f, 'val') })) };
      const t10 = kid(fc, 'top10');
      if (t10) c.top10 = { top: at(t10, 'top') !== '0', percent: bool(t10, 'percent', false), val: num(t10, 'val', 10), filterVal: num(t10, 'filterVal', undefined) };
      const dyn = kid(fc, 'dynamicFilter');
      if (dyn) c.dynamic = { type: at(dyn, 'type'), val: num(dyn, 'val', undefined), maxVal: num(dyn, 'maxVal', undefined) };
      const clr = kid(fc, 'colorFilter');
      if (clr) c.color = { dxf: num(clr, 'dxfId', -1), cell: at(clr, 'cellColor') !== '0' };
      const icon = kid(fc, 'iconFilter');
      if (icon) c.icon = { set: at(icon, 'iconSet'), id: num(icon, 'iconId', 0) };
      o.cols.push(c);
    }
    const ss = kid(af, 'sortState');
    if (ss) o.sort = readSortState(ss);
    return o;
  }
  function readSortState(ss) {
    return { ref: F.parseRange(at(ss, 'ref') || ''), columnSort: bool(ss, 'columnSort', false), caseSensitive: bool(ss, 'caseSensitive', false), keys: kids(ss, 'sortCondition').map((k) => ({ ref: F.parseRange(at(k, 'ref') || ''), desc: bool(k, 'descending', false), by: at(k, 'sortBy') || 'value', list: at(k, 'customList') || undefined, dxf: at(k, 'dxfId') != null ? num(k, 'dxfId', 0) : undefined })) };
  }
  function cfvo(e) { const o = { type: at(e, 'type') || 'num' }; if (at(e, 'val') != null) o.val = at(e, 'val'); if (at(e, 'gte') === '0') o.gte = false; return o; }
  function readCfRule(r, ctx) {
    const o = { type: at(r, 'type') || 'expression', priority: num(r, 'priority', 1) };
    if (at(r, 'dxfId') != null) o.dxf = num(r, 'dxfId', 0);
    if (bool(r, 'stopIfTrue', false)) o.stop = true;
    if (at(r, 'operator')) o.op = at(r, 'operator');
    if (at(r, 'text') != null) o.text = at(r, 'text');
    if (at(r, 'timePeriod')) o.period = at(r, 'timePeriod');
    if (at(r, 'rank') != null) o.rank = num(r, 'rank', 10);
    if (bool(r, 'percent', false)) o.percent = true;
    if (bool(r, 'bottom', false)) o.bottom = true;
    if (at(r, 'aboveAverage') === '0') o.below = true;
    if (bool(r, 'equalAverage', false)) o.equal = true;
    if (at(r, 'stdDev') != null) o.stdDev = num(r, 'stdDev', 0);
    const fs = kids(r, 'formula').map((f) => f.textContent);
    if (fs.length) o.f = fs;
    const cs = kid(r, 'colorScale');
    if (cs) o.scale = { cfvo: kids(cs, 'cfvo').map(cfvo), colors: kids(cs, 'color').map(color) };
    const db = kid(r, 'dataBar');
    if (db) o.bar = { cfvo: kids(db, 'cfvo').map(cfvo), color: color(kid(db, 'color')), showValue: at(db, 'showValue') !== '0', minLength: num(db, 'minLength', 10), maxLength: num(db, 'maxLength', 90) };
    const is = kid(r, 'iconSet');
    if (is) o.icons = { set: at(is, 'iconSet') || '3TrafficLights1', cfvo: kids(is, 'cfvo').map(cfvo), reverse: bool(is, 'reverse', false), showValue: at(is, 'showValue') !== '0', percent: at(is, 'percent') !== '0' };
    /* x14 link (data bar extensions) */
    const ext = kid(r, 'extLst');
    if (ext) { const id = ext.getElementsByTagName('*').find((e) => e.localName === 'id'); if (id) o.x14id = id.textContent.trim(); }
    void ctx;
    return o;
  }
  function readDV(d, sqref) {
    const ranges = M.parseSqref(sqref != null ? sqref : at(d, 'sqref') || '');
    if (!ranges.length) return null;
    const o = { ranges, type: at(d, 'type') || 'none' };
    if (at(d, 'operator')) o.op = at(d, 'operator');
    o.allowBlank = bool(d, 'allowBlank', false);
    o.showDrop = at(d, 'showDropDown') !== '1'; /* attribute name is inverted in the file format */
    o.showInput = bool(d, 'showInputMessage', false);
    o.showError = bool(d, 'showErrorMessage', false);
    if (at(d, 'errorStyle')) o.errorStyle = at(d, 'errorStyle');
    if (at(d, 'imeMode')) o.ime = at(d, 'imeMode');
    for (const k of ['promptTitle', 'prompt', 'errorTitle', 'error']) if (at(d, k) != null) o[k] = XML.unx(at(d, k));
    const f1 = kid(d, 'formula1'), f2 = kid(d, 'formula2');
    if (f1) o.f1 = f1.textContent;
    if (f2) o.f2 = f2.textContent;
    return o;
  }
  function readSheetExt(ext, sh, ctx) {
    for (const e of kids(ext, 'ext')) {
      for (const c of e.children) {
        if (c.namespaceURI !== L.sheetExtensions.NS) continue;
        if (c.localName === 'dataValidations') {
          for (const d of kids(c, 'dataValidation')) {
            const sq = kid(d, 'sqref');
            const ranges = M.parseSqref(sq ? sq.textContent : '');
            if (!ranges.length) continue;
            const o = readDV(d, sq.textContent);
            if (!o) continue;
            const f1 = kid(d, 'formula1'), f2 = kid(d, 'formula2');
            if (f1) o.f1 = f1.textContent.trim();
            if (f2) o.f2 = f2.textContent.trim();
            o.x14 = true;
            sh.dv.push(o);
            L.sheetExtensions.keep(o, d, sh);
          }
        } else if (c.localName === 'conditionalFormattings') {
          for (const cf of kids(c, 'conditionalFormatting')) {
            const sq = kid(cf, 'sqref');
            const ranges = M.parseSqref(sq ? sq.textContent : '');
            for (const r of kids(cf, 'cfRule')) {
              const id = at(r, 'id');
              /* data bar extensions refine an existing rule */
              const base = id && sh.cf.flatMap((x) => x.rules).find((x) => x.x14id === id);
              const db = kid(r, 'dataBar');
              if (base && db && base.bar) {
                for (const k of ['minLength', 'maxLength']) if (at(db, k) != null) base.bar[k] = num(db, k, base.bar[k]);
                if (at(db, 'showValue') != null) base.bar.showValue = bool(db, 'showValue', true);
                base.bar.gradient = at(db, 'gradient') !== '0';
                base.bar.border = bool(db, 'border', false);
                if (at(db, 'direction')) base.bar.direction = at(db, 'direction');
                const nf = kid(db, 'negativeFillColor'); if (nf) base.bar.negColor = color(nf);
                const bc = kid(db, 'borderColor'); if (bc) base.bar.borderColor = color(bc);
                const ac = kid(db, 'axisColor'); if (ac) base.bar.axisColor = color(ac);
                if (at(db, 'axisPosition')) base.bar.axis = at(db, 'axisPosition');
                const cfvos = kids(db, 'cfvo');
                if (cfvos.length === 2) base.bar.cfvo = cfvos.map((v) => { const o = { type: at(v, 'type') === 'autoMin' ? 'min' : at(v, 'type') === 'autoMax' ? 'max' : at(v, 'type') }; const f = kid(v, 'f'); if (f) o.val = f.textContent; return o; });
                L.sheetExtensions.keep(base, r, sh, cf);
                continue;
              }
              if (!ranges.length) continue;
              /* x14-only rules (icon sets with custom icons, rules referring to other sheets) */
              const o = readCfRule(r, ctx);
              o.x14id = id; o.x14Only = true;
              const fs = kids(r, 'f').map((f) => f.textContent);
              if (fs.length) o.f = fs;
              const dx = kid(r, 'dxf');
              if (dx) { const d = {}; const f = kid(dx, 'font'); if (f) d.font = font(f); const fl = kid(dx, 'fill'); if (fl) d.fill = fill(fl, true); const b = kid(dx, 'border'); if (b) d.border = border(b); const n = kid(dx, 'numFmt'); if (n) d.nf = at(n, 'formatCode'); o.dxf = ctx.wb.dxfs.push(d) - 1; }
              const is = kid(r, 'iconSet');
              if (is) {
                o.icons = { set: at(is, 'iconSet') || '3TrafficLights1', cfvo: kids(is, 'cfvo').map((v) => { const x = { type: at(v, 'type') || 'percent' }; const f = kid(v, 'f'); if (f) x.val = f.textContent; if (at(v, 'gte') === '0') x.gte = false; return x; }), reverse: bool(is, 'reverse', false), showValue: at(is, 'showValue') !== '0', percent: at(is, 'percent') !== '0' };
                const ci = kids(is, 'cfIcon');
                if (ci.length) o.icons.custom = ci.map((x) => ({ set: at(x, 'iconSet'), id: +(at(x, 'iconId') || 0) }));
              }
              sh.cf.push({ ranges, rules: [o] });
              L.sheetExtensions.keep(o, r, sh, cf);
            }
          }
        } else if (c.localName === 'sparklineGroups') {
          sh.sparklines = kids(c, 'sparklineGroup').map((g) => ({
            type: at(g, 'type') || 'line', markers: bool(g, 'markers', false), high: bool(g, 'high', false), low: bool(g, 'low', false), first: bool(g, 'first', false), last: bool(g, 'last', false), negative: bool(g, 'negative', false),
            color: color(kid(g, 'colorSeries')), negColor: color(kid(g, 'colorNegative')), markerColor: color(kid(g, 'colorMarkers')), highColor: color(kid(g, 'colorHigh')), lowColor: color(kid(g, 'colorLow')), firstColor: color(kid(g, 'colorFirst')), lastColor: color(kid(g, 'colorLast')),
            lineWeight: num(g, 'lineWeight', 0.75), displayEmptyCellsAs: at(g, 'displayEmptyCellsAs') || 'zero',
            axisColor: color(kid(g, 'colorAxis')), displayXAxis: bool(g, 'displayXAxis', false), displayHidden: bool(g, 'displayHidden', false), rightToLeft: bool(g, 'rightToLeft', false),
            minAxisType: at(g, 'minAxisType') || 'individual', maxAxisType: at(g, 'maxAxisType') || 'individual', manualMin: at(g, 'manualMin') != null ? +at(g, 'manualMin') : null, manualMax: at(g, 'manualMax') != null ? +at(g, 'manualMax') : null, dateAxis: bool(g, 'dateAxis', false),
            items: kids(kid(g, 'sparklines'), 'sparkline').map((s) => ({ f: (kid(s, 'f') || { textContent: '' }).textContent, sqref: (kid(s, 'sqref') || { textContent: '' }).textContent })),
          }));
          kids(c, 'sparklineGroup').forEach((g, i) => L.sheetExtensions.keep(sh.sparklines[i], g, sh));
        }
      }
    }
    L.sheetExtensions.read(sh, ext);
  }

  async function readComments(pkg, part, sh) {
    const el = await pkg.xml(part);
    if (!el) return;
    const authors = kids(kid(el, 'authors'), 'author').map((a) => a.textContent);
    for (const c of kids(kid(el, 'commentList'), 'comment')) {
      const p = F.parseCell(at(c, 'ref') || '');
      if (!p) continue;
      const t = kid(c, 'text');
      const runs = [];
      let text = '';
      if (t) {
        for (const ch of t.children) {
          if (ch.localName === 't') { runs.push({ t: XML.unx(ch.textContent) }); text += XML.unx(ch.textContent); }
          else if (ch.localName === 'r') {
            const tt = kid(ch, 't');
            const s = tt ? XML.unx(tt.textContent) : '';
            const rp = kid(ch, 'rPr');
            const run = { t: s };
            if (rp) { const f = font(rp); if (Object.keys(f).length) run.font = f; }
            runs.push(run);
            text += s;
          }
        }
      }
      sh.comments.set(M.key(p.r, p.c), { r: p.r, c: p.c, author: authors[num(c, 'authorId', 0)] || '', text, runs: runs.some((x) => x.font) ? runs : undefined, visible: false });
    }
    L.threads.notes(sh, part);
  }
  async function readVmlNotes(pkg, part, sh) {
    const t = await pkg.text(part);
    if (!t) return;
    /* VML is not always well-formed XML (unclosed <br> in old files): scan shapes with regexes */
    const re = /<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/g;
    let m;
    while ((m = re.exec(t))) {
      const body = m[2];
      if (!/ObjectType="Note"/.test(body)) continue;
      const row = /<x:Row>\s*(\d+)\s*<\/x:Row>/.exec(body), col = /<x:Column>\s*(\d+)\s*<\/x:Column>/.exec(body);
      if (!row || !col) continue;
      const cm = sh.comments.get(M.key(+row[1], +col[1]));
      if (!cm) continue;
      const style = (/style="([^"]*)"/.exec(m[1]) || [])[1] || '';
      cm.visible = /<x:Visible\s*\/>|<x:Visible>/.test(body) || /visibility:\s*visible/.test(style);
      const w = /width:\s*([\d.]+)pt/.exec(style), h = /height:\s*([\d.]+)pt/.exec(style);
      if (w) cm.w = +w[1];
      if (h) cm.h = +h[1];
      const anchor = /<x:Anchor>\s*([^<]+)<\/x:Anchor>/.exec(body);
      if (anchor) cm.anchor = anchor[1].split(',').map((x) => +x.trim());
      const fill = /fillcolor="([^"]+)"/.exec(m[1]);
      if (fill) cm.fill = fill[1];
    }
  }
  async function readThreaded(pkg, part, sh, ctx) {
    L.threads.read(sh, part, ctx.persons);
  }
  async function readTable(pkg, part, sh, ctx) {
    const el = await pkg.xml(part);
    if (!el) return;
    const ref = F.parseRange(at(el, 'ref') || '');
    if (!ref) return;
    const t = {
      ooxmlPart: part, id: num(el, 'id', ctx.wb.tables.length + 1), name: at(el, 'displayName') || at(el, 'name') || 'Table' + (ctx.wb.tables.length + 1), dname: at(el, 'name'), sheet: sh, ref,
      header: num(el, 'headerRowCount', 1) > 0, totals: num(el, 'totalsRowCount', 0) > 0 || bool(el, 'totalsRowShown', false) && num(el, 'totalsRowCount', 0) > 0,
      columns: [], style: null, autoFilter: !!kid(el, 'autoFilter'), comment: at(el, 'comment') || undefined,
    };
    if (at(el, 'tableType')) t.type = at(el, 'tableType');
    for (const c of kids(kid(el, 'tableColumns'), 'tableColumn')) {
      const o = { name: XML.unx(at(c, 'name') || ''), id: num(c, 'id', t.columns.length + 1) };
      if (at(c, 'totalsRowFunction')) o.totalsFn = at(c, 'totalsRowFunction');
      if (at(c, 'totalsRowLabel') != null) o.totalsLabel = at(c, 'totalsRowLabel');
      const cf = kid(c, 'calculatedColumnFormula'); if (cf) o.calc = cf.textContent;
      const tf = kid(c, 'totalsRowFormula'); if (tf) o.totalsFormula = tf.textContent;
      if (at(c, 'dataDxfId') != null) o.dataDxf = num(c, 'dataDxfId', 0);
      t.columns.push(o);
    }
    const si = kid(el, 'tableStyleInfo');
    if (si) t.style = { name: at(si, 'name') || 'TableStyleMedium2', rowStripes: bool(si, 'showRowStripes', false), colStripes: bool(si, 'showColumnStripes', false), first: bool(si, 'showFirstColumn', false), last: bool(si, 'showLastColumn', false) };
    const af = kid(el, 'autoFilter');
    if (af) { t.filter = readAutoFilter(af); if (ctx.markFiltered) ctx.markFiltered(t.filter); }
    ctx.wb.tables.push(t);
    L.tableKeep.read(t, el);
  }

  /* ------------------------------------------------------------ drawings */
  const EMU = 12700;
  function anchorPt(e) {
    if (!e) return null;
    const v = (n) => { const k = kid(e, n); return k ? parseFloat(k.textContent) || 0 : 0; };
    return { c: v('col'), cOff: v('colOff') / EMU, r: v('row'), rOff: v('rowOff') / EMU };
  }
  async function readDrawing(pkg, part, sh, ctx) {
    const el = await pkg.xml(part);
    if (!el) return;
    sh.extra.ooxmlDrawing = part;
    // Shape XML is saved as an opaque object. Capture before selecting a
    // display branch, so nested equations and effects keep Choice/Fallback.
    const originalShapes = new WeakMap(el.getElementsByTagName('*').filter(e => /^(sp|grpSp|cxnSp|pic)$/.test(e.localName)).map(e => [e, XML.serialize(e)]));
    const owned = new Set((sh.extra.objectGroups || []).flatMap(g => g.ids));
    const slicers = L.slicers.captureDrawing(el, sh, part);
    resolveAC(el, true);
    const rels = await pkg.rels(part);
    for (const a of el.children) {
      const kind = a.localName;
      if (!/Anchor$/.test(kind)) continue;
      const anchor = { type: kind === 'twoCellAnchor' ? 'two' : kind === 'oneCellAnchor' ? 'one' : 'abs' };
      if (kind === 'twoCellAnchor') { anchor.from = anchorPt(kid(a, 'from')); anchor.to = anchorPt(kid(a, 'to')); anchor.editAs = at(a, 'editAs') || 'twoCell'; }
      else if (kind === 'oneCellAnchor') { anchor.from = anchorPt(kid(a, 'from')); const ex = kid(a, 'ext'); anchor.w = num(ex, 'cx', 0) / EMU; anchor.h = num(ex, 'cy', 0) / EMU; }
      else { const p = kid(a, 'pos'), ex = kid(a, 'ext'); anchor.x = num(p, 'x', 0) / EMU; anchor.y = num(p, 'y', 0) / EMU; anchor.w = num(ex, 'cx', 0) / EMU; anchor.h = num(ex, 'cy', 0) / EMU; }
      const obj = a.children.find((c) => /^(pic|graphicFrame|sp|grpSp|cxnSp|contentPart)$/.test(c.localName));
      const record = slicers.get(a);
      if (!obj && !record) continue;
      let d = obj && await readDrawingObject(pkg, obj, rels, ctx, sh);
      if (!d && record) d = { kind: 'shape', text: record.names.join(', '), geom: 'rect' };
      if (!d) continue;
      if (d.kind === 'shape' && originalShapes.has(obj)) d.xml = originalShapes.get(obj);
      d.anchor = anchor;
      if (d.kind === 'shape' && ctx.wb.pkg) d.keep = { source: ctx.wb.pkg.id, part };
      if (d.kind === 'image' && owned.has(d.id) && originalShapes.has(obj)) d.keep = { source: ctx.wb.pkg.id, part,
        frame: K.fragment(K.parse(originalShapes.get(obj)), { pkg: ctx.wb.pkg, part }) };
      else if (d.kind === 'image' && ctx.wb.pkg && originalShapes.has(obj)) {
        d.keep = { source: ctx.wb.pkg.id, part };
        L.sheetObjects.keepPicture(d, K.fragment(K.parse(originalShapes.get(obj)), { pkg: ctx.wb.pkg, part }));
      }
      const cd = kid(a, 'clientData');
      if (cd && at(cd, 'fLocksWithSheet') === '0') d.unlocked = true;
      if (cd && at(cd, 'fPrintsWithSheet') === '0') d.noPrint = true;
      if (record) L.slicers.attach(d, record);
      sh.drawings.push(d);
    }
  }
  async function readDrawingObject(pkg, obj, rels, ctx, sh) {
    const nv = obj.children.find((c) => /^nv/.test(c.localName));
    const cNv = kid(nv, 'cNvPr');
    const base = { id: num(cNv, 'id', 0), name: at(cNv, 'name') || '', descr: at(cNv, 'descr') || undefined, hidden: bool(cNv, 'hidden', false) || undefined };
    if (obj.localName === 'pic') {
      const blip = obj.getElementsByTagName('*').find((e) => e.localName === 'blip');
      const re = blip && rid(blip, 'embed') ? rels.get(rid(blip, 'embed')) : null;
      const rl = blip && rid(blip, 'link') ? rels.get(rid(blip, 'link')) : null;
      let media = null;
      if (re && !re.external) { const bytes = await pkg.bytes(re.target); if (bytes) media = ctx.media(re.target, bytes); }
      /* a linked picture (Insert Picture ▸ Link to File, or a URL) keeps its link even when the image cannot be loaded here */
      const imgLink = rl && rl.external ? rl.raw : re && re.external ? re.raw : undefined;
      if (!media && !imgLink) return null;
      const d = Object.assign(base, { kind: 'image', media, imgLink });
      const sp = kid(obj, 'spPr');
      const ln = sp && kid(sp, 'ln');
      if (ln && !kid(ln, 'noFill')) d.lineXml = XML.serialize(ln);
      const src = kid(kid(obj, 'blipFill'), 'srcRect');
      if (src) d.crop = { l: num(src, 'l', 0) / 1000, t: num(src, 't', 0) / 1000, r: num(src, 'r', 0) / 1000, b: num(src, 'b', 0) / 1000 };
      const hl = kid(cNv, 'hlinkClick'); if (hl) { const hr = rels.get(rid(hl, 'id')); if (hr) d.link = hr.raw; }
      return d;
    }
    if (obj.localName === 'graphicFrame') {
      const gd = obj.getElementsByTagName('*').find((e) => e.localName === 'graphicData');
      const ch = gd && gd.children.find((e) => e.localName === 'chart');
      if (ch) {
        const r = rels.get(rid(ch, 'id'));
        if (!r) return null;
        const xml = await pkg.text(r.target);
        if (!xml) return null;
        const d = Object.assign(base, { kind: 'chart', xml, part: r.target });
        const crels = await pkg.rels(r.target);
        /* chart-specific parts (colours, style, user shapes) are kept as-is */
        const extras = [];
        for (const cr of crels.values()) {
          if (cr.external) continue;
          const b = await pkg.bytes(cr.target);
          if (b) extras.push({ type: cr.type, name: cr.target.split('/').pop(), bytes: b });
        }
        if (extras.length) d.chartParts = extras;
        if (L.xchart && L.xchart.read) { try { d.chart = L.xchart.read(XML.parse(xml), ctx.wb, sh); } catch (e) { d.chart = null; } }
        return d;
      }
      return null; /* SmartArt, slicers, timelines: not shown */
    }
    if (obj.localName === 'sp' || obj.localName === 'cxnSp' || obj.localName === 'grpSp') {
      /* shapes and text boxes: keep their XML; images inside groups are not resolved */
      const d = Object.assign(base, { kind: 'shape', xml: XML.serialize(obj), tag: obj.localName });
      const txb = obj.getElementsByTagName('*').filter((e) => e.localName === 't');
      if (txb.length) d.text = txb.map((e) => e.textContent).join('');
      if (obj.localName === 'sp') {
        const sp = kid(obj, 'spPr');
        const pg = kid(sp, 'prstGeom');
        d.geom = pg ? at(pg, 'prst') || 'rect' : 'rect';
        const macro = at(obj, 'macro'); if (macro) d.macro = macro;
        d.txBox = at(kid(nv, 'cNvSpPr'), 'txBox') === '1';
      }
      return d;
    }
    return null;
  }
  /** Markup Compatibility: drop the Choice branches we cannot render (slicers, chartex, ink) for their Fallback */
  function resolveAC(root, drawing) {
    const acs = root.getElementsByTagName('*').filter((e) => e.localName === 'AlternateContent').reverse();
    for (const ac of acs) {
      if (!ac.parentNode) continue;
      const ch = kids(ac, 'Choice');
      const fb = kid(ac, 'Fallback');
      let pick = null;
      for (const c of ch) {
        const req = (at(c, 'Requires') || '').split(/\s+/);
        const inner = c.getElementsByTagName('*');
        const bad = inner.some((e) => e.localName === 'graphicData' && /slicer|timeslicer|chartex|chartEx|ink/i.test(at(e, 'uri') || '')) || inner.some((e) => e.localName === 'contentPart');
        const ok = req.every((r) => /^(a14|x14|x15|xr|xr2|xr3|x14ac|c14|c16|c16r2|c16r3|a16|mc|v|o|x)$/.test(r) || !r);
        if (ok && !bad) { pick = c; break; }
      }
      if (!pick) pick = fb || null;
      if (pick) {
        /* the lifted children lose the namespace declarations made on mc:AlternateContent / mc:Choice */
        const decl = {};
        for (const e of [ac, pick]) if (e.attrs) for (const k in e.attrs) if (k === 'xmlns' || k.startsWith('xmlns:')) decl[k] = e.attrs[k];
        for (const n of pick.childNodes) if (n.nodeType === 1) for (const k in decl) if (!n.attrs || n.attrs[k] === undefined) { const p = k.slice(6); if (!p || n.lookupNamespaceURI(p) == null) n.setAttribute(k, decl[k]); }
        ac.replaceWith(...pick.childNodes.slice());
      } else ac.remove();
    }
    void drawing;
  }
  R.resolveAC = resolveAC;

  /* ------------------------------------------------------------ workbook */
  R.read = async function (input, opts) {
    opts = opts || {};
    let bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
    let zip;
    try { zip = await L.zip.read(bytes); } catch (e) {
      if (e.code === 'ole' && e.encrypted) {
        /* workbooks protected only against changes are encrypted with Excel's built-in default password */
        let inner = null;
        try { inner = await L.officeCrypto.decrypt(bytes, 'VelvetSweatshop'); } catch (e0) { inner = null; }
        if (!inner) {
          if (!opts.password) { const err = new Error('password'); err.code = 'password'; throw err; }
          try { inner = await L.officeCrypto.decrypt(bytes, opts.password); } catch (e2) {
            const unsup = e2 && e2.message === 'unsupported';
            const err = new Error(unsup ? 'This workbook is encrypted with a method Ledger does not support (only the AES encryption Excel 2007 and later use is supported).' : 'The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.');
            err.code = unsup ? 'unsupported' : 'badpassword';
            throw err;
          }
        }
        bytes = inner;
        zip = await L.zip.read(bytes);
      } else if (e.code === 'ole') {
        const err = new Error('xls'); err.code = 'xls'; err.bytes = bytes; throw err;
      } else throw e;
    }
    const pkg = new Pkg(zip);
    const rootRels = await pkg.rels('');
    const off = Array.from(rootRels.values()).find((r) => r.type === 'officeDocument');
    let wbPart = off ? off.target : 'xl/workbook.xml';
    if (!pkg.has(wbPart)) wbPart = pkg.find('xl/workbook.xml') || Array.from(zip.keys()).find((k) => /workbook\.xml$/.test(k));
    if (!wbPart) throw new Error('This file is not an Excel workbook.');
    const wbEl = await pkg.xml(wbPart);
    if (!wbEl) throw new Error('The workbook part is damaged.');
    const wrels = await pkg.rels(wbPart);
    const wb = new M.Workbook();
    try { L.opc.attach(wb, await L.opc.open(zip)); }
    catch (error) { L.opc.loss(wb, { id: 'package', what: 'Part of this file couldn\'t be read. Anything in it that VibeOffice can\'t edit itself, such as macros or embedded objects, won\'t be saved.', detail: error.message, where: wbPart, action: 'drop', notify: true }); }
    wb.ooxmlFormat = L.opc.variant(wb.pkg, 'xlsx');
    wb.strict = /purl\.oclc\.org/.test(wbEl.lookupNamespaceURI(wbEl.prefix) || '');
    const wpr = kid(wbEl, 'workbookPr');
    if (wpr) { wb.date1904 = bool(wpr, 'date1904', false); if (at(wpr, 'codeName')) wb.codeName = at(wpr, 'codeName'); if (bool(wpr, 'filterPrivacy', false)) wb.extra.filterPrivacy = true; if (at(wpr, 'defaultThemeVersion')) wb.themeVersion = at(wpr, 'defaultThemeVersion'); }
    const cp = kid(wbEl, 'calcPr');
    if (cp) {
      wb.calcPr.mode = at(cp, 'calcMode') || 'auto';
      wb.calcPr.iterate = bool(cp, 'iterate', false);
      wb.calcPr.iterateCount = num(cp, 'iterateCount', 100);
      wb.calcPr.iterateDelta = num(cp, 'iterateDelta', 0.001);
      wb.calcPr.fullPrecision = at(cp, 'fullPrecision') !== '0';
      wb.calcPr.fullCalcOnLoad = bool(cp, 'fullCalcOnLoad', false);
      wb.calcPr.calcId = at(cp, 'calcId');
      if (at(cp, 'refMode') === 'R1C1') wb.r1c1 = true;
    }
    const prot = kid(wbEl, 'workbookProtection');
    if (prot && (bool(prot, 'lockStructure', false) || bool(prot, 'lockWindows', false))) wb.protection = { attrs: Object.assign({}, prot.attrs), structure: bool(prot, 'lockStructure', false), windows: bool(prot, 'lockWindows', false) };
    const bv = path(wbEl, 'bookViews', 'workbookView');
    if (bv) { wb.active = num(bv, 'activeTab', 0); wb.firstTab = num(bv, 'firstSheet', 0); if (at(bv, 'showSheetTabs') === '0') wb.hideTabs = true; if (at(bv, 'showHorizontalScroll') === '0') wb.hideHScroll = true; if (at(bv, 'showVerticalScroll') === '0') wb.hideVScroll = true; if (at(bv, 'tabRatio')) wb.tabRatio = num(bv, 'tabRatio', 600); }
    /* theme, styles, strings */
    const themeRel = relOfType(wrels, 'theme');
    if (themeRel) { const te = await pkg.xml(themeRel.target); wb.theme = readTheme(te); wb.themeXml = await pkg.text(themeRel.target); }
    const stRel = relOfType(wrels, 'styles');
    const st = readStyles(stRel ? await pkg.xml(stRel.target) : null, wb);
    wb.dxfs = st.dxfs;
    wb.cellStyles = st.namedStyles.length ? st.namedStyles : wb.cellStyles;
    wb.tableStyles = st.tableStyles;
    const sstRel = relOfType(wrels, 'sharedStrings');
    const sst = sstRel ? readSST((await pkg.text(sstRel.target)) || '') : [];
    /* media dedupe: one entry per package part */
    const mediaMap = new Map();
    wb.media = new Map();
    const media = (partName, bytes2) => {
      if (mediaMap.has(partName)) return mediaMap.get(partName);
      const ext = partName.split('.').pop().toLowerCase();
      const id = 'img' + (wb.media.size + 1);
      wb.media.set(id, { part: partName, bytes: bytes2, ext, type: ({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', emf: 'image/x-emf', wmf: 'image/x-wmf', tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml', webp: 'image/webp' })[ext] || 'application/octet-stream', name: partName.split('/').pop() });
      mediaMap.set(partName, id);
      return id;
    };
    /* people for threaded comments */
    const personRel = relOfType(wrels, 'person');
    let persons = null;
    if (personRel) { const pe = await pkg.xml(personRel.target); persons = new Map(); for (const p of kids(pe, 'person')) persons.set(at(p, 'id'), at(p, 'displayName') || ''); }
    const ctx = { wb, sst, xf: st.xf, media, persons };
    /* sheets */
    const sheetEls = kids(kid(wbEl, 'sheets'), 'sheet');
    const sheetParts = [];
    for (const s of sheetEls) {
      const r = wrels.get(rid(s, 'id'));
      const kind = r ? (r.type === 'chartsheet' ? 'chartsheet' : r.type === 'dialogsheet' ? 'dialogsheet' : r.type === 'xlMacrosheet' || r.type === 'xlIntlMacrosheet' ? 'macrosheet' : 'worksheet') : 'worksheet';
      const sh = wb.addSheet(XML.unx(at(s, 'name') || 'Sheet' + (wb.sheets.length + 1)), null, kind);
      sh.sheetId = num(s, 'sheetId', wb.sheets.length);
      if (r && !r.external) sh.extra.ooxmlPart = r.target;
      const state = at(s, 'state');
      if (state === 'hidden' || state === 'veryHidden') sh.state = state;
      else if (state && state !== 'visible') sh.extra.invalidVisibility = state;
      sheetParts.push({ sh, part: r ? r.target : null });
    }
    /* defined names (before formulas are calculated) */
    for (const d of kids(kid(wbEl, 'definedNames'), 'definedName')) {
      const name = at(d, 'name');
      if (!name) continue;
      const o = { name, ref: d.textContent.trim(), scope: at(d, 'localSheetId') != null ? num(d, 'localSheetId', 0) : null };
      if (bool(d, 'hidden', false)) o.hidden = true;
      if (at(d, 'comment')) o.comment = at(d, 'comment');
      if (bool(d, 'function', false) || bool(d, 'vbProcedure', false) || bool(d, 'xlm', false)) o.macro = true;
      wb.names.push(o);
    }
    /* external links: keep the parts so references like [1]Sheet1!A1 survive a save */
    const ext = [];
    for (const er of kids(kid(wbEl, 'externalReferences'), 'externalReference')) {
      const r = wrels.get(rid(er, 'id'));
      if (!r) continue;
      const xml = await pkg.text(r.target);
      const lrels = await pkg.rels(r.target);
      const target = Array.from(lrels.values())[0];
      ext.push({ xml, part: r.target, relTarget: target ? target.raw : null, relType: target ? target.type : null, relExternal: target ? target.external : true });
      readExternalCache(xml, ext.length, target ? target.raw : '', wb, ctx);
    }
    if (ext.length) wb.externalLinks = ext;
    for (const { sh, part } of sheetParts) {
      if (!part) continue;
      if (sh.kind === 'chartsheet') { await readChartsheet(pkg, part, sh, ctx); continue; }
      await readWorksheet(pkg, part, sh, ctx);
    }
    /* document properties */
    const coreRel = Array.from(rootRels.values()).find((r) => r.type === 'core-properties' || /metadata\/core-properties$/.test(r.type));
    const coreEl = await pkg.xml(coreRel ? coreRel.target : 'docProps/core.xml');
    if (coreEl) {
      const g = (n) => { const e = kid(coreEl, n); return e ? e.textContent : ''; };
      Object.assign(wb.props, { title: g('title'), subject: g('subject'), creator: g('creator'), keywords: g('keywords'), description: g('description'), lastModifiedBy: g('lastModifiedBy'), category: g('category'), created: g('created') || null, modified: g('modified') || null });
    }
    const propertyPart = (type, fallback) => wb.pkg?.rels('').find(r => L.opc.relationshipType(r.type) === L.opc.NS.rel + '/' + type)?.part || fallback;
    const appEl = await pkg.xml(propertyPart('extended-properties', 'docProps/app.xml'));
    if (appEl) {
      const g = (n) => { const e = kid(appEl, n); return e ? e.textContent : ''; };
      wb.props.company = g('Company'); wb.props.manager = g('Manager'); wb.props.application = g('Application'); wb.props.appVersion = g('AppVersion');
    }
    const custEl = await pkg.xml(propertyPart('custom-properties', 'docProps/custom.xml'));
    if (custEl) wb.props.custom = kids(custEl, 'property').map((p) => ({ name: at(p, 'name'), type: p.children[0] ? p.children[0].localName : 'lpwstr', value: p.children[0] ? p.children[0].textContent : '' }));
    /* macros and other parts we carry through untouched */
    const vba = relOfType(wrels, 'vbaProject');
    if (vba) { wb.extra.vba = await pkg.bytes(vba.target); wb.macroEnabled = true; }
    if (/macroEnabled/i.test((await pkg.text('[Content_Types].xml')) || '')) wb.macroEnabled = true;
    const customXml = Array.from(zip.keys()).filter((k) => /^customXml\//.test(k));
    if (customXml.length) { wb.extra.customXml = []; for (const k of customXml) wb.extra.customXml.push({ name: k, bytes: await pkg.bytes(k) }); }
    if (zip.repaired) wb.repaired = true;
    /* active sheet must be visible */
    if (!wb.sheets.length) wb.addSheet('Sheet1');
    if (wb.active >= wb.sheets.length || wb.active < 0) wb.active = 0;
    if (wb.sheets[wb.active].state !== 'visible') wb.active = Math.max(0, wb.sheets.findIndex((s) => s.state === 'visible'));
    wb.extra.keepValues = L.preserve.values(wb);
    L.pivots.read(wb);
    L.slicers.read(wb);
    L.tableKeep.readWorkbook(wb);
    wb.extra.styleBaseline = { xf: st.xf.slice(), list: JSON.parse(JSON.stringify(wb.styles.list)), dxfs: JSON.parse(JSON.stringify(wb.dxfs)) };
    for (const sh of wb.sheets) if (sh.kind === 'macrosheet' || sh.kind === 'dialogsheet') sh.extra.opaqueBaseline = L.preserve.sheetContent(sh);
    return wb;
  };

  /** cached values of another workbook, so [1]Sheet1!A1 keeps showing what Excel saved */
  function readExternalCache(xml, idx, target, wb, ctx) {
    let el;
    try { el = XML.parse(xml || ''); } catch (e) { return; }
    const eb = kid(el, 'externalBook');
    if (!eb) return;
    const names = kids(kid(eb, 'sheetNames'), 'sheetName').map((s) => at(s, 'val') || '');
    const book = { target, sheets: {}, sheetList: [], names: {} };
    names.forEach((n) => { const s = new M.Sheet(wb, n); s.external = true; book.sheets[n.toLowerCase()] = s; book.sheetList.push(s); });
    for (const sd of kids(kid(eb, 'sheetDataSet'), 'sheetData')) {
      const s = book.sheetList[num(sd, 'sheetId', 0)];
      if (!s) continue;
      for (const row of kids(sd, 'row')) {
        for (const c of kids(row, 'cell')) {
          const p = F.parseCell(at(c, 'r') || '');
          if (!p) continue;
          const v = kid(c, 'v');
          if (!v) continue;
          const t = at(c, 't') || 'n';
          const txt = v.textContent;
          const val = t === 'n' ? +txt : t === 'b' ? txt === '1' : t === 'e' ? M.err(txt) : XML.unx(txt);
          s.put(p.r, p.c, { v: val });
        }
      }
    }
    for (const d of kids(kid(eb, 'definedNames'), 'definedName')) book.names[(at(d, 'name') || '').toLowerCase()] = { ref: (at(d, 'refersTo') || '').replace(/^=/, ''), sheetId: at(d, 'sheetId') };
    /* the book is addressed by its index ([1]) and by its file name */
    if (!wb.external) wb.external = {};
    wb.external[String(idx)] = book;
    const fname = decodeURIComponent(String(target).split(/[\\/]/).pop() || '');
    if (fname && !wb.external[fname]) wb.external[fname] = book;
    void ctx;
  }

  async function readChartsheet(pkg, part, sh, ctx) {
    const el = await pkg.xml(part);
    if (!el) return;
    const rels = await pkg.rels(part);
    const tc = color(path(el, 'sheetPr', 'tabColor')); if (tc) sh.tabColor = tc;
    const sv = path(el, 'sheetViews', 'sheetView');
    if (sv && at(sv, 'zoomScale')) sh.view.zoom = num(sv, 'zoomScale', 100);
    if (sv && at(sv, 'zoomToFit') === '1') sh.view.zoomToFit = true;
    const pm = kid(el, 'pageMargins');
    if (pm) sh.print.margins = { l: num(pm, 'left', 0.7), r: num(pm, 'right', 0.7), t: num(pm, 'top', 0.75), b: num(pm, 'bottom', 0.75), header: num(pm, 'header', 0.3), footer: num(pm, 'footer', 0.3) };
    const ps = kid(el, 'pageSetup');
    readPageSetup(ps, sh.print);
    L.preserve.keepPageSetup(sh, ps);
    const dr = kid(el, 'drawing');
    if (dr) { const r = rels.get(rid(dr, 'id')); if (r) await readDrawing(pkg, r.target, sh, ctx); }
  }

})(typeof window !== 'undefined' ? window : globalThis);
