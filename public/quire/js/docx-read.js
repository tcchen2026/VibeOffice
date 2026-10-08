/* Quire — WordprocessingML reader (.docx / .docm / .dotx / .dotm).
 * Builds the document model from the package: styles, numbering, settings, theme, sections,
 * headers & footers, footnotes/endnotes, comments, fields, revisions, tables, images, shapes,
 * text boxes, charts, SmartArt drawings, VML pictures and watermarks. Damaged packages are repaired
 * where possible.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, X = L.dml;
  const { kid, kids, at, num, rid, desc, descAll } = X;
  const bool = (el, d) => { if (!el) return d; const v = at(el, 'val'); return v == null ? true : !(v === '0' || v === 'false' || v === 'off'); };
  /** twips (or universal measure) → points */
  const tw = (v) => {
    if (v == null || v === '') return 0;
    const m = /^(-?\d*\.?\d+)\s*(mm|cm|in|pt|pc|pi)?$/.exec(String(v).trim());
    if (!m) return 0;
    const x = parseFloat(m[1]);
    switch (m[2]) { case 'mm': return x * 2.83465; case 'cm': return x * 28.3465; case 'in': return x * 72; case 'pt': return x; case 'pc': case 'pi': return x * 12; default: return x / 20; }
  };
  const twA = (el, n) => (at(el, n) == null ? undefined : L.round(tw(at(el, n)), 3));
  const emu = X.emu;

  const W = (L.docx = L.docx || {});

  W.read = async function (buffer, opts) {
    opts = opts || {};
    let zip;
    try { zip = await L.zip.read(buffer); }
    catch (e) {
      if (e.code === 'ole' && !e.encrypted && L.docbin) return L.docbin.read(e.bytes);
      throw e;
    }
    const warnings = [];
    if (zip.repaired) warnings.push('The file was damaged; Quire repaired what it could.');
    const get = (p) => zip.get(p) || zip.get(p.replace(/^\//, '')) || findCI(p);
    const findCI = (p) => { const lp = p.toLowerCase(); for (const [k, v] of zip) if (k.toLowerCase() === lp) return v; return null; };
    const xmlCache = new Map();
    const xmlParts = new WeakMap();
    const xml = async (p) => {
      if (!p) return null;
      if (xmlCache.has(p)) return xmlCache.get(p);
      const e = get(p);
      if (!e) return null;
      let x = null;
      try { x = X.parse(L.xmlTree.decode(await e.bytes()), L.preserve.captureProperties); } catch (err) { warnings.push('Part ' + p + ' could not be read.'); }
      xmlCache.set(p, x);
      if (x) xmlParts.set(x.ownerDocument || x, p);
      return x;
    };
    const rels = async (p) => {
      const out = {};
      const x = await xml(X.relsPath(p));
      if (!x) return out;
      for (const r of kids(x, 'Relationship')) {
        const type = (at(r, 'Type') || '').split('/').pop();
        const ext = at(r, 'TargetMode') === 'External';
        const target = at(r, 'Target') || '';
        out[at(r, 'Id')] = { type, target: ext ? target : X.resolvePath(p, target), external: ext, raw: target };
      }
      return out;
    };
    /* ---- main part ---- */
    const rootRelsX = await xml('_rels/.rels');
    let mainPath = 'word/document.xml';
    if (rootRelsX) for (const r of kids(rootRelsX, 'Relationship')) if (/\/officeDocument$/.test(at(r, 'Type') || '')) mainPath = X.resolvePath('', at(r, 'Target'));
    if (!get(mainPath)) {
      const cand = Array.from(zip.keys()).find((k) => /(^|\/)document\d*\.xml$/i.test(k));
      if (!cand) throw new Error('This file does not contain a Word document.');
      mainPath = cand;
    }
    const docX = await xml(mainPath);
    if (!docX) throw new Error('The document body could not be read; the file may be damaged.');
    const docRels = await rels(mainPath);
    const relOfType = (t) => Object.values(docRels).find((r) => r.type === t && !r.external);

    /* ---- custom XML data stores: content controls bound to them show the stored values (as Word does on open) ---- */
    const stores = new Map(), storeParts = {};
    for (const k of Array.from(zip.keys())) {
      if (!/\.xml$/i.test(k)) continue;
      const rr = await rels(k);
      const pr = Object.values(rr).find((r) => /customXmlProps$/.test(r.type));
      if (!pr) continue;
      const px = pr && await xml(pr.target);
      const root = px && (px.documentElement || px);
      const idAttr = root && Array.from(root.attributes).find((a) => a.localName === 'itemID');
      const dx = idAttr && await xml(k);
      if (idAttr && dx) { stores.set(idAttr.value.toUpperCase(), dx); storeParts[idAttr.value.toUpperCase()] = k; }
    }
    if (stores.size || get('docProps/core.xml')) {
      const core = await xml('docProps/core.xml');
      if (core) { stores.set('{6C3C8BC8-F283-45AE-878A-BAB7291924A1}', core); storeParts['{6C3C8BC8-F283-45AE-878A-BAB7291924A1}'] = 'docProps/core.xml'; }
      const appx = await xml('docProps/app.xml');
      if (appx) { stores.set('{6668398D-A668-4E3E-A5EB-62B293D839F1}', appx); storeParts['{6668398D-A668-4E3E-A5EB-62B293D839F1}'] = 'docProps/app.xml'; }
    }
    /** text a data-bound content control should show, or null to keep what the file has */
    const boundText = (sdtPr) => {
      const db = sdtPr && kid(sdtPr, 'dataBinding');
      if (!db || !stores.size || kid(sdtPr, 'picture')) return null;
      const xp = at(db, 'xpath');
      const store = stores.get(String(at(db, 'storeItemID') || '').toUpperCase());
      if (!xp || !store) return null;
      const ns = {};
      for (const m of String(at(db, 'prefixMappings') || '').matchAll(/xmlns:([\w.-]+)\s*=\s*['"]([^'"]*)['"]/g)) ns[m[1]] = m[2];
      let v;
      try {
        const sd = store.ownerDocument || store;
        const r = sd.evaluate(xp, sd, (pfx) => ns[pfx] || null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
        if (!r.singleNodeValue) return null;
        v = r.singleNodeValue.textContent;
      } catch (e) { return null; }
      if (v == null || v === '') return null;
      /* rich-text controls bind to a whole package (flat OPC / XHTML): keep the content the file has */
      if (!kid(sdtPr, 'text') && !kid(sdtPr, 'date') && !kid(sdtPr, 'dropDownList') && !kid(sdtPr, 'comboBox') && !kid(sdtPr, 'checkbox')) return null;
      if (/^\s*<(\?xml|pkg:package|html|w:)/.test(v)) return null;
      const dt = kid(sdtPr, 'date');
      if (dt) {
        const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(v);
        const fmt = at(kid(dt, 'dateFormat'), 'val');
        if (m && fmt && L.fields && L.fields.formatDate) { try { return L.fields.formatDate(new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)), fmt); } catch (e) { /* keep raw */ } }
        return v;
      }
      const list = kid(sdtPr, 'dropDownList') || kid(sdtPr, 'comboBox');
      if (list) { const li = kids(list, 'listItem').find((x) => at(x, 'value') === v); if (li) return at(li, 'displayText') || v; }
      const cb = kid(sdtPr, 'checkbox');
      if (cb) {
        const on = /^(1|true)$/i.test(v.trim());
        const stEl = kid(cb, on ? 'checkedState' : 'uncheckedState');
        const code = stEl && at(stEl, 'val');
        return code ? String.fromCharCode(parseInt(code, 16)) : on ? '\u2612' : '\u2610';
      }
      return v;
    };
    /** paragraph items for a bound value: line breaks for new lines, markers kept */
    const boundRuns = (value, from) => {
      if (from.some(r => r.t === 'sdts')) return from;
      const first = from.find((r) => r.t === 'text' || r.t === 'sym');
      const rPr = first ? Object.assign({}, first.rPr || {}) : {};
      const out = from.filter((r) => D.ilen(r) === 0 && !/^f[bse]$/.test(r.t));
      String(value).split(/\r?\n/).forEach((ln, i) => { if (i) out.push(D.item('br', { type: 'line' }, rPr)); if (ln) out.push(D.text(ln, rPr)); });
      return out;
    };

    const doc = D.newDoc();
    doc.main.blocks = [];
    doc.numbering = { abs: {}, nums: {} };
    doc.settings.compat = 12;
    doc.theme = { major: 'Calibri Light', minor: 'Calibri', colors: Object.assign({}, X.DEFAULT_THEME) };
    doc.src = { mainPath };
    try { L.opc.attach(doc, await L.opc.open(zip)); }
    catch (error) { L.opc.loss(doc, { id: 'package', what: 'Some original package data could not be retained: ' + error.message, where: mainPath, action: 'drop' }); }
    doc.keep = { media: {}, stores: storeParts };
    const propertyContext = el => ({ pkg: doc.pkg, part: xmlParts.get(el.ownerDocument || el) || mainPath });

    /* ---- theme ---- */
    const themeRel = relOfType('theme');
    if (themeRel) {
      const tx = await xml(themeRel.target);
      if (tx) {
        const cs = desc(tx, 'clrScheme');
        if (cs) for (const c of kids(cs)) { const v = X.color(X.colorEl(c), null); if (v) doc.theme.colors[c.localName] = v.c; }
        const fs = desc(tx, 'fontScheme');
        if (fs) {
          const mj = X.path(fs, 'majorFont', 'latin'), mn = X.path(fs, 'minorFont', 'latin');
          if (mj && at(mj, 'typeface')) doc.theme.major = at(mj, 'typeface');
          if (mn && at(mn, 'typeface')) doc.theme.minor = at(mn, 'typeface');
          const mjEa = X.path(fs, 'majorFont', 'ea'), mnEa = X.path(fs, 'minorFont', 'ea');
          doc.theme.majorEA = mjEa && at(mjEa, 'typeface'); doc.theme.minorEA = mnEa && at(mnEa, 'typeface');
        }
      }
    }
    const dctx = { theme: doc.theme };
    const themeFont = (v) => (/^major/.test(v) ? doc.theme.major : doc.theme.minor);

    /* ---- colours ---- */
    function wColor(el, attr, themeAttr) {
      if (!el) return undefined;
      const th = at(el, themeAttr || 'themeColor');
      let v = at(el, attr || 'val');
      if (th && th !== 'none') {
        let hex = X.schemeHex(th, dctx);
        const tint = at(el, (themeAttr || 'themeColor').replace('Color', '') + 'Tint') || at(el, 'themeTint') || (themeAttr === 'themeFill' ? at(el, 'themeFillTint') : null);
        const shade = at(el, (themeAttr || 'themeColor').replace('Color', '') + 'Shade') || at(el, 'themeShade') || (themeAttr === 'themeFill' ? at(el, 'themeFillShade') : null);
        if (tint || shade) {
          const [r, g, b] = L.color.hexToRgb(hex);
          let [hh, s, l] = L.color.rgbToHsl(r, g, b);
          if (tint) { const t = parseInt(tint, 16) / 255; l = l * t + (1 - t); }
          if (shade) l = l * (parseInt(shade, 16) / 255);
          const rgb = L.color.hslToRgb(hh, s, L.clamp(l, 0, 1));
          hex = L.color.rgbToHex(rgb[0], rgb[1], rgb[2]);
        }
        return hex.replace('#', '').toUpperCase();
      }
      if (v == null) return undefined;
      if (v === 'auto') return 'auto';
      if (/^[0-9a-f]{6}$/i.test(v)) return v.toUpperCase();
      const named = L.color.PRESET[v];
      return named ? named.replace('#', '') : undefined;
    }
    function border(el) {
      if (!el) return undefined;
      const val = at(el, 'val') || 'single';
      if (val === 'nil' || val === 'none') return { val: 'nil' };
      const o = { val, sz: num(el, 'sz', 4) / 8, color: wColor(el, 'color') || 'auto' };
      if (at(el, 'space') != null) o.space = num(el, 'space', 0);
      if (bool(kid(el, 'shadow')) && at(el, 'shadow')) o.shadow = true;
      return o;
    }
    function borders(el, names) {
      if (!el) return undefined;
      const o = {};
      for (const n of names) {
        const b = kid(el, n) || (n === 'left' ? kid(el, 'start') : n === 'right' ? kid(el, 'end') : null);
        if (b) o[n] = border(b);
      }
      return Object.keys(o).length ? o : undefined;
    }
    function shd(el) {
      if (!el) return undefined;
      const o = { val: at(el, 'val') || 'clear' };
      const fill = wColor(el, 'fill', 'themeFill');
      const color = wColor(el, 'color');
      if (fill) o.fill = fill;
      if (color) o.color = color;
      if (o.val === 'clear' && (!o.fill || o.fill === 'auto')) return o.val === 'clear' && !o.fill ? undefined : o;
      return o;
    }

    /* ---- run properties ---- */
    function rPr(el) {
      const r = {};
      if (!el) return r;
      for (const c of el.children) {
        if (![L.opc.NS.w, 'http://purl.oclc.org/ooxml/wordprocessingml/main'].includes(c.namespaceURI)) continue;
        switch (c.localName) {
          case 'rStyle': r.style = at(c, 'val'); break;
          case 'rFonts': {
            const f = at(c, 'ascii') || at(c, 'hAnsi');
            const ft = at(c, 'asciiTheme') || at(c, 'hAnsiTheme');
            if (ft) r.font = themeFont(ft); else if (f) r.font = f;
            if (at(c, 'eastAsia')) r.fontEA = at(c, 'eastAsia'); else if (at(c, 'eastAsiaTheme')) r.fontEA = themeFont(at(c, 'eastAsiaTheme'));
            if (at(c, 'cs')) r.fontCS = at(c, 'cs');
            if (at(c, 'hint')) r.hint = at(c, 'hint');
            break;
          }
          case 'b': r.b = bool(c); break;
          case 'i': r.i = bool(c); break;
          case 'caps': r.caps = bool(c); break;
          case 'smallCaps': r.smallCaps = bool(c); break;
          case 'strike': r.strike = bool(c); break;
          case 'dstrike': r.dstrike = bool(c); break;
          case 'outline': r.outline = bool(c); break;
          case 'shadow': r.shadow = bool(c); break;
          case 'emboss': r.emboss = bool(c); break;
          case 'imprint': r.imprint = bool(c); break;
          case 'noProof': r.noProof = bool(c); break;
          case 'vanish': r.hidden = bool(c); break;
          case 'specVanish': r.specVanish = bool(c); break;
          case 'webHidden': r.webHidden = bool(c); break;
          case 'color': { const v = wColor(c); if (v) r.color = v; break; }
          case 'spacing': r.spacing = tw(at(c, 'val')); break;
          case 'w': r.w = num(c, 'val', 100); break;
          case 'kern': r.kern = num(c, 'val', 0) / 2; break;
          case 'position': r.pos = num(c, 'val', 0) / 2; break;
          case 'sz': r.sz = num(c, 'val', 20) / 2; break;
          case 'szCs': r.szCs = num(c, 'val', 20) / 2; break;
          case 'highlight': if (at(c, 'val') && at(c, 'val') !== 'none') r.hl = at(c, 'val'); else r.hl = 'none'; break;
          case 'u': { const v = at(c, 'val') || 'single'; r.u = v; const uc = wColor(c, 'color'); if (uc && uc !== 'auto') r.uColor = uc; break; }
          case 'effect': if (at(c, 'val') && at(c, 'val') !== 'none') r.effect = at(c, 'val'); break;
          case 'bdr': r.border = border(c); break;
          case 'shd': { const s = shd(c); if (s) r.shd = s; break; }
          case 'vertAlign': { const v = at(c, 'val'); if (v === 'superscript' || v === 'subscript') r.vert = v; else if (v === 'baseline') r.vert = undefined; break; }
          case 'rtl': r.rtl = bool(c); break;
          case 'cs': r.cs = bool(c); break;
          case 'lang': if (at(c, 'val')) r.lang = at(c, 'val'); if (at(c, 'eastAsia')) r.langEA = at(c, 'eastAsia'); break;
          case 'em': if (at(c, 'val') && at(c, 'val') !== 'none') r.em = at(c, 'val'); break;
          case 'bCs': r.bCs = bool(c); break;
          case 'iCs': r.iCs = bool(c); break;
          case 'rPrChange': r.chg = Object.assign(rev(c), { old: rPr(kid(c, 'rPr')) }); break;
          case 'ins': r._ins = rev(c); break;
          case 'del': r._del = rev(c); break;
          default: break;
        }
      }
      return L.preserve.properties(r, el, 'rPr', propertyContext(el));
    }
    const rev = (el) => ({ author: at(el, 'author') || 'Unknown', date: at(el, 'date') || '', id: num(el, 'id', 0) });

    /* ---- paragraph properties ---- */
    const JC = { start: 'left', end: 'right', left: 'left', right: 'right', center: 'center', both: 'both', distribute: 'distribute', justify: 'both', lowKashida: 'both', mediumKashida: 'both', highKashida: 'both', thaiDistribute: 'distribute' };
    function pPr(el) {
      const p = {};
      if (!el) return p;
      for (const c of el.children) {
        if (![L.opc.NS.w, 'http://purl.oclc.org/ooxml/wordprocessingml/main'].includes(c.namespaceURI)) continue;
        switch (c.localName) {
          case 'pStyle': p.style = at(c, 'val'); break;
          case 'keepNext': p.keepNext = bool(c); break;
          case 'keepLines': p.keepLines = bool(c); break;
          case 'pageBreakBefore': p.pageBreakBefore = bool(c); break;
          case 'widowControl': p.widow = bool(c); break;
          case 'numPr': {
            const id = at(kid(c, 'numId'), 'val');
            const lvl = num(kid(c, 'ilvl'), 'val', 0);
            if (id != null) p.num = { id: String(id), lvl };
            else if (kid(c, 'ilvl')) p.num = { lvl, id: undefined };
            break;
          }
          case 'pBdr': p.borders = borders(c, ['top', 'left', 'bottom', 'right', 'between', 'bar']); break;
          case 'shd': { const s = shd(c); if (s) p.shd = s; break; }
          case 'tabs': p.tabs = kids(c, 'tab').map((t) => ({ pos: tw(at(t, 'pos')), al: { start: 'left', end: 'right' }[at(t, 'val')] || at(t, 'val') || 'left', leader: at(t, 'leader') && at(t, 'leader') !== 'none' ? at(t, 'leader') : undefined })); break;
          case 'spacing': {
            const s = {};
            if (at(c, 'before') != null) s.b = tw(at(c, 'before'));
            if (at(c, 'after') != null) s.a = tw(at(c, 'after'));
            if (at(c, 'beforeLines') != null && at(c, 'before') == null) s.b = num(c, 'beforeLines', 0) / 100 * 12;
            if (at(c, 'afterLines') != null && at(c, 'after') == null) s.a = num(c, 'afterLines', 0) / 100 * 12;
            if (at(c, 'beforeAutospacing') != null) s.bAuto = X.bool(at(c, 'beforeAutospacing'));
            if (at(c, 'afterAutospacing') != null) s.aAuto = X.bool(at(c, 'afterAutospacing'));
            if (at(c, 'line') != null) {
              const rule = at(c, 'lineRule') || 'auto';
              s.rule = rule;
              s.line = rule === 'auto' ? num(c, 'line', 240) / 240 : tw(at(c, 'line'));
            }
            p.sp = s;
            break;
          }
          case 'ind': {
            const o = {};
            const l = at(c, 'left') != null ? at(c, 'left') : at(c, 'start');
            const r = at(c, 'right') != null ? at(c, 'right') : at(c, 'end');
            if (l != null) o.l = tw(l);
            if (r != null) o.r = tw(r);
            if (at(c, 'hanging') != null) o.fl = -tw(at(c, 'hanging'));
            else if (at(c, 'firstLine') != null) o.fl = tw(at(c, 'firstLine'));
            if (at(c, 'leftChars') != null && l == null) o.l = num(c, 'leftChars', 0) / 100 * 10.5;
            if (at(c, 'firstLineChars') != null && at(c, 'firstLine') == null && at(c, 'hanging') == null) o.fl = num(c, 'firstLineChars', 0) / 100 * 10.5;
            p.ind = o;
            break;
          }
          case 'contextualSpacing': p.contextual = bool(c); break;
          case 'jc': p.jc = JC[at(c, 'val')] || 'left'; break;
          case 'textDirection': p.textDir = at(c, 'val'); break;
          case 'outlineLvl': p.outline = num(c, 'val', 9); break;
          case 'bidi': p.bidi = bool(c); break;
          case 'mirrorIndents': p.mirrorInd = bool(c); break;
          case 'suppressAutoHyphens': p.noHyphen = bool(c); break;
          case 'suppressLineNumbers': p.noLineNum = bool(c); break;
          case 'framePr': {
            const f = {};
            if (at(c, 'dropCap') && at(c, 'dropCap') !== 'none') { p.dropCap = { type: at(c, 'dropCap'), lines: num(c, 'lines', 3), dist: tw(at(c, 'hSpace') || 0), wrap: at(c, 'wrap') }; break; }
            if (at(c, 'w')) f.w = tw(at(c, 'w'));
            if (at(c, 'h')) f.h = tw(at(c, 'h'));
            if (at(c, 'hRule')) f.hRule = at(c, 'hRule');
            if (at(c, 'x')) f.x = tw(at(c, 'x'));
            if (at(c, 'y')) f.y = tw(at(c, 'y'));
            if (at(c, 'xAlign')) f.xAlign = at(c, 'xAlign');
            if (at(c, 'yAlign')) f.yAlign = at(c, 'yAlign');
            if (at(c, 'hAnchor')) f.hAnchor = at(c, 'hAnchor');
            if (at(c, 'vAnchor')) f.vAnchor = at(c, 'vAnchor');
            if (at(c, 'wrap')) f.wrap = at(c, 'wrap');
            if (at(c, 'hSpace')) f.hSpace = tw(at(c, 'hSpace'));
            if (at(c, 'vSpace')) f.vSpace = tw(at(c, 'vSpace'));
            p.frame = f;
            break;
          }
          case 'pPrChange': p.chg = Object.assign(rev(c), { old: pPr(kid(c, 'pPr')) }); break;
          case 'snapToGrid': if (!bool(c)) p.noSnap = true; break;
          case 'textAlignment': p.textAlign = at(c, 'val'); break;
          case 'kinsoku': case 'wordWrap': case 'overflowPunct': case 'autoSpaceDE': case 'autoSpaceDN': case 'adjustRightInd': case 'topLinePunct': break;
          default: break;
        }
      }
      return L.preserve.properties(p, el, 'pPr', propertyContext(el));
    }

    /* ---- styles ---- */
    const stylesRel = relOfType('styles');
    const stX = stylesRel ? await xml(stylesRel.target) : null;
    const defaults = { rPr: { font: 'Times New Roman', sz: 10 }, pPr: {} };
    const imported = {};
    if (stX) {
      const dd = kid(stX, 'docDefaults');
      if (dd) {
        const rd = X.path(dd, 'rPrDefault', 'rPr'), pd = X.path(dd, 'pPrDefault', 'pPr');
        Object.assign(defaults.rPr, rPr(rd));
        defaults.pPr = pPr(pd);
      }
      for (const s of kids(stX, 'style')) {
        /* some generators omit styleId: Word then matches the style by name */
        const id = at(s, 'styleId') || (at(kid(s, 'name'), 'val') || '').replace(/\s+/g, '');
        if (!id) continue;
        const type = at(s, 'type') || 'paragraph';
        const st = { id, name: at(kid(s, 'name'), 'val') || id, type, basedOn: at(kid(s, 'basedOn'), 'val') || null, next: at(kid(s, 'next'), 'val') || null, link: at(kid(s, 'link'), 'val') || null, pPr: pPr(kid(s, 'pPr')), rPr: rPr(kid(s, 'rPr')) };
        if (X.bool(at(s, 'default'))) st.isDefault = true;
        if (X.bool(at(s, 'customStyle'))) st.custom = true; else st.builtin = true;
        if (kid(s, 'qFormat')) st.q = true;
        if (kid(s, 'semiHidden')) st.semiHidden = true;
        if (kid(s, 'hidden')) st.hidden = true;
        if (kid(s, 'uiPriority')) st.prio = num(kid(s, 'uiPriority'), 'val', 99);
        if (kid(s, 'autoRedefine')) st.autoRedefine = true;
        if (type === 'table') {
          const tp = tblPr(kid(s, 'tblPr'));
          st.tblPr = tp;
          const tc = kid(s, 'tcPr');
          if (tc) st.tcPr = tcPr(tc);
          st.cond = {};
          for (const c of kids(s, 'tblStylePr')) {
            const ty = at(c, 'type');
            if (!ty) continue;
            const o = {};
            if (kid(c, 'pPr')) o.pPr = pPr(kid(c, 'pPr'));
            if (kid(c, 'rPr')) o.rPr = rPr(kid(c, 'rPr'));
            if (kid(c, 'tcPr')) o.tcPr = tcPr(kid(c, 'tcPr'));
            if (kid(c, 'tblPr')) o.tblPr = tblPr(kid(c, 'tblPr'));
            st.cond[ty] = o;
          }
        }
        if (type === 'numbering' && st.pPr.num) st.numStyle = true;
        imported[id] = st;
      }
    }
    /* map defaults to the canonical ids the engine uses */
    const rename = {};
    const CANON = { paragraph: 'Normal', character: 'DefaultParagraphFont', table: 'TableNormal', numbering: 'NoList' };
    const haveDefault = {};
    /* one default per type (files may mark several): the canonical id wins, otherwise the first one */
    for (const ty in CANON) if (imported[CANON[ty]] && imported[CANON[ty]].type === ty && imported[CANON[ty]].isDefault) haveDefault[ty] = true;
    for (const id in imported) {
      const s = imported[id];
      if (!s.isDefault || !CANON[s.type] || id === CANON[s.type]) continue;
      if (haveDefault[s.type] || imported[CANON[s.type]]) { delete s.isDefault; continue; }
      rename[id] = CANON[s.type];
      haveDefault[s.type] = true;
    }
    const rn = (id) => (id && rename[id] ? rename[id] : id);
    const styles = D.builtinStyles();
    /* imported documents bring their own look: built-ins they don't define are kept, but Normal etc. come from the file */
    for (const id in imported) {
      const s = imported[id];
      const nid = rn(id);
      s.id = nid;
      s.basedOn = rn(s.basedOn);
      s.next = rn(s.next);
      s.link = rn(s.link);
      if (nid !== id && styles[id]) delete styles[id];
      styles[nid] = s;
    }
    /* built-in heading styles not present in the file: inherit the document's Normal look but keep their emphasis */
    if (stX) {
      for (const k of Object.keys(styles)) {
        if (imported[k] || Object.values(rename).includes(k)) continue;
        const b = styles[k];
        if (/^Heading[1-9]$|^Title$|^Subtitle$/.test(k)) { b.rPr = Object.assign({}, b.rPr); }
      }
    }
    /* a basedOn loop (possible in hand-made files) would hang other readers when saved: cut it */
    for (const id in styles) {
      const seen = new Set([id]);
      let cur = styles[id];
      while (cur && cur.basedOn) {
        if (seen.has(cur.basedOn) || !styles[cur.basedOn]) { if (seen.has(cur.basedOn)) cur.basedOn = null; break; }
        seen.add(cur.basedOn);
        cur = styles[cur.basedOn];
      }
    }
    doc.styles = styles;
    doc.defaults = defaults;
    if (!styles.Normal) styles.Normal = { id: 'Normal', name: 'Normal', type: 'paragraph', pPr: {}, rPr: {}, builtin: true, isDefault: true };
    styles.Normal.isDefault = true;

    /* ---- media ---- */
    const mediaCache = new Map();
    async function loadMedia(p) {
      if (!p) return null;
      if (mediaCache.has(p)) return mediaCache.get(p);
      const e = get(p);
      if (!e) { mediaCache.set(p, null); return null; }
      let bytes;
      try { bytes = await e.bytes(); } catch (err) { mediaCache.set(p, null); return null; }
      const ext = (p.split('.').pop() || '').toLowerCase();
      let mime = L.extToMime(ext);
      if (mime === 'application/octet-stream') mime = sniff(bytes) || mime;
      let view = null;
      if (ext === 'emf' || ext === 'wmf' || mime === 'image/x-emf' || mime === 'image/x-wmf') {
        view = await L.metafile.toPNG(bytes, ext === 'wmf' || mime === 'image/x-wmf' ? 'wmf' : 'emf', 1600);
      }
      const id = L.media.add(new Blob([bytes], { type: mime }), p.split('/').pop(), view);
      doc.keep.media[id] = p;
      mediaCache.set(p, id);
      return id;
    }
    function sniff(b) {
      if (b[0] === 0x89 && b[1] === 0x50) return 'image/png';
      if (b[0] === 0xFF && b[1] === 0xD8) return 'image/jpeg';
      if (b[0] === 0x47 && b[1] === 0x49) return 'image/gif';
      if (b[0] === 0x42 && b[1] === 0x4D) return 'image/bmp';
      if (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x00 && b[3] === 0x00) return 'image/x-emf';
      if (b[0] === 0xD7 && b[1] === 0xCD) return 'image/x-wmf';
      return null;
    }
    /** preload every image/chart/diagram referenced by a part's relationships */
    async function preload(relMap) {
      for (const k in relMap) {
        const r = relMap[k];
        if (r.external) continue;
        if (r.type === 'image') await loadMedia(r.target);
        else if (r.type === 'chart') await loadChart(r.target);
        else if (r.type === 'diagramDrawing' || r.type === 'diagramData' || r.type === 'diagramLayout') await xml(r.target);
      }
    }
    const chartCache = new Map();
    async function loadChart(p) {
      if (chartCache.has(p)) return chartCache.get(p);
      const cx = await xml(p);
      let model = null;
      try { model = X.chart(cx, dctx); } catch (e) { model = null; }
      /* keep the original parts so the chart round-trips unchanged */
      const files = [];
      const add = async (path) => { const e = get(path); if (e) files.push({ path, data: await e.bytes() }); };
      await add(p);
      const relp = X.relsPath(p);
      if (get(relp)) {
        await add(relp);
        const cr = await rels(p);
        for (const k in cr) if (!cr[k].external) await add(cr[k].target);
      }
      // History and clipboard snapshots use JSON. Keep only a cache key in runs;
      // the original typed-array bytes must stay outside those snapshots.
      const out = { model, src: L.chart.keep({ files, path: p, source: doc.pkg?.id }) };
      chartCache.set(p, out);
      return out;
    }

    /* ---- numbering ---- */
    const numRel = relOfType('numbering');
    const numX = numRel ? await xml(numRel.target) : null;
    const numPicMedia = {};
    const numRels = numRel ? await rels(numRel.target) : {};
    if (numX) {
      for (const pb of kids(numX, 'numPicBullet')) {
        const im = desc(pb, 'imagedata') || desc(pb, 'blip');
        const r = im && (rid(im, 'id') || rid(im, 'embed'));
        if (r && numRels[r]) { const m = await loadMedia(numRels[r].target); if (m) numPicMedia[at(pb, 'numPicBulletId')] = m; }
      }
      const lvlOf = (lv) => {
        const fmtEl = kid(lv, 'numFmt');
        const o = {
          start: num(kid(lv, 'start'), 'val', 1), fmt: at(fmtEl, 'val') || 'decimal', text: kid(lv, 'lvlText') ? at(kid(lv, 'lvlText'), 'val') || '' : '%' + (num(lv, 'ilvl', 0) + 1) + '.',
          jc: JC[at(kid(lv, 'lvlJc'), 'val')] || 'left', suff: at(kid(lv, 'suff'), 'val') || 'tab', rPr: rPr(kid(lv, 'rPr')),
        };
        if (fmtEl && at(fmtEl, 'format')) o.custFmt = at(fmtEl, 'format');
        const pp = pPr(kid(lv, 'pPr'));
        if (pp.ind) o.ind = { l: pp.ind.l || 0, fl: pp.ind.fl || 0 };
        if (pp.tabs && pp.tabs.length) { const nt = pp.tabs.find((t) => t.al === 'num' || t.al === 'left'); if (nt) o.tabPos = nt.pos; }
        o.pPr = pp;
        o.pPrValues = { ind: L.clone(o.ind), tabPos: o.tabPos };
        if (kid(lv, 'isLgl')) o.isLgl = true;
        if (kid(lv, 'lvlRestart')) o.restart = num(kid(lv, 'lvlRestart'), 'val', 0) - 1;
        if (kid(lv, 'pStyle')) o.pStyle = rn(at(kid(lv, 'pStyle'), 'val'));
        if (kid(lv, 'lvlPicBulletId')) { const m = numPicMedia[at(kid(lv, 'lvlPicBulletId'), 'val')]; if (m) o.picture = m; }
        if (o.fmt === 'bullet') o.glyph = bulletGlyph(o.text, o.rPr.font);
        if (kid(lv, 'legacy')) o.legacy = true;
        return o;
      };
      for (const an of kids(numX, 'abstractNum')) {
        const id = at(an, 'abstractNumId');
        const levels = [];
        for (const lv of kids(an, 'lvl')) levels[num(lv, 'ilvl', 0)] = lvlOf(lv);
        for (let i = 0; i < 9; i++) if (!levels[i]) levels[i] = D.numLevel(i, 'decimal', `%${i + 1}.`);
        const a = { id, levels, multi: at(kid(an, 'multiLevelType'), 'val') || 'hybridMultilevel' };
        if (kid(an, 'numStyleLink')) a.styleLink = rn(at(kid(an, 'numStyleLink'), 'val'));
        if (kid(an, 'styleLink')) a.styleDef = rn(at(kid(an, 'styleLink'), 'val'));
        if (kid(an, 'name')) a.name = at(kid(an, 'name'), 'val');
        doc.numbering.abs[id] = a;
      }
      for (const n of kids(numX, 'num')) {
        const id = at(n, 'numId');
        const ov = {};
        for (const o of kids(n, 'lvlOverride')) {
          const l = num(o, 'ilvl', 0);
          const e = {};
          if (kid(o, 'startOverride')) e.start = num(kid(o, 'startOverride'), 'val', 1);
          if (kid(o, 'lvl')) e.lvl = lvlOf(kid(o, 'lvl'));
          ov[l] = e;
        }
        doc.numbering.nums[id] = { abs: at(kid(n, 'abstractNumId'), 'val'), ov };
      }
      /* numStyleLink: an abstract definition that defers to a numbering style */
      for (const k in doc.numbering.abs) {
        const a = doc.numbering.abs[k];
        if (!a.styleLink) continue;
        const st = styles[a.styleLink];
        const nid = st && st.pPr && st.pPr.num && st.pPr.num.id;
        const target = nid && doc.numbering.nums[nid] && doc.numbering.abs[doc.numbering.nums[nid].abs];
        if (target && target !== a) a.levels = target.levels;
      }
    }
    function bulletGlyph(text, font) {
      if (!text) return '';
      const f = String(font || '').toLowerCase();
      if (/symbol|wingdings|webdings/.test(f)) return L.mapSymbolChar(text, font);
      const c = text.charCodeAt(0);
      if (c >= 0xF000 && c <= 0xF0FF) return L.mapSymbolChar(text, font);
      return text;
    }

    /* ---- settings ---- */
    const setRel = relOfType('settings');
    const setX = setRel ? await xml(setRel.target) : null;
    if (setX) {
      const s = doc.settings;
      for (const c of setX.children) {
        switch (c.localName) {
          case 'defaultTabStop': s.defTab = tw(at(c, 'val')) || 36; break;
          case 'evenAndOddHeaders': s.evenOdd = bool(c); break;
          case 'trackRevisions': s.track = bool(c); break;
          case 'mirrorMargins': s.mirror = bool(c); break;
          case 'gutterAtTop': s.gutterTop = bool(c); break;
          case 'autoHyphenation': s.autoHyphen = bool(c); break;
          case 'displayBackgroundShape': s.showBg = bool(c); break;
          case 'updateFields': s.updateFields = bool(c); break;
          case 'documentProtection': {
            const ed = at(c, 'edit');
            const enf = X.bool(at(c, 'enforcement'));
            if (ed && ed !== 'none') s.protect = { kind: ed === 'trackedChanges' ? 'tracked' : ed, on: enf, hash: at(c, 'hashValue') || at(c, 'hash') || null, raw: Array.from(c.attributes).reduce((o, a) => { o[a.name] = a.value; return o; }, {}) };
            break;
          }
          case 'footnotePr': s.fnPr = Object.assign({}, s.fnPr, notePr(c)); break;
          case 'endnotePr': s.enPr = Object.assign({}, s.enPr, notePr(c)); break;
          case 'compat': {
            for (const cs of kids(c, 'compatSetting')) {
              if (at(cs, 'name') === 'compatibilityMode') s.compat = +at(cs, 'val') || 12;
              if (at(cs, 'name') === 'overrideTableStyleFontSizeAndJustification' && X.bool(at(cs, 'val'))) s.tblFsJcOverride = true;
            }
            /* keep legacy layout options (they change how Word and LibreOffice lay the document out) */
            const ser = new XMLSerializer();
            s.compatXML = kids(c).filter((k) => !(k.localName === 'compatSetting' && at(k, 'name') === 'compatibilityMode')).map((k) => ser.serializeToString(k).replace(/ xmlns(:\w+)?="[^"]*"/g, '')).join('');
            if (kid(c, 'doNotUseHTMLParagraphAutoSpacing')) s.noHtmlAuto = true;
            break;
          }
          case 'zoom': s.zoom = num(c, 'percent', 100); break;
          case 'docVars': s.docVars = kids(c, 'docVar').map((v) => ({ name: at(v, 'name'), val: at(v, 'val') })); break;
          case 'attachedTemplate': s.template = rid(c); break;
          case 'proofState': break;
          default: break;
        }
      }
    }
    function notePr(c) {
      const o = {};
      if (kid(c, 'numFmt')) o.fmt = at(kid(c, 'numFmt'), 'val');
      if (kid(c, 'numStart')) o.start = num(kid(c, 'numStart'), 'val', 1);
      if (kid(c, 'pos')) o.pos = at(kid(c, 'pos'), 'val');
      if (kid(c, 'numRestart')) o.restart = at(kid(c, 'numRestart'), 'val');
      return o;
    }

    /* ---- story parsing ---- */
    let fid = 0;
    const commentRanges = new Map();
    function storyCtx(partPath, relMap, story) {
      if (!commentRanges.has(partPath)) commentRanges.set(partPath, new Set(Array.from(xmlCache.get(partPath)?.getElementsByTagName('*') || []).filter(e => ['commentRangeStart', 'commentRangeEnd'].includes(e.localName)).map(e => at(e, 'id'))));
      return { part: partPath, rels: relMap, story, stores, vmlTypes, fstack: [], dropFids: new Set(), cmtOpen: new Set(), cmtRanges: commentRanges.get(partPath), pendingBm: [], hyper: new Map() };
    }

    function parseBlocks(el, ctx, out) {
      out = out || [];
      const ac = L.preserve.readAlternates(el, out, ctx, doc, 'block');
      for (const c of el.children) {
        const from = out.length;
        switch (c.localName) {
          case 'p': { const p = parseP(c, ctx); if (p) out.push(p); break; }
          case 'tbl': { const t = parseTbl(c, ctx); if (t) out.push(t); break; }
          case 'sdt': {
            const sc = kid(c, 'sdtContent');
            if (!sc) break;
            const n0 = out.length;
            parseBlocks(sc, ctx, out);
            const bv = boundText(kid(c, 'sdtPr'));
            if (bv != null) { const ps = out.slice(n0).filter((b) => b.t === 'p'); if (ps.length === 1) ps[0].runs = boundRuns(bv, ps[0].runs); }
            if (out.length === n0) out.push(D.para());
            const owned = out.slice(n0), control = L.preserve.control(c, ctx, doc);
            if (control) { const first = owned[0], last = owned.at(-1); const a = first.t === 'p' ? first.pPr : first.tblPr, b = last.t === 'p' ? last.pPr : last.tblPr; (a.sdts || (a.sdts = [])).unshift(control); (b.sdte || (b.sdte = [])).push(control.key); }
            break;
          }
          case 'customXml': case 'ins': case 'moveTo': case 'smartTag': parseBlocks(c, ctx, out); break;
          case 'del': case 'moveFrom': parseBlocks(c, ctx, out); break;
          case 'bookmarkStart': case 'permStart': ctx.pendingBm.push(L.preserve.rangeMarker(c, { pkg: doc.pkg, part: ctx.part })); break;
          case 'bookmarkEnd': case 'permEnd': { const last = lastParaIn(out), item = L.preserve.rangeMarker(c, { pkg: doc.pkg, part: ctx.part }); if (last) last.runs.push(item); else ctx.pendingBm.push(item); break; }
          case 'commentRangeStart': ctx.pendingBm.push(D.item('cs', { id: at(c, 'id') })); break;
          case 'commentRangeEnd': { const last = lastParaIn(out); const it = D.item('ce', { id: at(c, 'id') }); if (last) last.runs.push(it); else ctx.pendingBm.push(it); break; }
          case 'altChunk': out.push(L.preserve.opaqueBlock(c, ctx, doc, 'Embedded document')); break;
          default: break;
        }
        ac.add(c, from);
      }
      ac.finish();
      return out;
    }
    const lastParaIn = (bl) => { for (let i = bl.length - 1; i >= 0; i--) { if (bl[i].t === 'p') return bl[i]; if (bl[i].t === 'tbl') return D.lastPara([bl[i]]); } return null; };

    function parseP(el, ctx) {
      const pp = kid(el, 'pPr');
      const p = D.para([], pPr(pp), {});
      if (p.pPr.style) p.pPr.style = rn(p.pPr.style);
      if (p.pPr.style === 'Normal') delete p.pPr.style;
      if (pp) {
        const mr = kid(pp, 'rPr');
        if (mr) {
          const r = rPr(mr);
          if (r._ins) { p.mark = Object.assign(p.mark || {}, { ins: r._ins }); delete r._ins; }
          if (r._del) { p.mark = Object.assign(p.mark || {}, { del: r._del }); delete r._del; }
          if (r.style) r.style = rn(r.style);
          p.rPr = r;
        }
        const sp = kid(pp, 'sectPr');
        if (sp) p.sect = parseSect(sp, ctx);
      }
      if (ctx.pendingBm.length) { p.runs.push(...ctx.pendingBm); ctx.pendingBm = []; }
      parseInline(el, p, ctx, {});
      D.normalize(p);
      return p;
    }

    /* inline content */
    function parseInline(el, p, ctx, inh) {
      const ac = L.preserve.readAlternates(el, p.runs, ctx, doc, 'inline', inh);
      for (const c of el.children) {
        const from = p.runs.length;
        switch (c.localName) {
          case 'r': parseRun(c, p, ctx, inh); break;
          case 'hyperlink': {
            const r = rid(c);
            const link = {};
            if (r && ctx.rels[r]) link.url = ctx.rels[r].raw;
            if (at(c, 'anchor')) link.anchor = at(c, 'anchor');
            if (at(c, 'tooltip')) link.tip = at(c, 'tooltip');
            if (at(c, 'tgtFrame')) link.target = at(c, 'tgtFrame');
            parseInline(c, p, ctx, Object.assign({}, inh, { link: link.url || link.anchor ? link : undefined }));
            break;
          }
          case 'ins': case 'moveTo': parseInline(c, p, ctx, Object.assign({}, inh, { ins: rev(c) })); break;
          case 'del': case 'moveFrom': parseInline(c, p, ctx, Object.assign({}, inh, { del: rev(c) })); break;
          case 'fldSimple': {
            const instr = at(c, 'instr') || '';
            const f = 'f' + ++fid;
            const ty = D.fieldType(instr);
            if (ty === 'HYPERLINK') { parseInline(c, p, ctx, Object.assign({}, inh, { link: parseHyperInstr(instr) })); break; }
            const firstR = kid(c, 'r');
            const base = firstR ? rPr(kid(firstR, 'rPr')) : {};
            p.runs.push(D.item('fb', { fid: f, instr }, L.clone(base)));
            p.runs.push(D.item('fs', { fid: f }, L.clone(base)));
            parseInline(c, p, ctx, inh);
            p.runs.push(D.item('fe', { fid: f }, L.clone(base)));
            break;
          }
          case 'smartTag': case 'customXml': case 'dir': case 'bdo': parseInline(c, p, ctx, inh); break;
          case 'sdt': {
            const sc = kid(c, 'sdtContent');
            if (!sc) break;
            const bv = boundText(kid(c, 'sdtPr'));
            const tmp = D.para();
            parseInline(sc, tmp, ctx, inh);
            const control = L.preserve.control(c, ctx, doc);
            if (control) p.runs.push(D.item('sdts', { control }));
            p.runs.push(...(bv == null ? tmp.runs : boundRuns(bv, tmp.runs)));
            if (control) p.runs.push(D.item('sdte', { key: control.key }));
            break;
          }
          case 'bookmarkStart': case 'bookmarkEnd': case 'permStart': case 'permEnd': p.runs.push(L.preserve.rangeMarker(c, { pkg: doc.pkg, part: ctx.part })); break;
          case 'commentRangeStart': p.runs.push(D.item('cs', { id: at(c, 'id') })); ctx.cmtOpen.add(at(c, 'id')); break;
          case 'commentRangeEnd': p.runs.push(D.item('ce', { id: at(c, 'id') })); ctx.cmtSeen = ctx.cmtSeen || new Set(); ctx.cmtSeen.add(at(c, 'id')); break;
          case 'oMath': case 'oMathPara': p.runs.push(mathItem(c)); break;
          case 'r_': break;
          default: break;
        }
        ac.add(c, from);
      }
      ac.finish();
    }
    function mathItem(c) {
      const text = descAll(c, 't').map((t) => t.textContent).join('');
      const sz = L.omml ? L.omml.fontSize(c) : null;
      return D.item('raw', { math: true, text, xml: new XMLSerializer().serializeToString(c), para: c.localName === 'oMathPara' }, sz ? { sz } : {});
    }
    function parseHyperInstr(instr) {
      const link = {};
      const m = /HYPERLINK\s+(?:"([^"]*)"|(\S+))?/i.exec(instr);
      if (m && (m[1] || m[2]) && !/^\\/.test(m[2] || '')) link.url = m[1] || m[2];
      const l = /\\l\s+"([^"]*)"/.exec(instr);
      if (l) link.anchor = l[1];
      const o = /\\o\s+"([^"]*)"/.exec(instr);
      if (o) link.tip = o[1];
      return link.url || link.anchor ? link : undefined;
    }
    function topField(ctx) { return ctx.fstack[ctx.fstack.length - 1]; }
    function inInstr(ctx) { return ctx.fstack.some((f) => f.phase === 'instr'); }
    function fieldLink(ctx) { for (let i = ctx.fstack.length - 1; i >= 0; i--) if (ctx.fstack[i].link && ctx.fstack[i].phase === 'result') return ctx.fstack[i].link; return undefined; }

    function parseRun(r, p, ctx, inh) {
      const base = rPr(kid(r, 'rPr'));
      delete base._ins; delete base._del;
      if (base.style) base.style = rn(base.style);
      if (inh.link) base.link = inh.link;
      if (inh.ins) base.ins = inh.ins;
      if (inh.del) base.del = inh.del;
      const fl = fieldLink(ctx);
      if (fl && !base.link) base.link = fl;
      const push = (it) => {
        if (inInstr(ctx) && it.t !== 'fb' && it.t !== 'fs' && it.t !== 'fe') return;
        p.runs.push(it);
      };
      const text = s => { if (s) push(D.text(s, base)); };
      let customMark = false;
      const ac = L.preserve.readAlternates(r, p.runs, ctx, doc, 'run', base);
      for (const c of r.children) {
        const from = p.runs.length;
        switch (c.localName) {
          case 'rPr': break;
          case 't': {
            /* without xml:space="preserve", line breaks from pretty-printed XML (and the indentation around
               them) are dropped at the ends and folded to a space inside; plain spaces are kept, as in Word */
            let tx = c.textContent;
            if (/[\r\n]/.test(tx) && c.getAttribute('xml:space') !== 'preserve' && c.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'space') !== 'preserve') tx = tx.replace(/^[ \t]*[\r\n][ \t\r\n]*|[ \t\r\n]*[\r\n][ \t]*$/g, '').replace(/[ \t]*[\r\n]+[ \t]*/g, ' ');
            text(tx);
            break;
          }
          case 'delText': if (!inInstr(ctx)) text(c.textContent); break;
          case 'tab': push(D.item('tab', null, base)); break;
          case 'br': {
            const ty = at(c, 'type');
            push(D.item('br', { type: ty === 'page' ? 'page' : ty === 'column' ? 'column' : 'line', clear: at(c, 'clear') || undefined }, base));
            break;
          }
          case 'cr': push(D.item('br', { type: 'line' }, base)); break;
          case 'noBreakHyphen': text('‑'); break;
          case 'softHyphen': text('­'); break;
          case 'sym': {
            const font = at(c, 'font');
            const ch = String.fromCharCode(parseInt(at(c, 'char') || '0', 16) || 0x3F);
            push(D.item('sym', { font, char: ch, code: at(c, 'char') }, base));
            break;
          }
          case 'ptab': push(D.item('ptab', { al: at(c, 'alignment') || 'left', rel: at(c, 'relativeTo') || 'margin', leader: at(c, 'leader') || 'none' }, base)); break;
          case 'fldChar': {
            const ty = at(c, 'fldCharType');
            if (ty === 'begin') {
              if (inInstr(ctx)) { ctx.fstack.push({ nested: true, phase: 'instr' }); break; }
              const f = 'f' + ++fid;
              const item = D.item('fb', { fid: f, instr: '' }, L.clone(base));
              if (X.bool(at(c, 'fldLock'))) item.lock = true;
              const ffd = kid(c, 'ffData');
              if (ffd) {
                const ff = { name: at(kid(ffd, 'name'), 'val') || '' };
                const cb = kid(ffd, 'checkBox'), ti = kid(ffd, 'textInput'), dd = kid(ffd, 'ddList');
                if (cb) { ff.type = 'checkbox'; ff.def = bool(kid(cb, 'default'), false); ff.checked = kid(cb, 'checked') ? bool(kid(cb, 'checked')) : ff.def; ff.size = kid(cb, 'size') ? num(kid(cb, 'size'), 'val', 20) / 2 : null; }
                else if (ti) { ff.type = 'text'; ff.def = at(kid(ti, 'default'), 'val') || ''; ff.maxLength = num(kid(ti, 'maxLength'), 'val', 0); ff.format = at(kid(ti, 'format'), 'val') || ''; ff.textType = at(kid(ti, 'type'), 'val') || 'regular'; }
                else if (dd) { ff.type = 'dropdown'; ff.entries = kids(dd, 'listEntry').map((e) => at(e, 'val')); ff.result = num(kid(dd, 'result'), 'val', 0); }
                if (kid(ffd, 'helpText')) ff.help = at(kid(ffd, 'helpText'), 'val');
                if (kid(ffd, 'statusText')) ff.status = at(kid(ffd, 'statusText'), 'val');
                item.ff = ff;
              }
              if (X.bool(at(c, 'dirty'))) item.dirty = true;
              p.runs.push(item);
              ctx.fstack.push({ fid: f, item, phase: 'instr', p });
            } else if (ty === 'separate') {
              const t = topField(ctx);
              if (!t) break;
              if (t.nested) { t.phase = 'result'; break; }
              t.phase = 'result';
              const fty = D.fieldType(t.item.instr);
              if (fty === 'HYPERLINK') { t.link = parseHyperInstr(t.item.instr); if (t.link) { ctx.dropFids.add(t.fid); } }
              p.runs.push(D.item('fs', { fid: t.fid }, L.clone(base)));
            } else if (ty === 'end') {
              const t = ctx.fstack.pop();
              if (!t || t.nested) break;
              if (t.phase === 'instr' && D.fieldType(t.item.instr) === 'HYPERLINK') ctx.dropFids.add(t.fid);
              p.runs.push(D.item('fe', { fid: t.fid }, L.clone(base)));
            }
            break;
          }
          case 'instrText': { const t = topField(ctx); if (t && t.phase === 'instr' && !t.nested && ctx.fstack.filter((f) => f.phase === 'instr').length === 1) t.item.instr += c.textContent; break; }
          case 'delInstrText': { const t = topField(ctx); if (t && t.phase === 'instr' && !t.nested) t.item.instr += c.textContent; break; }
          case 'footnoteReference': case 'endnoteReference': {
            const kind = c.localName === 'footnoteReference' ? 'fn' : 'en';
            const it = D.item(kind, { id: at(c, 'id') }, base);
            if (X.bool(at(c, 'customMarkFollows'))) { customMark = it; }
            push(it);
            break;
          }
          case 'footnoteRef': push(D.item('fn', { id: ctx.noteId, self: true }, base)); break;
          case 'endnoteRef': push(D.item('en', { id: ctx.noteId, self: true }, base)); break;
          case 'separator': push(D.item('sep', null, base)); break;
          case 'continuationSeparator': push(D.item('csep', null, base)); break;
          case 'commentReference': {
            const id = at(c, 'id');
            /* comments without a range get an empty range at the reference */
            // A reference may precede its real range. Point comments have virtual
            // editor boundaries, but must not add duplicate OOXML range IDs.
            if (!ctx.cmtRanges.has(id)) { p.runs.push(D.item('cs', { id, point: true })); p.runs.push(D.item('ce', { id, point: true })); }
            break;
          }
          case 'drawing': { const it = parseDrawing(c, ctx, base); push(L.preserve.keepObject(it, c, ctx, doc, it?.diagram ? 'SmartArt' : it?.alt || 'Drawing')); break; }
          case 'pict': case 'object': {
            const it = parseVML(c, ctx, base);
            if (it) push(L.preserve.keepObject(it, c, ctx, doc, c.localName === 'object' ? 'Embedded object' : 'Drawing'));
            else if (!Array.from(c.getElementsByTagName('*')).some(e => /WaterMark/i.test(at(e, 'id') || ''))) push(L.preserve.keepObject(null, c, ctx, doc, 'Drawing'));
            break;
          }
          case 'control': case 'contentPart': { const it = L.preserve.keepObject(null, c, ctx, doc, c.localName === 'control' ? 'Form control' : 'Embedded content'); it.rPr = base; push(it); break; }
          case 'pgNum': { const f = 'f' + ++fid; push(D.item('fb', { fid: f, instr: ' PAGE ' }, base)); push(D.item('fs', { fid: f }, base)); text('1'); push(D.item('fe', { fid: f }, base)); break; }
          case 'lastRenderedPageBreak': case 'annotationRef': break;
          case 'ruby': {
            /* phonetic guide: kept as one inline object (base text with its annotation above) */
            const rb = kid(c, 'rubyBase'), rtEl = kid(c, 'rt');
            const txt = (el) => (el ? descAll(el, 't').map((t) => t.textContent).join('') : '');
            const firstRPr = (el) => { const rr = el && kid(el, 'r'); return rr ? rPr(kid(rr, 'rPr')) : {}; };
            const pr = kid(c, 'rubyPr');
            const o = {};
            if (pr) for (const k of pr.children) { const v = at(k, 'val'); if (v != null) o[k.localName] = v; }
            push(D.item('ruby', { base: txt(rb), rt: txt(rtEl), baseRPr: firstRPr(rb), rtRPr: firstRPr(rtEl), pr: o }, base));
            break;
          }
          case 'dayShort': case 'monthShort': case 'yearShort': case 'dayLong': case 'monthLong': case 'yearLong': break;
          default: break;
        }
        ac.add(c, from);
      }
      ac.finish();
      if (customMark) {
        /* the custom footnote mark is the run's own text */
        const txt = descAll(r, 't').map((x) => x.textContent).join('');
        if (txt) { customMark.custom = txt; const last = p.runs[p.runs.length - 1]; if (last && last.t === 'text' && last.text === txt) p.runs.pop(); }
      }
    }

    /* ---- tables ---- */
    function tblPr(el) {
      const o = {};
      if (!el) return o;
      for (const c of el.children) {
        switch (c.localName) {
          case 'tblStyle': o.style = rn(at(c, 'val')); break;
          case 'tblW': o.w = { type: at(c, 'type') || 'dxa', v: at(c, 'type') === 'pct' ? (String(at(c, 'w')).endsWith('%') ? parseFloat(at(c, 'w')) * 50 : num(c, 'w', 0)) : tw(at(c, 'w')) }; break;
          case 'jc': o.jc = JC[at(c, 'val')] || at(c, 'val'); break;
          case 'tblInd': o.ind = tw(at(c, 'w')); break;
          case 'tblBorders': o.borders = borders(c, ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']); break;
          case 'shd': { const s = shd(c); if (s) o.shd = s; break; }
          case 'tblLayout': o.layout = at(c, 'type') || 'autofit'; break;
          case 'tblCellMar': { const m = {}; for (const k of kids(c)) { const n = { left: 'l', start: 'l', right: 'r', end: 'r', top: 't', bottom: 'b' }[k.localName]; if (n) m[n] = tw(at(k, 'w')); } o.cellMar = m; break; }
          case 'tblCellSpacing': o.spacing = tw(at(c, 'w')); break;
          case 'tblLook': {
            const v = at(c, 'val');
            const look = {};
            if (v && /^[0-9a-f]+$/i.test(v)) { const n = parseInt(v, 16); look.firstRow = !!(n & 0x20); look.lastRow = !!(n & 0x40); look.firstCol = !!(n & 0x80); look.lastCol = !!(n & 0x100); look.noHBand = !!(n & 0x200); look.noVBand = !!(n & 0x400); }
            for (const k of ['firstRow', 'lastRow', 'firstColumn', 'lastColumn', 'noHBand', 'noVBand']) if (at(c, k) != null) look[k.replace('Column', 'Col')] = X.bool(at(c, k));
            o.look = look;
            break;
          }
          case 'tblStyleRowBandSize': o.rowBand = num(c, 'val', 1); break;
          case 'tblStyleColBandSize': o.colBand = num(c, 'val', 1); break;
          case 'tblpPr': o.float = { x: twA(c, 'tblpX'), y: twA(c, 'tblpY'), xAlign: at(c, 'tblpXSpec'), yAlign: at(c, 'tblpYSpec'), hAnchor: at(c, 'horzAnchor'), vAnchor: at(c, 'vertAnchor'), l: twA(c, 'leftFromText'), r: twA(c, 'rightFromText'), t: twA(c, 'topFromText'), b: twA(c, 'bottomFromText') }; break;
          case 'tblCaption': o.caption = at(c, 'val'); break;
          case 'tblDescription': o.desc = at(c, 'val'); break;
          case 'bidiVisual': o.bidi = bool(c); break;
          case 'tblOverlap': break;
          default: break;
        }
      }
      return o;
    }
    function tcPr(el) {
      const o = {};
      if (!el) return o;
      for (const c of el.children) {
        switch (c.localName) {
          case 'tcW': if (at(c, 'type') !== 'auto' && at(c, 'type') !== 'nil') o.w = at(c, 'type') === 'pct' ? undefined : tw(at(c, 'w')); break;
          case 'gridSpan': o.span = num(c, 'val', 1); break;
          case 'hMerge': o.hMerge = at(c, 'val') || 'continue'; break;
          case 'vMerge': o.vMerge = at(c, 'val') === 'restart' ? 'restart' : 'continue'; break;
          case 'tcBorders': o.borders = borders(c, ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl']); break;
          case 'shd': { const s = shd(c); if (s) o.shd = s; break; }
          case 'noWrap': o.noWrap = bool(c); break;
          case 'tcMar': { const m = {}; for (const k of kids(c)) { const n = { left: 'l', start: 'l', right: 'r', end: 'r', top: 't', bottom: 'b' }[k.localName]; if (n) m[n] = tw(at(k, 'w')); } o.mar = m; break; }
          case 'textDirection': o.textDir = at(c, 'val'); break;
          case 'vAlign': o.vAlign = at(c, 'val') === 'both' ? 'center' : at(c, 'val'); break;
          case 'hideMark': o.hideMark = true; break;
          case 'tcFitText': o.fit = bool(c); break;
          default: break;
        }
      }
      return o;
    }
    function trPr(el) {
      const o = {};
      if (!el) return o;
      for (const c of el.children) {
        switch (c.localName) {
          case 'trHeight': o.h = tw(at(c, 'val')); o.hRule = at(c, 'hRule') || 'atLeast'; break;
          case 'cantSplit': o.cantSplit = bool(c); break;
          case 'tblHeader': o.header = bool(c); break;
          case 'gridBefore': o.gridBefore = num(c, 'val', 0); break;
          case 'gridAfter': o.gridAfter = num(c, 'val', 0); break;
          case 'jc': o.jc = at(c, 'val'); break;
          case 'hidden': o.hidden = bool(c); break;
          case 'ins': o.ins = rev(c); break;
          case 'del': o.del = rev(c); break;
          default: break;
        }
      }
      return o;
    }
    function parseTbl(el, ctx) {
      const tp = tblPr(kid(el, 'tblPr'));
      L.preserve.properties(tp, kid(el, 'tblPr'), 'tblPr', { pkg: doc.pkg, part: ctx.part });
      const grid = kids(kid(el, 'tblGrid'), 'gridCol').map((g) => tw(at(g, 'w')));
      const rows = [];
      const rowAlternates = L.preserve.readAlternates(el, rows, ctx, doc, 'row');
      const rowEls = [];
      const rowControls = [], rowModels = new Map(), emptyRows = [];
      const collectRows = (parent) => { for (const c of parent.children) { if (c.localName === 'tr') rowEls.push(c); else if (c.localName === 'sdt') { const sc = kid(c, 'sdtContent'); if (sc) { const n = rowEls.length; collectRows(sc); rowControls.push({ control: L.preserve.control(c, ctx, doc), elements: rowEls.slice(n) }); } } else if (c.localName === 'customXml' || c.localName === 'ins' || c.localName === 'del') collectRows(c); else if (c.namespaceURI === 'urn:vibeoffice:keep' && L.opc.alternate(c) && L.opc.alternate(c) !== L.opc.alternate(el)) emptyRows.push({ el: c, index: rowEls.length }); } };
      collectRows(el);
      for (const tr of rowEls) {
        const rp = trPr(kid(tr, 'trPr'));
        const ex = kid(tr, 'tblPrEx');
        const cells = [];
        const cellAlternates = L.preserve.readAlternates(tr, cells, ctx, doc, 'cell');
        const cellEls = [];
        const cellControls = [], cellModels = new Map(), emptyCells = [];
        const cellBound = new Map(); /* cell-level content controls bound to a data store */
        const collectCells = (parent) => { for (const c of parent.children) { if (c.localName === 'tc') cellEls.push(c); else if (c.localName === 'sdt') { const sc = kid(c, 'sdtContent'); if (sc) { const n0 = cellEls.length; collectCells(sc); const bv = boundText(kid(c, 'sdtPr')); if (bv != null && cellEls.length === n0 + 1) cellBound.set(cellEls[n0], bv); cellControls.push({ control: L.preserve.control(c, ctx, doc), elements: cellEls.slice(n0) }); } } else if (c.localName === 'customXml') collectCells(c); else if (c.namespaceURI === 'urn:vibeoffice:keep' && L.opc.alternate(c) && L.opc.alternate(c) !== L.opc.alternate(tr)) emptyCells.push({ el: c, index: cellEls.length }); } };
        collectCells(tr);
        for (const tc of cellEls) {
          const cp = tcPr(kid(tc, 'tcPr'));
          const blocks = parseBlocks(tc, ctx);
          if (cellBound.has(tc)) { const ps = blocks.filter((b) => b.t === 'p'); if (ps.length === 1) ps[0].runs = boundRuns(cellBound.get(tc), ps[0].runs); }
          if (!blocks.length || blocks[blocks.length - 1].t !== 'p') { const sp = D.para(); if (blocks.length) sp.synth = true; blocks.push(sp); }
          cells.push({ id: D.nid(), tcPr: cp, blocks });
          cellModels.set(tc, cells.at(-1));
          cellAlternates.add(tc, cells.length - 1);
        }
        cellAlternates.finish();
        /* legacy horizontal merges → gridSpan */
        for (let i = cells.length - 1; i > 0; i--) {
          if (cells[i].tcPr.hMerge === 'continue') {
            const prev = cells[i - 1];
            prev.tcPr.span = (prev.tcPr.span || 1) + (cells[i].tcPr.span || 1);
            if (prev.tcPr.w && cells[i].tcPr.w) prev.tcPr.w += cells[i].tcPr.w;
            cells.splice(i, 1);
          }
        }
        for (const c of cells) delete c.tcPr.hMerge;
        for (const c of cellControls) L.preserve.markControl(c.elements.map(e => cellModels.get(e)).filter(c => cells.includes(c)), 'tcPr', c.control);
        if (!cells.length && emptyCells.length) return L.preserve.opaqueBlock(el, ctx, doc, 'Compatibility table');
        if (!cells.length) continue;
        L.preserve.emptyAlternates(emptyCells, cellEls.map(e => cellModels.get(e)).map(c => cells.includes(c) ? c : null), ctx, doc);
        const row = { id: D.nid(), trPr: rp, cells };
        if (ex) { const xp = tblPr(ex); if (xp.borders) row.exBorders = xp.borders; }
        rows.push(row);
        rowModels.set(tr, row);
        rowAlternates.add(tr, rows.length - 1);
      }
      rowAlternates.finish();
      for (const c of rowControls) L.preserve.markControl(c.elements.map(e => rowModels.get(e)).filter(Boolean), 'trPr', c.control);
      if (!rows.length) return emptyRows.length ? L.preserve.opaqueBlock(el, ctx, doc, 'Compatibility table') : null;
      L.preserve.emptyAlternates(emptyRows, rowEls.map(e => rowModels.get(e)), ctx, doc);
      /* grid sanity: compute from cell widths when missing or inconsistent */
      const nCols = Math.max(...rows.map((r) => (r.trPr.gridBefore || 0) + r.cells.reduce((s, c) => s + (c.tcPr.span || 1), 0) + (r.trPr.gridAfter || 0)));
      let g = grid.slice();
      if (g.length < nCols || g.every((x) => !x)) {
        const r0 = rows.find((r) => r.cells.reduce((s, c) => s + (c.tcPr.span || 1), 0) === nCols) || rows[0];
        const tw0 = tp.w && tp.w.type === 'dxa' && tp.w.v ? tp.w.v : 432;
        g = [];
        for (const c of r0.cells) { const sp = c.tcPr.span || 1; const w = c.tcPr.w || tw0 / nCols * sp; for (let k = 0; k < sp; k++) g.push(w / sp); }
        while (g.length < nCols) g.push(tw0 / nCols);
      }
      /* zero-width columns get a minimal width */
      g = g.map((x) => (x > 0 ? L.round(x, 2) : 4));
      const t = D.table(rows, g, tp);
      return t;
    }

    /* ---- sections ---- */
    function parseSect(el, ctx) {
      const s = D.defaultSect();
      s.refs = { hdr: {}, ftr: {} };
      for (const c of el.children) {
        switch (c.localName) {
          case 'type': s.type = at(c, 'val') || 'nextPage'; break;
          case 'pgSz': s.pgW = tw(at(c, 'w')) || 612; s.pgH = tw(at(c, 'h')) || 792; s.orient = at(c, 'orient') || (s.pgW > s.pgH ? 'landscape' : 'portrait'); if (at(c, 'code')) s.paperCode = num(c, 'code', 0); break;
          case 'pgMar': s.mt = tw(at(c, 'top')); s.mb = tw(at(c, 'bottom')); s.ml = tw(at(c, 'left') || at(c, 'start')); s.mr = tw(at(c, 'right') || at(c, 'end')); s.hdr = tw(at(c, 'header')); s.ftr = tw(at(c, 'footer')); s.gutter = tw(at(c, 'gutter')); break;
          case 'cols': {
            const cols = { n: num(c, 'num', 1) || 1, space: at(c, 'space') != null ? tw(at(c, 'space')) : 36, sep: X.bool(at(c, 'sep')), eq: at(c, 'equalWidth') == null ? true : X.bool(at(c, 'equalWidth')), w: [] };
            const cs = kids(c, 'col');
            if (cs.length) { cols.w = cs.map((x) => ({ w: tw(at(x, 'w')), space: tw(at(x, 'space')) })); if (at(c, 'equalWidth') == null) cols.eq = false; cols.n = Math.max(cols.n, cs.length); }
            s.cols = cols;
            break;
          }
          case 'titlePg': s.titlePg = bool(c); break;
          case 'vAlign': s.vAlign = at(c, 'val') || 'top'; break;
          case 'pgNumType': s.pgNum = { fmt: at(c, 'fmt') || 'decimal', start: at(c, 'start') != null ? num(c, 'start', 1) : null, chapStyle: at(c, 'chapStyle') || undefined, chapSep: at(c, 'chapSep') || undefined }; break;
          case 'headerReference': case 'footerReference': {
            const k = c.localName === 'headerReference' ? 'hdr' : 'ftr';
            const r = rid(c);
            const sid = ctx.hfMap && ctx.hfMap[r];
            if (sid) s.refs[k][at(c, 'type') || 'default'] = sid;
            break;
          }
          case 'pgBorders': {
            const b = borders(c, ['top', 'left', 'bottom', 'right']) || {};
            b.offsetFrom = at(c, 'offsetFrom') || 'text';
            b.display = at(c, 'display') || 'allPages';
            b.zOrder = at(c, 'zOrder') || 'front';
            s.borders = b;
            break;
          }
          case 'lnNumType': s.lnNum = { countBy: num(c, 'countBy', 1), start: num(c, 'start', 0), distance: twA(c, 'distance'), restart: at(c, 'restart') || 'newPage' }; break;
          case 'textDirection': s.textDir = at(c, 'val'); break;
          case 'bidi': s.bidi = bool(c); break;
          case 'footnotePr': s.fnPr = notePr(c); break;
          case 'endnotePr': s.enPr = notePr(c); break;
          case 'paperSrc': s.paperSrc = { first: num(c, 'first', 0), other: num(c, 'other', 0) }; break;
          case 'docGrid': s.docGrid = { type: at(c, 'type'), linePitch: num(c, 'linePitch', 0), charSpace: num(c, 'charSpace', 0) }; break;
          case 'formProt': s.formProt = bool(c); break;
          default: break;
        }
      }
      if (s.pgW <= 0) s.pgW = 612;
      if (s.pgH <= 0) s.pgH = 792;
      return s;
    }

    /* ---- drawings (DrawingML) ---- */
    function floatOf(an) {
      const f = { posH: { rel: 'column' }, posV: { rel: 'paragraph' } };
      const ph = kid(an, 'positionH'), pv = kid(an, 'positionV');
      const pos = (pe, o) => {
        if (!pe) return;
        o.rel = at(pe, 'relativeFrom') || o.rel;
        const al = kid(pe, 'align'), off = kid(pe, 'posOffset'), pct = kid(pe, 'pctPosHOffset') || kid(pe, 'pctPosVOffset');
        if (al) o.align = al.textContent.trim();
        else if (off) o.off = L.round(emu(off.textContent), 2);
        else if (pct) o.pct = (+pct.textContent || 0) / 1000;
      };
      pos(ph, f.posH); pos(pv, f.posV);
      if (X.bool(at(an, 'simplePos'))) { const sp = kid(an, 'simplePos'); f.posH = { rel: 'page', off: emu(num(sp, 'x', 0)) }; f.posV = { rel: 'page', off: emu(num(sp, 'y', 0)) }; }
      f.dist = { t: emu(num(an, 'distT', 0)), b: emu(num(an, 'distB', 0)), l: emu(num(an, 'distL', 0)), r: emu(num(an, 'distR', 0)) };
      f.z = num(an, 'relativeHeight', 0);
      f.behind = X.bool(at(an, 'behindDoc'));
      f.allowOverlap = at(an, 'allowOverlap') == null ? true : X.bool(at(an, 'allowOverlap'));
      f.layoutInCell = at(an, 'layoutInCell') == null ? true : X.bool(at(an, 'layoutInCell'));
      f.locked = X.bool(at(an, 'locked'));
      if (kid(an, 'wrapNone')) f.wrap = f.behind ? 'behind' : 'none';
      else if (kid(an, 'wrapSquare')) { f.wrap = 'square'; f.side = at(kid(an, 'wrapSquare'), 'wrapText') || 'bothSides'; }
      else if (kid(an, 'wrapTight')) { f.wrap = 'tight'; f.side = at(kid(an, 'wrapTight'), 'wrapText') || 'bothSides'; }
      else if (kid(an, 'wrapThrough')) { f.wrap = 'through'; f.side = at(kid(an, 'wrapThrough'), 'wrapText') || 'bothSides'; }
      else if (kid(an, 'wrapTopAndBottom')) f.wrap = 'topBottom';
      else f.wrap = f.behind ? 'behind' : 'none';
      return f;
    }
    function parseDrawing(dEl, ctx, base) {
      const inl = kid(dEl, 'inline') || kid(dEl, 'anchor');
      if (!inl) return null;
      const ext = kid(inl, 'extent');
      const w = emu(num(ext, 'cx', 0)), h = emu(num(ext, 'cy', 0));
      const docPr = kid(inl, 'docPr');
      const gd = X.path(inl, 'graphic', 'graphicData');
      if (!gd) return null;
      const content = gd.firstElementChild;
      let it = null;
      const mctx = mediaCtx(ctx);
      if (content) {
        switch (content.localName) {
          case 'pic': it = picItem(content, mctx, w, h); break;
          case 'wsp': it = shapeItem(content, mctx, w, h, ctx); break;
          case 'wgp': case 'wpc': case 'lockedCanvas': it = groupItem(content, mctx, w, h, ctx); break;
          case 'chart': {
            const r = rid(content);
            const rel = r && ctx.rels[r];
            const c = rel && chartCache.get(rel.target);
            it = { t: 'chart', w, h, chart: c && c.model, src: c ? c.src : null };
            if (!c || !c.model) { it.t = 'img'; it.media = null; it.alt = 'Chart'; }
            break;
          }
          case 'relIds': it = smartArtItem(content, mctx, w, h, ctx); break;
          default: it = { t: 'img', media: null, w, h, alt: content.localName };
        }
      }
      if (!it) return null;
      it.rPr = base;
      it.w = it.w || w; it.h = it.h || h;
      if (docPr) { if (at(docPr, 'descr')) it.alt = at(docPr, 'descr'); if (at(docPr, 'name')) it.name = at(docPr, 'name'); if (at(docPr, 'title')) it.title = at(docPr, 'title'); if (X.bool(at(docPr, 'hidden'))) it.hidden = true; const hl = kid(docPr, 'hlinkClick'); if (hl) { const r = rid(hl); if (r && ctx.rels[r]) it.link = { url: ctx.rels[r].raw }; } }
      if (inl.localName === 'anchor') it.float = floatOf(inl);
      return it;
    }
    function mediaCtx(ctx) {
      return { theme: doc.theme, media: (r) => { const rel = r && ctx.rels[r]; if (!rel) return null; if (rel.external) return null; return mediaCache.get(rel.target) || null; } };
    }
    function picItem(pic, mctx, w, h) {
      const bf = kid(pic, 'blipFill');
      const blip = kid(bf, 'blip');
      let media = null;
      if (blip) {
        const svg = descAll(blip, 'svgBlip')[0];
        const r = (svg && rid(svg, 'embed')) || rid(blip, 'embed') || rid(blip, 'link');
        media = mctx.media(r) || mctx.media(rid(blip, 'embed'));
      }
      const it = { t: 'img', media, w, h };
      const sr = kid(bf, 'srcRect');
      if (sr) { const c = { l: num(sr, 'l', 0) / 100000, t: num(sr, 't', 0) / 100000, r: num(sr, 'r', 0) / 100000, b: num(sr, 'b', 0) / 100000 }; if (c.l || c.t || c.r || c.b) it.crop = c; }
      const sp = kid(pic, 'spPr');
      const xf = X.xfrm(kid(sp, 'xfrm'));
      if (xf) { if (xf.rot) it.rot = xf.rot; if (xf.flipH) it.flipH = true; if (xf.flipV) it.flipV = true; }
      const ln = sp && kid(sp, 'ln');
      if (ln) { const l = X.line(ln, mctx); if (l && l.t !== 'none') it.border = { val: 'single', sz: l.w, color: (l.c || '#000000').replace('#', '') }; }
      if (blip) {
        if (desc(blip, 'grayscl')) it.gray = true;
        const lum = desc(blip, 'lum');
        if (lum) { if (at(lum, 'bright')) it.bright = num(lum, 'bright', 0) / 100000; if (at(lum, 'contrast')) it.contrast = num(lum, 'contrast', 0) / 100000; }
        const bi = desc(blip, 'biLevel');
        if (bi) it.bw = true;
        const amf = desc(blip, 'alphaModFix');
        if (amf) it.alpha = num(amf, 'amt', 100000) / 100000;
      }
      const geom = X.geometry(sp);
      if (geom.geom !== 'rect') it.geom = geom;
      return it;
    }
    function textBoxBlocks(txbx, ctx) {
      const content = kid(txbx, 'txbxContent');
      if (!content) return null;
      const story = { kind: 'tb', blocks: [] };
      const sctx = Object.assign({}, ctx, { pendingBm: [], story });
      story.blocks = parseBlocks(content, sctx);
      if (!story.blocks.length) story.blocks.push(D.para());
      return story;
    }
    function shapeItem(wsp, mctx, w, h, ctx) {
      const sp = kid(wsp, 'spPr');
      const g = X.geometry(sp);
      const refs = X.styleRefs(kid(wsp, 'style'), mctx);
      const xf = X.xfrm(kid(sp, 'xfrm'));
      const fill = X.fillIn(sp, mctx) || refs.fill || { t: 'none' };
      const ln = kid(sp, 'ln');
      const line = ln ? X.line(ln, mctx, refs.line) : refs.line || { t: 'none' };
      const it = { t: 'shape', w, h, geom: g.geom, adj: g.adj, path: g.path, fill, line };
      if (xf) { if (xf.rot) it.rot = xf.rot; if (xf.flipH) it.flipH = true; if (xf.flipV) it.flipV = true; }
      const sh = X.shadow(kid(sp, 'effectLst'), mctx);
      if (sh) it.shadow = sh;
      const tx = kid(wsp, 'txbx');
      if (tx) { it.tb = textBoxBlocks(tx, ctx); if (at(tx, 'id') != null) it.tbId = at(tx, 'id'); }
      const ltx = kid(wsp, 'linkedTxbx');
      if (ltx && !tx) it.tbLink = { id: at(ltx, 'id'), seq: +at(ltx, 'seq') || 1 };
      const bp = kid(wsp, 'bodyPr');
      if (bp) {
        it.ins = { l: at(bp, 'lIns') != null ? emu(at(bp, 'lIns')) : 7.2, t: at(bp, 'tIns') != null ? emu(at(bp, 'tIns')) : 3.6, r: at(bp, 'rIns') != null ? emu(at(bp, 'rIns')) : 7.2, b: at(bp, 'bIns') != null ? emu(at(bp, 'bIns')) : 3.6 };
        if (at(bp, 'anchor')) it.anchor = at(bp, 'anchor');
        if (at(bp, 'vert') && at(bp, 'vert') !== 'horz') it.vert = at(bp, 'vert');
        if (kid(bp, 'spAutoFit')) it.autofit = true;
        const warp = kid(bp, 'prstTxWarp');
        if (warp && at(warp, 'prst') && at(warp, 'prst') !== 'textNoShape') it.warp = at(warp, 'prst');
        if (at(bp, 'wrap') === 'none') it.noWrap = true;
      }
      if (refs.fontColor) it.fontColor = refs.fontColor;
      if (X.at(kid(wsp, 'cNvPr'), 'name')) it.name = X.at(kid(wsp, 'cNvPr'), 'name');
      /* WordArt: the text box holds the letters, whose own fill/outline become the shape's */
      if (bp && at(bp, 'fromWordArt') === '1' && it.tb && tx) {
        const paras = it.tb.blocks.filter((b) => b.t === 'p');
        const text = paras.map((p) => D.plainText(p)).join('\n');
        const firstRun = paras.length && paras[0].runs.find((r) => r.t === 'text');
        const rp = firstRun ? firstRun.rPr || {} : {};
        const rEl = descAll(tx, 'r')[0];
        const rPrEl = rEl && kid(rEl, 'rPr');
        const tf = rPrEl && kid(rPrEl, 'textFill');
        const to = rPrEl && kid(rPrEl, 'textOutline');
        const tfill = tf ? X.fillIn(tf, mctx) : null;
        it.wordart = { text, font: rp.font || 'Arial Black', sz: rp.sz || 36, b: rp.b || undefined, i: rp.i || undefined };
        it.fill = tfill || (rp.color && rp.color !== 'auto' ? { t: 'solid', c: '#' + rp.color, a: 1 } : { t: 'solid', c: '#000000', a: 1 });
        if (to) { const ol = X.line(to, mctx); if (ol && ol.t !== 'none') it.line = ol; else it.line = { t: 'none' }; } else it.line = { t: 'none' };
        if (!it.warp) it.warp = 'textPlain';
        delete it.tb;
        delete it.ins;
      }
      return it;
    }
    function groupItem(g, mctx, w, h, ctx) {
      const gsp = kid(g, 'grpSpPr');
      const xf = X.xfrm(kid(gsp, 'xfrm')) || { x: 0, y: 0, w, h, chx: 0, chy: 0, chw: w, chh: h };
      const sx = xf.chw ? w / xf.chw : 1, sy = xf.chh ? h / xf.chh : 1;
      const it = { t: 'group', w, h, kids: [] };
      if (g.localName === 'wpc') { const bg = kid(g, 'bg'); const f = bg && X.fillIn(bg, mctx); if (f) it.fill = f; }
      const walk = (parent, map) => {
        for (const c of parent.children) {
          const nm = c.localName;
          if (nm === 'wsp' || nm === 'sp' || nm === 'pic' || nm === 'grpSp' || nm === 'wgp' || nm === 'cxnSp') {
            const spPr = kid(c, 'spPr') || kid(c, 'grpSpPr');
            const cx = X.xfrm(kid(spPr, 'xfrm'));
            if (!cx) continue;
            const r = map(cx);
            if (nm === 'grpSp' || nm === 'wgp') {
              const csx = cx.chw ? cx.w / cx.chw : 1, csy = cx.chh ? cx.h / cx.chh : 1;
              walk(c, (b) => map({ x: cx.x + (b.x - (cx.chx || 0)) * csx, y: cx.y + (b.y - (cx.chy || 0)) * csy, w: b.w * csx, h: b.h * csy, rot: b.rot, flipH: b.flipH, flipV: b.flipV }));
              continue;
            }
            let k;
            if (nm === 'pic') k = picItem(c, mctx, r.w, r.h);
            else k = shapeItem(c, mctx, r.w, r.h, ctx);
            if (!k) continue;
            /* plain DrawingML shapes (locked canvases pasted from PowerPoint) carry DrawingML text */
            if (nm === 'sp' && !k.tb) {
              const txb = kid(c, 'txBody') || X.path(c, 'txSp', 'txBody');
              if (txb && txb.textContent.trim()) {
                const refs = X.styleRefs(kid(c, 'style'), mctx);
                const dr = X.path(txb, 'lstStyle', 'lvl1pPr', 'defRPr');
                k.tb = { kind: 'tb', blocks: dmlText(txb, refs.fontColor, { scale: g.localName === 'lockedCanvas' ? Math.min(sx, sy) : 1, defSz: dr && at(dr, 'sz') ? num(dr, 'sz', 1800) / 100 : 18 }) };
                const bp = kid(txb, 'bodyPr');
                k.anchor = (bp && at(bp, 'anchor')) || 't';
                k.ins = { l: bp && at(bp, 'lIns') != null ? emu(at(bp, 'lIns')) : 7.2, t: bp && at(bp, 'tIns') != null ? emu(at(bp, 'tIns')) : 3.6, r: bp && at(bp, 'rIns') != null ? emu(at(bp, 'rIns')) : 7.2, b: bp && at(bp, 'bIns') != null ? emu(at(bp, 'bIns')) : 3.6 };
                if (bp && at(bp, 'wrap') === 'none') k.noWrap = true;
                k.overflowText = true; /* PowerPoint-style text is not clipped by its shape */
              }
            }
            k.x = r.x; k.y = r.y; k.w = r.w; k.h = r.h;
            if (cx.rot) k.rot = cx.rot;
            it.kids.push(k);
          }
        }
      };
      walk(g, (b) => ({ x: (b.x - (xf.chx || 0)) * sx, y: (b.y - (xf.chy || 0)) * sy, w: b.w * sx, h: b.h * sy }));
      it.tb = null;
      return it;
    }
    function smartArtItem(relIds, mctx, w, h, ctx) {
      /* SmartArt: use the pre-rendered drawing part (shapes with DrawingML text) */
      let dpath = null;
      const dm = ctx.rels[rid(relIds, 'dm')];
      for (const r of Object.values(ctx.rels)) if (r.type === 'diagramDrawing') { dpath = r.target; break; }
      void dm;
      const dx = dpath && xmlCache.get(dpath);
      const tree = dx && desc(dx, 'spTree');
      if (!tree) return smartArtFromData(relIds, w, h, ctx) || { t: 'img', media: null, w, h, alt: 'Diagram' };
      const it = { t: 'group', w, h, kids: [], diagram: true };
      for (const sp of kids(tree, 'sp')) {
        const spPr = kid(sp, 'spPr');
        const xf = X.xfrm(kid(spPr, 'xfrm'));
        if (!xf) continue;
        const g = X.geometry(spPr);
        const refs = X.styleRefs(kid(sp, 'style'), mctx);
        const k = { t: 'shape', x: xf.x, y: xf.y, w: xf.w, h: xf.h, geom: g.geom, adj: g.adj, path: g.path, fill: X.fillIn(spPr, mctx) || refs.fill || { t: 'none' }, line: kid(spPr, 'ln') ? X.line(kid(spPr, 'ln'), mctx, refs.line) : refs.line || { t: 'none' } };
        if (xf.rot) k.rot = xf.rot;
        const txb = kid(sp, 'txBody');
        if (txb) {
          k.tb = { kind: 'tb', blocks: dmlText(txb, refs.fontColor) };
          const bp = kid(txb, 'bodyPr');
          k.anchor = (bp && at(bp, 'anchor')) || 'ctr';
          const txXf = X.xfrm(kid(sp, 'txXfrm'));
          k.ins = { l: 3.6, t: 3.6, r: 3.6, b: 3.6 };
          if (txXf) k.txRect = { x: txXf.x - xf.x, y: txXf.y - xf.y, w: txXf.w, h: txXf.h };
        }
        it.kids.push(k);
      }
      return it;
    }
    /**
     * SmartArt saved without its drawing (Word 2007 and some generators): lay the data model out
     * ourselves — a row of boxes joined by arrows for processes, a ring for cycles, a tree for
     * hierarchies, otherwise a column of boxes with each node's children as bullet text.
     */
    function smartArtFromData(relIds, w, h, ctx) {
      const dmRel = ctx.rels[rid(relIds, 'dm')], loRel = ctx.rels[rid(relIds, 'lo')];
      const dm = dmRel && xmlCache.get(dmRel.target);
      if (!dm) return null;
      const lo = loRel && xmlCache.get(loRel.target);
      const kind = String((lo && (at(lo, 'uniqueId') || (lo.documentElement && lo.documentElement.getAttribute('uniqueId')))) || '').toLowerCase();
      const pts = new Map();
      for (const pt of descAll(dm, 'pt')) {
        const type = at(pt, 'type') || 'node';
        if (type !== 'node' && type !== 'doc') continue;
        const tEl = kid(pt, 't');
        const text = tEl ? kids(tEl, 'p').map((ap) => descAll(ap, 't').map((t) => t.textContent).join('')).join(' ').trim() : '';
        pts.set(at(pt, 'modelId'), { id: at(pt, 'modelId'), type, text, kids: [] });
      }
      const parent = new Map();
      for (const c of descAll(dm, 'cxn')) {
        const t = at(c, 'type') || 'parOf';
        if (t !== 'parOf') continue;
        const a = pts.get(at(c, 'srcId')), b = pts.get(at(c, 'destId'));
        if (a && b) { a.kids.push({ n: b, ord: +at(c, 'srcOrd') || 0 }); parent.set(b.id, a); }
      }
      for (const p of pts.values()) p.kids = p.kids.sort((x, y) => x.ord - y.ord).map((k) => k.n);
      const root = [...pts.values()].find((p) => p.type === 'doc') || null;
      const tops = root ? root.kids : [...pts.values()].filter((p) => !parent.has(p.id));
      if (!tops.length) return null;
      const accent = (doc.theme.colors && doc.theme.colors.accent1) || '4F81BD';
      const fill = { t: 'solid', c: '#' + String(accent).replace('#', ''), a: 1 };
      const line = { t: 'solid', c: '#FFFFFF', w: 1.5 };
      const it = { t: 'group', w, h, kids: [], diagram: true };
      const box = (x, y, bw, bh, text, sub, geom) => {
        const lines = [text].concat(sub || []).filter((t) => t != null && t !== '');
        const fs = Math.max(7, Math.min(18, bh / (lines.length * 1.6 + 0.6), (bw / Math.max(4, ...lines.map((l) => l.length))) * 1.7));
        const paras = lines.map((t, i) => D.para([D.text(i && sub ? '• ' + t : t, { sz: Math.round(fs * (i && sub ? 0.85 : 1) * 2) / 2, color: 'FFFFFF' })], { jc: i && sub ? 'left' : 'center', sp: { a: 0, b: 0, line: 0.95, rule: 'auto' } }));
        it.kids.push({ t: 'shape', x, y, w: bw, h: bh, geom: geom || 'roundRect', adj: {}, fill, line, tb: { kind: 'tb', blocks: paras.length ? paras : [D.para()] }, anchor: 'ctr', ins: { l: 4, t: 3, r: 4, b: 3 } });
      };
      const arrow = (x, y, aw, ah, rot) => it.kids.push({ t: 'shape', x, y, w: aw, h: ah, geom: 'rightArrow', adj: {}, fill: { t: 'solid', c: '#' + String(accent).replace('#', ''), a: 0.45 }, line: { t: 'none' }, rot: rot || 0 });
      const n = tops.length;
      if (/process|chevron|arrow/.test(kind)) {
        const gap = Math.min(w / (n * 4), 30), bw = (w - gap * (n - 1)) / n, bh = Math.min(h, bw * 0.75), y = (h - bh) / 2;
        tops.forEach((p, i) => { box(i * (bw + gap), y, bw, bh, p.text, p.kids.map((k) => k.text)); if (i < n - 1) arrow(i * (bw + gap) + bw + gap * 0.15, h / 2 - gap * 0.35, gap * 0.7, gap * 0.7); });
      } else if (/cycle/.test(kind)) {
        const r = Math.min(w, h) / 2, bw = Math.min(r * 0.8, (2 * Math.PI * r * 0.55) / n), bh = bw * 0.6;
        tops.forEach((p, i) => { const a = -Math.PI / 2 + (i * 2 * Math.PI) / n; box(w / 2 + (r - bw / 2) * Math.cos(a) - bw / 2, h / 2 + (r - bh / 2) * Math.sin(a) - bh / 2, bw, bh, p.text, null, 'ellipse'); });
      } else if (/hierarchy|orgchart/.test(kind)) {
        const levels = [];
        const walk = (p, d) => { (levels[d] = levels[d] || []).push(p); for (const k of p.kids) walk(k, d + 1); };
        for (const t of tops) walk(t, 0);
        const lh = h / levels.length;
        levels.forEach((row, d) => { const bw = Math.min(w / row.length * 0.85, w * 0.3), gap = (w - bw * row.length) / (row.length + 1); row.forEach((p, i) => box(gap + i * (bw + gap), d * lh + lh * 0.12, bw, lh * 0.7, p.text)); });
      } else {
        const gap = Math.min(h / (n * 6), 8), bh = (h - gap * (n - 1)) / n;
        tops.forEach((p, i) => box(0, i * (bh + gap), w, bh, p.text, p.kids.map((k) => k.text)));
      }
      return it;
    }
    /** DrawingML text body → paragraphs (for SmartArt and charts) */
    function dmlText(txb, fontColor, opts) {
      const out = [];
      const scale = (opts && opts.scale) || 1;
      const defSz = opts && opts.defSz;
      for (const ap of kids(txb, 'p')) {
        const ppr = kid(ap, 'pPr');
        const jc = { l: 'left', ctr: 'center', r: 'right', just: 'both' }[at(ppr, 'algn')] || 'center';
        const p = D.para([], { jc, sp: { a: 0, b: 0, line: 0.9, rule: 'auto' } });
        for (const r of kids(ap)) {
          if (r.localName !== 'r' && r.localName !== 'br') continue;
          if (r.localName === 'br') { p.runs.push(D.item('br', { type: 'line' })); continue; }
          const rp = kid(r, 'rPr');
          const o = {};
          if (rp) {
            if (at(rp, 'sz')) o.sz = num(rp, 'sz', 1800) / 100;
            if (at(rp, 'b') != null) o.b = X.bool(at(rp, 'b'));
            if (at(rp, 'i') != null) o.i = X.bool(at(rp, 'i'));
            const c = X.colorOf(kid(rp, 'solidFill'), { theme: doc.theme });
            if (c) o.color = c.c.replace('#', '');
            const lat = kid(rp, 'latin');
            if (lat && at(lat, 'typeface') && !/^\+/.test(at(lat, 'typeface'))) o.font = at(lat, 'typeface');
          }
          if (!o.color && fontColor) o.color = fontColor.replace('#', '');
          if (o.sz == null && defSz) o.sz = defSz;
          if (scale !== 1 && o.sz) o.sz = Math.max(1, Math.round(o.sz * scale * 2) / 2);
          const t = kid(r, 't');
          if (t && t.textContent) p.runs.push(D.text(t.textContent, o));
        }
        out.push(p);
      }
      return out.length ? out : [D.para()];
    }

    /* ---- VML (w:pict, w:object) ---- */
    function cssLen(v) {
      if (v == null) return 0;
      const m = /(-?\d*\.?\d+)\s*(pt|in|cm|mm|px|pc|em)?/.exec(String(v));
      if (!m) return 0;
      const x = parseFloat(m[1]);
      switch (m[2]) { case 'in': return x * 72; case 'cm': return x * 28.3465; case 'mm': return x * 2.83465; case 'px': return x * 0.75; case 'pc': return x * 12; case 'em': return x * 12; case 'pt': return x; default: return x * 0.75; }
    }
    function styleMap(s) { const o = {}; for (const kv of String(s || '').split(';')) { const i = kv.indexOf(':'); if (i > 0) o[kv.slice(0, i).trim().toLowerCase()] = kv.slice(i + 1).trim(); } return o; }
    function vColor(v) {
      if (!v) return null;
      const m = /#?([0-9a-f]{6}|[0-9a-f]{3})\b/i.exec(v);
      if (m) { let h = m[1]; if (h.length === 3) h = h.split('').map((c) => c + c).join(''); return '#' + h.toUpperCase(); }
      const n = String(v).split(/\s/)[0];
      return L.color.PRESET[n] || null;
    }
    /* VML shape types seen so far (a <v:shapetype> is usually defined once and referenced afterwards) */
    const vmlTypes = new Map();
    const VML_ARROW = { block: 'triangle', classic: 'stealth', open: 'arrow', oval: 'oval', diamond: 'diamond' };
    const VML_DASH = { solid: 'solid', dot: 'sysDot', dash: 'dash', dashdot: 'dashDot', longdash: 'lgDash', longdashdot: 'lgDashDot', longdashdotdot: 'lgDashDotDot', shortdash: 'sysDash', shortdot: 'sysDot', shortdashdot: 'sysDashDot', shortdashdotdot: 'sysDashDotDot', '1 1': 'sysDot', '3 1': 'sysDash' };
    const vBool = (v, d) => (v == null ? d : !/^(f|false|0|off)$/i.test(v));
    function vmlType(s) {
      const t = (at(s, 'type') || '').replace(/^#/, '');
      return t ? vmlTypes.get(t) || null : null;
    }
    /** fill of a VML element ({t:'none'} when unfilled) */
    function vmlFill(s, def) {
      const ty = vmlType(s);
      const filled = vBool(at(s, 'filled'), ty ? vBool(at(ty, 'filled'), def) : def);
      if (!filled) return { t: 'none' };
      const fe = kid(s, 'fill');
      const c1 = vColor(at(s, 'fillcolor')) || '#FFFFFF';
      let a = 1;
      if (fe && at(fe, 'opacity') != null) { const o = at(fe, 'opacity'); a = /f$/.test(o) ? parseFloat(o) / 65536 : parseFloat(o); if (!(a >= 0)) a = 1; }
      if (fe && /gradient/.test(at(fe, 'type') || '')) {
        const c2 = vColor(at(fe, 'color2')) || '#FFFFFF';
        const ang = parseFloat(at(fe, 'angle') || '0') || 0;
        return { t: 'grad', ang: (90 - ang + 360) % 360, stops: [{ p: 0, c: c1, a }, { p: 1, c: c2, a }] };
      }
      if (fe && at(fe, 'on') === 'f') return { t: 'none' };
      return { t: 'solid', c: c1, a };
    }
    function vmlLine(s, def) {
      const ty = vmlType(s);
      const stroked = vBool(at(s, 'stroked'), ty ? vBool(at(ty, 'stroked'), def) : def);
      const se = kid(s, 'stroke');
      if (!stroked || (se && at(se, 'on') === 'f')) return { t: 'none' };
      const ln = { c: vColor(at(s, 'strokecolor')) || (se && vColor(at(se, 'color'))) || '#000000', w: at(s, 'strokeweight') ? cssLen(at(s, 'strokeweight')) : 0.75, dash: 'solid' };
      if (se) {
        const ds = (at(se, 'dashstyle') || '').toLowerCase();
        if (ds && VML_DASH[ds]) ln.dash = VML_DASH[ds];
        const ea = at(se, 'endarrow'), sa = at(se, 'startarrow');
        if (ea && VML_ARROW[ea]) ln.tail = { type: VML_ARROW[ea], w: 'med', len: 'med' };
        if (sa && VML_ARROW[sa]) ln.head = { type: VML_ARROW[sa], w: 'med', len: 'med' };
        if (at(se, 'linestyle') && /thin|thick/i.test(at(se, 'linestyle'))) ln.cmpd = 'dbl';
      }
      return ln;
    }
    /** one VML element (not a group) sized w×h points */
    function vmlShape(s, ctx, w, h) {
      const st = styleMap(at(s, 'style'));
      const ty = vmlType(s);
      const spt = at(s, 'spt') || (ty && at(ty, 'spt')) || '';
      const im = desc(s, 'imagedata');
      const tb = desc(s, 'textbox');
      let it;
      if (im && !tb) {
        const r = rid(im, 'id') || rid(im, 'pict');
        const rel = r && ctx.rels[r];
        it = { t: 'img', media: rel && !rel.external ? mediaCache.get(rel.target) || null : null, w: w || 72, h: h || 72 };
        const ct = at(im, 'croptop'), cb = at(im, 'cropbottom'), cl = at(im, 'cropleft'), cr = at(im, 'cropright');
        const frac = (v) => { if (!v) return 0; if (/f$/.test(v)) return parseFloat(v) / 65536; return parseFloat(v) || 0; };
        if (ct || cb || cl || cr) it.crop = { t: frac(ct), b: frac(cb), l: frac(cl), r: frac(cr) };
        if (X.bool(at(im, 'grayscale'))) it.gray = true;
        if (X.bool(at(im, 'bilevel'))) it.bw = true;
        if (at(im, 'blacklevel')) it.bright = frac(at(im, 'blacklevel'));
        if (at(im, 'gain')) it.contrast = frac(at(im, 'gain')) - 1;
        if (at(s, 'alt')) it.alt = at(s, 'alt');
      } else {
        let geom = { rect: 'rect', roundrect: 'roundRect', oval: 'ellipse', line: 'line', arc: 'arc' }[s.localName] || 'rect';
        if (s.localName === 'shape') geom = { 1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'triangle', 6: 'rtTriangle', 9: 'hexagon', 10: 'octagon', 13: 'rightArrow', 32: 'straightConnector1', 33: 'bentConnector2', 34: 'bentConnector3', 35: 'bentConnector4', 37: 'curvedConnector2', 38: 'curvedConnector3', 20: 'line', 66: 'leftArrow', 67: 'downArrow', 68: 'upArrow', 183: 'sun', 184: 'moon', 96: 'smileyFace', 74: 'heart', 12: 'star5' }[spt] || 'rect';
        const isText = spt === '202' || !!tb;
        it = { t: 'shape', w: w || 72, h: h || 36, geom, fill: vmlFill(s, geom !== 'line'), line: vmlLine(s, true) };
        if (geom === 'line' || /Connector/.test(geom)) {
          it.fill = { t: 'none' };
          const adj = at(s, 'adj');
          const a1 = parseFloat(((adj || '').split(',')[0]) || (ty && (at(ty, 'adj') || '').split(',')[0]) || '');
          if (isFinite(a1) && /bentConnector[34]/.test(geom)) it.adj = { adj1: Math.round((a1 / 21600) * 100000) };
        }
        if (tb) {
          it.tb = textBoxBlocks(tb, ctx);
          const ins = at(tb, 'inset');
          if (ins) { const v = ins.split(',').map((x) => (x && x.trim() ? cssLen(x) : null)); it.ins = { l: v[0] != null ? v[0] : 7.2, t: v[1] != null ? v[1] : 3.6, r: v[2] != null ? v[2] : 7.2, b: v[3] != null ? v[3] : 3.6 }; }
          else it.ins = { l: 7.2, t: 3.6, r: 7.2, b: 3.6 };
          if (/mso-fit-shape-to-text:\s*t/.test(at(tb, 'style') || '')) it.autofit = true;
          const va = st['v-text-anchor'];
          if (va) it.anchor = /middle/.test(va) ? 'ctr' : /bottom/.test(va) ? 'b' : undefined;
          if (/layout-flow:\s*vertical/.test(at(tb, 'style') || '')) it.vert = 'vert';
        }
        const tp = desc(s, 'textpath');
        if (tp && at(tp, 'string')) { const ts = styleMap(at(tp, 'style')); it.wordart = { text: at(tp, 'string'), font: (ts['font-family'] || 'Arial').replace(/["']/g, ''), b: /bold/.test(ts['font-weight'] || '') || undefined, i: /italic/.test(ts['font-style'] || '') || undefined }; it.warp = it.warp || 'textPlain'; }
        const sh = kid(s, 'shadow');
        if (sh && vBool(at(sh, 'on'), false)) { const off = (at(sh, 'offset') || '2pt,2pt').split(','); it.shadow = { c: vColor(at(sh, 'color')) || '#808080', a: 0.6, dx: cssLen(off[0] || '2pt'), dy: cssLen(off[1] || '2pt'), blur: 0 }; }
      }
      if (st.rotation) it.rot = parseFloat(st.rotation) || 0;
      if (st.flip) { if (/x/.test(st.flip)) it.flipH = true; if (/y/.test(st.flip)) it.flipV = true; }
      if (st.visibility === 'hidden') it.hidden = true;
      return it;
    }
    /** a VML group flattened into a Quire group: children mapped from the group's coordinate space into points */
    function vmlGroup(g, ctx, w, h) {
      const it = { t: 'group', w, h, kids: [] };
      const walk = (grp, map) => {
        const co = (at(grp, 'coordorigin') || '0,0').split(',').map((v) => parseFloat(v) || 0);
        const cs = (at(grp, 'coordsize') || '1000,1000').split(',').map((v) => parseFloat(v) || 1000);
        const gst = styleMap(at(grp, 'style'));
        const gw = cssLen(gst.width) || cs[0], gh = cssLen(gst.height) || cs[1];
        void gw; void gh;
        const kx = 1 / (cs[0] || 1), ky = 1 / (cs[1] || 1);
        /* local (child coordinate) → unit box of this group → parent via map */
        const local = (x, y) => map((x - co[0]) * kx, (y - co[1]) * ky);
        for (const c of grp.children) {
          const nm = c.localName;
          if (nm === 'shapetype') { if (at(c, 'id')) vmlTypes.set(at(c, 'id'), c); continue; }
          if (!/^(shape|rect|roundrect|oval|line|group|image|polyline|arc)$/.test(nm)) continue;
          const cst = styleMap(at(c, 'style'));
          if (nm === 'line' || nm === 'polyline') {
            let pts;
            if (nm === 'line') pts = [(at(c, 'from') || '0,0'), (at(c, 'to') || '0,0')].map((v) => v.split(',').map((n) => parseFloat(n) || 0));
            else { const raw = (at(c, 'points') || '').split(/[\s,]+/).filter(Boolean).map((n) => parseFloat(n) || 0); pts = []; for (let i = 0; i + 1 < raw.length; i += 2) pts.push([raw[i], raw[i + 1]]); }
            for (let i = 0; i + 1 < pts.length; i++) {
              const [x1, y1] = local(pts[i][0], pts[i][1]), [x2, y2] = local(pts[i + 1][0], pts[i + 1][1]);
              const k = vmlShape(c, ctx, Math.abs(x2 - x1) || 0.5, Math.abs(y2 - y1) || 0.5);
              k.geom = 'line'; k.fill = { t: 'none' };
              k.x = Math.min(x1, x2); k.y = Math.min(y1, y2);
              if (x2 < x1) k.flipH = true;
              if (y2 < y1) k.flipV = true;
              if (i > 0 && k.line) delete k.line.head;
              if (i + 2 < pts.length && k.line) delete k.line.tail;
              it.kids.push(k);
            }
            continue;
          }
          const left = parseFloat(cst.left || cst['margin-left'] || '0') || 0, top = parseFloat(cst.top || cst['margin-top'] || '0') || 0;
          const cw = parseFloat(cst.width || '0') || 0, ch = parseFloat(cst.height || '0') || 0;
          const [x1, y1] = local(left, top), [x2, y2] = local(left + cw, top + ch);
          if (nm === 'group') {
            const sub = (at(c, 'coordsize') || '1000,1000').split(',').map((v) => parseFloat(v) || 1000);
            void sub;
            walk(c, (u, v) => [x1 + u * (x2 - x1), y1 + v * (y2 - y1)]);
            continue;
          }
          /* the canvas background frame of an editas="canvas" group has no ink of its own */
          const k = vmlShape(c, ctx, Math.abs(x2 - x1), Math.abs(y2 - y1));
          if (!k) continue;
          if (k.t === 'img' && !k.media) continue;
          if (k.t === 'shape' && k.fill.t === 'none' && k.line.t === 'none' && !k.tb && !k.wordart) continue;
          k.x = Math.min(x1, x2); k.y = Math.min(y1, y2);
          it.kids.push(k);
        }
      };
      walk(g, (u, v) => [u * w, v * h]);
      return it.kids.length ? it : null;
    }
    function parseVML(pict, ctx, base) {
      for (const t of pict.getElementsByTagName('*')) if (t.localName === 'shapetype' && at(t, 'id')) vmlTypes.set(at(t, 'id'), t);
      const shapes = Array.from(pict.children).filter((c) => /^(shape|rect|roundrect|oval|line|group|image|polyline|arc)$/.test(c.localName));
      if (!shapes.length) return null;
      const s = shapes[shapes.length - 1];
      const st = styleMap(at(s, 'style'));
      const id = at(s, 'id') || '';
      const w = cssLen(st.width), h = cssLen(st.height);
      /* watermarks live in headers as VML WordArt or pictures */
      if (/WaterMark/i.test(id)) {
        const tp = desc(s, 'textpath');
        if (tp) {
          const ts = styleMap(at(tp, 'style'));
          doc.watermark = { type: 'text', text: at(tp, 'string') || '', font: (ts['font-family'] || 'Calibri').replace(/["']/g, ''), size: ts['font-size'] ? cssLen(ts['font-size']) : 0, color: (vColor(at(s, 'fillcolor')) || '#C0C0C0').replace('#', ''), semi: !!desc(s, 'fill') && /\.5|50/.test(at(desc(s, 'fill'), 'opacity') || ''), layout: /rotation:\s*315|rotation:-45/.test(at(s, 'style') || '') ? 'diagonal' : 'horizontal' };
          return D.item('raw', { zero: true, watermark: true }, base);
        }
        const im = desc(s, 'imagedata');
        if (im) {
          const r = rid(im, 'id');
          const rel = r && ctx.rels[r];
          doc.watermark = { type: 'picture', media: rel ? mediaCache.get(rel.target) : null, w, h, washout: true };
          return D.item('raw', { zero: true, watermark: true }, base);
        }
      }
      let it;
      if (s.localName === 'group') it = vmlGroup(s, ctx, w || 72, h || 72);
      else if (s.localName === 'line') {
        it = vmlShape(s, ctx, 1, 1);
        const from = (at(s, 'from') || '0,0').split(','), to = (at(s, 'to') || '0,0').split(',');
        const x1 = cssLen(from[0]), y1 = cssLen(from[1]), x2 = cssLen(to[0]), y2 = cssLen(to[1]);
        it.w = Math.abs(x2 - x1) || 0.5; it.h = Math.abs(y2 - y1) || 0.5;
        it.geom = 'line'; it.fill = { t: 'none' };
        if (x2 < x1) it.flipH = true;
        if (y2 < y1) it.flipV = true;
        if (st.position === 'absolute' && !st['margin-left'] && !st.left) st['margin-left'] = Math.min(x1, x2) + 'pt';
        if (st.position === 'absolute' && !st['margin-top'] && !st.top) st['margin-top'] = Math.min(y1, y2) + 'pt';
      } else it = vmlShape(s, ctx, w, h);
      if (!it) return null;
      if (it.t === 'img') { const ole = kid(pict, 'OLEObject'); if (ole) it.ole = { progId: at(ole, 'ProgID') }; }
      const absolute = st.position === 'absolute';
      if (absolute) {
        const wrapEl = desc(s, 'wrap');
        const wt = wrapEl ? at(wrapEl, 'type') : null;
        const z = parseInt(st['z-index'] || '0', 10);
        const f = {
          posH: { rel: st['mso-position-horizontal-relative'] || 'column' }, posV: { rel: st['mso-position-vertical-relative'] || 'paragraph' },
          dist: { t: 0, b: 0, l: 9, r: 9 }, z: z + 251658240, behind: z < 0,
        };
        if (st['mso-position-horizontal'] && st['mso-position-horizontal'] !== 'absolute') f.posH.align = st['mso-position-horizontal']; else f.posH.off = cssLen(st['margin-left'] || st.left || '0');
        if (st['mso-position-vertical'] && st['mso-position-vertical'] !== 'absolute') f.posV.align = st['mso-position-vertical']; else f.posV.off = cssLen(st['margin-top'] || st.top || '0');
        if (f.posH.rel === 'text') f.posH.rel = 'column';
        if (f.posV.rel === 'text') f.posV.rel = 'paragraph';
        f.wrap = wt === 'square' ? 'square' : wt === 'tight' ? 'tight' : wt === 'through' ? 'through' : wt === 'topAndBottom' ? 'topBottom' : f.behind ? 'behind' : 'none';
        it.float = f;
      }
      it.rPr = base;
      it.vml = true;
      return it;
    }

    /* ---- headers & footers, notes, comments ---- */
    const hfMap = {};
    let hfN = 0;
    for (const r of Object.keys(docRels)) {
      const rel = docRels[r];
      if ((rel.type === 'header' || rel.type === 'footer') && !rel.external) {
        const x = await xml(rel.target);
        if (!x) continue;
        const hr = await rels(rel.target);
        await preload(hr);
        const sid = (rel.type === 'header' ? 'h' : 'f') + ++hfN;
        const story = { kind: rel.type === 'header' ? 'hdr' : 'ftr', id: sid, blocks: [], keep: { part: rel.target } };
        const ctx = storyCtx(rel.target, hr, story);
        story.blocks = parseBlocks(x, ctx);
        if (!story.blocks.length || story.blocks[story.blocks.length - 1].t !== 'p') { const sp = D.para(); sp.synth = true; story.blocks.push(sp); }
        finishStory(story, ctx);
        doc.hf[sid] = story;
        hfMap[r] = sid;
      }
    }
    async function notesPart(type, kind, tag) {
      const rel = relOfType(type);
      if (!rel) return;
      const x = await xml(rel.target);
      if (!x) return;
      const nr = await rels(rel.target);
      await preload(nr);
      if (doc.pkg) {
        doc.keep.notes = doc.keep.notes || {};
        doc.keep.notes[kind] = Array.from(doc.pkg.xml(rel.target)?.children || []).filter(n => ['separator', 'continuationSeparator', 'continuationNotice'].includes(at(n, 'type')))
          .map(n => ({ id: at(n, 'id'), fragment: L.opc.fragment(n, { pkg: doc.pkg, part: rel.target }) }));
      }
      for (const n of kids(x, tag)) {
        const ty = at(n, 'type');
        const id = at(n, 'id');
        if (ty === 'separator' || ty === 'continuationSeparator' || ty === 'continuationNotice') continue;
        const story = { kind, id, blocks: [] };
        const ctx = storyCtx(rel.target, nr, story);
        ctx.noteId = id;
        story.blocks = parseBlocks(n, ctx);
        if (!story.blocks.length) story.blocks.push(D.para());
        finishStory(story, ctx);
        (kind === 'fn' ? doc.fn : doc.en)[id] = story;
      }
    }
    await notesPart('footnotes', 'fn', 'footnote');
    await notesPart('endnotes', 'en', 'endnote');
    const cmRel = relOfType('comments');
    if (cmRel) {
      const x = await xml(cmRel.target);
      const cr = await rels(cmRel.target);
      await preload(cr);
      /* threading & resolved state (Word 2013+) */
      const exRel = relOfType('commentsExtended');
      const exX = exRel ? await xml(exRel.target) : null;
      const exMap = new Map();
      if (exX) for (const e of kids(exX, 'commentEx')) exMap.set(at(e, 'paraId'), { done: X.bool(at(e, 'done')), parent: at(e, 'paraIdParent') });
      const paraOwner = new Map();
      if (x) for (const c of kids(x, 'comment')) {
        const id = at(c, 'id');
        const story = { kind: 'cmt', id, author: at(c, 'author') || 'Unknown', initials: at(c, 'initials') || '', date: at(c, 'date') || '', blocks: [] };
        const ctx = storyCtx(cmRel.target, cr, story);
        story.blocks = parseBlocks(c, ctx);
        if (!story.blocks.length) story.blocks.push(D.para());
        finishStory(story, ctx);
        const ps = kids(c, 'p');
        const lastP = ps[ps.length - 1];
        const pid = lastP && at(lastP, 'paraId');
        if (pid) { story.keep = { paraId: pid }; paraOwner.set(pid, id); const ex = exMap.get(pid); if (ex) { story.done = ex.done; story._parentPara = ex.parent; } }
        doc.comments[id] = story;
      }
      for (const k in doc.comments) { const c = doc.comments[k]; if (c._parentPara) { c.parent = paraOwner.get(c._parentPara); delete c._parentPara; } }
      const exRoot = exX && (exX.documentElement || exX);
      if (exRoot && exRoot.namespaceURI !== 'http://schemas.microsoft.com/office/word/2012/wordml') {
        // Pre-release comment extensions are not interchangeable with the
        // released Word 2013 vocabulary. Word repairs an automatic promotion.
        doc.keep.commentExtension = { part: exRel.target, namespace: exRoot.namespaceURI, values: L.preserve.commentMetadata(doc) };
      }
    }

    /* ---- body ---- */
    await preload(docRels);
    const body = kid(docX, 'body');
    const bg = kid(docX, 'background');
    if (bg) { const c = wColor(bg, 'color', 'themeColor'); if (c && c !== 'auto' && c !== 'FFFFFF') doc.bg = c; doc.keep.background = { fragment: L.opc.fragment(bg, { pkg: doc.pkg, part: mainPath }), value: doc.bg }; }
    const mctx = storyCtx(mainPath, docRels, doc.main);
    mctx.hfMap = hfMap;
    if (body) {
      doc.main.blocks = parseBlocks(body, mctx);
      const fs = kid(body, 'sectPr');
      doc.sect = fs ? parseSect(fs, mctx) : D.defaultSect();
    }
    /* sections inside paragraphs need the header map too (parsed during parseBlocks) */
    if (!doc.main.blocks.length || doc.main.blocks[doc.main.blocks.length - 1].t !== 'p') doc.main.blocks.push(D.para());
    if (mctx.pendingBm.length) doc.main.blocks[doc.main.blocks.length - 1].runs.push(...mctx.pendingBm);
    finishStory(doc.main, mctx);

    function finishStory(story, ctx) {
      /* drop hyperlink field markers (links live on the runs) and repair unmatched field markers */
      const seen = new Map();
      D.walk(story, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.fid) { const s = seen.get(it.fid) || {}; s[it.t] = true; seen.set(it.fid, s); } });
      const bad = new Set(ctx.dropFids);
      for (const [f, s] of seen) if (!s.fb || !s.fe) bad.add(f);
      if (bad.size) D.walk(story, (b) => { if (b.t === 'p' && b.runs.some((it) => bad.has(it.fid))) b.runs = b.runs.filter((it) => !bad.has(it.fid)); });
      /* clean up empty instr fields & trim */
      D.walk(story, (b) => { if (b.t === 'p') { for (const it of b.runs) if (it.t === 'fb' && it.instr) it.instr = ' ' + it.instr.trim() + ' '; D.normalize(b); } });
    }

    /* ---- properties ---- */
    const propertyPart = (type, fallback) => doc.pkg?.rels('').find(r => L.opc.relationshipType(r.type) === type && !r.external)?.part || fallback;
    const coreX = await xml(propertyPart(L.opc.NS.pkg + '/metadata/core-properties', 'docProps/core.xml'));
    if (coreX) {
      const g = (n) => { const e = desc(coreX, n); return e ? e.textContent : ''; };
      Object.assign(doc.props, { title: g('title'), subject: g('subject'), creator: g('creator'), keywords: g('keywords'), description: g('description'), lastModifiedBy: g('lastModifiedBy'), revision: +g('revision') || 1, created: g('created') || doc.props.created, modified: g('modified') || doc.props.modified, category: g('category'), lastPrinted: g('lastPrinted'), contentStatus: g('contentStatus'), language: g('language'), identifier: g('identifier'), version: g('version') });
    }
    const appX = await xml(propertyPart(L.opc.NS.rel + '/extended-properties', 'docProps/app.xml'));
    if (appX) { const g = (n) => { const e = desc(appX, n); return e ? e.textContent : ''; }; doc.props.company = g('Company'); doc.props.manager = g('Manager'); doc.props.template = g('Template'); }
    const custX = await xml(propertyPart(L.opc.NS.rel + '/custom-properties', 'docProps/custom.xml'));
    if (custX) for (const p of kids(custX, 'property')) { const v = p.firstElementChild; if (at(p, 'name')) doc.custom[at(p, 'name')] = v ? v.textContent : ''; }
    const mainCT = await (async () => { const ct = await xml('[Content_Types].xml'); if (!ct) return ''; const o = kids(ct, 'Override').find((x) => at(x, 'PartName') === '/' + mainPath); return o ? at(o, 'ContentType') : ''; })();
    if (/template/.test(mainCT)) doc.isTemplate = true;
    if (/macroEnabled/.test(mainCT)) doc.hasMacros = true;
    doc.ooxmlFormat = L.opc.variant(doc.pkg, doc.isTemplate ? 'dotx' : 'docx');
    doc.keep.values = Object.fromEntries(['settings', 'theme', 'props', 'custom', 'watermark'].map(k => [k, JSON.parse(JSON.stringify(doc[k]))]));

    /* numbering instances referenced but undefined → plain paragraphs */
    D.walk(doc.main, (b) => { if (b.t === 'p' && b.pPr.num && b.pPr.num.id && b.pPr.num.id !== '0' && !doc.numbering.nums[b.pPr.num.id]) delete b.pPr.num; });
    D.stylesChanged();
    doc._idxDirty = true;
    L.preserve.finishObjects(doc);
    return { doc, warnings };
  };
})();
