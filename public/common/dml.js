/* VibeOffice — DrawingML helpers shared by Quire's .docx reader and Ledger: colours, fills, lines, geometry, charts. */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const X = (L.dml = {});
  const emu = (v) => (+v || 0) / 12700;
  X.emu = emu;

  /* ---------- XML helpers ---------- */
  X.parse = function (s, beforeResolve) {
    if (s && s.charCodeAt(0) === 0xFEFF) s = s.slice(1); /* byte-order mark left in by some writers */
    /* characters XML 1.0 forbids are stripped and the parse retried (seen in damaged files) */
    const strip = (t) => t.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
    if (L.xml) {   // Ledger's own parser (also in Node)
      let el;
      try { el = L.xml.parse(s); } catch (e) { el = L.xml.parse(strip(s)); }
      if (beforeResolve) beforeResolve(el.parentNode);
      X.resolveAC(el.parentNode, s);
      return el.parentNode.children[0];
    }
    let d = new DOMParser().parseFromString(s, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) {
      d = new DOMParser().parseFromString(strip(s), 'application/xml');
      if (d.getElementsByTagName('parsererror').length) throw new Error('XML parse error');
    }
    if (beforeResolve) beforeResolve(d);
    X.resolveAC(d, s);
    return d.documentElement;
  };
  /** Markup Compatibility: keep the Choice we understand (wps/wpg/wpc/w14), else the Fallback */
  const UNDERSTOOD = /^(wps|wpg|wp14|w14|w15|a14|wpc|v|o|w10|mc|wne|m)$/;
  X.resolveAC = function (d, source) {
    if (L.opc) L.opc.captureAC(d, source);
    const acs = Array.from(d.getElementsByTagName('*')).filter((e) => e.localName === 'AlternateContent').reverse();
    for (const ac of acs) {
      if (!ac.parentNode) continue;
      const kidsEl = Array.from(ac.children);
      const fb = kidsEl.find((c) => c.localName === 'Fallback');
      let pick = null;
      for (const ch of kidsEl.filter((c) => c.localName === 'Choice')) {
        const req = (ch.getAttribute('Requires') || '').split(/\s+/).filter(Boolean);
        if (req.every((r) => UNDERSTOOD.test(r))) { pick = ch; break; }
      }
      if (!pick) pick = fb || kidsEl.find((c) => c.localName === 'Choice');
      if (pick) ac.replaceWith(...Array.from(pick.childNodes)); else ac.remove();
    }
  };
  X.kids = (el, name) => (el ? Array.from(el.children).filter((c) => !name || c.localName === name) : []);
  X.kid = (el, name) => { if (!el) return null; for (const c of el.children) if (c.localName === name) return c; return null; };
  X.path = (el, ...names) => { let e = el; for (const n of names) { e = X.kid(e, n); if (!e) return null; } return e; };
  X.desc = (el, name) => { if (!el) return null; const all = el.getElementsByTagName('*'); for (const e of all) if (e.localName === name) return e; return null; };
  X.descAll = (el, name) => (el ? Array.from(el.getElementsByTagName('*')).filter((e) => e.localName === name) : []);
  /** attribute by local name (namespace-agnostic) */
  X.at = (el, n) => {
    if (!el) return null;
    const pre = el.prefix;
    let v;
    /* WordprocessingML attributes carry the element's own prefix: try that first (one lookup instead of two) */
    if (pre === 'w') { v = el.getAttribute('w:' + n); if (v != null) return v; }
    v = el.getAttribute(n);
    if (v != null) return v;
    if (pre !== 'w') { v = el.getAttribute('w:' + n); if (v != null) return v; }
    const at = el.attributes;
    for (let i = 0; i < at.length; i++) if (at[i].localName === n) return at[i].value;
    return null;
  };
  X.num = (el, n, d) => { const v = X.at(el, n); if (v == null || v === '') return d; const x = parseFloat(v); return isNaN(x) ? d : x; };
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_R_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
  X.rid = (el, n) => (el ? el.getAttributeNS(NS_R, n || 'id') || el.getAttributeNS(NS_R_STRICT, n || 'id') || el.getAttribute('r:' + (n || 'id')) : null);
  X.bool = (v, d) => (v == null ? d : v === '1' || v === 'true' || v === 'on' || v === 't');
  X.resolvePath = function (base, target) {
    if (/^\//.test(target)) return target.slice(1);
    const parts = base.split('/');
    parts.pop();
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.' && seg !== '') parts.push(seg);
    }
    return parts.join('/');
  };
  X.relsPath = (p) => { const i = p.lastIndexOf('/'); return p.slice(0, i) + '/_rels/' + p.slice(i + 1) + '.rels'; };

  /* ---------- colours ---------- */
  const SYS = { windowText: '#000000', window: '#FFFFFF', btnFace: '#ECE9D8', highlight: '#316AC5', highlightText: '#FFFFFF', grayText: '#808080', menu: '#FFFFFF', menuText: '#000000' };
  const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);
  X.colorEl = (parent) => (parent ? Array.from(parent.children).find((c) => COLOR_TAGS.has(c.localName)) || null : null);
  const SCHEME_ALIAS = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2', dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2', hyperlink: 'hlink', followedHyperlink: 'folHlink' };
  X.schemeHex = (name, ctx) => {
    const t = (ctx && ctx.theme && ctx.theme.colors) || X.DEFAULT_THEME;
    const k = SCHEME_ALIAS[name] || name;
    return t[k] || X.DEFAULT_THEME[k] || '#000000';
  };
  X.DEFAULT_THEME = { dk1: '#000000', lt1: '#FFFFFF', dk2: '#1F497D', lt2: '#EEECE1', accent1: '#4F81BD', accent2: '#C0504D', accent3: '#9BBB59', accent4: '#8064A2', accent5: '#4BACC6', accent6: '#F79646', hlink: '#0000FF', folHlink: '#800080' };
  /** → {c:'#hex', a} */
  X.color = function (el, ctx) {
    if (!el) return null;
    let base = null;
    switch (el.localName) {
      case 'srgbClr': base = '#' + (X.at(el, 'val') || '000000').toUpperCase(); break;
      case 'schemeClr': { const v = X.at(el, 'val'); base = v === 'phClr' ? (ctx && ctx.phClr) || '#000000' : X.schemeHex(v, ctx); break; }
      case 'sysClr': base = '#' + (X.at(el, 'lastClr') || '').toUpperCase(); if (base === '#') base = SYS[X.at(el, 'val')] || '#000000'; break;
      case 'prstClr': base = L.color.PRESET[X.at(el, 'val')] || '#000000'; break;
      case 'scrgbClr': { const lin = (v) => { v = (+v || 0) / 100000; return v <= 0.0031308 ? v * 12.92 * 255 : (1.055 * Math.pow(v, 1 / 2.4) - 0.055) * 255; }; base = L.color.rgbToHex(lin(X.at(el, 'r')), lin(X.at(el, 'g')), lin(X.at(el, 'b'))); break; }
      case 'hslClr': { const [r, g, b] = L.color.hslToRgb((+X.at(el, 'hue') || 0) / 21600000, (+X.at(el, 'sat') || 0) / 100000, (+X.at(el, 'lum') || 0) / 100000); base = L.color.rgbToHex(r, g, b); break; }
      default: return null;
    }
    let alpha = 1;
    const mods = [];
    for (const m of el.children) {
      const v = +X.at(m, 'val');
      if (m.localName === 'alpha') alpha = v / 100000;
      else if (m.localName === 'alphaMod') alpha *= v / 100000;
      else if (m.localName === 'alphaOff') alpha += v / 100000;
      else mods.push({ name: m.localName, val: isNaN(v) ? 0 : v });
    }
    return { c: L.color.applyMods(base, mods), a: L.clamp(alpha, 0, 1) };
  };
  X.colorOf = (parent, ctx) => X.color(X.colorEl(parent), ctx);

  /* ---------- fills & lines ---------- */
  const FILL_TAGS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
  X.fill = function (el, ctx) {
    if (!el) return undefined;
    switch (el.localName) {
      case 'noFill': return { t: 'none' };
      case 'solidFill': { const c = X.colorOf(el, ctx); return c ? { t: 'solid', c: c.c, a: c.a } : { t: 'solid', c: '#000000', a: 1 }; }
      case 'gradFill': {
        const stops = X.kids(X.kid(el, 'gsLst'), 'gs').map((g) => { const c = X.colorOf(g, ctx) || { c: '#000000', a: 1 }; return { p: X.num(g, 'pos', 0) / 100000, c: c.c, a: c.a }; });
        const lin = X.kid(el, 'lin'), p = X.kid(el, 'path');
        return { t: 'grad', stops: stops.length ? stops : [{ p: 0, c: '#FFFFFF', a: 1 }, { p: 1, c: '#000000', a: 1 }], ang: lin ? X.num(lin, 'ang', 0) / 60000 : 90, path: p ? (X.at(p, 'path') === 'rect' ? 'rect' : X.at(p, 'path') === 'shape' ? 'shape' : 'circle') : 'lin' };
      }
      case 'blipFill': {
        const blip = X.kid(el, 'blip');
        const id = blip && ctx.media ? ctx.media(X.rid(blip, 'embed') || X.rid(blip, 'link')) : null;
        if (!id) return { t: 'none' };
        const amt = X.desc(blip, 'alphaModFix');
        return { t: 'img', media: id, tile: !!X.kid(el, 'tile'), a: amt ? X.num(amt, 'amt', 100000) / 100000 : 1 };
      }
      case 'pattFill': {
        const fg = X.colorOf(X.kid(el, 'fgClr'), ctx), bg = X.colorOf(X.kid(el, 'bgClr'), ctx);
        return { t: 'patt', prst: X.at(el, 'prst') || 'pct50', fg: fg ? fg.c : '#000000', bg: bg ? bg.c : '#FFFFFF' };
      }
      case 'grpFill': return ctx && ctx.groupFill ? L.clone(ctx.groupFill) : { t: 'none' };
      default: return undefined;
    }
  };
  X.fillIn = (parent, ctx) => { if (!parent) return undefined; const f = Array.from(parent.children).find((c) => FILL_TAGS.has(c.localName)); return f ? X.fill(f, ctx) : undefined; };
  X.line = function (el, ctx, base) {
    if (!el) return base;
    const ln = Object.assign({ c: '#000000', w: 0.75, dash: 'solid' }, base && base.t !== 'none' ? base : {});
    delete ln.t;
    if (X.at(el, 'w') != null) ln.w = L.round(emu(X.at(el, 'w')), 3);
    if (X.at(el, 'cmpd')) ln.cmpd = X.at(el, 'cmpd');
    const f = X.fillIn(el, ctx);
    if (f) {
      if (f.t === 'none') return { t: 'none', w: ln.w };
      if (f.t === 'solid') { ln.c = f.c; ln.a = f.a; }
      else if (f.t === 'grad') ln.c = f.stops[0].c;
      else if (f.t === 'patt') ln.c = f.fg;
    } else if (base && base.t === 'none') return { t: 'none', w: ln.w };
    const dash = X.kid(el, 'prstDash');
    if (dash) ln.dash = X.at(dash, 'val') || 'solid';
    const he = X.kid(el, 'headEnd'), te = X.kid(el, 'tailEnd');
    if (he && X.at(he, 'type') && X.at(he, 'type') !== 'none') ln.head = { type: X.at(he, 'type'), w: X.at(he, 'w') || 'med', len: X.at(he, 'len') || 'med' };
    if (te && X.at(te, 'type') && X.at(te, 'type') !== 'none') ln.tail = { type: X.at(te, 'type'), w: X.at(te, 'w') || 'med', len: X.at(te, 'len') || 'med' };
    return ln;
  };
  X.shadow = function (effectLst, ctx) {
    const o = effectLst && X.kid(effectLst, 'outerShdw');
    if (!o) return null;
    const dist = emu(X.num(o, 'dist', 0)), dir = (X.num(o, 'dir', 0) / 60000) * Math.PI / 180;
    const c = X.colorOf(o, ctx) || { c: '#000000', a: 0.5 };
    return { c: c.c, a: c.a, dx: L.round(Math.cos(dir) * dist, 2), dy: L.round(Math.sin(dir) * dist, 2), blur: L.round(emu(X.num(o, 'blurRad', 0)), 2) };
  };
  /** style matrix references (wps:style lnRef/fillRef/fontRef) */
  X.styleRefs = function (styleEl, ctx) {
    if (!styleEl) return {};
    const o = {};
    const fr = X.kid(styleEl, 'fillRef'), lr = X.kid(styleEl, 'lnRef'), fo = X.kid(styleEl, 'fontRef');
    if (fr && X.num(fr, 'idx', 0) > 0) { const c = X.colorOf(fr, ctx); if (c) o.fill = { t: 'solid', c: c.c, a: c.a }; }
    if (lr && X.num(lr, 'idx', 0) > 0) { const c = X.colorOf(lr, ctx); if (c) o.line = { c: c.c, w: [0, 0.75, 1.5, 2.25][X.num(lr, 'idx', 1)] || 0.75, dash: 'solid' }; }
    if (fo) { const c = X.colorOf(fo, ctx); if (c) o.fontColor = c.c; }
    return o;
  };

  /* ---------- geometry ---------- */
  function guideEnv(wE, hE) {
    const w = wE || 0, h = hE || 0, ssv = Math.min(w, h), ls = Math.max(w, h);
    const v = { w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, ls, ss: ssv, cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000 };
    for (const d of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) { v['wd' + d] = w / d; v['hd' + d] = h / d; v['ssd' + d] = ssv / d; }
    return v;
  }
  function evalGuides(list, env) {
    const val = (tok) => { if (tok == null) return 0; if (tok in env) return env[tok]; const x = parseFloat(tok); return isNaN(x) ? 0 : x; };
    const rad = (a) => (a / 60000) * Math.PI / 180;
    for (const gd of list) {
      const f = (X.at(gd, 'fmla') || '').trim().split(/\s+/);
      const [op, a1, b1, c1] = f;
      const x = val(a1), y = val(b1), z = val(c1);
      let r = 0;
      switch (op) {
        case 'val': r = x; break;
        case '*/': r = z ? (x * y) / z : 0; break;
        case '+-': r = x + y - z; break;
        case '+/': r = z ? (x + y) / z : 0; break;
        case '?:': r = x > 0 ? y : z; break;
        case 'abs': r = Math.abs(x); break;
        case 'at2': r = (Math.atan2(y, x) * 180 / Math.PI) * 60000; break;
        case 'cat2': r = x * Math.cos(Math.atan2(z, y)); break;
        case 'sat2': r = x * Math.sin(Math.atan2(z, y)); break;
        case 'cos': r = x * Math.cos(rad(y)); break;
        case 'sin': r = x * Math.sin(rad(y)); break;
        case 'tan': r = x * Math.tan(rad(y)); break;
        case 'max': r = Math.max(x, y); break;
        case 'min': r = Math.min(x, y); break;
        case 'mod': r = Math.sqrt(x * x + y * y + z * z); break;
        case 'pin': r = y < x ? x : y > z ? z : y; break;
        case 'sqrt': r = Math.sqrt(Math.max(0, x)); break;
        default: r = 0;
      }
      env[X.at(gd, 'name')] = isFinite(r) ? r : 0;
    }
    return env;
  }
  function custGeom(cg, extW, extH) {
    const env = guideEnv(extW, extH);
    evalGuides(X.kids(X.kid(cg, 'avLst'), 'gd'), env);
    evalGuides(X.kids(X.kid(cg, 'gdLst'), 'gd'), env);
    const V = (tok) => { if (tok == null) return 0; if (tok in env) return env[tok]; const x = parseFloat(tok); return isNaN(x) ? 0 : x; };
    const paths = [];
    for (const p of X.kids(X.kid(cg, 'pathLst'), 'path')) {
      const pw = X.num(p, 'w', 0) || extW || 0, ph = X.num(p, 'h', 0) || extH || 0;
      const W = emu(pw), H = emu(ph);
      const cmds = [];
      const P = (e) => [L.round(emu(V(X.at(e, 'x'))), 3), L.round(emu(V(X.at(e, 'y'))), 3)];
      for (const c of p.children) {
        const pts = X.kids(c, 'pt').map(P);
        switch (c.localName) {
          case 'moveTo': if (pts[0]) cmds.push(['M', ...pts[0]]); break;
          case 'lnTo': if (pts[0]) cmds.push(['L', ...pts[0]]); break;
          case 'cubicBezTo': if (pts.length >= 3) cmds.push(['C', ...pts[0], ...pts[1], ...pts[2]]); break;
          case 'quadBezTo': if (pts.length >= 2) cmds.push(['Q', ...pts[0], ...pts[1]]); break;
          case 'arcTo': cmds.push(['A', L.round(emu(V(X.at(c, 'wR'))), 3), L.round(emu(V(X.at(c, 'hR'))), 3), V(X.at(c, 'stAng')) / 60000, V(X.at(c, 'swAng')) / 60000]); break;
          case 'close': cmds.push(['Z']); break;
          default: break;
        }
      }
      if (cmds.length) paths.push({ w: Math.max(W, 0.01), h: Math.max(H, 0.01), cmds, fill: X.at(p, 'fill') === 'none' ? 'none' : 'norm', stroke: X.at(p, 'stroke') !== '0' && X.at(p, 'stroke') !== 'false' });
    }
    return paths.length ? { paths } : null;
  }
  X.geometry = function (spPr) {
    const pg = X.kid(spPr, 'prstGeom');
    const xe = spPr && X.path(spPr, 'xfrm', 'ext');
    const extW = xe ? X.num(xe, 'cx', 0) : 0, extH = xe ? X.num(xe, 'cy', 0) : 0;
    if (pg) {
      const adj = {};
      for (const gd of X.kids(X.kid(pg, 'avLst'), 'gd')) { const m = /val\s+(-?\d+)/.exec(X.at(gd, 'fmla') || ''); if (m) adj[X.at(gd, 'name')] = +m[1]; }
      return { geom: X.at(pg, 'prst') || 'rect', adj: Object.keys(adj).length ? adj : undefined };
    }
    const cg = X.kid(spPr, 'custGeom');
    if (cg) { const p = custGeom(cg, extW, extH); if (p) return { geom: 'custom', path: p }; }
    return { geom: 'rect' };
  };
  X.xfrm = function (el) {
    if (!el) return null;
    const off = X.kid(el, 'off'), ext = X.kid(el, 'ext');
    const o = { x: emu(X.num(off, 'x', 0)), y: emu(X.num(off, 'y', 0)), w: emu(X.num(ext, 'cx', 0)), h: emu(X.num(ext, 'cy', 0)), rot: X.num(el, 'rot', 0) / 60000, flipH: X.bool(X.at(el, 'flipH')), flipV: X.bool(X.at(el, 'flipV')) };
    const co = X.kid(el, 'chOff'), ce = X.kid(el, 'chExt');
    if (co && ce) { o.chx = emu(X.num(co, 'x', 0)); o.chy = emu(X.num(co, 'y', 0)); o.chw = emu(X.num(ce, 'cx', 0)); o.chh = emu(X.num(ce, 'cy', 0)); }
    return o;
  };

  /* ---------- charts (c:chartSpace → chart model used by L.chart) ---------- */
  X.chart = function (cx, ctx) {
    if (!cx) return null;
    const { kid, kids, desc, descAll, at, num, path } = X;
    const chart = kid(cx, 'chart');
    const plot = kid(chart, 'plotArea');
    if (!plot) return null;
    const types = ['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'pieChart', 'pie3DChart', 'ofPieChart', 'doughnutChart', 'areaChart', 'area3DChart', 'scatterChart', 'radarChart', 'bubbleChart', 'stockChart'];
    const ct = kids(plot).find((c) => types.includes(c.localName));
    if (!ct) return null;
    const t = ct.localName;
    const grouping = at(kid(ct, 'grouping'), 'val') || 'clustered';
    const barDir = at(kid(ct, 'barDir'), 'val') || 'col';
    let kind = 'col';
    if (/bar/.test(t)) kind = barDir === 'bar' ? (grouping === 'clustered' ? 'bar' : 'barStacked') : grouping === 'percentStacked' ? 'colPct' : grouping === 'stacked' ? 'colStacked' : 'col';
    else if (/line|radar|stock/.test(t)) kind = 'lineMarkers';
    else if (/pie/i.test(t)) kind = 'pie';
    else if (/doughnut/.test(t)) kind = 'doughnut';
    else if (/area/.test(t)) kind = grouping === 'stacked' ? 'areaStacked' : 'area';
    else if (/scatter|bubble/.test(t)) kind = 'scatter';
    if (kind === 'lineMarkers' && /line/.test(t)) { const s0 = kid(ct, 'ser'); const mk = s0 && path(s0, 'marker', 'symbol'); if (mk && at(mk, 'val') === 'none') kind = 'line'; }
    const strVals = (el) => {
      if (!el) return [];
      const ml = desc(el, 'multiLvlStrCache');
      const cache = ml || desc(el, 'strCache') || desc(el, 'numCache') || desc(el, 'strLit') || desc(el, 'numLit');
      if (!cache) { const v = desc(el, 'v'); return v ? [v.textContent] : []; }
      const n = num(kid(cache, 'ptCount'), 'val', 0);
      const out = new Array(n).fill('');
      for (const p2 of kids(ml ? kid(cache, 'lvl') || cache : cache, 'pt')) out[num(p2, 'idx', 0)] = (kid(p2, 'v') || { textContent: '' }).textContent;
      return out;
    };
    const pal = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((k) => X.schemeHex(k, ctx));
    const others = kids(plot).filter((c) => types.includes(c.localName) && c !== ct);
    const sers = kids(ct, 'ser').concat(...others.map((o) => kids(o, 'ser').map((x) => Object.assign(x, { __group: o }))));
    let cats = [];
    const series = sers.map((s, i) => {
      const name = strVals(kid(s, 'tx'))[0] || 'Series ' + (i + 1);
      const c = kid(s, 'cat') || kid(s, 'xVal');
      const v = kid(s, 'val') || kid(s, 'yVal');
      const cv = strVals(c);
      if (cv.length > cats.length) cats = cv;
      const vals = strVals(v).map((x) => +x || 0);
      const sp = kid(s, 'spPr');
      let col = null;
      const sf = sp && kid(sp, 'solidFill');
      if (sf) { const cc = X.colorOf(sf, ctx); if (cc) col = cc.c; }
      else if (sp && kid(sp, 'ln') && kid(kid(sp, 'ln'), 'solidFill')) { const cc = X.colorOf(kid(kid(sp, 'ln'), 'solidFill'), ctx); if (cc) col = cc.c; }
      const out = { name, vals, color: col || pal[i % pal.length] };
      if (s.__group) {
        const gt = s.__group.localName;
        const noLine = sp && kid(sp, 'ln') && kid(kid(sp, 'ln'), 'noFill');
        const mk = path(s, 'marker', 'symbol');
        out.overlay = /line|scatter|radar/.test(gt) ? (noLine ? 'markers' : mk && at(mk, 'val') === 'none' ? 'line' : 'lineMarkers') : 'line';
      }
      return out;
    });
    if (!cats.length && series[0]) cats = series[0].vals.map((_, i) => String(i + 1));
    const model = { kind, title: '', legend: 'none', gridY: !!desc(plot, 'majorGridlines'), labels: false, fsz: 10, cats, series };
    const title = kid(chart, 'title');
    if (title) model.title = kid(title, 'tx') ? descAll(title, 't').map((x) => x.textContent).join('').trim() : series.length === 1 ? series[0].name : '';
    /* axis titles: the one on a horizontal axis (axPos b/t) and the one on a vertical axis (l/r) */
    for (const ax of kids(plot).filter((k) => /^(catAx|valAx|dateAx)$/.test(k.localName))) {
      const tt = kid(ax, 'title');
      if (!tt || (kid(ax, 'delete') && X.bool(at(kid(ax, 'delete'), 'val')))) continue;
      const txt = descAll(tt, 't').map((x) => x.textContent).join('').trim();
      if (!txt) continue;
      const pos = at(kid(ax, 'axPos'), 'val') || 'b';
      if (/[bt]/.test(pos)) { if (!model.axTitleX) model.axTitleX = txt; } else if (!model.axTitleY) {
        model.axTitleY = txt;
        const bp = desc(tt, 'bodyPr');
        if (bp && at(bp, 'rot') != null && +at(bp, 'rot') === 0) model.axTitleYHoriz = true;
      }
    }
    const leg = kid(chart, 'legend');
    if (leg) model.legend = at(kid(leg, 'legendPos'), 'val') || 'r';
    const dl = desc(ct, 'dLbls');
    if (dl && (at(kid(dl, 'showVal'), 'val') === '1' || at(kid(dl, 'showPercent'), 'val') === '1')) model.labels = true;
    if ((kind === 'pie' || kind === 'doughnut') && sers[0]) {
      model.pieColors = cats.map((_, k) => {
        const dp = kids(sers[0], 'dPt').find((d) => num(kid(d, 'idx'), 'val', -1) === k);
        const sf = dp && path(dp, 'spPr', 'solidFill');
        if (sf) { const cc = X.colorOf(sf, ctx); if (cc) return cc.c; }
        return pal[k % pal.length];
      });
    }
    const txPr = desc(cx, 'defRPr');
    if (txPr && at(txPr, 'sz')) model.fsz = +at(txPr, 'sz') / 100;
    return model;
  };
})(typeof window !== 'undefined' ? window : globalThis);
