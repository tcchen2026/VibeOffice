/* Lectern — PresentationML reader (.pptx / .ppsx / .potx / .pptm). */
(function () {
  'use strict';
  const L = window.L;
  const pt = L.emu2pt;
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /* ---------- XML helpers ---------- */
  const parse = (s) => {
    const d = new DOMParser().parseFromString(s, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw new Error('XML parse error');
    resolveAlternateContent(d);
    return d.documentElement;
  };
  /** Markup Compatibility: replace every mc:AlternateContent with its Fallback (or its first Choice when there is no fallback). */
  function resolveAlternateContent(d) {
    const acs = Array.from(d.getElementsByTagName('*')).filter((e) => e.localName === 'AlternateContent').reverse();
    for (const ac of acs) {
      if (!ac.parentNode) continue;
      const kidsEl = Array.from(ac.children);
      const fb = kidsEl.find((c) => c.localName === 'Fallback');
      const ch = kidsEl.find((c) => c.localName === 'Choice');
      const pick = fb && fb.children.length ? fb : ch || fb;
      if (pick) ac.replaceWith(...Array.from(pick.childNodes)); else ac.remove();
    }
  }
  const kids = (el, name) => (el ? Array.from(el.children).filter((c) => !name || c.localName === name) : []);
  const kid = (el, name) => (el ? Array.from(el.children).find((c) => c.localName === name) || null : null);
  const path = (el, ...names) => { let e = el; for (const n of names) { e = kid(e, n); if (!e) return null; } return e; };
  const desc = (el, name) => { if (!el) return null; const all = el.getElementsByTagName('*'); for (const e of all) if (e.localName === name) return e; return null; };
  const descAll = (el, name) => (el ? Array.from(el.getElementsByTagName('*')).filter((e) => e.localName === name) : []);
  const at = (el, n) => (el ? el.getAttribute(n) : null);
  const num = (el, n, d) => { const v = at(el, n); return v == null || v === '' ? d : +v; };
  const rid = (el, n) => (el ? el.getAttributeNS(NS_R, n || 'id') || el.getAttribute('r:' + (n || 'id')) : null);
  const bool = (v) => v === '1' || v === 'true' || v === 'on';
  function resolvePath(base, target) {
    if (/^\//.test(target)) return target.slice(1);
    const parts = base.split('/');
    parts.pop();
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.' && seg !== '') parts.push(seg);
    }
    return parts.join('/');
  }
  const relsPath = (p) => { const i = p.lastIndexOf('/'); return p.slice(0, i) + '/_rels/' + p.slice(i + 1) + '.rels'; };

  /* ---------- colours ---------- */
  const SYS = { windowText: '#000000', window: '#FFFFFF', btnFace: '#ECE9D8', highlight: '#316AC5', highlightText: '#FFFFFF', grayText: '#808080', menu: '#FFFFFF', menuText: '#000000' };
  const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);
  function colorEl(parent) { return parent ? Array.from(parent.children).find((c) => COLOR_TAGS.has(c.localName)) || null : null; }
  /** → {c, a} where c is '#hex' or a scheme slot name */
  function color(el, ctx) {
    if (!el) return null;
    let base = null, scheme = null;
    switch (el.localName) {
      case 'srgbClr': base = '#' + (at(el, 'val') || '000000').toUpperCase(); break;
      case 'schemeClr': {
        const v = at(el, 'val');
        if (v === 'phClr') { if (ctx.phClr) { if (ctx.phClr.c[0] === '#') base = ctx.phClr.c; else scheme = ctx.phClr.c; } else scheme = 'accent1'; }
        else scheme = v;
        break;
      }
      case 'sysClr': base = '#' + (at(el, 'lastClr') || '').toUpperCase(); if (base === '#') base = SYS[at(el, 'val')] || '#000000'; break;
      case 'prstClr': base = L.color.PRESET[at(el, 'val')] || '#000000'; break;
      case 'scrgbClr': { const lin = (v) => { v = (+v || 0) / 100000; return v <= 0.0031308 ? v * 12.92 * 255 : (1.055 * Math.pow(v, 1 / 2.4) - 0.055) * 255; }; base = L.color.rgbToHex(lin(at(el, 'r')), lin(at(el, 'g')), lin(at(el, 'b'))); break; }
      case 'hslClr': { const [r, g, b] = L.color.hslToRgb((+at(el, 'hue') || 0) / 21600000, (+at(el, 'sat') || 0) / 100000, (+at(el, 'lum') || 0) / 100000); base = L.color.rgbToHex(r, g, b); break; }
      default: return null;
    }
    let alpha = 1;
    const mods = [];
    for (const m of el.children) {
      const v = +at(m, 'val');
      if (m.localName === 'alpha') alpha = v / 100000;
      else if (m.localName === 'alphaMod') alpha *= v / 100000;
      else if (m.localName === 'alphaOff') alpha += v / 100000;
      else mods.push({ name: m.localName, val: isNaN(v) ? 0 : v });
    }
    alpha = L.clamp(alpha, 0, 1);
    if (scheme && !mods.length && ctx.keepScheme !== false) {
      const map = ctx.clrMapOvr || {};
      return { c: map[scheme] ? map[scheme] : normScheme(scheme), a: alpha };
    }
    const hex = scheme ? L.model.resolveColor(normScheme((ctx.clrMapOvr || {})[scheme] || scheme), ctx.design) : base;
    return { c: L.color.applyMods(hex, mods), a: alpha };
  }
  const normScheme = (s) => s;
  function colorOf(parent, ctx) { return color(colorEl(parent), ctx); }

  /* ---------- fills / lines / effects ---------- */
  function fill(el, ctx) {
    if (!el) return undefined;
    switch (el.localName) {
      case 'noFill': return { t: 'none' };
      case 'solidFill': { const c = colorOf(el, ctx); return c ? { t: 'solid', c: c.c, a: c.a } : { t: 'solid', c: '#000000', a: 1 }; }
      case 'gradFill': {
        const stops = kids(kid(el, 'gsLst'), 'gs').map((g) => { const c = colorOf(g, ctx) || { c: '#000000', a: 1 }; return { p: num(g, 'pos', 0) / 100000, c: c.c, a: c.a }; });
        const lin = kid(el, 'lin'), p = kid(el, 'path');
        const f = { t: 'grad', stops: stops.length ? stops : [{ p: 0, c: '#FFFFFF', a: 1 }, { p: 1, c: '#000000', a: 1 }], ang: lin ? num(lin, 'ang', 0) / 60000 : 90, path: p ? (at(p, 'path') === 'rect' ? 'rect' : at(p, 'path') === 'shape' ? 'shape' : 'circle') : 'lin' };
        return f;
      }
      case 'blipFill': {
        const blip = kid(el, 'blip');
        const id = blip ? ctx.media(rid(blip, 'embed')) : null;
        if (!id) return { t: 'none' };
        const amt = desc(blip, 'alphaModFix');
        return { t: 'img', media: id, tile: !!kid(el, 'tile'), a: amt ? num(amt, 'amt', 100000) / 100000 : 1 };
      }
      case 'pattFill': {
        const fg = colorOf(kid(el, 'fgClr'), ctx), bg = colorOf(kid(el, 'bgClr'), ctx);
        return { t: 'patt', prst: at(el, 'prst') || 'pct50', fg: fg ? fg.c : '#000000', bg: bg ? bg.c : '#FFFFFF' };
      }
      case 'grpFill': return ctx.groupFill ? L.clone(ctx.groupFill) : { t: 'none' };
      default: return undefined;
    }
  }
  const FILL_TAGS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
  function fillIn(parent, ctx) { if (!parent) return undefined; const f = Array.from(parent.children).find((c) => FILL_TAGS.has(c.localName)); return f ? fill(f, ctx) : undefined; }
  function line(el, ctx, base) {
    if (!el) return base;
    const ln = Object.assign({ c: 'tx1', w: 0.75, dash: 'solid' }, base && base.t !== 'none' ? base : {});
    delete ln.t;
    if (at(el, 'w') != null) ln.w = L.round(pt(+at(el, 'w')), 3);
    if (at(el, 'cap')) ln.cap = at(el, 'cap');
    if (at(el, 'cmpd')) ln.cmpd = at(el, 'cmpd');
    const f = fillIn(el, ctx);
    if (f) {
      if (f.t === 'none') return { t: 'none', w: ln.w };
      if (f.t === 'solid') { ln.c = f.c; ln.a = f.a; }
      else if (f.t === 'grad') ln.c = f.stops[0].c;
      else if (f.t === 'patt') ln.c = f.fg;
    } else if (base && base.t === 'none') return { t: 'none', w: ln.w };
    const dash = kid(el, 'prstDash');
    if (dash) ln.dash = at(dash, 'val') || 'solid';
    if (kid(el, 'round')) ln.join = 'round'; else if (kid(el, 'bevel')) ln.join = 'bevel';
    const he = kid(el, 'headEnd'), te = kid(el, 'tailEnd');
    if (he && at(he, 'type') && at(he, 'type') !== 'none') ln.head = { type: at(he, 'type'), w: at(he, 'w') || 'med', len: at(he, 'len') || 'med' };
    if (te && at(te, 'type') && at(te, 'type') !== 'none') ln.tail = { type: at(te, 'type'), w: at(te, 'w') || 'med', len: at(te, 'len') || 'med' };
    return ln;
  }
  function shadow(effectLst, ctx) {
    const o = effectLst && kid(effectLst, 'outerShdw');
    if (!o) return null;
    const dist = pt(num(o, 'dist', 0)), dir = (num(o, 'dir', 0) / 60000) * Math.PI / 180;
    const c = colorOf(o, ctx) || { c: '#000000', a: 0.5 };
    return { c: c.c, a: c.a, dx: L.round(Math.cos(dir) * dist, 2), dy: L.round(Math.sin(dir) * dist, 2), blur: L.round(pt(num(o, 'blurRad', 0)), 2) };
  }

  /* ---------- text ---------- */
  const fontName = (f) => (!f ? undefined : f === '+mj-lt' || f === '+mj-ea' || f === '+mj-cs' ? '+mj' : f === '+mn-lt' || f === '+mn-ea' || f === '+mn-cs' ? '+mn' : f);
  function runProps(rPr, ctx) {
    const r = {};
    if (!rPr) return r;
    if (at(rPr, 'sz')) r.sz = +at(rPr, 'sz') / 100;
    if (at(rPr, 'b') != null) r.b = bool(at(rPr, 'b'));
    if (at(rPr, 'i') != null) r.i = bool(at(rPr, 'i'));
    if (at(rPr, 'u') && at(rPr, 'u') !== 'none') r.u = at(rPr, 'u');
    if (at(rPr, 'strike') && at(rPr, 'strike') !== 'noStrike') r.strike = at(rPr, 'strike');
    if (at(rPr, 'cap') && at(rPr, 'cap') !== 'none') r.cap = at(rPr, 'cap');
    if (at(rPr, 'spc')) r.spc = +at(rPr, 'spc') / 100;
    if (at(rPr, 'baseline') && +at(rPr, 'baseline')) r.base = +at(rPr, 'baseline');
    for (const c of rPr.children) {
      switch (c.localName) {
        case 'solidFill': { const v = colorOf(c, ctx); if (v) { r.color = v.c; if (v.a < 1) r.fill = { t: 'solid', c: v.c, a: v.a }; } break; }
        case 'gradFill': r.fill = fill(c, ctx); if (r.fill && r.fill.stops[0]) r.color = r.fill.stops[0].c; break;
        case 'noFill': r.fill = { t: 'none' }; break;
        case 'pattFill': { const f = fill(c, ctx); r.color = f.fg; break; }
        case 'ln': { const l = line(c, ctx); if (l && l.t !== 'none') r.ln = l; break; }
        case 'effectLst': { const s = shadow(c, ctx); if (s) r.shd = true; break; }
        case 'highlight': { const v = colorOf(c, ctx); if (v) r.hl = v.c; break; }
        case 'latin': { const f = fontName(at(c, 'typeface')); if (f) r.font = f; break; }
        case 'hlinkClick': { const l = ctx.link(c); if (l) r.link = l; break; }
        default: break;
      }
    }
    return r;
  }
  function spacing(el) {
    if (!el) return undefined;
    const p = kid(el, 'spcPct'), s = kid(el, 'spcPts');
    if (p) return { pct: num(p, 'val', 0) / 1000 };
    if (s) return { pts: num(s, 'val', 0) / 100 };
    return undefined;
  }
  /** a:pPr or a:lvlNpPr → level style object */
  function paraProps(pPr, ctx, withDef) {
    const o = {};
    if (!pPr) return o;
    if (at(pPr, 'algn')) o.algn = at(pPr, 'algn');
    if (at(pPr, 'marL') != null) o.marL = L.round(pt(+at(pPr, 'marL')), 3);
    if (at(pPr, 'indent') != null) o.indent = L.round(pt(+at(pPr, 'indent')), 3);
    const ls = spacing(kid(pPr, 'lnSpc')), sb = spacing(kid(pPr, 'spcBef')), sa = spacing(kid(pPr, 'spcAft'));
    if (ls) o.lnSpc = ls;
    if (sb) o.spcBef = sb;
    if (sa) o.spcAft = sa;
    let bu = null;
    const extra = {};
    for (const c of pPr.children) {
      switch (c.localName) {
        case 'buNone': bu = { t: 'none' }; break;
        case 'buChar': bu = { t: 'char', ch: at(c, 'char') || '•' }; break;
        case 'buAutoNum': bu = { t: 'num', scheme: at(c, 'type') || 'arabicPeriod', start: num(c, 'startAt', 1) }; break;
        case 'buBlip': bu = { t: 'char', ch: '■' }; break;
        case 'buClr': { const v = colorOf(c, ctx); if (v) extra.c = v.c; break; }
        case 'buSzPct': extra.sz = num(c, 'val', 100000) / 100000; break;
        case 'buFont': extra.font = at(c, 'typeface'); break;
        default: break;
      }
    }
    if (bu || Object.keys(extra).length) {
      o.bu = Object.assign({}, bu || {}, extra);
      if (bu && bu.t === 'char' && extra.font) o.bu.ch = L.mapSymbolChar(bu.ch, extra.font);
      if (o.bu.font && /wingdings|symbol/i.test(o.bu.font)) delete o.bu.font;
    }
    if (withDef) { const d = kid(pPr, 'defRPr'); if (d) { const r = runProps(d, ctx); delete r.link; if (Object.keys(r).length) o.rPr = r; } }
    return o;
  }
  function listStyle(lst, ctx) {
    if (!lst) return null;
    const out = {};
    let any = false;
    for (let i = 1; i <= 9; i++) {
      const lv = kid(lst, `lvl${i}pPr`);
      if (lv) { const p = paraProps(lv, ctx, true); if (Object.keys(p).length) { out[i - 1] = p; any = true; } }
    }
    return any ? out : null;
  }
  function mergeLst(...lsts) {
    let out = null;
    for (const l of lsts) {
      if (!l) continue;
      out = out || {};
      for (const k in l) out[k] = L.deepMerge(out[k] || {}, l[k]);
    }
    return out;
  }
  function paragraphs(txBody, ctx) {
    const out = [];
    for (const p of kids(txBody, 'p')) {
      const pPr = kid(p, 'pPr');
      const para = { lvl: pPr ? num(pPr, 'lvl', 0) : 0, pp: paraProps(pPr, ctx, false), rs: [] };
      for (const c of p.children) {
        if (c.localName === 'r') {
          const t = kid(c, 't');
          const rp = runProps(kid(c, 'rPr'), ctx);
          para.rs.push(Object.assign({ t: t ? t.textContent : '' }, rp));
        } else if (c.localName === 'br') {
          const rp = runProps(kid(c, 'rPr'), ctx);
          const last = para.rs[para.rs.length - 1];
          if (last && !last.fld && L.equal(L.txt.runProps(last), rp)) last.t += '\n';
          else para.rs.push(Object.assign({ t: '\n' }, rp));
        } else if (c.localName === 'fld') {
          const t = kid(c, 't');
          const type = at(c, 'type') || '';
          const rp = runProps(kid(c, 'rPr'), ctx);
          if (type === 'slidenum' || /^datetime/.test(type)) para.rs.push(Object.assign({ t: t ? t.textContent : '', fld: type }, rp));
          else para.rs.push(Object.assign({ t: t ? t.textContent : '' }, rp));
        } else if (c.localName === 'endParaRPr') {
          const rp = runProps(c, ctx);
          delete rp.link;
          if (Object.keys(rp).length) para.end = rp;
        }
      }
      L.txt.normalize(para);
      out.push(para);
    }
    return out.length ? out : [L.txt.para('')];
  }
  function bodyPr(bp, base) {
    const o = Object.assign({}, base || {});
    if (!bp) return o;
    const ins = (n, d) => (at(bp, n) != null ? L.round(pt(+at(bp, n)), 3) : d);
    const b0 = o.ins || [7.2, 3.6, 7.2, 3.6];
    o.ins = [ins('lIns', b0[0]), ins('tIns', b0[1]), ins('rIns', b0[2]), ins('bIns', b0[3])];
    if (at(bp, 'anchor')) o.anchor = { t: 't', ctr: 'ctr', b: 'b', just: 'ctr', dist: 'ctr' }[at(bp, 'anchor')] || 't';
    if (at(bp, 'anchorCtr') != null) o.anchorCtr = bool(at(bp, 'anchorCtr'));
    if (at(bp, 'wrap')) o.wrap = at(bp, 'wrap') !== 'none';
    if (at(bp, 'vert') && at(bp, 'vert') !== 'horz') o.vert = at(bp, 'vert') === 'vert270' ? 'vert270' : 'vert'; else if (at(bp, 'vert') === 'horz') delete o.vert;
    if (at(bp, 'numCol') && +at(bp, 'numCol') > 1) { o.cols = +at(bp, 'numCol'); o.colGap = pt(num(bp, 'spcCol', 0)); }
    if (at(bp, 'rot')) o.rot = num(bp, 'rot', 0) / 60000;
    const na = kid(bp, 'normAutofit'), sa = kid(bp, 'spAutoFit'), noa = kid(bp, 'noAutofit');
    if (na) { o.autofit = 'norm'; o.fontScale = at(na, 'fontScale') ? +at(na, 'fontScale') / 100000 : undefined; o.lnSpcRed = at(na, 'lnSpcReduction') ? +at(na, 'lnSpcReduction') / 100000 : undefined; if (!o.fontScale) delete o.fontScale; if (!o.lnSpcRed) delete o.lnSpcRed; }
    else if (sa) o.autofit = 'shape';
    else if (noa) o.autofit = 'none';
    const warp = kid(bp, 'prstTxWarp');
    if (warp && at(warp, 'prst') && at(warp, 'prst') !== 'textNoShape') o.warp = at(warp, 'prst');
    if (bool(at(bp, 'fromWordArt'))) o.fromWordArt = true;
    return o;
  }

  /* ---------- geometry ---------- */
  function xfrmOf(el) {
    if (!el) return null;
    const off = kid(el, 'off'), ext = kid(el, 'ext');
    const o = {
      x: pt(num(off, 'x', 0)), y: pt(num(off, 'y', 0)), w: pt(num(ext, 'cx', 0)), h: pt(num(ext, 'cy', 0)),
      rot: num(el, 'rot', 0) / 60000, flipH: bool(at(el, 'flipH')), flipV: bool(at(el, 'flipV')),
    };
    const co = kid(el, 'chOff'), ce = kid(el, 'chExt');
    if (co && ce) { o.chx = pt(num(co, 'x', 0)); o.chy = pt(num(co, 'y', 0)); o.chw = pt(num(ce, 'cx', 0)); o.chh = pt(num(ce, 'cy', 0)); }
    return o;
  }
  /* DrawingML shape-guide evaluator (custGeom formulas) */
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
      const f = (at(gd, 'fmla') || '').trim().split(/\s+/);
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
      env[at(gd, 'name')] = isFinite(r) ? r : 0;
    }
    return env;
  }
  function custGeom(cg, extW, extH) {
    const env = guideEnv(extW, extH);
    evalGuides(kids(kid(cg, 'avLst'), 'gd'), env);
    evalGuides(kids(kid(cg, 'gdLst'), 'gd'), env);
    const V = (tok) => { if (tok == null) return 0; if (tok in env) return env[tok]; const x = parseFloat(tok); return isNaN(x) ? 0 : x; };
    const paths = [];
    for (const p of kids(kid(cg, 'pathLst'), 'path')) {
      const pw = num(p, 'w', 0) || extW || 0, ph = num(p, 'h', 0) || extH || 0;
      const W = pt(pw), H = pt(ph);
      const cmds = [];
      const P = (e) => [L.round(pt(V(at(e, 'x'))), 3), L.round(pt(V(at(e, 'y'))), 3)];
      for (const c of p.children) {
        const pts = kids(c, 'pt').map(P);
        switch (c.localName) {
          case 'moveTo': if (pts[0]) cmds.push(['M', ...pts[0]]); break;
          case 'lnTo': if (pts[0]) cmds.push(['L', ...pts[0]]); break;
          case 'cubicBezTo': if (pts.length >= 3) cmds.push(['C', ...pts[0], ...pts[1], ...pts[2]]); break;
          case 'quadBezTo': if (pts.length >= 2) cmds.push(['Q', ...pts[0], ...pts[1]]); break;
          case 'arcTo': cmds.push(['A', L.round(pt(V(at(c, 'wR'))), 3), L.round(pt(V(at(c, 'hR'))), 3), V(at(c, 'stAng')) / 60000, V(at(c, 'swAng')) / 60000]); break;
          case 'close': cmds.push(['Z']); break;
          default: break;
        }
      }
      if (cmds.length && W > 0 && H > 0) paths.push({ w: W, h: H, cmds, fill: at(p, 'fill') === 'none' ? 'none' : 'norm', stroke: at(p, 'stroke') !== '0' && at(p, 'stroke') !== 'false' });
      else if (cmds.length) paths.push({ w: Math.max(W, 0.01), h: Math.max(H, 0.01), cmds, fill: at(p, 'fill') === 'none' ? 'none' : 'norm', stroke: at(p, 'stroke') !== '0' && at(p, 'stroke') !== 'false' });
    }
    return paths.length ? { paths } : null;
  }
  function geometry(spPr) {
    const pg = kid(spPr, 'prstGeom');
    const xe = spPr && path(spPr, 'xfrm', 'ext');
    const extW = xe ? num(xe, 'cx', 0) : 0, extH = xe ? num(xe, 'cy', 0) : 0;
    if (pg) {
      const adj = {};
      for (const gd of kids(kid(pg, 'avLst'), 'gd')) { const m = /val\s+(-?\d+)/.exec(at(gd, 'fmla') || ''); if (m) adj[at(gd, 'name')] = +m[1]; }
      return { geom: at(pg, 'prst') || 'rect', adj: Object.keys(adj).length ? adj : undefined };
    }
    const cg = kid(spPr, 'custGeom');
    if (cg) { const p = custGeom(cg, extW, extH); if (p) return { geom: 'custom', path: p }; }
    return { geom: 'rect' };
  }

  /* ---------- reader ---------- */
  async function read(buffer, opts) {
    opts = opts || {};
    const zip = await L.zip.read(buffer);
    const damaged = [];
    const textCache = new Map();
    const getText = async (p) => {
      if (textCache.has(p)) return textCache.get(p);
      let f = zip.get(p);
      if (!f) { try { f = zip.get(decodeURIComponent(p)); } catch (e) { f = null; } }
      let t = null;
      if (f) { try { t = await f.text(); } catch (e) { damaged.push(p); } }
      textCache.set(p, t);
      return t;
    };
    const xml = async (p) => { const t = await getText(p); if (!t) return null; try { return parse(t); } catch (e) { damaged.push(p); return null; } };
    const relsCache = new Map();
    const rels = async (p) => {
      if (relsCache.has(p)) return relsCache.get(p);
      const r = await xml(relsPath(p));
      const map = {};
      if (r) for (const e of kids(r, 'Relationship')) {
        const ext = at(e, 'TargetMode') === 'External';
        map[at(e, 'Id')] = { type: (at(e, 'Type') || '').split('/').pop(), target: ext ? at(e, 'Target') : resolvePath(p, at(e, 'Target')), external: ext };
      }
      relsCache.set(p, map);
      return map;
    };
    const mediaCache = new Map();
    const loadMedia = async (p) => {
      if (mediaCache.has(p)) return mediaCache.get(p);
      const f = zip.get(p);
      if (!f) { mediaCache.set(p, null); return null; }
      let bytes;
      try { bytes = await f.bytes(); } catch (e) { damaged.push(p); mediaCache.set(p, null); return null; }
      const ext = p.split('.').pop().toLowerCase();
      let view = null;
      if ((ext === 'wmf' || ext === 'emf') && L.metafile) view = await L.metafile.toPNG(bytes, ext);
      const id = L.media.add(new Blob([bytes], { type: L.extToMime(ext) }), p.split('/').pop(), view);
      mediaCache.set(p, id);
      return id;
    };

    const root = await rels('');
    const rootRel = Object.values(await (async () => { const r = await xml('_rels/.rels'); const m = {}; if (r) for (const e of kids(r, 'Relationship')) m[at(e, 'Id')] = { type: (at(e, 'Type') || '').split('/').pop(), target: resolvePath('', at(e, 'Target')) }; return m; })()).find((r) => r.type === 'officeDocument');
    void root;
    const presPath = rootRel ? rootRel.target : 'ppt/presentation.xml';
    const presX = await xml(presPath);
    /* table style definitions (PowerPoint embeds every style a table uses) */
    const tblStyles = new Map();
    {
      const pr = await rels(presPath);
      const tsRel = Object.values(pr || {}).find((r) => r.type === 'tableStyles');
      const tsX = await xml(tsRel ? tsRel.target : 'ppt/tableStyles.xml');
      if (tsX) for (const st of kids(tsX, 'tblStyle')) tblStyles.set((at(st, 'styleId') || '').toUpperCase(), st);
    }
    if (!presX || presX.localName !== 'presentation') throw new Error('This file does not contain a PowerPoint presentation.');
    const presRels = await rels(presPath);
    const sz = kid(presX, 'sldSz');
    const W = L.round(pt(num(sz, 'cx', 9144000)), 3), H = L.round(pt(num(sz, 'cy', 6858000)), 3);
    const pres = L.model.newPresentation({ w: W, h: H, empty: true });
    pres.designs = {};
    pres.firstNum = num(presX, 'firstSlideNum', 1);

    /* document properties */
    try {
      const core = await xml('docProps/core.xml');
      if (core) {
        const g = (n) => { const e = desc(core, n); return e ? e.textContent : ''; };
        Object.assign(pres.props, { title: g('title'), subject: g('subject'), author: g('creator'), keywords: g('keywords'), comments: g('description'), category: g('category'), created: g('created') || pres.props.created });
        pres.title = pres.props.title;
      }
      const app = await xml('docProps/app.xml');
      if (app) { const c = desc(app, 'Company'); if (c) pres.props.company = c.textContent; }
    } catch (e) { /* optional parts */ }
    /* show settings */
    try {
      const pp = Object.values(presRels).find((r) => r.type === 'presProps');
      if (pp) {
        const px = await xml(pp.target);
        const sp = px && kid(px, 'showPr');
        if (sp) {
          pres.show.loop = bool(at(sp, 'loop'));
          pres.show.noAnim = at(sp, 'showAnimation') === '0';
          pres.show.useTimings = at(sp, 'useTimings') !== '0';
          pres.show.kiosk = !!kid(sp, 'kiosk');
          const rg = kid(sp, 'sldRg');
          if (rg) { pres.show.from = num(rg, 'st', 0); pres.show.to = num(rg, 'end', 0); }
          const pc = colorOf(kid(sp, 'penClr'), { design: null });
          if (pc && pc.c[0] === '#') pres.show.penColor = pc.c;
        }
      }
    } catch (e) { /* ignore */ }

    /* ---- themes & masters ---- */
    const masters = new Map(); /* master path → info */
    const layouts = new Map(); /* layout path → info */
    async function loadTheme(p) {
      const t = await xml(p);
      const theme = { colors: Object.assign({}, L.model.COLOR_SCHEMES[0]), fonts: { major: 'Arial', minor: 'Arial' }, fills: [], lines: [], bgFills: [], name: t ? at(t, 'name') : 'Theme' };
      if (!t) return theme;
      const cs = desc(t, 'clrScheme');
      if (cs) for (const c of cs.children) { const v = color(colorEl(c), { design: null }); if (v) theme.colors[c.localName] = v.c[0] === '#' ? v.c : theme.colors[v.c] || '#000000'; }
      const fs = desc(t, 'fontScheme');
      if (fs) { const mj = path(fs, 'majorFont', 'latin'), mn = path(fs, 'minorFont', 'latin'); if (mj && at(mj, 'typeface')) theme.fonts.major = at(mj, 'typeface'); if (mn && at(mn, 'typeface')) theme.fonts.minor = at(mn, 'typeface'); }
      const fm = desc(t, 'fmtScheme');
      if (fm) {
        theme.fills = kids(kid(fm, 'fillStyleLst'));
        theme.lines = kids(kid(fm, 'lnStyleLst'));
        theme.bgFills = kids(kid(fm, 'bgFillStyleLst'));
      }
      return theme;
    }
    function clrMapOf(el) {
      if (!el) return null;
      const m = {};
      for (const k of ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) if (at(el, k)) m[k] = at(el, k);
      return m;
    }
    function bgOf(cSld, ctx, theme) {
      const bg = kid(cSld, 'bg');
      if (!bg) return undefined;
      const bp = kid(bg, 'bgPr');
      if (bp) return fillIn(bp, ctx);
      const br = kid(bg, 'bgRef');
      if (br) {
        const idx = num(br, 'idx', 0);
        const c = colorOf(br, ctx);
        const list = idx >= 1001 ? theme.bgFills : theme.fills;
        const el = list[(idx >= 1001 ? idx - 1001 : idx - 1)];
        if (el) return fill(el, Object.assign({}, ctx, { phClr: c }));
        if (c) return { t: 'solid', c: c.c, a: c.a };
      }
      return undefined;
    }
    function phOf(sp) {
      const nv = kid(sp, 'nvSpPr') || kid(sp, 'nvPicPr') || kid(sp, 'nvGraphicFramePr') || kid(sp, 'nvCxnSpPr');
      const ph = nv && path(nv, 'nvPr', 'ph');
      if (!ph) return null;
      return { type: at(ph, 'type') || 'obj', idx: num(ph, 'idx', 0) };
    }
    function collectPh(spTree) {
      const list = [];
      for (const sp of kids(spTree)) {
        const ph = phOf(sp);
        if (!ph) continue;
        const spPr = kid(sp, 'spPr');
        const txb = kid(sp, 'txBody');
        list.push({ ph, el: sp, xfrm: xfrmOf(spPr && kid(spPr, 'xfrm')), bodyPr: txb ? kid(txb, 'bodyPr') : null, lst: txb ? kid(txb, 'lstStyle') : null, spPr });
      }
      return list;
    }
    const titleLike = (t) => t === 'title' || t === 'ctrTitle';
    function matchPh(list, ph) {
      if (!list) return null;
      if (ph.idx) { const m = list.find((x) => x.ph.idx === ph.idx && (x.ph.type === ph.type || !titleLike(x.ph.type))); if (m) return m; }
      let m = list.find((x) => x.ph.type === ph.type);
      if (m) return m;
      if (ph.type === 'ctrTitle') m = list.find((x) => x.ph.type === 'title');
      else if (ph.type === 'title') m = list.find((x) => x.ph.type === 'ctrTitle');
      else if (['subTitle', 'obj', 'body', 'tbl', 'chart', 'pic', 'dgm', 'media', 'clipArt'].includes(ph.type)) m = list.find((x) => x.ph.type === 'body' || x.ph.type === 'obj');
      return m || null;
    }
    /** layout → master inheritance is by type only (idx values are unrelated between the two) */
    function matchMasterPh(list, ph) {
      if (!list || !ph) return null;
      const t = ph.type || 'obj';
      const want = t === 'ctrTitle' || t === 'title' ? ['title', 'ctrTitle'] : ['dt', 'ftr', 'sldNum', 'hdr'].includes(t) ? [t] : ['body', 'obj'];
      for (const w of want) { const m = list.find((x) => (x.ph.type || 'obj') === w); if (m) return m; }
      return null;
    }
    function levelsFrom(lstEl, ctx) {
      const out = [];
      if (!lstEl) return out;
      for (let i = 1; i <= 9; i++) { const lv = kid(lstEl, `lvl${i}pPr`); out.push(lv ? paraProps(lv, ctx, true) : {}); }
      return out;
    }
    async function loadMaster(p) {
      if (masters.has(p)) return masters.get(p);
      const mx = await xml(p);
      const mr = await rels(p);
      const themeRel = Object.values(mr).find((r) => r.type === 'theme');
      const theme = await loadTheme(themeRel ? themeRel.target : '');
      const design = L.model.buildDesign('default', W, H);
      design.name = theme.name || 'Imported Design';
      design.key = 'imported';
      design.colors = Object.assign({}, design.colors, theme.colors);
      design.fonts = theme.fonts;
      const cm = clrMapOf(kid(mx, 'clrMap'));
      if (cm) design.clrMap = cm;
      design.deco = [];
      design.titleDeco = null;
      const ctx = mkCtx(p, mr, design, theme);
      const cSld = kid(mx, 'cSld');
      design.bg = bgOf(cSld, ctx, theme) || { t: 'solid', c: 'bg1' };
      const tree = kid(cSld, 'spTree');
      const phs = collectPh(tree);
      for (const ph of phs) {
        const k = ph.ph.type;
        if (ph.xfrm && ['title', 'body', 'dt', 'ftr', 'sldNum'].includes(k)) design.ph[k] = { x: ph.xfrm.x, y: ph.xfrm.y, w: ph.xfrm.w, h: ph.xfrm.h };
        if (k === 'obj' && ph.xfrm && !phs.some((q) => q.ph.type === 'body')) design.ph.body = { x: ph.xfrm.x, y: ph.xfrm.y, w: ph.xfrm.w, h: ph.xfrm.h };
      }
      /* text styles */
      const txs = kid(mx, 'txStyles');
      const base = L.clone(L.model.DEFAULT_TX);
      const t = levelsFrom(kid(txs, 'titleStyle'), ctx), b = levelsFrom(kid(txs, 'bodyStyle'), ctx), o = levelsFrom(kid(txs, 'otherStyle'), ctx);
      design.tx = {
        title: [L.deepMerge(base.title[0], t[0] || {})],
        body: base.body.map((lv, i) => L.deepMerge(lv, b[i] || {})),
        other: base.other.map((lv, i) => L.deepMerge(lv, o[i] || {})),
      };
      /* title anchors from master placeholder bodyPr */
      design.deco = await shapesFrom(tree, ctx, { skipPh: true });
      const info = { design, theme, ctx, phs, mr, path: p, dft: null };
      masters.set(p, info);
      pres.designs[design.id] = design;
      return info;
    }
    async function loadLayout(p) {
      if (layouts.has(p)) return layouts.get(p);
      const lx = await xml(p);
      const lr = await rels(p);
      const mRel = Object.values(lr).find((r) => r.type === 'slideMaster');
      const master = await loadMaster(mRel ? mRel.target : '');
      const ctx = mkCtx(p, lr, master.design, master.theme);
      /* a layout may remap colours (e.g. a dark variant: bg1→dk1, tx1→lt1); slides on it inherit that */
      const lovr = path(lx, 'clrMapOvr', 'overrideClrMapping');
      if (lovr) ctx.clrMapOvr = ovrDiff(clrMapOf(lovr), master.design);
      const cSld = kid(lx, 'cSld');
      const tree = kid(cSld, 'spTree');
      const phs = collectPh(tree);
      const deco = await shapesFrom(tree, ctx, { skipPh: true });
      const bg = bgOf(cSld, ctx, master.theme);
      const showMaster = at(lx, 'showMasterSp') !== '0';
      const type = at(lx, 'type') || 'cust';
      const name = cSld ? at(cSld, 'name') : '';
      const info = { master, phs, deco, bg, showMaster, type, name, path: p, ctx, clrMapOvr: ctx.clrMapOvr };
      /* the title-slide layout provides ctrTitle/subTitle frames */
      for (const ph of phs) if ((ph.ph.type === 'ctrTitle' || ph.ph.type === 'subTitle') && ph.xfrm && !master.design._ctrSet) master.design.ph[ph.ph.type] = { x: ph.xfrm.x, y: ph.xfrm.y, w: ph.xfrm.w, h: ph.xfrm.h };
      if (phs.some((x) => x.ph.type === 'ctrTitle')) master.design._ctrSet = true;
      const d = master.design;
      if (deco.length || bg || !showMaster) {
        const key = p.split('/').pop().replace('.xml', '');
        d.layoutDecos = d.layoutDecos || {};
        d.layoutBgs = d.layoutBgs || {};
        d.layoutNames = d.layoutNames || {};
        d.layoutDecos[key] = (showMaster ? L.clone(d.deco) : []).concat(deco);
        if (bg) d.layoutBgs[key] = bg;
        d.layoutNames[key] = name;
        info.lkey = key;
      }
      layouts.set(p, info);
      return info;
    }

    /* per-part parsing context */
    function mkCtx(partPath, partRels, design, theme) {
      const ctx = {
        design, theme, partPath, rels: partRels, phClr: null, clrMapOvr: null, mediaQueue: [],
        media: (id) => { const r = partRels[id]; if (!r || r.external) return null; const m = mediaCache.get(r.target); if (m) return m; ctx.pendingMedia.add(r.target); return '__pending__:' + r.target; },
        link: (el) => {
          const action = at(el, 'action') || '';
          const id = rid(el);
          const tip = at(el, 'tooltip') || undefined;
          if (/hlinksldjump/.test(action)) { const r = partRels[id]; return r ? { slidePath: r.target } : null; }
          const jm = /jump=(\w+)/.exec(action);
          if (jm) return { action: { nextslide: 'next', previousslide: 'prev', firstslide: 'first', lastslide: 'last', lastslideviewed: 'lastViewed', endshow: 'end' }[jm[1]] || 'next' };
          if (/hlinkfile|program|macro|ole/.test(action)) return null;
          if (id && partRels[id] && partRels[id].external) return Object.assign({ url: partRels[id].target }, tip ? { tip } : {});
          return null;
        },
        pendingMedia: new Set(),
      };
      return ctx;
    }

    /** keep only the entries of a colour-map override that differ from the master's map */
    function ovrDiff(m, design) {
      if (!m) return null;
      const out = {};
      for (const k in m) if (m[k] !== (design.clrMap && design.clrMap[k] ? design.clrMap[k] : { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }[k] || k)) out[k] = m[k];
      return Object.keys(out).length ? out : null;
    }
    /* style refs (p:style) */
    function styleRefs(styleEl, ctx) {
      if (!styleEl) return {};
      const out = {};
      const lr = kid(styleEl, 'lnRef'), fr = kid(styleEl, 'fillRef'), er = kid(styleEl, 'effectRef'), fo = kid(styleEl, 'fontRef');
      if (lr && num(lr, 'idx', 0) > 0) {
        const c = colorOf(lr, ctx);
        const el = ctx.theme.lines[num(lr, 'idx', 1) - 1];
        out.line = el ? line(el, Object.assign({}, ctx, { phClr: c })) : c ? { c: c.c, w: 0.75 } : undefined;
      }
      if (fr && num(fr, 'idx', 0) > 0) {
        const c = colorOf(fr, ctx);
        const idx = num(fr, 'idx', 1);
        const el = idx >= 1001 ? ctx.theme.bgFills[idx - 1001] : ctx.theme.fills[idx - 1];
        out.fill = el ? fill(el, Object.assign({}, ctx, { phClr: c })) : c ? { t: 'solid', c: c.c, a: c.a } : undefined;
        if (out.fill && out.fill.t === 'grad' && c) out.fill = { t: 'solid', c: c.c, a: c.a };
      }
      void er;
      if (fo) {
        const c = colorOf(fo, ctx);
        out.fontColor = c ? c.c : undefined;
        out.font = at(fo, 'idx') === 'major' ? '+mj' : at(fo, 'idx') === 'minor' ? '+mn' : undefined;
      }
      return out;
    }

    /* ---- shapes ---- */
    async function shapesFrom(tree, ctx, opts) {
      opts = opts || {};
      const out = [];
      for (const el of kids(tree)) {
        try {
          const s = await shapeFrom(el, ctx, opts);
          if (!s) continue;
          if (Array.isArray(s)) out.push(...s); else out.push(s);
        } catch (e) { console.warn('shape skipped', e); }
      }
      return out;
    }
    /** resolve '__pending__:<part>' media references anywhere inside obj (also in copies made before loading) */
    async function finalizeMedia(ctx, obj) {
      for (const p of ctx.pendingMedia) await loadMedia(p);
      ctx.pendingMedia.clear();
      const refs = [];
      const scan = (o) => {
        if (!o || typeof o !== 'object') return;
        for (const k in o) {
          const v = o[k];
          if (typeof v === 'string' && v.startsWith('__pending__:')) refs.push([o, k, v.slice(12)]);
          else if (v && typeof v === 'object') scan(v);
        }
      };
      scan(obj);
      for (const [o, k, part] of refs) { if (!mediaCache.has(part)) await loadMedia(part); o[k] = mediaCache.get(part) || null; }
    }
    async function shapeFrom(el, ctx, opts) {
      const ln = el.localName;
      if (ln === 'AlternateContent') {
        const fb = kid(el, 'Fallback');
        const ch = kid(el, 'Choice');
        const src = fb && fb.children.length ? fb : ch;
        if (!src) return null;
        const res = [];
        for (const c of src.children) { const s = await shapeFrom(c, ctx, opts); if (s) Array.isArray(s) ? res.push(...s) : res.push(s); }
        return res;
      }
      if (!['sp', 'pic', 'cxnSp', 'grpSp', 'graphicFrame'].includes(ln)) return null;
      const ph = phOf(el);
      if (opts.skipPh && ph) return null;
      const nv = kids(el).find((c) => /^nv/.test(c.localName));
      const cNvPr = nv ? kid(nv, 'cNvPr') : null;
      const base = { id: L.uid('s'), name: at(cNvPr, 'name') || ln, numId: num(cNvPr, 'id', 0) };
      if (at(cNvPr, 'descr')) base.alt = at(cNvPr, 'descr');
      if (bool(at(cNvPr, 'hidden'))) base.hidden = true;
      const sl = cNvPr && kid(cNvPr, 'hlinkClick');
      if (sl) { const l = ctx.link(sl); if (l) base.link = l; }
      const map = ctx.map || ((x) => x);
      let sh = null;
      if (ln === 'grpSp') {
        const gp = kid(el, 'grpSpPr');
        const xf = xfrmOf(gp && kid(gp, 'xfrm'));
        const gfill = fillIn(gp, ctx);
        const g = Object.assign(base, { type: 'group', kids: [], rot: 0 });
        if (xf) {
          const p0 = map({ x: xf.x, y: xf.y, w: xf.w, h: xf.h });
          Object.assign(g, p0, { rot: xf.rot || 0 });
          const sx = xf.chw ? xf.w / xf.chw : 1, sy = xf.chh ? xf.h / xf.chh : 1;
          const inner = (b) => map({ x: xf.x + (b.x - (xf.chx || 0)) * sx, y: xf.y + (b.y - (xf.chy || 0)) * sy, w: b.w * sx, h: b.h * sy });
          const cctx = Object.assign({}, ctx, { map: inner, groupFill: gfill || ctx.groupFill, flipParent: xf.flipH || xf.flipV });
          g.kids = await shapesFrom(el, cctx, {});
          ctx.pendingMedia = cctx.pendingMedia;
        } else {
          g.kids = await shapesFrom(el, ctx, {});
          const b = L.model.groupBounds(g);
          Object.assign(g, b);
        }
        if (!g.kids.length) return null;
        return g;
      }
      const spPr = kid(el, 'spPr');
      let xf = xfrmOf(spPr && kid(spPr, 'xfrm'));
      if (ln === 'graphicFrame') xf = xfrmOf(kid(el, 'xfrm'));
      /* inherit placeholder frame & text styles from layout and master */
      let lay = null, mas = null;
      if (ph && ctx.layout) { lay = matchPh(ctx.layout.phs, ph); mas = matchMasterPh(ctx.layout.master.phs, lay ? lay.ph : ph); }
      if (!xf && lay && lay.xfrm) xf = Object.assign({}, lay.xfrm);
      if (!xf && mas && mas.xfrm) xf = Object.assign({}, mas.xfrm);
      if (!xf) {
        if (!ph) return null;
        const fr = ph.type === 'title' || ph.type === 'ctrTitle' ? ctx.design.ph.title : ctx.design.ph.body;
        xf = Object.assign({ rot: 0, flipH: false, flipV: false }, fr);
      }
      const geo = map({ x: xf.x, y: xf.y, w: xf.w, h: xf.h });
      Object.assign(base, geo, { rot: xf.rot || 0 });
      if (xf.flipH) base.flipH = true;
      if (xf.flipV) base.flipV = true;
      const refs = styleRefs(kid(el, 'style'), ctx);

      if (ln === 'pic') {
        const bf = kid(el, 'blipFill');
        const blip = bf && kid(bf, 'blip');
        const mid = blip ? ctx.media(rid(blip, 'embed')) || ctx.media(rid(blip, 'link')) : null;
        sh = Object.assign(base, { type: 'image', geom: 'rect', media: mid, crop: { l: 0, t: 0, r: 0, b: 0 }, lockAspect: true, img: {} });
        const sr = bf && kid(bf, 'srcRect');
        if (sr) sh.crop = { l: num(sr, 'l', 0) / 100000, t: num(sr, 't', 0) / 100000, r: num(sr, 'r', 0) / 100000, b: num(sr, 'b', 0) / 100000 };
        if (blip) {
          if (desc(blip, 'grayscl')) sh.img.mode = 'gray';
          if (desc(blip, 'biLevel')) sh.img.mode = 'bw';
          const lum = desc(blip, 'lum');
          if (lum) { const br = num(lum, 'bright', 0) / 100000, co = num(lum, 'contrast', 0) / 100000; if (br >= 0.6 && co <= -0.6) sh.img.mode = 'wash'; else { if (br) sh.img.bright = br; if (co) sh.img.contrast = co; } }
          const am = desc(blip, 'alphaModFix');
          if (am) sh.img.alpha = num(am, 'amt', 100000) / 100000;
        }
        const g = geometry(spPr);
        sh.geom = g.geom; if (g.adj) sh.adj = g.adj; if (g.path) sh.path = g.path;
        sh.line = line(spPr && kid(spPr, 'ln'), ctx, refs.line ? refs.line : { t: 'none' }) || { t: 'none' };
        const ef = spPr && kid(spPr, 'effectLst');
        const shd = shadow(ef, ctx); if (shd) sh.shadow = shd;
        if (ph) sh.ph = ph;
        if (!mid) sh.missingLabel = 'Picture';
        return sh;
      }
      if (ln === 'graphicFrame') {
        const gd = path(el, 'graphic', 'graphicData');
        const uri = at(gd, 'uri') || '';
        if (/table/.test(uri)) return tableFrom(desc(gd, 'tbl'), base, ctx);
        if (/chart/.test(uri)) {
          const c = kids(gd).find((x) => x.localName === 'chart');
          const r = c ? ctx.rels[rid(c)] : null;
          const model = r ? await chartFrom(r.target, ctx) : null;
          return Object.assign(base, { type: 'chart', chart: model || L.chart.sample() });
        }
        if (/diagram/.test(uri)) {
          const g = await smartArtFrom(gd, base, ctx);
          if (g) return g;
        }
        /* OLE object or unknown: try fallback picture */
        const pic = desc(el, 'pic');
        if (pic) {
          const bl = desc(pic, 'blip');
          const mid = bl ? ctx.media(rid(bl, 'embed')) : null;
          if (mid) return Object.assign(base, { type: 'image', geom: 'rect', media: mid, crop: { l: 0, t: 0, r: 0, b: 0 }, line: { t: 'none' }, lockAspect: true, img: {} });
        }
        /* 2007-era OLE objects keep their preview picture in the legacy VML drawing, keyed by spid */
        const ole = desc(el, 'oleObj');
        const spid = ole && at(ole, 'spid');
        if (spid && ctx.rels) {
          const vr = Object.values(ctx.rels).find((r) => r.type === 'vmlDrawing' && !r.external);
          const vml = vr ? await getText(vr.target) : null;
          if (vml) {
            const re = new RegExp('<v:shape\\b[^>]*o:spid="' + spid.replace(/[^\w]/g, '') + '"[\\s\\S]*?</v:shape>');
            const m = re.exec(vml);
            const relid = m && /<v:imagedata\b[^>]*o:relid="([^"]+)"/.exec(m[0]);
            if (relid) {
              const vrels = await rels(vr.target);
              const ir = vrels[relid[1]];
              const mid = ir && !ir.external ? await loadMedia(ir.target) : null;
              if (mid) return Object.assign(base, { type: 'image', geom: 'rect', media: mid, crop: { l: 0, t: 0, r: 0, b: 0 }, line: { t: 'none' }, lockAspect: true, img: {} });
            }
          }
        }
        return Object.assign(base, { type: 'shape', geom: 'rect', fill: { t: 'solid', c: '#F2F2F2', a: 1 }, line: { c: '#A6A6A6', w: 0.75, dash: 'dash' }, tx: L.txt.body([L.txt.para('Embedded object', { algn: 'ctr' }, { sz: 12, color: '#7F7F7F' })], { anchor: 'ctr' }) });
      }
      /* sp & cxnSp */
      const g = geometry(spPr);
      const isLine = ln === 'cxnSp' || ['line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5', 'curvedConnector2', 'curvedConnector3', 'curvedConnector4', 'curvedConnector5'].includes(g.geom);
      if (isLine) {
        const lnEl = spPr && kid(spPr, 'ln');
        sh = Object.assign(base, { type: 'line', geom: /^bent/.test(g.geom) ? 'bentConnector3' : /^curved/.test(g.geom) ? 'curvedConnector3' : g.geom === 'straightConnector1' ? 'straightConnector1' : 'line' });
        sh.line = line(lnEl, ctx, refs.line || { c: 'tx1', w: 0.75, dash: 'solid' }) || { c: 'tx1', w: 0.75 };
        if (sh.rot) {
          /* express rotated lines as plain endpoints */
          const [a, b] = L.geom.lineEnds(sh);
          const cx = sh.x + sh.w / 2, cy = sh.y + sh.h / 2;
          const p1 = L.rotPt(a[0], a[1], cx, cy, sh.rot), p2 = L.rotPt(b[0], b[1], cx, cy, sh.rot);
          sh.rot = 0;
          L.geom.setLineEnds(sh, p1, p2);
        }
        const ef = spPr && kid(spPr, 'effectLst');
        const shd = shadow(ef, ctx); if (shd) sh.shadow = shd;
        return sh;
      }
      const txBody = kid(el, 'txBody');
      const isTxBox = !!(nv && kid(nv, 'cNvSpPr') && bool(at(kid(nv, 'cNvSpPr'), 'txBox')));
      sh = Object.assign(base, { type: ph || isTxBox ? 'text' : 'shape', geom: g.geom });
      if (g.adj) sh.adj = g.adj;
      if (g.path) sh.path = g.path;
      /* precedence: own spPr > own p:style > layout placeholder > master placeholder */
      let f = fillIn(spPr, ctx);
      if (f === undefined && refs.fill) f = refs.fill;
      if (f === undefined && ph && lay && lay.spPr) f = fillIn(lay.spPr, ctx);
      if (f === undefined && ph && mas && mas.spPr) f = fillIn(mas.spPr, ctx);
      sh.fill = f || { t: 'none' };
      let lnBase = { t: 'none' };
      if (ph && mas && mas.spPr && kid(mas.spPr, 'ln')) lnBase = line(kid(mas.spPr, 'ln'), ctx, lnBase) || lnBase;
      if (ph && lay && lay.spPr && kid(lay.spPr, 'ln')) lnBase = line(kid(lay.spPr, 'ln'), ctx, lnBase) || lnBase;
      if (refs.line) lnBase = refs.line;
      sh.line = line(spPr && kid(spPr, 'ln'), ctx, lnBase) || { t: 'none' };
      const ef = spPr && kid(spPr, 'effectLst');
      const shd = shadow(ef, ctx); if (shd) sh.shadow = shd;
      if (ph) sh.ph = ph;
      if (txBody || ph) {
        let bp = {};
        if (mas && mas.bodyPr) bp = bodyPr(mas.bodyPr, bp);
        if (lay && lay.bodyPr) bp = bodyPr(lay.bodyPr, bp);
        bp = bodyPr(txBody ? kid(txBody, 'bodyPr') : null, bp);
        const tx = Object.assign({ anchor: 't', wrap: true, autofit: 'none', ins: [7.2, 3.6, 7.2, 3.6] }, bp);
        tx.ps = txBody ? paragraphs(txBody, ctx) : [L.txt.para('')];
        /* p:style fontRef sits above inherited placeholder styles but below the shape's own list style */
        let refLst = null;
        if (refs.fontColor || refs.font) { refLst = {}; for (let i = 0; i < 9; i++) refLst[i] = { rPr: Object.assign({}, refs.fontColor ? { color: refs.fontColor } : {}, refs.font ? { font: refs.font } : {}) }; }
        const lstX = mergeLst(mas ? listStyle(mas.lst, ctx) : null, lay ? listStyle(lay.lst, ctx) : null, refLst, txBody ? listStyle(kid(txBody, 'lstStyle'), ctx) : null);
        if (lstX) tx.lst = lstX;
        /* colour-map override (dark layouts): master text colours that name a remapped slot follow the remap */
        if (ctx.clrMapOvr) {
          const cls = L.style.cls(ph ? { ph, type: 'text' } : null);
          const lv = (ctx.design.tx && ctx.design.tx[cls]) || [];
          for (let i = 0; i < 9; i++) {
            const own = tx.lst && tx.lst[i] && tx.lst[i].rPr && tx.lst[i].rPr.color;
            const st = lv[Math.min(i, lv.length - 1)];
            const base = own || (st && st.rPr && st.rPr.color);
            if (base && ctx.clrMapOvr[base]) { tx.lst = tx.lst || {}; tx.lst[i] = Object.assign({}, tx.lst[i]); tx.lst[i].rPr = Object.assign({}, tx.lst[i].rPr, { color: ctx.clrMapOvr[base] }); }
          }
        }
        delete tx.fromWordArt;
        const warp = tx.warp;
        delete tx.warp;
        /* placeholder frames: title anchors default to the master */
        sh.tx = tx;
        if (warp && bp.fromWordArt !== false && tx.ps.length <= 3 && !ph) {
          /* WordArt */
          const r0 = (tx.ps[0] && tx.ps[0].rs[0]) || {};
          const fl = r0.fill || (r0.color ? { t: 'solid', c: r0.color, a: 1 } : { t: 'solid', c: 'accent1', a: 1 });
          sh.type = 'wordart';
          sh.wa = { text: L.txt.plain(tx), font: L.style.font(r0.font || (lstX && lstX[0] && lstX[0].rPr && lstX[0].rPr.font) || '+mj', ctx.design), b: !!r0.b, i: !!r0.i, fill: fl, line: r0.ln || { t: 'none' }, shadow: r0.shd ? { c: '#000000', a: 0.4, dx: 2, dy: 2, blur: 2 } : null, warp, algn: (tx.ps[0] && tx.ps[0].pp && tx.ps[0].pp.algn) || 'ctr' };
          delete sh.tx;
          sh.fill = { t: 'none' }; sh.line = { t: 'none' };
          return sh;
        }
      }
      return sh;
    }

    /* ---- table styles ---- */
    const MED2 = { '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}': 'accent1', '{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}': 'accent2', '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}': 'accent3', '{00A15C55-8517-42AA-B614-E9B94910E393}': 'accent4', '{7DF18680-E054-41AD-8BC1-D1AEF772440D}': 'accent5', '{93296810-A885-4BE3-A3E7-6D5BEEA58F35}': 'accent6', '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}': 'dk1' };
    /** a style part → { fill, bd: {left,right,top,bottom,insideH,insideV}, tx: {b,i,color} } */
    function stylePart(el, ctx) {
      if (!el) return null;
      const out = { bd: {}, tx: {} };
      const ts = kid(el, 'tcTxStyle');
      if (ts) {
        if (at(ts, 'b')) out.tx.b = at(ts, 'b') === 'on';
        if (at(ts, 'i')) out.tx.i = at(ts, 'i') === 'on';
        const c = colorOf(ts, ctx) || (kid(ts, 'fontRef') && colorOf(kid(ts, 'fontRef'), ctx));
        if (c) out.tx.color = c.c;
      }
      const tcs = kid(el, 'tcStyle');
      if (tcs) {
        const bdr = kid(tcs, 'tcBdr');
        if (bdr) for (const side of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) {
          const e = kid(bdr, side);
          if (!e) continue;
          const ln = kid(e, 'ln'), lr = kid(e, 'lnRef');
          if (ln) out.bd[side] = line(ln, ctx, { c: 'tx1', w: 1 });
          else if (lr) { const c = colorOf(lr, ctx); const th = ctx.theme.lines[num(lr, 'idx', 1) - 1]; out.bd[side] = th ? line(th, Object.assign({}, ctx, { phClr: c })) : { c: c ? c.c : 'tx1', w: 1 }; }
        }
        const f = fillIn(kid(tcs, 'fill') || tcs, ctx);
        if (f) out.fill = f;
        else { const fr = kid(tcs, 'fillRef'); if (fr) { const c = colorOf(fr, ctx); out.fill = c ? { t: 'solid', c: c.c, a: c.a } : undefined; } }
      }
      return out;
    }
    function builtinStyle(id, ctx) {
      const tint = (k, t) => L.color.applyMods(L.model.resolveColor(k, ctx.design), [{ name: 'tint', val: t }]);
      if (id === '{5940675A-B579-460E-94D1-54222C63F5DA}') { /* No Style, Table Grid */
        const b = { c: 'tx1', w: 1, dash: 'solid' };
        return { wholeTbl: { bd: { left: b, right: b, top: b, bottom: b, insideH: b, insideV: b }, tx: { color: 'tx1' } } };
      }
      if (id === '{2D5ABB26-0587-4C30-8999-92F81FD0307C}') return { wholeTbl: { bd: {}, tx: { color: 'tx1' } } };
      const a = MED2[id] || 'accent1';
      const w = { c: 'lt1', w: 1, dash: 'solid' }, thick = { c: 'lt1', w: 3, dash: 'solid' };
      const strong = { fill: { t: 'solid', c: a, a: 1 }, bd: {}, tx: { b: true, color: 'lt1' } };
      return {
        wholeTbl: { fill: { t: 'solid', c: tint(a, 20000), a: 1 }, bd: { left: w, right: w, top: w, bottom: w, insideH: w, insideV: w }, tx: { color: 'dk1' } },
        band1H: { fill: { t: 'solid', c: tint(a, 40000), a: 1 }, bd: {}, tx: {} },
        band1V: { fill: { t: 'solid', c: tint(a, 40000), a: 1 }, bd: {}, tx: {} },
        firstCol: strong, lastCol: strong,
        lastRow: Object.assign({}, strong, { bd: { top: thick } }),
        firstRow: Object.assign({}, strong, { bd: { bottom: thick } }),
      };
    }
    const styleCache = new Map();
    function tableStyleParts(id, ctx) {
      const key = id + '|' + (ctx.design && ctx.design.id);
      if (styleCache.has(key)) return styleCache.get(key);
      let parts = null;
      const el = tblStyles.get(id);
      if (el) {
        parts = {};
        for (const name of ['wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'firstCol', 'lastCol', 'firstRow', 'lastRow', 'seCell', 'swCell', 'neCell', 'nwCell']) { const p2 = stylePart(kid(el, name), ctx); if (p2) parts[name] = p2; }
      } else parts = builtinStyle(id, ctx);
      styleCache.set(key, parts);
      return parts;
    }
    function applyTableStyle(rows, nCols, id, fl, ctx) {
      const parts = tableStyleParts(id, ctx);
      if (!parts) return;
      const nRows = rows.length;
      const r0 = fl.firstRow ? 1 : 0, r1 = nRows - 1 - (fl.lastRow ? 1 : 0);
      const c0 = fl.firstCol ? 1 : 0, c1 = nCols - 1 - (fl.lastCol ? 1 : 0);
      for (let r = 0; r < nRows; r++) for (let c = 0; c < nCols; c++) {
        const cell = rows[r].cells[c];
        if (!cell) continue;
        /* parts in increasing precedence, each with the region it covers */
        const apply = [];
        apply.push(['wholeTbl', 0, nRows - 1, 0, nCols - 1]);
        if (fl.bandCol && c >= c0 && c <= c1) apply.push([(c - c0) % 2 === 0 ? 'band1V' : 'band2V', 0, nRows - 1, c, c]);
        if (fl.bandRow && r >= r0 && r <= r1) apply.push([(r - r0) % 2 === 0 ? 'band1H' : 'band2H', r, r, 0, nCols - 1]);
        if (fl.lastCol && c === nCols - 1) apply.push(['lastCol', 0, nRows - 1, c, c]);
        if (fl.firstCol && c === 0) apply.push(['firstCol', 0, nRows - 1, 0, 0]);
        if (fl.lastRow && r === nRows - 1) apply.push(['lastRow', r, r, 0, nCols - 1]);
        if (fl.firstRow && r === 0) apply.push(['firstRow', 0, 0, 0, nCols - 1]);
        let fill = null; const bd = {}; const tx = {};
        for (const [name, ra, rb, ca, cb] of apply) {
          const pt2 = parts[name];
          if (!pt2) continue;
          if (pt2.fill) fill = pt2.fill;
          Object.assign(tx, pt2.tx);
          const pick = (edgeOuter, outer, inner) => (edgeOuter ? pt2.bd[outer] : pt2.bd[inner]);
          const L2 = pick(c === ca, 'left', 'insideV'), R2 = pick(c === cb, 'right', 'insideV'), T2 = pick(r === ra, 'top', 'insideH'), B2 = pick(r === rb, 'bottom', 'insideH');
          if (L2) bd.l = L2; if (R2) bd.r = R2; if (T2) bd.t = T2; if (B2) bd.b = B2;
        }
        if (!cell.own.fill && fill) cell.fill = L.clone(fill);
        for (const k of ['l', 'r', 't', 'b']) if (!cell.own.bd.includes(k) && bd[k]) cell.bd[k] = L.clone(bd[k]);
        if (tx.color || tx.b != null || tx.i != null) {
          for (const p of cell.tx.ps) {
            for (const run of p.rs) { if (run.color == null && tx.color) run.color = tx.color; if (run.b == null && tx.b != null) run.b = tx.b; if (run.i == null && tx.i != null) run.i = tx.i; }
            const e = Object.assign({}, p.end || {});
            if (e.color == null && tx.color) e.color = tx.color; if (e.b == null && tx.b != null) e.b = tx.b;
            p.end = e;
          }
        }
      }
    }
    function tableFrom(tbl, base, ctx) {
      if (!tbl) return null;
      const tp = kid(tbl, 'tblPr');
      const styled = !!(tp && kid(tp, 'tableStyleId'));
      const firstRow = bool(at(tp, 'firstRow')), bandRow = bool(at(tp, 'bandRow')), firstCol = bool(at(tp, 'firstCol'));
      const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => L.round(pt(num(g, 'w', 0)), 3));
      const rows = kids(tbl, 'tr').map((tr, ri) => ({
        h: L.round(pt(num(tr, 'h', 0)), 3),
        cells: kids(tr, 'tc').map((tc, ci) => {
          const txb = kid(tc, 'txBody');
          const pr = kid(tc, 'tcPr');
          const tx = Object.assign({ anchor: 't', wrap: true, autofit: 'none', ins: [7.2, 3.6, 7.2, 3.6] }, bodyPr(txb ? kid(txb, 'bodyPr') : null, {}));
          if (pr) {
            tx.ins = [pr.hasAttribute('marL') ? pt(num(pr, 'marL', 91440)) : 7.2, pr.hasAttribute('marT') ? pt(num(pr, 'marT', 45720)) : 3.6, pr.hasAttribute('marR') ? pt(num(pr, 'marR', 91440)) : 7.2, pr.hasAttribute('marB') ? pt(num(pr, 'marB', 45720)) : 3.6].map((v) => L.round(v, 3));
            if (at(pr, 'anchor')) tx.anchor = { t: 't', ctr: 'ctr', b: 'b' }[at(pr, 'anchor')] || 't';
          }
          tx.ps = txb ? paragraphs(txb, ctx) : [L.txt.para('')];
          const cell = { tx, fill: { t: 'none' }, bd: {} };
          const f = fillIn(pr, ctx);
          if (f) cell.fill = f;
          for (const [k, tag] of [['l', 'lnL'], ['r', 'lnR'], ['t', 'lnT'], ['b', 'lnB']]) { const e = pr && kid(pr, tag); if (e) cell.bd[k] = line(e, ctx, { c: 'tx1', w: 1 }); }
          if (num(tc, 'gridSpan', 1) > 1) cell.gs = num(tc, 'gridSpan', 1);
          if (num(tc, 'rowSpan', 1) > 1) cell.rs = num(tc, 'rowSpan', 1);
          if (bool(at(tc, 'hMerge'))) cell.hm = true;
          if (bool(at(tc, 'vMerge'))) cell.vm = true;
          cell.own = { fill: !!f, bd: Object.keys(cell.bd) };
          return cell;
        }),
      }));
      const lastRow = bool(at(tp, 'lastRow')), lastCol = bool(at(tp, 'lastCol')), bandCol = bool(at(tp, 'bandCol'));
      if (styled) applyTableStyle(rows, cols.length, (at(kid(tp, 'tableStyleId'), 'val') || kid(tp, 'tableStyleId').textContent || '').trim().toUpperCase(), { firstRow, lastRow, firstCol, lastCol, bandRow, bandCol }, ctx);
      for (const r of rows) for (const c of r.cells) delete c.own;
      const t = Object.assign(base, { type: 'table', tbl: { cols, rows, firstRow, bandRow } });
      const sumW = cols.reduce((a, b) => a + b, 0);
      if (sumW > 0) t.w = sumW;
      const sumH = rows.reduce((a, r) => a + r.h, 0);
      if (sumH > 0) t.h = Math.max(t.h, sumH);
      return t;
    }

    async function chartFrom(p, ctx) {
      const cx = await xml(p);
      if (!cx) return null;
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
      if (kind === 'lineMarkers' && /line/.test(t)) {
        const s0 = kid(ct, 'ser');
        const mk = s0 && path(s0, 'marker', 'symbol');
        if (mk && at(mk, 'val') === 'none') kind = 'line';
      }
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
      const pal = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((k) => L.model.resolveColor(k, ctx.design));
      /* combo charts: series from the other chart groups are overlaid as lines / markers */
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
        if (sf) { const cc = colorOf(sf, Object.assign({}, ctx, { keepScheme: false })); if (cc) col = cc.c; }
        else if (sp && kid(sp, 'ln') && kid(kid(sp, 'ln'), 'solidFill')) { const cc = colorOf(kid(kid(sp, 'ln'), 'solidFill'), Object.assign({}, ctx, { keepScheme: false })); if (cc) col = cc.c; }
        const out = { name, vals, color: col || pal[i % pal.length] };
        if (s.__group) {
          const gt = s.__group.localName;
          const noLine = sp && kid(sp, 'ln') && kid(kid(sp, 'ln'), 'noFill');
          const mk = path(s, 'marker', 'symbol');
          const mkCol = path(s, 'marker', 'spPr', 'solidFill');
          if (mkCol) { const cc = colorOf(mkCol, Object.assign({}, ctx, { keepScheme: false })); if (cc) out.color = cc.c; }
          out.overlay = /line|scatter|radar/.test(gt) ? (noLine ? 'markers' : mk && at(mk, 'val') === 'none' ? 'line' : 'lineMarkers') : 'line';
        }
        return out;
      });
      if (!cats.length && series[0]) cats = series[0].vals.map((_, i) => String(i + 1));
      const model = { kind, title: '', legend: 'none', gridY: !!desc(plot, 'majorGridlines'), labels: false, fsz: 10, cats, series };
      const title = kid(chart, 'title');
      /* an explicit (possibly empty) rich title wins; only a title without text falls back to the series name */
      if (title) model.title = kid(title, 'tx') ? descAll(title, 't').map((x) => x.textContent).join('').trim() : series.length === 1 ? series[0].name : '';
      const leg = kid(chart, 'legend');
      if (leg) model.legend = at(kid(leg, 'legendPos'), 'val') || 'r';
      const dl = desc(ct, 'dLbls');
      if (dl && (at(kid(dl, 'showVal'), 'val') === '1' || at(kid(dl, 'showPercent'), 'val') === '1')) model.labels = true;
      /* number formats: data labels, value axis; category label rotation */
      const fmtOf = (el) => { const nf = el && kid(el, 'numFmt'); return nf && at(nf, 'formatCode') && at(nf, 'formatCode') !== 'General' ? at(nf, 'formatCode') : null; };
      const cacheFmt = (() => { const s0 = kid(ct, 'ser'); const fc = s0 && desc(kid(s0, 'val') || kid(s0, 'yVal'), 'formatCode'); return fc && fc.textContent !== 'General' ? fc.textContent : null; })();
      const lblFmt = fmtOf(dl) || cacheFmt;
      if (lblFmt) model.numFmt = lblFmt;
      const valAx = kid(plot, 'valAx'), catAx = kid(plot, 'catAx') || kid(plot, 'dateAx');
      const axFmt = fmtOf(valAx) || (valAx && kid(valAx, 'numFmt') && at(kid(valAx, 'numFmt'), 'sourceLinked') === '1' ? cacheFmt : null);
      if (axFmt) model.axisFmt = axFmt;
      const rotEl = catAx && path(catAx, 'txPr', 'bodyPr');
      if (rotEl && at(rotEl, 'rot') && +at(rotEl, 'rot')) model.catRot = +at(rotEl, 'rot') / 60000;
      if ((kind === 'pie' || kind === 'doughnut') && sers[0]) {
        model.pieColors = cats.map((_, k) => {
          const dp = kids(sers[0], 'dPt').find((d) => num(kid(d, 'idx'), 'val', -1) === k);
          const sf = dp && path(dp, 'spPr', 'solidFill');
          if (sf) { const cc = colorOf(sf, Object.assign({}, ctx, { keepScheme: false })); if (cc) return cc.c; }
          return pal[k % pal.length];
        });
      }
      const txPr = desc(cx, 'defRPr');
      if (txPr && at(txPr, 'sz')) model.fsz = +at(txPr, 'sz') / 100;
      return model;
    }

    async function smartArtFrom(gd, base, ctx) {
      const relIds = kids(gd).find((c) => c.localName === 'relIds');
      if (!relIds) return null;
      const dm = ctx.rels[rid(relIds, 'dm')];
      let drawingPath = null;
      if (dm) {
        const dx = await xml(dm.target);
        const ext = dx && desc(dx, 'dataModelExt');
        if (ext && at(ext, 'relId') && ctx.rels[at(ext, 'relId')]) drawingPath = ctx.rels[at(ext, 'relId')].target;
      }
      if (!drawingPath) { const r = Object.values(ctx.rels).find((x) => x.type === 'diagramDrawing'); if (r) drawingPath = r.target; }
      if (!drawingPath) return null;
      const dx = await xml(drawingPath);
      const tree = dx && desc(dx, 'spTree');
      if (!tree) return null;
      const dRels = await rels(drawingPath);
      const dctx = Object.assign({}, ctx, { rels: dRels, partPath: drawingPath, map: (b) => ({ x: base.x + b.x, y: base.y + b.y, w: b.w, h: b.h }), layout: null, pendingMedia: ctx.pendingMedia });
      dctx.media = (id) => { const r = dRels[id]; if (!r || r.external) return null; const m = mediaCache.get(r.target); if (m) return m; ctx.pendingMedia.add(r.target); return '__pending__:' + r.target; };
      const kidsS = await shapesFrom(tree, dctx, {});
      /* dsp shapes may carry a separate text frame */
      for (const sp of kids(tree, 'sp')) void sp;
      if (!kidsS.length) return null;
      return Object.assign(base, { type: 'group', kids: kidsS, name: base.name || 'Diagram' });
    }

    /* ---- slides ---- */
    const sldIds = kids(kid(presX, 'sldIdLst'), 'sldId');
    const slidePaths = sldIds.map((s) => presRels[rid(s)]).filter(Boolean).map((r) => r.target);
    const pathToId = new Map();
    const LAYOUT_MAP = { title: 'title', titleOnly: 'titleOnly', tx: 'text', twoColTx: 'twoText', blank: 'blank', objOnly: 'contentOnly', obj: 'content', twoObj: 'twoContent', fourObj: 'fourContent', txAndObj: 'textContent', objAndTx: 'contentText', txAndTwoObj: 'textTwoContent', twoObjAndTx: 'twoContentText', tbl: 'table', chart: 'chart', txOverObj: 'textOverContent', objOverTx: 'contentOverText', vertTx: 'vertText', vertTitleAndTx: 'vertTitleText', objAndTwoObj: 'contentTwoContent', twoObjAndObj: 'twoContentContent', secHead: 'titleOnly', twoTxTwoObj: 'twoContent', objTx: 'textContent', picTx: 'textContent', txAndChart: 'textContent', chartAndTx: 'contentText', txAndClipArt: 'textContent', clipArtAndTx: 'contentText', txAndMedia: 'textContent', mediaAndTx: 'contentText', dgm: 'content', txOverObj2: 'textOverContent' };
    let n = 0;
    for (const sp of slidePaths) {
      n++;
      if (opts.progress) opts.progress(n, slidePaths.length);
      const sx = await xml(sp);
      if (!sx) { if (zip.get(sp)) damaged.push(sp); continue; }
      try {
      const sr = await rels(sp);
      const layRel = Object.values(sr).find((r) => r.type === 'slideLayout');
      const lay = layRel ? await loadLayout(layRel.target) : null;
      const mRel = Object.values(presRels).find((r) => r.type === 'slideMaster');
      const master = lay ? lay.master : mRel ? await loadMaster(mRel.target) : null;
      if (!master) throw new Error('slide has no master');
      const design = master.design;
      const ctx = mkCtx(sp, sr, design, master.theme);
      ctx.layout = lay;
      const ovr = path(sx, 'clrMapOvr', 'overrideClrMapping');
      if (ovr) ctx.clrMapOvr = ovrDiff(clrMapOf(ovr), design);
      else if (lay && lay.clrMapOvr) ctx.clrMapOvr = lay.clrMapOvr;
      const cSld = kid(sx, 'cSld');
      const slide = {
        id: L.uid('sl'), layout: 'blank', design: design.id, shapes: [], notes: '',
        trans: { type: 'none', spd: 'fast', click: true, after: null }, anims: [], hidden: at(sx, 'show') === '0',
      };
      if (lay) {
        slide.layout = LAYOUT_MAP[lay.type] || (lay.phs.some((x) => x.ph.type === 'ctrTitle') ? 'title' : lay.phs.some((x) => x.ph.type === 'title') ? (lay.phs.length > 1 ? 'text' : 'titleOnly') : 'blank');
        if (lay.lkey) slide.lkey = lay.lkey;
      }
      if (at(sx, 'showMasterSp') === '0') slide.hideMaster = true;
      const bg = bgOf(cSld, ctx, master.theme);
      if (bg) slide.bg = bg;
      slide.shapes = await shapesFrom(kid(cSld, 'spTree'), ctx, {});
      await finalizeMedia(ctx, slide);
      /* slide-level text that is a dt/ftr/sldNum placeholder stays as shapes */
      /* notes */
      const nr = Object.values(sr).find((r) => r.type === 'notesSlide');
      if (nr) {
        const nx = await xml(nr.target);
        if (nx) {
          const body = kids(desc(nx, 'spTree')).find((s) => { const ph = phOf(s); return ph && (ph.type === 'body' || (ph.type === 'obj' && ph.idx === 1)); });
          if (body) { const tb = kid(body, 'txBody'); if (tb) slide.notes = kids(tb, 'p').map((p) => descAll(p, 't').map((t) => t.textContent).join('')).join('\n').replace(/\s+$/, ''); }
        }
      }
      /* transition */
      slide.trans = transitionFrom(sx) || slide.trans;
      /* animations */
      try { slide.anims = timingFrom(kid(sx, 'timing'), slide); } catch (e) { console.warn('timing skipped', e); slide.anims = []; }
      pathToId.set(sp, slide.id);
      pres.slides.push(slide);
      } catch (e) {
        /* keep going: one unreadable slide must not lose the whole deck */
        console.warn('slide skipped', sp, e);
        damaged.push(sp);
      }
    }
    /* resolve slide-jump hyperlinks and strip numeric ids */
    const fixLinks = (o) => {
      if (!o || typeof o !== 'object') return;
      if (o.link && o.link.slidePath) { const id = pathToId.get(o.link.slidePath); o.link = id ? { slide: id } : undefined; if (!o.link) delete o.link; }
      if (Array.isArray(o)) o.forEach(fixLinks); else for (const k in o) if (o[k] && typeof o[k] === 'object') fixLinks(o[k]);
    };
    for (const s of pres.slides) { fixLinks(s.shapes); L.model.walk(s.shapes, (x) => { delete x.numId; return true; }); }
    for (const id in pres.designs) {
      const d = pres.designs[id];
      delete d._ctrSet;
      const all = [d.deco, d.titleDeco].concat(d.layoutDecos ? Object.values(d.layoutDecos) : []);
      for (const list of all) if (list) { fixLinks(list); L.model.walk(list, (x) => { delete x.numId; if (x.link) delete x.link; return true; }); }
      for (const m of masters.values()) if (m.design === d) await finalizeMedia(m.ctx, d);
      for (const l of layouts.values()) if (l.master.design === d) await finalizeMedia(l.ctx, d);
    }
    if (!Object.keys(pres.designs).length) { const d = L.model.buildDesign('default', W, H); pres.designs[d.id] = d; }
    if (zip.repaired || damaged.length) pres.repaired = { rebuilt: !!zip.repaired, parts: damaged.slice(0, 20) };
    return pres;
  }

  /* ---------- transitions ---------- */
  function transitionFrom(sx) {
    let tr = kid(sx, 'transition');
    const ac = kids(sx).find((c) => c.localName === 'AlternateContent');
    let p14 = null;
    if (ac) {
      const ch = kid(ac, 'Choice'), fb = kid(ac, 'Fallback');
      p14 = ch && kid(ch, 'transition');
      tr = (fb && kid(fb, 'transition')) || tr || p14;
    }
    if (!tr) return null;
    const out = { type: 'none', spd: at(tr, 'spd') || 'fast', click: at(tr, 'advClick') !== '0', after: at(tr, 'advTm') != null ? +at(tr, 'advTm') : null };
    const durAttr = (p14 && (p14.getAttribute('p14:dur') || at(p14, 'dur'))) || tr.getAttribute('p14:dur');
    if (durAttr) out.dur = +durAttr;
    const el = Array.from(tr.children).find((c) => !['sndAc', 'extLst'].includes(c.localName));
    if (!el) return out;
    const t = el.localName;
    const known = ['blinds', 'checker', 'circle', 'comb', 'cover', 'cut', 'diamond', 'dissolve', 'fade', 'newsflash', 'plus', 'pull', 'push', 'random', 'randomBar', 'split', 'strips', 'wedge', 'wheel', 'wipe', 'zoom'];
    if (!known.includes(t)) { out.type = 'fade'; return out; }
    out.type = t;
    if (at(el, 'dir')) out.dir = at(el, 'dir');
    if (at(el, 'orient')) out.orient = at(el, 'orient');
    if (at(el, 'spokes')) out.spokes = +at(el, 'spokes');
    if (bool(at(el, 'thruBlk'))) out.thruBlk = true;
    /* defaults per spec */
    if (['cover', 'pull', 'push', 'wipe'].includes(t) && !out.dir) out.dir = 'l';
    if (['blinds', 'checker', 'comb', 'randomBar'].includes(t) && !out.dir) out.dir = 'horz';
    if (t === 'split') { out.orient = out.orient || 'horz'; out.dir = out.dir || 'out'; }
    if (t === 'wheel' && !out.spokes) out.spokes = 4;
    if (t === 'zoom' && !out.dir) out.dir = 'out';
    if (t === 'strips' && !out.dir) out.dir = 'lu';
    return out;
  }

  /* ---------- animation timing ---------- */
  function timingFrom(timing, slide) {
    if (!timing) return [];
    const mainSeq = descAll(timing, 'cTn').find((c) => at(c, 'nodeType') === 'mainSeq');
    if (!mainSeq) return [];
    const idMap = new Map();
    L.model.walk(slide.shapes, (s) => { if (s.numId) idMap.set(String(s.numId), s.id); return true; });
    const out = [];
    const clickPars = kids(kid(mainSeq, 'childTnLst'), 'par');
    for (const cp of clickPars) {
      const effects = descAll(cp, 'cTn').filter((c) => at(c, 'presetClass'));
      for (const e of effects) {
        const cls = at(e, 'presetClass');
        if (!['entr', 'exit', 'emph', 'path'].includes(cls)) continue;
        const tgt = desc(e, 'spTgt');
        if (!tgt) continue;
        const sid = idMap.get(at(tgt, 'spid'));
        if (!sid) continue;
        const pRg = desc(tgt, 'pRg');
        const pid = num(e, 'presetID', 10), sub = num(e, 'presetSubtype', 0);
        const node = at(e, 'nodeType');
        const cond = path(e, 'stCondLst', 'cond');
        let dur = 0;
        for (const c of descAll(e, 'cTn')) { const d = at(c, 'dur'); if (d && d !== 'indefinite') dur = Math.max(dur, +d + (num(path(c, 'stCondLst', 'cond'), 'delay', 0) || 0)); }
        const a = { id: L.uid('a'), sid, cls, start: node === 'clickEffect' ? 'click' : node === 'afterEffect' ? 'after' : 'with', dur: dur || 500, delay: cond && at(cond, 'delay') !== 'indefinite' ? num(cond, 'delay', 0) : 0, by: 'all' };
        const ENTR = { 1: 'appear', 2: 'flyIn', 3: 'blinds', 4: 'box', 5: 'checkerboard', 6: 'circle', 8: 'diamond', 9: 'dissolve', 10: 'fade', 12: 'peekIn', 13: 'plus', 14: 'randomBars', 16: 'split', 18: 'strips', 20: 'wedge', 21: 'wheel', 22: 'wipe', 23: 'zoom', 42: 'ascend', 47: 'descend', 55: 'expand', 49: 'spinner' };
        if (cls === 'entr' || cls === 'exit') {
          let eff = ENTR[pid];
          if (!eff) { const ae = desc(e, 'animEffect'); eff = ae ? L.anim.fromFilter(at(ae, 'filter')).eff : 'fade'; }
          if (cls === 'exit') eff = { appear: 'disappear', flyIn: 'flyOut', ascend: 'descend', expand: 'fade', spinner: 'zoom' }[eff] || eff;
          a.eff = eff;
          if (eff === 'flyIn' || eff === 'flyOut') a.dir = L.anim.dirFromSubtype('entr', 'flyIn', sub) || 'b';
          else {
            const ae = desc(e, 'animEffect');
            if (ae && at(ae, 'filter')) { const f = L.anim.fromFilter(at(ae, 'filter')); if (f.eff === eff && f.dir) a.dir = f.dir; }
            if (!a.dir) a.dir = L.anim.dirFromSubtype(cls === 'exit' ? 'exit' : 'entr', eff, sub);
          }
          if (a.eff === 'appear' || a.eff === 'disappear') a.dur = 1;
        } else if (cls === 'emph') {
          a.eff = { 6: 'growShrink', 8: 'spin', 9: 'transparency', 32: 'teeter', 35: 'blink' }[pid] || 'growShrink';
          const sc = desc(e, 'by');
          if (a.eff === 'growShrink' && sc) a.amount = Math.round(num(sc, 'x', 150000) / 1000);
          if (!{ 6: 1, 8: 1, 9: 1, 32: 1, 35: 1 }[pid]) a.amount = 110;
          const rc = descAll(e, 'cTn').find((c) => at(c, 'repeatCount'));
          if (rc && at(rc, 'repeatCount') !== 'indefinite') a.repeat = Math.max(1, Math.round(+at(rc, 'repeatCount') / 1000));
        } else {
          a.eff = 'custom';
          const am = desc(e, 'animMotion');
          const pth = am ? at(am, 'path') || '' : '';
          const nums = pth.replace(/[A-Za-z]/g, ' ').trim().split(/[\s,]+/).map(Number).filter((x) => !isNaN(x));
          const pts = [];
          for (let i = 0; i + 1 < nums.length; i += 2) pts.push([nums[i], nums[i + 1]]);
          a.pts = pts.length >= 2 ? pts : [[0, 0], [0.25, 0]];
        }
        if (pRg) {
          a.para = num(pRg, 'st', 0);
          const prev = out[out.length - 1];
          if (prev && prev.sid === sid && prev.cls === cls && prev.eff === a.eff && prev._para != null && a.para === prev._para + 1) { prev._para = a.para; prev.by = 'para'; continue; }
          a._para = a.para;
          delete a.para;
        }
        out.push(a);
      }
    }
    for (const a of out) delete a._para;
    return out;
  }

  L.pptx = L.pptx || {};
  L.pptx.read = read;
})();
