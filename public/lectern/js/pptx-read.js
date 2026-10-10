/* Lectern — PresentationML reader (.pptx / .ppsx / .potx / .pptm). */
(function () {
  'use strict';
  const L = window.L, K = L.opc;
  const pt = L.emu2pt;
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /* ---------- XML helpers ---------- */
  const parse = (s) => {
    const d = new DOMParser().parseFromString(s, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw new Error('XML parse error');
    if (L.opc) L.opc.captureAC(d, s);
    resolveAlternateContent(d);
    if (s.indexOf('%"') >= 0) percentsToThousandths(d);
    return d.documentElement;
  };
  const PCT_ATTRS = new Set(['val', 'pos', 'amt', 'lim', 'l', 't', 'r', 'b', 'sx', 'sy', 'thresh', 'fontScale', 'lnSpcReduction', 'stA', 'stPos', 'endA', 'endPos', 'sat', 'lum', 'bright', 'contrast']);
  /** Strict OOXML (and some generators) write percentages as "90%"; the transitional form is 90000 */
  function percentsToThousandths(d) {
    for (const el of d.getElementsByTagName('*')) {
      for (const a of Array.from(el.attributes)) {
        const v = a.value;
        if (v.length > 1 && v.charCodeAt(v.length - 1) === 37 && PCT_ATTRS.has(a.localName) && /^-?\d+(\.\d+)?%$/.test(v)) el.setAttribute(a.name, String(Math.round(parseFloat(v) * 1000)));
      }
    }
  }
  /** Markup Compatibility: replace every mc:AlternateContent with its Fallback (or its first Choice when there is no fallback). */
  function resolveAlternateContent(d) {
    const acs = Array.from(d.getElementsByTagName('*')).filter((e) => e.localName === 'AlternateContent').reverse();
    for (const ac of acs) {
      if (!ac.parentNode) continue;
      const kidsEl = Array.from(ac.children);
      const fb = kidsEl.find((c) => c.localName === 'Fallback');
      const ch = kidsEl.find((c) => c.localName === 'Choice');
      const pick = fb && fb.children.length ? fb : ch || fb;
      if (pick?.children.length) ac.replaceWith(...Array.from(pick.childNodes));
      else if (!['spTree', 'grpSp'].includes(ac.parentNode.localName)) ac.remove();
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
    try { return L.opc.resolve(base, target).part; } catch (e) { /* retain the damaged-package reader's recovery path */ }
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
        /* path gradients radiate from the focus rectangle (insets from the edges, in %) */
        const ftr = p && kid(p, 'fillToRect');
        if (ftr) {
          const l = num(ftr, 'l', 0) / 100000, t = num(ftr, 't', 0) / 100000, r = num(ftr, 'r', 0) / 100000, b = num(ftr, 'b', 0) / 100000;
          const cx = L.round((l + 1 - r) / 2, 4), cy = L.round((t + 1 - b) / 2, 4);
          if (Math.abs(cx - 0.5) > 0.001 || Math.abs(cy - 0.5) > 0.001) f.focus = [cx, cy];
        }
        return f;
      }
      case 'blipFill': {
        const blip = kid(el, 'blip');
        const id = blip ? ctx.media(rid(blip, 'embed')) : null;
        if (!id) return { t: 'none' };
        const amt = desc(blip, 'alphaModFix');
        const tl = kid(el, 'tile');
        const f = { t: 'img', media: id, tile: !!tl, a: amt ? num(amt, 'amt', 100000) / 100000 : 1 };
        /* duotone: the picture recoloured between two colours (theme background textures use it) */
        const duo = kid(blip, 'duotone');
        if (duo) { const cs = kids(duo).filter((x) => COLOR_TAGS.has(x.localName)).map((x) => color(x, Object.assign({}, ctx, { keepScheme: false }))); if (cs.length === 2 && cs[0] && cs[1]) f.duotone = [cs[0].c, cs[1].c]; }
        /* tiling: the picture at its own size scaled by sx/sy, offset by tx/ty from the alignment corner */
        if (tl) f.tileOpts = { sx: num(tl, 'sx', 100000) / 100000, sy: num(tl, 'sy', 100000) / 100000, tx: pt(num(tl, 'tx', 0)), ty: pt(num(tl, 'ty', 0)), algn: at(tl, 'algn') || 'tl', flip: at(tl, 'flip') || 'none' };
        const sr = kid(el, 'srcRect');
        if (!tl && sr && (num(sr, 'l', 0) || num(sr, 't', 0) || num(sr, 'r', 0) || num(sr, 'b', 0))) f.crop = { l: num(sr, 'l', 0) / 100000, t: num(sr, 't', 0) / 100000, r: num(sr, 'r', 0) / 100000, b: num(sr, 'b', 0) / 100000 };
        /* stretched into a rectangle inset from (or, negative, reaching past) the edges: "fill" cropping */
        const fr = !tl && path(el, 'stretch', 'fillRect');
        if (fr && (num(fr, 'l', 0) || num(fr, 't', 0) || num(fr, 'r', 0) || num(fr, 'b', 0))) f.fillRect = { l: num(fr, 'l', 0) / 100000, t: num(fr, 't', 0) / 100000, r: num(fr, 'r', 0) / 100000, b: num(fr, 'b', 0) / 100000 };
        return f;
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
        case 'effectLst': { const s = shadow(c, ctx); if (s) { r.shd = true; r.shdX = s; } break; }
        case 'highlight': { const v = colorOf(c, ctx); if (v) r.hl = v.c; break; }
        case 'latin': { const f = fontName(at(c, 'typeface')); if (f) r.font = f; break; }
        case 'hlinkClick': { const l = ctx.link(c); if (l) r.link = l; break; }
        default: break;
      }
    }
    if (ctx?.pkg && ctx.partPath) L.properties.capture(r, rPr, ctx, 'run');
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
    if (ctx?.pkg && ctx.partPath) L.properties.capture(o, pPr, ctx, 'para');
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

  /* ---------- charts (c:chartSpace → chart model v2, drawn by L.chart.render) ---------- */
  const CHART_TYPES = /^(bar|bar3D|line|line3D|pie|pie3D|ofPie|doughnut|area|area3D|scatter|radar|bubble|stock|surface|surface3D)Chart$/;
  const autoSeriesColor = (i, n, style, design) => L.chart.autoColor(i, n, style, design);
  function chartModel(cx, ctx) {
    const chart = kid(cx, 'chart');
    const plot = kid(chart, 'plotArea');
    if (!plot) return null;
    const cctx = Object.assign({}, ctx, { keepScheme: false, media: () => null, link: () => null, phClr: null });
    const design = ctx.design;
    const style = num(kid(cx, 'style'), 'val', 2);
    const yes = (el, d) => { if (!el) return d; const v = at(el, 'val'); return v == null ? true : v === '1' || v === 'true'; };
    const val = (el, n) => at(kid(el, n), 'val');
    const nval = (el, n, d) => { const v = val(el, n); return v == null || v === '' || isNaN(+v) ? d : +v; };
    const fillOf = (sp) => (sp ? fillIn(sp, cctx) : undefined);
    /* a line without its own colour or width keeps Office's automatic one (c and w left out) */
    const lineOf = (sp) => {
      const ln = sp && kid(sp, 'ln');
      if (!ln) return undefined;
      const l = line(ln, cctx, {});
      if (l && l.t !== 'none') { if (!fillIn(ln, cctx)) delete l.c; if (at(ln, 'w') == null) delete l.w; }
      return l;
    };
    const textOf = (tx) => {
      if (!tx) return null;
      const o = {};
      const d = path(tx, 'p', 'pPr', 'defRPr') || (kid(tx, 'p') && desc(kid(tx, 'p'), 'rPr'));
      if (d) { const r = runProps(d, cctx); if (r.sz) o.sz = r.sz; if (r.b != null) o.b = r.b; if (r.i != null) o.i = r.i; if (r.color) o.color = r.color; if (r.font) o.font = r.font; }
      const bp = kid(tx, 'bodyPr');
      if (bp && at(bp, 'rot') != null && +at(bp, 'rot') !== -60000000) o.rot = +at(bp, 'rot') / 60000;
      if (bp && /^(vert|vert270|wordArtVert|eaVert)$/.test(at(bp, 'vert') || '')) o.rot = at(bp, 'vert') === 'vert270' ? -90 : 90;
      return o;
    };
    const richText = (el) => {
      const rich = el && path(el, 'tx', 'rich');
      if (!rich) return null;
      const lines = kids(rich, 'p').map((p) => kids(p).filter((r) => r.localName === 'r' || r.localName === 'fld').map((r) => (kid(r, 't') || { textContent: '' }).textContent).join(''));
      const t = textOf(rich) || {};
      const r1 = desc(rich, 'r') && kid(desc(rich, 'r'), 'rPr');
      if (r1) { const r = runProps(r1, cctx); if (r.sz) t.sz = r.sz; if (r.b != null) t.b = r.b; if (r.i != null) t.i = r.i; if (r.color) t.color = r.color; if (r.font) t.font = r.font; }
      return { text: lines.join('\n'), tx: t };
    };
    /* cached values of a reference or literal: strings, numbers (null when absent) and the format code */
    const cacheOf = (el) => {
      if (!el) return null;
      const ml = desc(el, 'multiLvlStrCache');
      const cache = ml || desc(el, 'numCache') || desc(el, 'strCache') || desc(el, 'numLit') || desc(el, 'strLit');
      if (!cache) { const v = kid(el, 'v') || desc(el, 'v'); return v ? { str: [v.textContent], num: [+v.textContent], fmt: null, isNum: false } : null; }
      const isNum = /^num/.test(cache.localName);
      const lvl = ml ? kid(cache, 'lvl') || cache : cache;
      const pts = kids(lvl, 'pt');
      let n = num(kid(cache, 'ptCount'), 'val', 0);
      for (const p of pts) n = Math.max(n, num(p, 'idx', 0) + 1);
      const str = new Array(n).fill(''), nums = new Array(n).fill(null);
      const fmt = (kid(cache, 'formatCode') || { textContent: '' }).textContent || null;
      for (const p of pts) {
        const i = num(p, 'idx', 0), t = (kid(p, 'v') || { textContent: '' }).textContent;
        str[i] = t;
        const x = parseFloat(t);
        nums[i] = t !== '' && isFinite(x) ? x : null;
        if (isNum && at(p, 'formatCode')) (str.fmts = str.fmts || [])[i] = at(p, 'formatCode');
      }
      /* the outer levels of a multi-level category axis, innermost first */
      let outer = null;
      if (ml) outer = kids(cache, 'lvl').slice(1).map((lv) => { const a = new Array(n).fill(null); for (const p of kids(lv, 'pt')) a[num(p, 'idx', 0)] = (kid(p, 'v') || { textContent: '' }).textContent; return a; });
      return { str, num: nums, fmt: fmt && fmt !== 'General' ? fmt : null, isNum, outer };
    };
    const fmtCat = (c) => {
      if (!c) return [];
      if (c.isNum && c.fmt && L.numfmt) return c.num.map((v, i) => (v == null ? c.str[i] : L.numfmt.text((c.str.fmts && c.str.fmts[i]) || c.fmt, v)));
      return c.str.slice();
    };
    const lblsOf = (el, inherit) => {
      if (!el) return inherit || null;
      if (yes(kid(el, 'delete'), false)) return { del: true };
      const o = Object.assign({}, inherit && !inherit.del ? inherit : {});
      delete o.pts;
      for (const k of ['showVal', 'showPercent', 'showCatName', 'showSerName', 'showLegendKey', 'showBubbleSize']) { const e = kid(el, k); if (e) o[k] = yes(e, false); }
      if (val(el, 'dLblPos')) o.pos = val(el, 'dLblPos');
      const nf = kid(el, 'numFmt');
      if (nf && at(nf, 'formatCode') && at(nf, 'sourceLinked') !== '1') o.fmt = at(nf, 'formatCode');
      if (kid(el, 'separator')) o.sep = kid(el, 'separator').textContent;
      const rng = kids(path(el, 'extLst')).map((x) => kid(x, 'showDataLabelsRange')).find(Boolean);
      if (rng) o.showRange = yes(rng, false);
      const tx = textOf(kid(el, 'txPr')); if (tx) o.tx = tx;
      const sp = kid(el, 'spPr');
      if (sp) { const f = fillOf(sp), l = lineOf(sp); if (f && f.t !== 'none') o.fill = f; if (l && l.t !== 'none') o.line = l; }
      for (const d of kids(el, 'dLbl')) {
        const i = nval(d, 'idx', -1);
        if (i < 0) continue;
        o.pts = o.pts || {};
        if (yes(kid(d, 'delete'), false)) { o.pts[i] = { del: true }; continue; }
        const one = lblsOf(d, o);
        /* custom label text; fields ([VALUE], [SERIES NAME], [CELLRANGE]…) are filled in when drawn */
        const rich = path(d, 'tx', 'rich');
        if (rich) {
          const runs = [];
          kids(rich, 'p').forEach((p, pi) => {
            if (pi) runs.push('\n');
            for (const r of kids(p)) {
              if (r.localName === 'r') runs.push((kid(r, 't') || { textContent: '' }).textContent);
              else if (r.localName === 'fld') { const ty = String(at(r, 'type') || '').toUpperCase(); runs.push(/^(VALUE|SERIESNAME|CATEGORYNAME|PERCENTAGE|XVALUE|YVALUE|BUBBLESIZE|CELLRANGE)$/.test(ty) ? '\u0001' + ty + '\u0001' : (kid(r, 't') || { textContent: '' }).textContent); }
            }
          });
          one.text = runs.join('');
          const r1 = desc(rich, 'rPr') || path(rich, 'p', 'pPr', 'defRPr');
          if (r1) { const rp = runProps(r1, cctx); one.tx = Object.assign({}, one.tx || {}); if (rp.sz) one.tx.sz = rp.sz; if (rp.b != null) one.tx.b = rp.b; if (rp.i != null) one.tx.i = rp.i; if (rp.color) one.tx.color = rp.color; }
        }
        o.pts[i] = one;
      }
      return o;
    };
    const markerOf = (m) => {
      if (!m) return null;
      const o = {};
      if (val(m, 'symbol')) o.sym = val(m, 'symbol');
      if (val(m, 'size')) o.size = +val(m, 'size');
      const sp = kid(m, 'spPr');
      if (sp) { const f = fillOf(sp), l = lineOf(sp); if (f) o.fill = f; if (l) o.line = l; }
      return o;
    };
    /* axes by id */
    const axes = new Map();
    for (const a of kids(plot)) {
      if (!/^(valAx|catAx|dateAx|serAx)$/.test(a.localName)) continue;
      const o = { kind: a.localName === 'valAx' ? 'val' : a.localName === 'serAx' ? 'ser' : 'cat', id: val(a, 'axId'), cross: val(a, 'crossAx'), date: a.localName === 'dateAx' };
      o.del = yes(kid(a, 'delete'), false);
      o.pos = val(a, 'axPos') || 'b';
      const scl = kid(a, 'scaling');
      if (scl) {
        if (val(scl, 'orientation') === 'maxMin') o.rev = true;
        if (val(scl, 'min') != null) o.min = nval(scl, 'min');
        if (val(scl, 'max') != null) o.max = nval(scl, 'max');
        if (val(scl, 'logBase') != null) o.log = nval(scl, 'logBase', 10);
      }
      if (val(a, 'majorUnit') != null) o.major = nval(a, 'majorUnit');
      const nf = kid(a, 'numFmt');
      if (nf && at(nf, 'formatCode')) { o.fmt = at(nf, 'formatCode'); o.linked = at(nf, 'sourceLinked') === '1'; }
      const gl = kid(a, 'majorGridlines');
      if (gl) { const l = lineOf(kid(gl, 'spPr')); o.grid = l || {}; }
      const gm = kid(a, 'minorGridlines');
      if (gm) { const l = lineOf(kid(gm, 'spPr')); o.minorGrid = l || {}; }
      o.tick = val(a, 'majorTickMark') || 'cross';
      o.lblPos = val(a, 'tickLblPos') || 'nextTo';
      o.crosses = val(a, 'crosses') || (kid(a, 'crossesAt') ? 'at' : 'autoZero');
      if (kid(a, 'crossesAt')) o.crossesAt = nval(a, 'crossesAt', 0);
      if (val(a, 'crossBetween')) o.between = val(a, 'crossBetween') !== 'midCat';
      const sp = kid(a, 'spPr');
      const l = lineOf(sp); if (l) o.line = l;
      const tx = textOf(kid(a, 'txPr')); if (tx) o.tx = tx;
      if (val(a, 'tickLblSkip')) o.skip = nval(a, 'tickLblSkip', 1);
      const t = kid(a, 'title');
      if (t) { const rt = richText(t); o.title = rt ? rt.text : a.localName === 'valAx' ? 'Axis Title' : 'Axis Title'; o.titleTx = (rt && rt.tx) || textOf(kid(t, 'txPr')) || {}; }
      const du = kid(a, 'dispUnits');
      if (du) { const bu = val(du, 'builtInUnit'); const U = { hundreds: 1e2, thousands: 1e3, tenThousands: 1e4, hundredThousands: 1e5, millions: 1e6, tenMillions: 1e7, hundredMillions: 1e8, billions: 1e9, trillions: 1e12 }; o.unit = U[bu] || nval(du, 'custUnit', 1); }
      axes.set(o.id, o);
    }
    /* chart groups */
    const groups = [];
    const series = [];
    let catCache = null;
    for (const ct of kids(plot).filter((c) => CHART_TYPES.test(c.localName))) {
      const ln = ct.localName.replace(/3D|Chart$/g, '').replace('Chart', '');
      let type = { bar: 'bar', line: 'line', pie: 'pie', ofPie: 'pie', doughnut: 'doughnut', area: 'area', scatter: 'scatter', radar: 'radar', bubble: 'bubble', stock: 'stock', surface: 'area' }[ln] || 'bar';
      const g = { type, d3: /3D/.test(ct.localName) };
      g.grouping = val(ct, 'grouping') || (type === 'line' || type === 'area' ? 'standard' : 'clustered');
      if (type === 'bar') {
        g.dir = val(ct, 'barDir') || 'col';
        /* 3-D "standard" columns stand in rows one behind the other */
        if (g.grouping === 'standard') { g.grouping = 'clustered'; if (g.d3) g.deep = true; }
        g.gap = nval(ct, 'gapWidth', 150);
        g.overlap = /stacked/i.test(g.grouping) ? 100 : nval(ct, 'overlap', 0);
      }
      g.vary = yes(kid(ct, 'varyColors'), type === 'pie' || type === 'doughnut');
      if (type === 'pie' || type === 'doughnut') { g.firstAng = nval(ct, 'firstSliceAng', 0); g.hole = type === 'doughnut' ? nval(ct, 'holeSize', 50) : 0; }
      if (type === 'scatter') g.scatterStyle = val(ct, 'scatterStyle') || 'marker';
      if (type === 'radar') g.radarStyle = val(ct, 'radarStyle') || 'standard';
      if (type === 'bubble') { g.bubbleScale = nval(ct, 'bubbleScale', 100); g.sizeArea = val(ct, 'sizeRepresents') !== 'w'; }
      if (type === 'line' || type === 'stock') { g.marker = yes(kid(ct, 'marker'), true); if (kid(ct, 'hiLowLines')) g.hiLow = lineOf(kid(kid(ct, 'hiLowLines'), 'spPr')) || {}; if (kid(ct, 'dropLines')) g.drop = lineOf(kid(kid(ct, 'dropLines'), 'spPr')) || {}; if (kid(ct, 'upDownBars')) g.upDown = true; }
      const ax = kids(ct, 'axId').map((e) => at(e, 'val'));
      g.axIds = ax;
      const gl = lblsOf(kid(ct, 'dLbls'), null);
      g.si = [];
      for (const s of kids(ct, 'ser')) {
        const i = series.length;
        const sr = { g: groups.length, idx: nval(s, 'idx', i), order: nval(s, 'order', i) };
        const tx = kid(s, 'tx');
        const nc = cacheOf(tx);
        sr.name = (nc && nc.str[0]) || (tx && kid(tx, 'v') ? kid(tx, 'v').textContent : '') || 'Series ' + (sr.idx + 1);
        const cEl = kid(s, 'cat') || kid(s, 'xVal');
        const cc = cacheOf(cEl);
        if (cc && (!catCache || cc.str.length > catCache.str.length)) catCache = cc;
        if (kid(s, 'xVal')) sr.xs = cc ? (cc.isNum ? cc.num.slice() : cc.str.map((_, k) => k + 1)) : null;
        const vc = cacheOf(kid(s, 'val') || kid(s, 'yVal'));
        sr.vals = vc ? vc.num.slice() : [];
        if (vc && vc.fmt) sr.fmt = vc.fmt;
        if (type === 'bubble') { const bc = cacheOf(kid(s, 'bubbleSize')); sr.sizes = bc ? bc.num.slice() : sr.vals.map(() => 1); }
        const sp = kid(s, 'spPr');
        const f = fillOf(sp), l = lineOf(sp);
        if (f) sr.fill = f;
        if (l) sr.line = l;
        const mk = markerOf(kid(s, 'marker')); if (mk) sr.marker = mk;
        if (kid(s, 'smooth')) sr.smooth = yes(kid(s, 'smooth'), false);
        if (kid(s, 'explosion')) sr.expl = nval(s, 'explosion', 0);
        if (yes(kid(s, 'invertIfNegative'), false)) sr.invNeg = true;
        for (const dp of kids(s, 'dPt')) {
          const k = nval(dp, 'idx', -1);
          if (k < 0) continue;
          const o = {};
          const dsp = kid(dp, 'spPr');
          const df = fillOf(dsp), dl = lineOf(dsp);
          if (df) o.fill = df; if (dl) o.line = dl;
          if (kid(dp, 'explosion')) o.expl = nval(dp, 'explosion', 0);
          const dm = markerOf(kid(dp, 'marker')); if (dm) o.marker = dm;
          (sr.pts = sr.pts || {})[k] = o;
        }
        const sl = lblsOf(kid(s, 'dLbls'), gl);
        /* labels taken from a cell range (Office 2013 "Value From Cells") */
        const dr = desc(s, 'dlblRangeCache');
        if (sl && dr) { const rc = cacheOf(dr.parentNode); if (rc) sl.range = rc.str; }
        if (sl && !sl.del && (sl.showVal || sl.showPercent || sl.showCatName || sl.showSerName || sl.showBubbleSize || sl.pts)) sr.lbl = sl;
        /* the colour that stands for the series: legend, datasheet, editing */
        const solid = (x) => (x && x.t === 'solid' ? x.c : x && x.t === 'grad' && x.stops[0] ? x.stops[0].c : x && x.t === 'patt' ? x.fg : null);
        const lineType = type === 'line' || type === 'scatter' || type === 'radar' && g.radarStyle !== 'filled' || type === 'stock';
        sr.color = (lineType ? (l && l.t !== 'none' && l.c && l.c[0] === '#' ? l.c : null) || solid(mk && mk.fill) : solid(f)) || null;
        g.si.push(i);
        series.push(sr);
      }
      groups.push(g);
    }
    if (!groups.length) return null;
    /* secondary axes: a group whose value axis is not the first group's */
    /* XY charts list the horizontal value axis first */
    const isXY = (g) => g.type === 'scatter' || g.type === 'bubble';
    const valAxOf = (g) => (isXY(g) ? axes.get(g.axIds[1]) || null : g.axIds.map((id) => axes.get(id)).find((a) => a && a.kind === 'val') || null);
    const catAxOf = (g) => (isXY(g) ? axes.get(g.axIds[0]) || null : g.axIds.map((id) => axes.get(id)).find((a) => a && a.kind === 'cat') || null);
    const v1 = valAxOf(groups[0]), c1 = catAxOf(groups[0]);
    let v2 = null, c2 = null;
    for (const g of groups) {
      const va = valAxOf(g);
      if (va && v1 && va !== v1) { g.sec = true; v2 = va; c2 = catAxOf(g); }
    }
    const nSer = series.length;
    /* colours left to Office's automatic choice */
    for (const sr of series) {
      const auto = autoSeriesColor(sr.idx, Math.max(nSer, sr.idx + 1), style, design);
      if (!sr.color) sr.color = auto;
      sr.auto = auto;
    }
    const cats = catCache ? fmtCat(catCache) : (series[0] ? series[0].vals.map((_, i) => String(i + 1)) : []);
    const main = groups[0];
    const m = { v: 2, style, groups, series, cats };
    if (catCache && catCache.outer && catCache.outer.length) m.catOuter = catCache.outer;
    if (catCache && catCache.isNum) m.catNums = catCache.num.slice();
    /* legacy fields: chart type, datasheet and editing */
    const k = main.type;
    if (k === 'bar') m.kind = main.dir === 'bar' ? (main.grouping === 'percentStacked' ? 'barPct' : main.grouping === 'stacked' ? 'barStacked' : 'bar') : main.grouping === 'percentStacked' ? 'colPct' : main.grouping === 'stacked' ? 'colStacked' : 'col';
    else if (k === 'line' || k === 'stock') m.kind = series.some((s) => s.g === 0 && (!s.marker || s.marker.sym !== 'none')) && main.marker !== false ? 'lineMarkers' : 'line';
    else if (k === 'area') m.kind = main.grouping === 'percentStacked' ? 'areaPct' : main.grouping === 'stacked' ? 'areaStacked' : 'area';
    else if (k === 'pie') m.kind = 'pie';
    else if (k === 'doughnut') m.kind = 'doughnut';
    else if (k === 'radar') m.kind = main.radarStyle === 'filled' ? 'radarFilled' : 'radar';
    else if (k === 'bubble') m.kind = 'bubble';
    else m.kind = 'scatter';
    for (const sr of series) if (sr.g !== 0) sr.overlay = groups[sr.g].type === 'line' || groups[sr.g].type === 'scatter' ? (sr.line && sr.line.t === 'none' ? 'markers' : sr.marker && sr.marker.sym === 'none' ? 'line' : 'lineMarkers') : groups[sr.g].type;
    m.ax = {};
    if (c1) m.ax.c = c1; if (v1) m.ax.v = v1; if (v2) m.ax.v2 = v2; if (c2 && c2 !== c1) m.ax.c2 = c2;
    m.gridY = !!(v1 && v1.grid);
    if (c1 && c1.grid) m.gridX = true;
    /* text: chart-wide default, title, legend */
    const tx0 = textOf(kid(cx, 'txPr')) || {};
    m.fsz = tx0.sz || 10; if (tx0.font) m.font = tx0.font; if (tx0.color) m.textColor = tx0.color; if (tx0.b) m.bold = true;
    const title = kid(chart, 'title');
    const autoDel = yes(kid(chart, 'autoTitleDeleted'), false);
    m.title = '';
    if (title) {
      const rt = richText(title);
      m.title = rt ? rt.text : series.length === 1 ? series[0].name : 'Chart Title';
      m.titleTx = (rt && rt.tx) || textOf(kid(title, 'txPr')) || {};
      if (yes(kid(title, 'overlay'), false)) m.titleOverlay = true;
      const tsp = kid(title, 'spPr'); if (tsp) { const f = fillOf(tsp), l = lineOf(tsp); if (f && f.t !== 'none') m.titleFill = f; if (l && l.t !== 'none') m.titleLine = l; }
    } else if (!autoDel && series.length === 1) { m.title = series[0].name; m.titleTx = {}; } /* Office's automatic title for a single series */
    const leg = kid(chart, 'legend');
    m.legend = 'none';
    if (leg) {
      m.legend = val(leg, 'legendPos') || 'r';
      const lt = textOf(kid(leg, 'txPr')); if (lt) m.legendTx = lt;
      if (yes(kid(leg, 'overlay'), false)) m.legendOverlay = true;
      const lsp = kid(leg, 'spPr'); if (lsp) { const f = fillOf(lsp), l = lineOf(lsp); if (f && f.t !== 'none') m.legendFill = f; if (l && l.t !== 'none') m.legendLine = l; }
      const del = kids(leg, 'legendEntry').filter((e) => yes(kid(e, 'delete'), false)).map((e) => nval(e, 'idx', -1));
      if (del.length) m.legendDel = del;
      const ml = path(leg, 'layout', 'manualLayout');
      if (ml && val(ml, 'x') != null && val(ml, 'w') != null) m.legendBox = { x: nval(ml, 'x', 0), y: nval(ml, 'y', 0), w: nval(ml, 'w', 0), h: nval(ml, 'h', 0), edge: (val(ml, 'xMode') || 'factor') === 'edge' };
    }
    /* chart area and plot area */
    const csp = kid(cx, 'spPr');
    if (csp) { const f = fillOf(csp), l = lineOf(csp); if (f && f.t !== 'none') m.bg = f; if (l && l.t !== 'none') m.border = l; }
    const psp = kid(plot, 'spPr');
    if (psp) { const f = fillOf(psp), l = lineOf(psp); if (f && f.t !== 'none') m.plotBg = f; if (l && l.t !== 'none') m.plotBorder = l; }
    const pml = path(plot, 'layout', 'manualLayout');
    if (pml && val(pml, 'w') != null && val(pml, 'h') != null) m.plotBox = { x: nval(pml, 'x', 0), y: nval(pml, 'y', 0), w: nval(pml, 'w', 1), h: nval(pml, 'h', 1), inner: val(pml, 'layoutTarget') === 'inner', edge: (val(pml, 'xMode') || 'factor') === 'edge' };
    if (val(chart, 'dispBlanksAs')) m.blanks = val(chart, 'dispBlanksAs');
    /* 3-D charts: the view angles and the walls */
    if (groups.some((g) => g.d3)) {
      const v3 = kid(chart, 'view3D');
      m.view3D = { rotX: nval(v3, 'rotX', groups[0].type === 'pie' ? 30 : 15), rotY: nval(v3, 'rotY', groups[0].type === 'pie' ? 0 : 20), depth: nval(v3, 'depthPercent', 100), rAng: yes(kid(v3, 'rAngAx'), false) };
      const wall = (n) => { const w = kid(chart, n); const sp = w && kid(w, 'spPr'); if (!sp) return null; const f = fillOf(sp), l = lineOf(sp); return { fill: f, line: l }; };
      m.walls = { back: wall('backWall'), side: wall('sideWall'), floor: wall('floor') };
    }
    /* legacy summary fields */
    const anyLbl = series.find((s) => s.lbl);
    m.labels = !!anyLbl;
    const fmt1 = series[0] && series[0].fmt;
    if (anyLbl && anyLbl.lbl.fmt) m.numFmt = anyLbl.lbl.fmt; else if (fmt1) m.numFmt = fmt1;
    if (v1 && v1.fmt && (!v1.linked || !fmt1) && v1.fmt !== 'General') m.axisFmt = v1.fmt; else if (v1 && v1.linked && fmt1) m.axisFmt = fmt1;
    if (c1 && c1.tx && c1.tx.rot != null) m.catRot = c1.tx.rot;
    if (k === 'pie' || k === 'doughnut') {
      const s0 = series.find((s) => s.g === 0);
      if (s0) m.pieColors = cats.map((_, i) => { const p = s0.pts && s0.pts[i]; const f = p && p.fill; return f && f.t === 'solid' ? f.c : main.vary ? autoSeriesColor(i, cats.length, style, design) : s0.color; });
    }
    return m;
  }

  /* ---------- reader ---------- */
  async function read(buffer, opts) {
    opts = opts || {};
    let zip;
    let password = null;
    try { zip = await L.zip.read(buffer); } catch (e) {
      if (e && e.code === 'ole' && !e.encrypted) e.message = 'This is an older binary PowerPoint 97–2003 file (.ppt) with a .pptx name. Save it as .pptx in PowerPoint, Keynote, Google Slides or LibreOffice first.';
      if (!(e && e.code === 'ole' && e.encrypted && L.officeCrypto)) throw e;
      /* a presentation with a password to open: decrypt it here, in the browser */
      if (!opts.password) { const err = new Error('This presentation is password-protected.'); err.code = 'password'; throw err; }
      let inner;
      try { inner = await L.officeCrypto.decrypt(buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer), opts.password); } catch (e2) {
        const unsup = e2 && e2.message === 'unsupported';
        const err = new Error(unsup ? 'This presentation is encrypted with a method that is not supported (only the AES encryption of PowerPoint 2007 and later is).' : e2 && /secure/.test(e2.message) ? e2.message : 'The password is incorrect. PowerPoint cannot open the file.');
        err.code = unsup ? 'unsupported' : /secure/.test(String(e2 && e2.message)) ? 'insecure' : 'badpassword';
        throw err;
      }
      zip = await L.zip.read(inner);
      password = opts.password;
    }
    const damaged = [];
    const file = part => zip.get(part) || zip.get(L.opc.relative('', part));
    let pkg = null, packageError = null;
    try { pkg = await L.opc.open(zip); } catch (e) { packageError = e.message; }
    const originals = new Map();
    const original = part => {
      if (!pkg) return null;
      if (!originals.has(part)) {
        try { originals.set(part, L.preserve.originalShapes(pkg, part)); } catch (e) { originals.set(part, null); }
      }
      return originals.get(part);
    };
    const textCache = new Map();
    const getText = async (p) => {
      if (textCache.has(p)) return textCache.get(p);
      const f = file(p);
      let t = null;
      if (f) { try { t = pkg?.has(p) ? pkg.text(p) : L.xmlTree.decode(await f.bytes()); } catch (e) { damaged.push(p); } }
      textCache.set(p, t);
      return t;
    };
    const xml = async (p) => { const t = await getText(p); if (!t) return null; try { return parse(t); } catch (e) { damaged.push(p); return null; } };
    /* content types, for parts that are carried over unchanged (chart workbooks, chart styles) */
    let ctTable = null;
    const contentType = async (part) => {
      if (!ctTable) {
        ctTable = { ext: {}, over: {} };
        const ct = await xml('[Content_Types].xml');
        if (ct) for (const e of kids(ct)) {
          if (e.localName === 'Default') ctTable.ext[String(at(e, 'Extension') || '').toLowerCase()] = at(e, 'ContentType');
          else if (e.localName === 'Override') ctTable.over[String(at(e, 'PartName') || '').replace(/^\//, '').toLowerCase()] = at(e, 'ContentType');
        }
      }
      return ctTable.over[part.toLowerCase()] || ctTable.ext[part.split('.').pop().toLowerCase()] || null;
    };
    const relsCache = new Map();
    const rels = async (p) => {
      if (relsCache.has(p)) return relsCache.get(p);
      const r = await xml(relsPath(p));
      const map = {};
      if (r) for (const e of kids(r, 'Relationship')) {
        const ext = at(e, 'TargetMode') === 'External';
        map[at(e, 'Id')] = { type: (at(e, 'Type') || '').split('/').pop(), fullType: at(e, 'Type') || '', target: ext ? at(e, 'Target') : resolvePath(p, at(e, 'Target')), external: ext };
      }
      relsCache.set(p, map);
      return map;
    };
    const mediaCache = new Map();
    const loadMedia = async (p) => {
      if (mediaCache.has(p)) return mediaCache.get(p);
      const f = file(p);
      if (!f) { mediaCache.set(p, null); return null; }
      let bytes;
      try { bytes = await f.bytes(); } catch (e) { damaged.push(p); mediaCache.set(p, null); return null; }
      const ext = p.split('.').pop().toLowerCase();
      let view = null;
      if ((ext === 'wmf' || ext === 'emf') && L.metafile) view = await L.metafile.toPNG(bytes, ext);
      const id = L.media.add(new Blob([bytes], { type: L.extToMime(ext) }), p.split('/').pop(), view);
      const dims = L.imageSize ? L.imageSize(bytes) : null;
      if (dims) L.media.get(id).size = dims;
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
    if (pkg) L.opc.attach(pres, pkg);
    if (packageError) L.opc.loss(pres, { id: 'package:unreadable', what: 'Part of this file couldn\'t be read. Anything in it that VibeOffice can\'t edit itself, such as macros or embedded objects, won\'t be saved.', detail: String(packageError), where: presPath, action: 'drop', notify: true });
    pres.designs = {};
    pres.firstNum = num(presX, 'firstSlideNum', 1);

    /* document properties */
    try {
      const propertyPart = (type, fallback) => pkg?.rels('').find(r => L.opc.relationshipType(r.type) === type)?.part || fallback;
      const core = await xml(propertyPart(L.opc.NS.pkg + '/metadata/core-properties', 'docProps/core.xml'));
      if (core) {
        const g = (n) => { const e = desc(core, n); return e ? e.textContent : ''; };
        Object.assign(pres.props, { title: g('title'), subject: g('subject'), author: g('creator'), keywords: g('keywords'), comments: g('description'), category: g('category'), created: g('created') || pres.props.created });
        pres.title = pres.props.title;
      }
      const app = await xml(propertyPart(L.opc.NS.rel + '/extended-properties', 'docProps/app.xml'));
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
      theme.rels = await rels(p);
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
    /** fill and line styles of the theme refer to pictures through the theme part's own relationships */
    function themeCtx(ctx, theme, phClr) {
      const trels = (theme && theme.rels) || {};
      const c = Object.assign({}, ctx, { phClr });
      c.media = (id) => { const r = trels[id]; if (!r || r.external) return null; const m = mediaCache.get(r.target); if (m) return m; ctx.pendingMedia.add(r.target); return '__pending__:' + r.target; };
      return c;
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
        if (el) return fill(el, themeCtx(ctx, theme, c));
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
      if (!mx) throw new Error('Missing slide master: ' + p);
      const mr = await rels(p);
      const themeRel = Object.values(mr).find((r) => r.type === 'theme');
      const theme = await loadTheme(themeRel ? themeRel.target : '');
      const design = L.model.buildDesign('default', W, H);
      design.keep = { source: pkg?.id, part: p, theme: themeRel?.target, layouts: {}, layoutParts: [] };
      if (pkg?.has(p)) design.keep.master = K.fragment(pkg.xml(p), { pkg, part: p });
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
      if (!lx) throw new Error('Missing slide layout: ' + p);
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
      {
        const key = p.split('/').pop().replace('.xml', '');
        d.layoutDecos = d.layoutDecos || {};
        d.layoutBgs = d.layoutBgs || {};
        d.layoutNames = d.layoutNames || {};
        d.layoutDecos[key] = deco;
        (d.layoutShowMaster ||= {})[key] = showMaster;
        if (bg) d.layoutBgs[key] = bg;
        d.layoutNames[key] = name;
        info.lkey = key;
      }
      if (pkg?.has(p)) d.keep.layoutParts.push({ part: p, lkey: info.lkey, type,
        fragment: K.fragment(pkg.xml(p), { pkg, part: p }) });
      layouts.set(p, info);
      return info;
    }

    /* per-part parsing context */
    function mkCtx(partPath, partRels, design, theme) {
      const ctx = {
        pkg, design, theme, partPath, rels: partRels, phClr: null, clrMapOvr: null, mediaQueue: [],
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
        out.fill = el ? fill(el, themeCtx(ctx, ctx.theme, c)) : c ? { t: 'solid', c: c.c, a: c.a } : undefined;
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
          const source = original(ctx.partPath);
          const cnv = desc(el, 'cNvPr'), raw = source?.map.get(at(cnv, 'id'));
          let s = await shapeFrom(el, ctx, opts);
          if ((!s || Array.isArray(s) && !s.length) && source) s = L.frames.placeholder(el, raw, tree, ctx);
          if (!s) continue;
          if (source) {
            const models = Array.isArray(s) ? s : [s];
            for (const shape of models) L.preserve.shape(shape, source.map.get(String(shape.numId)), el, pkg, ctx.partPath);
            L.frames.attach(models, el, raw, tree, pkg, ctx.partPath);
          }
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
      /* an SVG picture whose PNG copy is missing shows the SVG itself */
      const fix = (o) => {
        if (!o || typeof o !== 'object') return;
        if (o.type === 'image' && !o.media && o.svgMedia) { o.media = o.svgMedia; delete o.svgMedia; delete o.missingLabel; }
        if (o.type === 'image' && !o.media && !o.missingLabel) o.missingLabel = 'Picture';
        for (const k in o) { const v = o[k]; if (v && typeof v === 'object') fix(v); }
      };
      fix(obj);
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
        /* Office 2016 SVG pictures: the vector original sits in an extension, a PNG copy in r:embed (sometimes left out) */
        const svgEl = blip && desc(blip, 'svgBlip');
        const svgId = svgEl ? ctx.media(rid(svgEl, 'embed')) : null;
        const mid = blip ? ctx.media(rid(blip, 'embed')) || svgId || ctx.media(rid(blip, 'link')) : null;
        sh = Object.assign(base, { type: 'image', geom: 'rect', media: mid, crop: { l: 0, t: 0, r: 0, b: 0 }, lockAspect: true, img: {} });
        if (svgId && svgId !== mid) sh.svgMedia = svgId;
        /* a picture linked to a file outside the presentation keeps its link */
        const lrel = blip && rid(blip, 'link') ? ctx.rels[rid(blip, 'link')] : null;
        if (lrel && lrel.external) sh.linkUrl = lrel.target;
        const sr = bf && kid(bf, 'srcRect');
        if (sr) sh.crop = { l: num(sr, 'l', 0) / 100000, t: num(sr, 't', 0) / 100000, r: num(sr, 'r', 0) / 100000, b: num(sr, 'b', 0) / 100000 };
        if (blip) {
          if (desc(blip, 'grayscl')) sh.img.mode = 'gray';
          if (desc(blip, 'biLevel')) sh.img.mode = 'bw';
          const lum = desc(blip, 'lum');
          if (lum) { const br = num(lum, 'bright', 0) / 100000, co = num(lum, 'contrast', 0) / 100000; if (br >= 0.6 && co <= -0.6) sh.img.mode = 'wash'; else { if (br) sh.img.bright = br; if (co) sh.img.contrast = co; } }
          const am = desc(blip, 'alphaModFix');
          if (am) sh.img.alpha = num(am, 'amt', 100000) / 100000;
          /* Set Transparent Color: one colour of the picture shows through */
          const cc = kid(blip, 'clrChange');
          if (cc) {
            const kctx = Object.assign({}, ctx, { keepScheme: false });
            const from = colorOf(kid(cc, 'clrFrom'), kctx), to = colorOf(kid(cc, 'clrTo'), kctx);
            if (from && to && to.a < 0.5) sh.img.clear = from.c;
          }
        }
        const g = geometry(spPr);
        sh.geom = g.geom; if (g.adj) sh.adj = g.adj; if (g.path) sh.path = g.path;
        sh.line = line(spPr && kid(spPr, 'ln'), ctx, refs.line ? refs.line : { t: 'none' }) || { t: 'none' };
        const ef = spPr && kid(spPr, 'effectLst');
        const shd = shadow(ef, ctx); if (shd) sh.shadow = shd;
        if (ph) sh.ph = ph;
        if (!mid) sh.missingLabel = sh.linkUrl ? 'Linked picture: ' + decodeURIComponent(String(sh.linkUrl).split(/[\\/]/).pop() || '') : 'Picture';
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
          /* a chart whose part is missing is kept as an empty frame rather than shown with made-up data */
          return Object.assign(base, { type: 'chart', chart: model || { v: 2, kind: 'col', empty: true, title: '', legend: 'none', cats: [], series: [], groups: [] } });
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
            /* PowerPoint's VML names the shape by id="_x0000_s1038"; Word-style VML uses o:spid */
            const re = new RegExp('<v:shape\\b[^>]*\\b(?:o:spid|id)="' + spid.replace(/[^\w]/g, '') + '"[\\s\\S]*?</v:shape>');
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
      /* "Background" fill: the shape shows the slide background */
      if (f === undefined && bool(at(el, 'useBgFill'))) f = { t: 'bg' };
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
      let model = null;
      /* a chart pasted with its source formatting carries its own theme colours and fonts */
      let cctx = ctx;
      try {
        const tov = Object.values(await rels(p)).find((x) => x.type === 'themeOverride' && !x.external);
        if (tov) {
          const tx = await xml(tov.target);
          const th = tx ? await loadTheme(tov.target) : null;
          if (th) {
            const d = Object.assign({}, ctx.design, { colors: Object.assign({}, (ctx.design && ctx.design.colors) || {}, desc(tx, 'clrScheme') ? th.colors : {}) });
            if (desc(tx, 'fontScheme')) d.fonts = Object.assign({}, (ctx.design && ctx.design.fonts) || {}, th.fonts);
            cctx = Object.assign({}, ctx, { design: d });
          }
        }
      } catch (e) { /* keep the slide's theme */ }
      try { model = chartModel(cx, cctx); } catch (e) { console.warn('chart', p, e); }
      if (model && cctx !== ctx && cctx.design.fonts !== (ctx.design && ctx.design.fonts)) {
        /* theme fonts of the override are resolved here, the slide's theme would give others */
        const f = cctx.design.fonts;
        const res = (o) => { if (o && (o.font === '+mn' || o.font === '+mj')) o.font = o.font === '+mj' ? f.major : f.minor; };
        if (!model.font) model.font = f.minor; else res(model);
        [model.titleTx, model.legendTx].forEach(res);
        Object.values(model.ax || {}).forEach((a) => { res(a.tx); res(a.titleTx); });
        model.series.forEach((sr) => sr.lbl && res(sr.lbl.tx));
      }
      if (!model) return null;
      /* keep the chart part with its embedded workbook, style and colour parts, so that a chart that is not
         edited here is saved exactly as it came (PowerPoint can still open its data in Excel) */
      try {
        let text = await getText(p);
        const r = await rels(p);
        const parts = [];
        for (const [id, rel] of Object.entries(r)) {
          const ref = new RegExp('<(\\w+:)?(userShapes|externalData)\\b[^>]*?r:id="' + id.replace(/[^\w-]/g, '') + '"[^>]*?(?:/>|>[\\s\\S]*?</(\\w+:)?(userShapes|externalData)>)');
          if (rel.type === 'chartUserShapes') { text = text.replace(ref, ''); continue; }
          if (rel.external) { parts.push({ id, fullType: rel.fullType, target: rel.target, external: true }); continue; }
          const f = file(rel.target);
          if (!f) { text = text.replace(ref, ''); continue; }
          let bytes;
          try { bytes = await f.bytes(); } catch (e) { text = text.replace(ref, ''); continue; }
          parts.push({ id, fullType: rel.fullType, type: rel.type, name: rel.target.split('/').pop(), bytes, ct: await contentType(rel.target) });
        }
        /* The package writer carries the original graph, including Strict charts. */
        if (text) model.srcId = L.chart.keep({ xml: text, parts, source: pkg?.id, part: p, owner: ctx.partPath });
      } catch (e) { /* the chart is then written from its model */ }
      return model;
    }

    async function smartArtFrom(gd, base, ctx) {
      const relIds = kids(gd).find((c) => c.localName === 'relIds');
      if (!relIds) return null;
      const dm = ctx.rels[rid(relIds, 'dm')];
      let drawingPath = null, sa = null;
      const withSa = (g) => { if (g && sa) g.sa = sa; return g; };
      if (dm) {
        const dx = await xml(dm.target);
        /* a diagram Lectern wrote (our layout ids): editable again as SmartArt (js/diagram.js) */
        if (dx && L.saIO) try { sa = L.saIO.read(dx); } catch (e) { sa = null; }
        /* an item's own formatting: its point's fill and outline, its text's colour, bold and italic */
        /* a shape filled with a picture: not something our layouts carry, so the diagram stays as it came */
        if (sa && sa.points.some(([pt]) => desc(pt, 'blipFill'))) sa = null;
        if (sa) for (const [pt, item] of sa.points) {
          const sp = kid(pt, 'spPr'), f = {};
          const fe = sp && kids(sp).find((c) => /Fill$/.test(c.localName));
          if (fe) f.fill = fill(fe, ctx);
          if (sp && kid(sp, 'ln')) f.line = line(kid(sp, 'ln'), ctx);
          const rp = desc(pt, 'rPr') || desc(pt, 'endParaRPr');
          if (rp) {
            if (at(rp, 'b') === '1') f.b = true;
            if (at(rp, 'i') === '1') f.i = true;
            const sf = kid(rp, 'solidFill'), c = sf && colorOf(sf, ctx);
            if (c) f.color = c.c;
          }
          if (Object.keys(f).length) item.fmt = { node: f };
        }
        const ext = dx && desc(dx, 'dataModelExt');
        if (ext && at(ext, 'relId') && ctx.rels[at(ext, 'relId')]) drawingPath = ctx.rels[at(ext, 'relId')].target;
      }
      if (!drawingPath) { const r = Object.values(ctx.rels).find((x) => x.type === 'diagramDrawing'); if (r) drawingPath = r.target; }
      if (!drawingPath) return withSa(await smartArtFromData(relIds, base, ctx));
      const dx = await xml(drawingPath);
      const tree = dx && desc(dx, 'spTree');
      if (!tree) return withSa(await smartArtFromData(relIds, base, ctx));
      const dRels = await rels(drawingPath);
      const dctx = Object.assign({}, ctx, { rels: dRels, partPath: drawingPath, map: (b) => ({ x: base.x + b.x, y: base.y + b.y, w: b.w, h: b.h }), layout: null, pendingMedia: ctx.pendingMedia });
      dctx.media = (id) => { const r = dRels[id]; if (!r || r.external) return null; const m = mediaCache.get(r.target); if (m) return m; ctx.pendingMedia.add(r.target); return '__pending__:' + r.target; };
      const kidsS = await shapesFrom(tree, dctx, {});
      /* dsp shapes may carry a separate text frame */
      for (const sp of kids(tree, 'sp')) void sp;
      if (!kidsS.length) return withSa(await smartArtFromData(relIds, base, ctx));
      const g = Object.assign(base, { type: 'group', kids: kidsS, name: base.name || 'Diagram' });
      if (sa) g.sa = sa;
      return g;
    }
    /**
     * SmartArt saved without its drawing (PowerPoint 2007 and some generators keep only the data model):
     * lay the nodes out ourselves — chevrons or boxes joined by arrows for processes, a ring for cycles,
     * a tree for hierarchies, stacked bands for pyramids, otherwise a column of boxes with each node's
     * children as bullet text — in the theme's accent colour, as PowerPoint's default style draws them.
     */
    async function smartArtFromData(relIds, base, ctx) {
      const dmRel = ctx.rels[rid(relIds, 'dm')], loRel = ctx.rels[rid(relIds, 'lo')];
      const dm = dmRel ? await xml(dmRel.target) : null;
      if (!dm) return null;
      const lo = loRel ? await xml(loRel.target) : null;
      const kind = String((lo && at(lo, 'uniqueId')) || '').toLowerCase();
      const pts = new Map();
      for (const p of descAll(dm, 'pt')) {
        const type = at(p, 'type') || 'node';
        if (type !== 'node' && type !== 'doc') continue;
        const tEl = kid(p, 't');
        const text = tEl ? kids(tEl, 'p').map((ap) => descAll(ap, 't').map((t) => t.textContent).join('')).join(' ').trim() : '';
        pts.set(at(p, 'modelId'), { id: at(p, 'modelId'), type, text, kids: [] });
      }
      const parent = new Map();
      for (const c of descAll(dm, 'cxn')) {
        if ((at(c, 'type') || 'parOf') !== 'parOf') continue;
        const a = pts.get(at(c, 'srcId')), b = pts.get(at(c, 'destId'));
        if (a && b) { a.kids.push({ n: b, ord: num(c, 'srcOrd', 0) }); parent.set(b.id, a); }
      }
      for (const p of pts.values()) p.kids = p.kids.sort((x, y) => x.ord - y.ord).map((k) => k.n);
      const root = [...pts.values()].find((p) => p.type === 'doc');
      const tops = root ? root.kids : [...pts.values()].filter((p) => !parent.has(p.id));
      if (!tops.length) return null;
      const W = base.w, H = base.h, X = base.x, Y = base.y;
      /* colours of the diagram's colour style (node1: the main shapes) */
      const csRel = ctx.rels[rid(relIds, 'cs')];
      const cs = csRel ? await xml(csRel.target) : null;
      const lbl = (name) => cs && kids(cs, 'styleLbl').find((e) => at(e, 'name') === name);
      const clrList = (st, list) => (st ? kids(kid(st, list)).map((e) => color(e, ctx)).filter(Boolean) : []);
      const node1 = lbl('node1');
      const fills = clrList(node1, 'fillClrLst'), lines = clrList(node1, 'linClrLst'), txs = clrList(node1, 'txFillClrLst');
      const cycle = fills.length > 1;
      const fillOf = (i) => { const c = fills.length ? fills[cycle ? i % fills.length : 0] : { c: 'accent1', a: 1 }; return { t: 'solid', c: c.c, a: c.a }; };
      const lineC = lines.length ? lines[0].c : 'bg1';
      const textC = txs.length ? txs[0].c : 'bg1';
      const out = [];
      /* one text size per role, the largest at which every node of that role fits its box (as SmartArt does) */
      const fits = (text, w, h, sz, bullets) => {
        const cw = sz * 0.5, lh = sz * 1.2;
        let lines = 0;
        const perLine = Math.max(1, (w - 8) / cw - (bullets ? 2 : 0));
        for (const t of text) {
          const words = String(t).split(/\s+/);
          lines++;
          let line = 0;
          for (const wd of words) {
            /* words are not broken: the size must let the longest word fit on a line */
            if (wd.length > perLine) return false;
            const add = (line ? 1 : 0) + wd.length;
            if (line + add > perLine && line) { lines++; line = wd.length; } else line += add;
          }
        }
        return lines * lh <= h - 6;
      };
      const fitSize = (items, max) => {
        let sz = max;
        for (const it of items) { while (sz > 6 && !fits(it.lines, it.w, it.h, sz, it.bullets)) sz -= 0.5; }
        return sz;
      };
      const pending = [];
      const box = (x, y, bw, bh, text, sub, geom, o) => {
        o = o || {};
        const sh = { id: L.uid('s'), name: 'Diagram shape', type: 'shape', geom: geom || 'roundRect', x: X + x, y: Y + y, w: bw, h: bh, rot: 0, fill: o.fill || fillOf(o.i || 0), line: o.noLine ? { t: 'none' } : { c: lineC, w: 1.5 } };
        if (geom === 'roundRect' || !geom) sh.adj = { adj: 10000 };
        pending.push({ sh, text, sub: sub || [], role: o.role || 'node', color: o.color || textC, w: bw * (o.textW || 1), h: o.textH || bh, algn: o.algn || 'ctr', anchor: o.anchor || 'ctr' });
        out.push(sh);
        return sh;
      };
      const arrow = (x, y, aw, ah) => out.push({ id: L.uid('s'), name: 'Diagram arrow', type: 'shape', geom: 'rightArrow', x: X + x, y: Y + y, w: aw, h: ah, rot: 0, fill: Object.assign(fillOf(0), { a: 0.5 }), line: { t: 'none' } });
      const connector = (x1, y1, x2, y2) => out.push({ id: L.uid('s'), name: 'Diagram connector', type: 'line', geom: 'line', x: X + Math.min(x1, x2), y: Y + Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), rot: 0, flipH: x2 < x1, fill: { t: 'none' }, line: { c: fills.length ? fills[0].c : 'accent1', w: 1.5 } });
      const light = (i) => ({ t: 'solid', c: L.color.lighten(L.model.resolveColor(fillOf(i).c, ctx.design), 0.8), a: 1 });
      const n = tops.length;
      const kidsText = (p) => p.kids.map((k) => k.text);
      const fam = kind.replace(/^.*\//, '').replace(/#.*$/, '');
      if (/chevron/.test(fam)) {
        const bw = W / (n - (n - 1) * 0.15), bh = Math.min(H, bw * 0.45), y = (H - bh) / 2;
        tops.forEach((p, i) => box(i * bw * 0.85, y, bw, bh, p.text, kidsText(p), i === 0 && !/chevron2/.test(fam) ? 'homePlate' : 'chevron', { i, textW: 0.7 }));
      } else if (/^(hlist|lprocess|hprocess)/.test(fam)) {
        /* columns: a heading box with the children below it */
        const gap = Math.min(W / (n * 8), 16), bw = (W - gap * (n - 1)) / n;
        const grouped = /^lprocess/.test(fam);
        tops.forEach((p, i) => {
          const x = i * (bw + gap);
          if (grouped) {
            const kidsP = p.kids, hh = H * 0.2, ch = (H - hh - gap * (kidsP.length + 1)) / Math.max(1, kidsP.length);
            box(x, 0, bw, H, p.text, null, 'roundRect', { i, fill: light(i), color: 'dk1', role: 'head', anchor: 't', textH: hh });
            kidsP.forEach((k, j) => box(x + bw * 0.1, hh + gap + j * (ch + gap), bw * 0.8, ch, k.text, null, 'roundRect', { i, role: 'child' }));
          } else {
            /* the body box is drawn even when it has no text, as Office does */
            const hh = H * 0.2;
            box(x, 0, bw, hh, p.text, null, 'rect', { i, role: 'head' });
            box(x, hh, bw, H - hh, '', kidsText(p), 'rect', { i, fill: light(i), color: 'dk1', role: 'child', algn: 'l', anchor: 't', noLine: true });
          }
        });
      } else if (/^vlist(2|5)/.test(fam) || /^vlist/.test(fam) && tops.some((p) => p.kids.length)) {
        /* a heading per row with its children beside (vList5) or under (vList2) it */
        const gap = Math.min(H / (n * 8), 8), bh = (H - gap * (n - 1)) / n;
        tops.forEach((p, i) => {
          const y = i * (bh + gap);
          if (/^vlist5/.test(fam)) {
            const hw = W * 0.3;
            box(0, y, hw, bh, p.text, null, 'roundRect', { i, role: 'head' });
            if (p.kids.length) box(hw, y + bh * 0.08, W - hw, bh * 0.84, '', kidsText(p), 'rect', { i, fill: light(i), color: 'dk1', role: 'child', algn: 'l' });
          } else {
            const hh = p.kids.length ? bh * 0.45 : bh;
            box(0, y, W, hh, p.text, null, 'roundRect', { i, role: 'head', algn: 'l' });
            if (p.kids.length) box(W * 0.03, y + hh, W * 0.97, bh - hh, '', kidsText(p), 'rect', { i, fill: { t: 'none' }, noLine: true, color: 'dk1', role: 'child', algn: 'l', anchor: 't' });
          }
        });
      } else if (/process|arrow/.test(fam)) {
        const gap = Math.min(W / (n * 4), 40), bw = (W - gap * (n - 1)) / n, bh = Math.min(H, bw * 0.75), y = (H - bh) / 2;
        tops.forEach((p, i) => { box(i * (bw + gap), y, bw, bh, p.text, kidsText(p), 'roundRect', { i }); if (i < n - 1) arrow(i * (bw + gap) + bw + gap * 0.15, H / 2 - gap * 0.35, gap * 0.7, gap * 0.7); });
      } else if (/cycle|radial|target/.test(fam)) {
        const r = Math.min(W, H) / 2, bw = Math.min(r * 0.8, (2 * Math.PI * r * 0.55) / n), bh = bw * 0.6;
        tops.forEach((p, i) => { const a = -Math.PI / 2 + (i * 2 * Math.PI) / n; box(W / 2 + (r - bw / 2) * Math.cos(a) - bw / 2, H / 2 + (r - bh / 2) * Math.sin(a) - bh / 2, bw, bh, p.text, null, 'ellipse', { i }); });
      } else if (/hierarchy|orgchart/.test(fam)) {
        const levels = [];
        const pos = new Map();
        const walk = (p, d) => { (levels[d] = levels[d] || []).push(p); for (const k of p.kids) walk(k, d + 1); };
        for (const t of tops) walk(t, 0);
        const lh = H / levels.length;
        levels.forEach((row, d) => { const bw = Math.min((W / row.length) * 0.85, W * 0.3), gap = (W - bw * row.length) / (row.length + 1); row.forEach((p, i) => { const x = gap + i * (bw + gap), y = d * lh + lh * 0.12; box(x, y, bw, lh * 0.7, p.text, null, 'rect', { i: d, role: 'node' }); pos.set(p, { cx: x + bw / 2, top: y, bottom: y + lh * 0.7 }); }); });
        /* connectors: down from the parent, across, and down to each child */
        for (const [p, pp] of pos) for (const k of p.kids) { const kp = pos.get(k); if (!kp) continue; const mid = (pp.bottom + kp.top) / 2; connector(pp.cx, pp.bottom, pp.cx, mid); connector(pp.cx, mid, kp.cx, mid); connector(kp.cx, mid, kp.cx, kp.top); }
        out.unshift(...out.splice(out.length - [...pos].reduce((a, [p]) => a + p.kids.filter((k) => pos.has(k)).length * 3, 0)));
      } else if (/pyramid/.test(fam)) {
        const bh = H / n;
        tops.forEach((p, i) => { const bw = W * (i + 1) / n; box((W - bw) / 2, i * bh, bw, bh * 0.94, p.text, kidsText(p), 'rect', { i }); });
      } else if (/matrix/.test(fam)) {
        const cols = 2, rows = Math.ceil(n / 2), gap = Math.min(W, H) * 0.04, bw = (W - gap) / cols, bh = (H - gap * (rows - 1)) / rows;
        tops.forEach((p, i) => box((i % cols) * (bw + gap), Math.floor(i / cols) * (bh + gap), bw, bh, p.text, kidsText(p), 'roundRect', { i, algn: p.kids.length ? 'l' : 'ctr' }));
      } else if (/^(default|blist|snake)/.test(fam) || (!/^vlist/.test(fam) && n > 3 && tops.every((p) => !p.kids.length))) {
        /* Basic Block List: rows of 3:2 boxes, the last row centred */
        let best = null;
        for (let cols = 1; cols <= n; cols++) {
          const rows = Math.ceil(n / cols);
          const bw = Math.min(W / (cols + (cols - 1) * 0.1), (H / (rows + (rows - 1) * 0.1)) / 0.6);
          if (!best || bw > best.bw) best = { cols, rows, bw };
        }
        const { cols, rows, bw } = best, bh = bw * 0.6, gx = bw * 0.1, gy = bh * 0.1 / 0.6;
        const top0 = (H - (rows * bh + (rows - 1) * gy)) / 2;
        tops.forEach((p, i) => {
          const r = Math.floor(i / cols), inRow = Math.min(cols, n - r * cols), c = i % cols;
          const left0 = (W - (inRow * bw + (inRow - 1) * gx)) / 2;
          box(left0 + c * (bw + gx), top0 + r * (bh + gy), bw, bh, p.text, kidsText(p), 'rect', { i, algn: p.kids.length ? 'l' : 'ctr' });
        });
      } else {
        const gap = Math.min(H / (n * 6), 8), bh = (H - gap * (n - 1)) / n;
        tops.forEach((p, i) => box(0, i * (bh + gap), W, bh, p.text, kidsText(p), 'roundRect', { i }));
      }
      /* the diagram's own background and outline (dgm:bg, dgm:whole) */
      const dbg = kid(dm, 'bg'), dwhole = kid(dm, 'whole');
      const bgFill = dbg ? fillIn(dbg, ctx) : undefined, bgLine = dwhole ? line(kid(dwhole, 'ln'), ctx) : undefined;
      if ((bgFill && bgFill.t !== 'none') || (bgLine && bgLine.t !== 'none')) out.unshift({ id: L.uid('s'), name: 'Diagram background', type: 'shape', geom: 'rect', x: X, y: Y, w: W, h: H, rot: 0, fill: bgFill || { t: 'none' }, line: bgLine && bgLine.t !== 'none' ? bgLine : { t: 'none' } });
      /* text: a common size per role */
      const roles = new Map();
      for (const it of pending) { const r = roles.get(it.role) || []; r.push(it); roles.set(it.role, r); }
      const sizes = new Map();
      for (const [role, items] of roles) sizes.set(role, fitSize(items.map((it) => ({ lines: [it.text].concat(it.sub).filter((t) => t), w: it.w, h: it.h, bullets: it.sub.length > 0 })), role === 'child' ? 24 : 36));
      for (const it of pending) {
        const sz = Math.round(sizes.get(it.role) * 2) / 2;
        const lines = [];
        if (it.text) lines.push(L.txt.para(it.text, { algn: it.algn }, { sz, color: it.color }));
        for (const t of it.sub) lines.push(L.txt.para(t, { algn: 'l', marL: sz * 0.9, indent: -sz * 0.9, bu: { t: 'char', ch: '•' } }, { sz: Math.round(sz * (it.text ? 0.85 : 1) * 2) / 2, color: it.color }));
        if (!lines.length) lines.push(L.txt.para('', { algn: 'ctr' }, { color: it.color }));
        it.sh.tx = L.txt.body(lines, { anchor: it.anchor === 't' ? 't' : 'ctr', ins: [5, 4, 5, 4] });
      }
      return Object.assign(base, { type: 'group', kids: out, name: base.name || 'Diagram' });
    }

    /* ---- slides ---- */
    const sldIds = kids(kid(presX, 'sldIdLst'), 'sldId');
    const slideRefs = sldIds.map(el => ({ el, path: presRels[rid(el)]?.target })).filter(s => s.path);
    const pathToId = new Map();
    const LAYOUT_MAP = { title: 'title', titleOnly: 'titleOnly', tx: 'text', twoColTx: 'twoText', blank: 'blank', objOnly: 'contentOnly', obj: 'content', twoObj: 'twoContent', fourObj: 'fourContent', txAndObj: 'textContent', objAndTx: 'contentText', txAndTwoObj: 'textTwoContent', twoObjAndTx: 'twoContentText', tbl: 'table', chart: 'chart', txOverObj: 'textOverContent', objOverTx: 'contentOverText', vertTx: 'vertText', vertTitleAndTx: 'vertTitleText', objAndTwoObj: 'contentTwoContent', twoObjAndObj: 'twoContentContent', secHead: 'titleOnly', twoTxTwoObj: 'twoContent', objTx: 'textContent', picTx: 'textContent', txAndChart: 'textContent', chartAndTx: 'contentText', txAndClipArt: 'textContent', clipArtAndTx: 'contentText', txAndMedia: 'textContent', mediaAndTx: 'contentText', dgm: 'content', txOverObj2: 'textOverContent' };
    let n = 0;
    for (const sourceSlide of slideRefs) {
      const sp = sourceSlide.path;
      n++;
      if (opts.progress) opts.progress(n, slideRefs.length);
      const sx = await xml(sp);
      if (!sx) { if (file(sp)) damaged.push(sp); continue; }
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
      /* ActiveX controls are shown by their preview pictures (they are not run here) */
      const ctrls = kid(cSld, 'controls');
      if (ctrls) for (const c of kids(ctrls, 'control')) {
        try {
          const pic = kid(c, 'pic');
          let sh = pic ? await shapeFrom(pic, ctx, {}) : null;
          if (!sh && at(c, 'spid')) {
            /* the 2007 form: the preview lives in the legacy VML drawing */
            const vr = Object.values(ctx.rels).find((r) => r.type === 'vmlDrawing' && !r.external);
            const vml = vr ? await getText(vr.target) : null;
            const m = vml && new RegExp('<v:shape\\b[^>]*\\b(?:o:spid|id)="_x0000_s' + String(at(c, 'spid')).replace(/[^\w]/g, '') + '"[\\s\\S]*?</v:shape>').exec(vml);
            const im = m && /<v:imagedata\b[^>]*o:relid="([^"]+)"/.exec(m[0]);
            const st = m && /style="([^"]*)"/.exec(m[0]);
            if (im && st) {
              const vrels = await rels(vr.target);
              const r = vrels[im[1]];
              const mid = r && !r.external ? await loadMedia(r.target) : null;
              const css = {}; st[1].split(';').forEach((kv) => { const [k, v] = kv.split(':'); if (k && v) css[k.trim()] = v.trim(); });
              const ptv = (v) => (/pt$/.test(v) ? parseFloat(v) : /in$/.test(v) ? parseFloat(v) * 72 : /px$/.test(v) ? parseFloat(v) * 0.75 : parseFloat(v) || 0);
              if (mid) sh = { id: L.uid('s'), name: at(c, 'name') || 'Control', type: 'image', geom: 'rect', media: mid, x: ptv(css['margin-left'] || css.left), y: ptv(css['margin-top'] || css.top), w: ptv(css.width), h: ptv(css.height), rot: 0, crop: { l: 0, t: 0, r: 0, b: 0 }, line: { t: 'none' }, lockAspect: true, img: {} };
            }
          }
          if (sh && !Array.isArray(sh)) { if (!sh.name) sh.name = at(c, 'name') || 'Control'; slide.shapes.push(sh); }
        } catch (e) { /* skip the control */ }
      }
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
      if (original(sp)) L.preserve.slide(slide, original(sp).tree, pkg, sp, at(sourceSlide.el, 'id'), sx);
      if (lay && design.keep) design.keep.layouts[slide.layout + (slide.lkey ? '|' + slide.lkey : '')] = lay.path;
      if (nr && slide.keep) slide.keep.notes = { part: nr.target, text: slide.notes };
      pathToId.set(sp, slide.id);
      pres.slides.push(slide);
      } catch (e) {
        /* keep going: one unreadable slide must not lose the whole deck */
        console.warn('slide skipped', sp, e);
        damaged.push(sp);
      }
    }
    // Unused masters and layouts are still part of the document's design library.
    for (const ref of Object.values(presRels).filter(r => r.type === 'slideMaster')) {
      try {
        const master = await loadMaster(ref.target);
        for (const rel of Object.values(master.mr).filter(r => r.type === 'slideLayout')) {
          try { await loadLayout(rel.target); } catch (_) { damaged.push(rel.target); }
        }
      } catch (_) { damaged.push(ref.target); }
    }
    for (const layout of layouts.values()) {
      const key = LAYOUT_MAP[layout.type] || (layout.phs.some(p => p.ph.type === 'ctrTitle') ? 'title' : layout.phs.some(p => p.ph.type === 'title') ? (layout.phs.length > 1 ? 'text' : 'titleOnly') : 'blank');
      layout.master.design.keep.layouts[key + '|' + layout.lkey] = layout.path;
      layout.master.design.keep.layouts[key] ||= layout.path;
      const record = layout.master.design.keep.layoutParts.find(r => r.part === layout.path); if (record) record.key = key;
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
    /* saved again with the same password to open, as PowerPoint does */
    if (password) pres.password = password;
    await applyDuotones(pres);
    /* pictures with a transparent colour get a display copy; the original is what is saved */
    const clearPics = [];
    const pick = (sh) => { if (sh.type === 'image' && sh.img && sh.img.clear && sh.media) clearPics.push(sh); return true; };
    for (const s of pres.slides) L.model.walk(s.shapes, pick);
    for (const d of Object.values(pres.designs)) [d.deco, d.titleDeco].concat(d.layoutDecos ? Object.values(d.layoutDecos) : []).forEach((list) => list && L.model.walk(list, pick));
    for (const sh of clearPics) { const v = await L.media.withTransparent(sh.media, sh.img.clear, 3); if (v) sh.img.view = v; }
    pres.keep = { values: L.preserve.values(pres), media: Object.fromEntries(Array.from(mediaCache, ([part, id]) => [id, part])) };
    await L.frames.seal(pres, mediaCache);
    L.preserve.seal(pres);
    L.preserve.readSlideLists(pres);
    pres.ooxmlFormat = L.opc.variant(pres.pkg, 'pptx');
    return pres;
  }

  /** duotone pictures (theme background textures) are recoloured once, after loading */
  async function applyDuotones(root) {
    const jobs = [];
    const seen = new Set();
    const scan = (o) => {
      if (!o || typeof o !== 'object' || seen.has(o)) return;
      seen.add(o);
      if (o.t === 'img' && o.duotone && o.media) jobs.push(o);
      for (const k in o) { const v = o[k]; if (v && typeof v === 'object') scan(v); }
    };
    scan(root);
    const cache = new Map();
    for (const f of jobs) {
      const key = f.media + '|' + f.duotone.join('|');
      if (!cache.has(key)) cache.set(key, duotone(f.media, f.duotone));
      const id = await cache.get(key);
      if (id) f.media = id;
      delete f.duotone;
    }
  }
  async function duotone(mediaId, cols) {
    try {
      if (typeof document === 'undefined') return null;
      const img = await L.loadImage(L.media.url(mediaId));
      const k = Math.min(1, 1024 / Math.max(img.naturalWidth || 1, img.naturalHeight || 1));
      const w = Math.max(1, Math.round(img.naturalWidth * k)), h = Math.max(1, Math.round(img.naturalHeight * k));
      const cv = document.createElement('canvas');
      cv.width = w; cv.height = h;
      const g = cv.getContext('2d');
      g.drawImage(img, 0, 0, w, h);
      const d = g.getImageData(0, 0, w, h);
      const [a, b] = cols.map((c) => L.color.hexToRgb(L.model.resolveColor(c)));
      const px = d.data;
      for (let i = 0; i < px.length; i += 4) {
        const t = (0.299 * px[i] + 0.587 * px[i + 1] + 0.114 * px[i + 2]) / 255;
        px[i] = a[0] + (b[0] - a[0]) * t; px[i + 1] = a[1] + (b[1] - a[1]) * t; px[i + 2] = a[2] + (b[2] - a[2]) * t;
      }
      g.putImageData(d, 0, 0);
      const blob = await new Promise((res) => cv.toBlob(res, 'image/png'));
      if (!blob) return null;
      const id = L.media.add(blob, 'duotone.png');
      const o = L.media.get(mediaId).size;
      L.media.get(id).size = { w, h, dpiX: ((o && o.dpiX) || 96) * k, dpiY: ((o && o.dpiY) || 96) * k };
      return id;
    } catch (e) { return null; }
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
