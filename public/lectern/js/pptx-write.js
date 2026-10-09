/* Lectern — PresentationML (.pptx/.ppsx/.potx) writer. */
(function () {
  'use strict';
  const L = window.L;
  const K = L.opc;
  const X = L.xesc;
  const emu = L.pt2emu;
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const RT = (n) => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + n;
  const CT = {
    pres: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    show: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
    tmpl: 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
    slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    layout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    master: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
    notesMaster: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
    notes: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    presProps: 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
    viewProps: 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
    tableStyles: 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
    core: 'application/vnd.openxmlformats-package.core-properties+xml',
    app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  };
  const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const NSDECL = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;
  const guid = () => '{' + 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => { const r = (Math.random() * 16) | 0; return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16).toUpperCase(); }) + '}';
  const pct = (v) => Math.round((v || 0) * 1000); /* % → 1/1000 % */

  const Rels = K.Rels;

  /* ---------- colours / fills / lines ---------- */
  const SCHEME_OK = new Set(['bg1', 'tx1', 'bg2', 'tx2', 'dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink', 'phClr']);
  function clr(c, alpha, design) {
    const a = alpha != null && alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : '';
    if (c && c[0] !== '#' && SCHEME_OK.has(c)) return a ? `<a:schemeClr val="${c}">${a}</a:schemeClr>` : `<a:schemeClr val="${c}"/>`;
    const hex = L.model.resolveColor(c || '#000000', design).replace('#', '').toUpperCase();
    return a ? `<a:srgbClr val="${hex}">${a}</a:srgbClr>` : `<a:srgbClr val="${hex}"/>`;
  }
  function fillXML(f, ctx, design) {
    if (!f) return '';
    switch (f.t) {
      case 'none': return '<a:noFill/>';
      case 'solid': return `<a:solidFill>${clr(f.c, f.a, design)}</a:solidFill>`;
      case 'grad': {
        let stops = (f.stops || []).slice().sort((a, b) => a.p - b.p);
        if (stops.length < 2) stops = [stops[0] || { p: 0, c: '#FFFFFF' }, Object.assign({}, stops[0] || { c: '#000000' }, { p: 1 })];
        const gs = stops.map((s) => `<a:gs pos="${Math.round(L.clamp(s.p, 0, 1) * 100000)}">${clr(s.c, s.a, design)}</a:gs>`).join('');
        const shade = f.path && f.path !== 'lin'
          ? `<a:path path="${f.path === 'rect' ? 'rect' : f.path === 'shape' ? 'shape' : 'circle'}">${(() => { const [fx, fy] = f.focus || [0.5, 0.5]; return `<a:fillToRect l="${Math.round(fx * 100000)}" t="${Math.round(fy * 100000)}" r="${Math.round((1 - fx) * 100000)}" b="${Math.round((1 - fy) * 100000)}"/>`; })()}</a:path>`
          : `<a:lin ang="${Math.round((((f.ang || 0) % 360) + 360) % 360 * 60000)}" scaled="1"/>`;
        return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst>${shade}</a:gradFill>`;
      }
      case 'patt': return `<a:pattFill prst="${f.prst || 'pct50'}"><a:fgClr>${clr(f.fg || '#000000', null, design)}</a:fgClr><a:bgClr>${clr(f.bg || '#FFFFFF', null, design)}</a:bgClr></a:pattFill>`;
      case 'img': {
        const rid = ctx.media(f.media);
        if (!rid) return '<a:noFill/>';
        return `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rid}"${f.a != null && f.a < 1 ? `><a:alphaModFix amt="${Math.round(f.a * 100000)}"/></a:blip>` : '/>'}${f.crop && !f.tile ? `<a:srcRect l="${Math.round(f.crop.l * 100000)}" t="${Math.round(f.crop.t * 100000)}" r="${Math.round(f.crop.r * 100000)}" b="${Math.round(f.crop.b * 100000)}"/>` : '<a:srcRect/>'}${f.tile ? tileXML(f.tileOpts) : f.fillRect ? `<a:stretch><a:fillRect l="${Math.round(f.fillRect.l * 100000)}" t="${Math.round(f.fillRect.t * 100000)}" r="${Math.round(f.fillRect.r * 100000)}" b="${Math.round(f.fillRect.b * 100000)}"/></a:stretch>` : '<a:stretch><a:fillRect/></a:stretch>'}</a:blipFill>`;
      }
      default: return '';
    }
  }
  function lineXML(ln, design, tag) {
    tag = tag || 'a:ln';
    if (!ln) return '';
    if (ln.t === 'none' || !ln.c) return `<${tag} w="${emu(ln && ln.w != null ? ln.w : 0.75)}"><a:noFill/></${tag}>`;
    const attrs = [`w="${emu(ln.w == null ? 0.75 : ln.w)}"`];
    if (ln.cap) attrs.push(`cap="${ln.cap}"`);
    if (ln.cmpd && ln.cmpd !== 'sng') attrs.push(`cmpd="${ln.cmpd}"`);
    let x = `<${tag} ${attrs.join(' ')}><a:solidFill>${clr(ln.c, ln.a, design)}</a:solidFill>`;
    if (ln.dash && ln.dash !== 'solid') x += `<a:prstDash val="${ln.dash}"/>`;
    if (ln.join === 'round') x += '<a:round/>'; else if (ln.join === 'bevel') x += '<a:bevel/>';
    const end = (e, name) => (e && e.type && e.type !== 'none' ? `<a:${name} type="${e.type}" w="${e.w || 'med'}" len="${e.len || 'med'}"/>` : '');
    x += end(ln.head, 'headEnd') + end(ln.tail, 'tailEnd');
    return x + `</${tag}>`;
  }
  function shadowXML(sh, design) {
    if (!sh) return '';
    const dx = sh.dx == null ? 3 : sh.dx, dy = sh.dy == null ? 3 : sh.dy;
    const dist = emu(Math.hypot(dx, dy));
    const dir = Math.round(((((Math.atan2(dy, dx) * 180) / Math.PI) + 360) % 360) * 60000);
    return `<a:effectLst><a:outerShdw blurRad="${emu(sh.blur || 0)}" dist="${dist}" dir="${dir}" algn="tl" rotWithShape="0">${clr(sh.c || '#000000', sh.a == null ? 0.5 : sh.a, design)}</a:outerShdw></a:effectLst>`;
  }
  function xfrm(sh, tag, extra) {
    const a = [];
    // Presentation graphic frames support position and extent only.
    if (tag !== 'p:xfrm') {
      if (sh.rot) a.push(`rot="${Math.round((((sh.rot % 360) + 360) % 360) * 60000)}"`);
      if (sh.flipH) a.push('flipH="1"');
      if (sh.flipV) a.push('flipV="1"');
    }
    return `<${tag || 'a:xfrm'}${a.length ? ' ' + a.join(' ') : ''}><a:off x="${emu(sh.x)}" y="${emu(sh.y)}"/><a:ext cx="${Math.max(0, emu(sh.w))}" cy="${Math.max(0, emu(sh.h))}"/>${extra || ''}</${tag || 'a:xfrm'}>`;
  }
  function geomXML(sh) {
    if (sh.geom === 'custom' && sh.path) {
      const paths = sh.path.paths.map((p) => {
        const sx = L.EMU, sy = L.EMU;
        const pt = (x, y) => `<a:pt x="${Math.round(x * sx)}" y="${Math.round(y * sy)}"/>`;
        const cmds = p.cmds.map((c) => {
          switch (c[0]) {
            case 'M': return `<a:moveTo>${pt(c[1], c[2])}</a:moveTo>`;
            case 'L': return `<a:lnTo>${pt(c[1], c[2])}</a:lnTo>`;
            case 'C': return `<a:cubicBezTo>${pt(c[1], c[2])}${pt(c[3], c[4])}${pt(c[5], c[6])}</a:cubicBezTo>`;
            case 'Q': return `<a:quadBezTo>${pt(c[1], c[2])}${pt(c[3], c[4])}</a:quadBezTo>`;
            case 'A': return `<a:arcTo wR="${Math.round(c[1] * sx)}" hR="${Math.round(c[2] * sy)}" stAng="${Math.round(c[3] * 60000)}" swAng="${Math.round(c[4] * 60000)}"/>`;
            case 'Z': return '<a:close/>';
            default: return '';
          }
        }).join('');
        return `<a:path w="${Math.round(p.w * sx)}" h="${Math.round(p.h * sy)}"${p.fill === 'none' ? ' fill="none"' : ''}${p.stroke === false ? ' stroke="0"' : ''}>${cmds}</a:path>`;
      }).join('');
      return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`;
    }
    const g = sh.geom || 'rect';
    const av = L.opc.presetAdjust(g, sh.adj);
    return `<a:prstGeom prst="${g}"><a:avLst>${av}</a:avLst></a:prstGeom>`;
  }

  /* ---------- text ---------- */
  const fontRef = (f) => (f === '+mj' ? '+mj-lt' : f === '+mn' ? '+mn-lt' : f);
  function rPrXML(r, tag, ctx, design) {
    tag = tag || 'a:rPr';
    const a = ['lang="en-US"'];
    if (r.sz) a.push(`sz="${Math.round(r.sz * 100)}"`);
    if (r.b != null) a.push(`b="${r.b ? 1 : 0}"`);
    if (r.i != null) a.push(`i="${r.i ? 1 : 0}"`);
    if (r.u) a.push(`u="${r.u === true ? 'sng' : r.u}"`);
    if (r.strike) a.push(`strike="${r.strike === true ? 'sngStrike' : r.strike}"`);
    if (r.cap) a.push(`cap="${r.cap}"`);
    if (r.spc) a.push(`spc="${Math.round(r.spc * 100)}"`);
    if (r.base) a.push(`baseline="${Math.round(r.base)}"`);
    a.push('dirty="0"');
    let kids = '';
    if (r.ln) kids += lineXML(r.ln, design);
    if (r.fill) kids += fillXML(r.fill, ctx, design);
    else if (r.color) kids += `<a:solidFill>${clr(r.color, null, design)}</a:solidFill>`;
    if (r.shd && r.shdX) kids += shadowXML(r.shdX, design);
    else if (r.shd) kids += `<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw></a:effectLst>`;
    if (r.hl) kids += `<a:highlight>${clr(r.hl, null, design)}</a:highlight>`;
    if (r.font) { const f = X(fontRef(r.font)); kids += `<a:latin typeface="${f}"/><a:ea typeface="${f}"/><a:cs typeface="${f}"/>`; }
    if (r.link && ctx && tag === 'a:rPr') kids += hlinkXML(r.link, ctx);
    const xml = kids ? `<${tag} ${a.join(' ')}>${kids}</${tag}>` : `<${tag} ${a.join(' ')}/>`;
    return L.properties.text(r, ctx, xml, 'run', tag);
  }
  function hlinkXML(link, ctx, tag) {
    tag = tag || 'a:hlinkClick';
    if (!link) return '';
    if (link.url) { const rid = ctx.rels.add(RT('hyperlink'), link.url, true); return `<${tag} r:id="${rid}"${link.tip ? ` tooltip="${X(link.tip)}"` : ''}/>`; }
    if (link.slide) {
      const idx = ctx.slideIndex(link.slide);
      if (idx >= 0) { const rid = ctx.rels.add(RT('slide'), ctx.slideTarget ? ctx.slideTarget(idx) : `slide${idx + 1}.xml`); return `<${tag} r:id="${rid}" action="ppaction://hlinksldjump"/>`; }
      return '';
    }
    const jump = { next: 'nextslide', prev: 'previousslide', first: 'firstslide', last: 'lastslide', lastViewed: 'lastslideviewed', end: 'endshow' }[link.action];
    return jump ? `<${tag} r:id="" action="ppaction://hlinkshowjump?jump=${jump}"/>` : '';
  }
  function buXML(bu, design) {
    if (!bu) return '';
    let x = '';
    if (bu.c) x += `<a:buClr>${clr(bu.c, null, design)}</a:buClr>`;
    if (bu.sz && bu.sz !== 1) x += `<a:buSzPct val="${Math.round(bu.sz * 100000)}"/>`;
    if (bu.t === 'none') return x + '<a:buNone/>';
    if (bu.t === 'num') return x + (bu.font ? `<a:buFont typeface="${X(bu.font)}"/>` : '<a:buFont typeface="+mj-lt"/>') + `<a:buAutoNum type="${bu.scheme || 'arabicPeriod'}"${bu.start && bu.start !== 1 ? ` startAt="${bu.start}"` : ''}/>`;
    if (bu.t === 'char') return x + `<a:buFont typeface="${X(bu.font || 'Arial')}"/><a:buChar char="${X(bu.ch || '•')}"/>`;
    return x;
  }
  const spc = (s, tag) => (s ? (s.pts != null ? `<${tag}><a:spcPts val="${Math.round(s.pts * 100)}"/></${tag}>` : `<${tag}><a:spcPct val="${pct(s.pct)}"/></${tag}>`) : '');
  /** paragraph or level properties */
  function pPrXML(pp, tag, lvl, ctx, design, withDefRPr) {
    const a = [];
    if (pp.marL != null) a.push(`marL="${emu(pp.marL)}"`);
    if (lvl) a.push(`lvl="${lvl}"`);
    if (pp.indent != null) a.push(`indent="${emu(pp.indent)}"`);
    if (pp.algn) a.push(`algn="${pp.algn}"`);
    let kids = spc(pp.lnSpc, 'a:lnSpc') + spc(pp.spcBef, 'a:spcBef') + spc(pp.spcAft, 'a:spcAft') + buXML(pp.bu, design);
    if (withDefRPr && pp.rPr && Object.keys(pp.rPr).length) kids += rPrXML(pp.rPr, 'a:defRPr', ctx, design);
    const xml = !a.length && !kids ? '' : kids ? `<${tag}${a.length ? ' ' + a.join(' ') : ''}>${kids}</${tag}>` : `<${tag} ${a.join(' ')}/>`;
    return L.properties.text(pp, ctx, xml, 'para', tag, lvl);
  }
  function lstStyleXML(lst, ctx, design) {
    if (!lst) return '<a:lstStyle/>';
    let x = '';
    for (let i = 0; i < 9; i++) if (lst[i]) x += pPrXML(lst[i], `a:lvl${i + 1}pPr`, 0, ctx, design, true);
    return x ? `<a:lstStyle>${x}</a:lstStyle>` : '<a:lstStyle/>';
  }
  function parasXML(ps, ctx, design) {
    return ps.map((p) => {
      let x = '<a:p>';
      x += pPrXML(p.pp || {}, 'a:pPr', p.lvl || 0, ctx, design, false) || '';
      for (const r of p.rs) {
        const props = L.txt.runProps(r);
        if (r.fld) { x += `<a:fld id="${ctx.fieldId ? ctx.fieldId() : guid()}" type="${X(r.fld)}">${rPrXML(props, 'a:rPr', ctx, design)}<a:t>${X(r.t || '')}</a:t></a:fld>`; continue; }
        const parts = String(r.t).split('\n');
        parts.forEach((t, i) => {
          if (i) x += `<a:br>${rPrXML(Object.assign({}, props, { link: undefined }), 'a:rPr', ctx, design)}</a:br>`;
          if (t || r.keep && parts.length === 1) x += `<a:r>${rPrXML(props, 'a:rPr', ctx, design)}<a:t>${X(t)}</a:t></a:r>`;
        });
      }
      const end = p.end || (p.rs.length ? Object.fromEntries(Object.entries(L.txt.runProps(p.rs[p.rs.length - 1])).filter(([key]) => key !== 'keep')) : {});
      x += rPrXML(Object.assign({}, end, { link: undefined }), 'a:endParaRPr', ctx, design);
      return x + '</a:p>';
    }).join('');
  }
  function bodyPrXML(tx, extra) {
    const a = [];
    const ins = tx.ins || [7.2, 3.6, 7.2, 3.6];
    a.push(`wrap="${tx.wrap === false ? 'none' : 'square'}"`);
    a.push(`lIns="${emu(ins[0])}" tIns="${emu(ins[1])}" rIns="${emu(ins[2])}" bIns="${emu(ins[3])}"`);
    a.push(`anchor="${tx.anchor || 't'}"`);
    if (tx.anchorCtr) a.push('anchorCtr="1"');
    if (tx.vert) a.push(`vert="${tx.vert}"`);
    if (tx.rot) a.push(`rot="${Math.round(tx.rot * 60000)}"`);
    if (tx.cols > 1) a.push(`numCol="${tx.cols}" spcCol="${emu(tx.colGap || 0)}"`);
    if (extra && extra.attrs) a.push(extra.attrs);
    a.push('rtlCol="0"');
    let kids = extra && extra.warp ? extra.warp : '';
    if (tx.autofit === 'shape') kids += '<a:spAutoFit/>';
    else if (tx.autofit === 'norm') kids += tx.fontScale && tx.fontScale < 1 ? `<a:normAutofit fontScale="${Math.round(tx.fontScale * 100000)}"${tx.lnSpcRed ? ` lnSpcReduction="${Math.round(tx.lnSpcRed * 100000)}"` : ''}/>` : '<a:normAutofit/>';
    return `<a:bodyPr ${a.join(' ')}>${kids}</a:bodyPr>`;
  }
  function txBodyXML(tx, ctx, design, tag, extra) {
    tag = tag || 'p:txBody';
    const ps = tx.ps && tx.ps.length ? tx.ps : [L.txt.para('')];
    return `<${tag}>${L.properties.text(tx, ctx, bodyPrXML(tx, extra), 'body', 'a:bodyPr')}${L.properties.list(tx, ctx, lstStyleXML(tx.lst, ctx, design))}${parasXML(ps, ctx, design)}</${tag}>`;
  }

  /* ---------- shapes ---------- */
  function cNvPr(id, sh, ctx) {
    const kids = sh.link ? hlinkXML(sh.link, ctx) : '';
    const attrs = `id="${id}" name="${X(sh.name || 'Shape ' + id)}"${sh.alt ? ` descr="${X(sh.alt)}"` : ''}${sh.hidden ? ' hidden="1"' : ''}`;
    return L.preserve.nonVisual(sh, ctx, kids ? `<p:cNvPr ${attrs}>${kids}</p:cNvPr>` : `<p:cNvPr ${attrs}/>`);
  }
  function phXML(ph) {
    if (!ph) return '<p:nvPr/>';
    const a = [];
    if (ph.type && ph.type !== 'obj') a.push(`type="${ph.type}"`);
    if (ph.type === 'dt') a.push('sz="half"');
    if (ph.type === 'ftr' || ph.type === 'sldNum') a.push('sz="quarter"');
    if (ph.idx) a.push(`idx="${ph.idx}"`);
    return `<p:nvPr><p:ph${a.length ? ' ' + a.join(' ') : ''}/></p:nvPr>`;
  }
  function shapeXML(sh, ctx, design) {
    if (ctx.designTemplate) return '';
    const id = ctx.nextId(sh.id);
    if (ctx.writer && sh.keep?.frame) {
      const kept = L.frames.emit(sh, ctx);
      if (kept != null) return kept;
      (ctx.convertedFrames ||= new Map()).set(String(id), { key: sh.keep.frame.fragment.key, label: sh.keep.frame.label });
    }
    if (ctx.writer && sh.keep?.media) {
      const kept = L.preserve.emitShape(sh, ctx, () => shapeXML({ ...sh, keep: null }, ctx, design));
      if (kept != null) return kept;
    }
    if (ctx.writer && sh.keep?.designFrame) {
      const kept = L.designs.shape(sh, ctx);
      if (kept != null) return kept;
    }
    return L.properties.apply(sh, ctx, shapeContentXML(sh, id, ctx, design));
  }
  function shapeContentXML(sh, id, ctx, design) {
    switch (sh.type) {
      case 'group': {
        const kids = sh.kids.map((k) => shapeXML(k, { ...ctx, inGroup: true }, design)).join('');
        const ext = `<a:chOff x="${emu(sh.x)}" y="${emu(sh.y)}"/><a:chExt cx="${emu(sh.w)}" cy="${emu(sh.h)}"/>`;
        return `<p:grpSp><p:nvGrpSpPr>${cNvPr(id, sh, ctx)}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${xfrm(sh, 'a:xfrm', ext)}</p:grpSpPr>${kids}</p:grpSp>`;
      }
      case 'line': {
        const simple = sh.keep?.element === 'sp' || !!sh.tx, tag = simple ? 'sp' : 'cxnSp', nv = simple ? 'SpPr' : 'CxnSpPr';
        return `<p:${tag}><p:nv${nv}>${cNvPr(id, sh, ctx)}<p:cNv${nv}/><p:nvPr/></p:nv${nv}><p:spPr>${xfrm(sh)}<a:prstGeom prst="${sh.geom || 'line'}"><a:avLst/></a:prstGeom>${lineXML(sh.line || { c: 'tx1', w: 0.75 }, design)}${shadowXML(sh.shadow, design)}</p:spPr>${sh.tx ? txBodyXML(sh.tx, ctx, design) : ''}</p:${tag}>`;
      }
      case 'image': {
        /* a picture whose image is missing or linked from outside keeps its frame (PowerPoint shows its own placeholder) */
        const rid = ctx.media(sh.media) || (sh.linkUrl ? ctx.rels.add(RT('image'), sh.linkUrl, true) : null);
        const blipAttr = !rid ? '' : ctx.media(sh.media) ? ` r:embed="${rid}"` : ` r:link="${rid}"`;
        const svgRid = rid && ctx.svg ? ctx.svg(sh.svgMedia || sh.media) : null;
        const c = sh.crop || {};
        const crop = c.l || c.t || c.r || c.b ? `<a:srcRect${c.l ? ` l="${Math.round(c.l * 100000)}"` : ''}${c.t ? ` t="${Math.round(c.t * 100000)}"` : ''}${c.r ? ` r="${Math.round(c.r * 100000)}"` : ''}${c.b ? ` b="${Math.round(c.b * 100000)}"` : ''}/>` : '';
        const im = sh.img || {};
        let fx = '';
        if (im.alpha != null && im.alpha < 1) fx += `<a:alphaModFix amt="${Math.round(im.alpha * 100000)}"/>`;
        if (im.clear) fx += `<a:clrChange><a:clrFrom>${clr(im.clear, null, design)}</a:clrFrom><a:clrTo>${clr(im.clear, 0, design)}</a:clrTo></a:clrChange>`;
        if (im.mode === 'gray') fx += '<a:grayscl/>';
        if (im.mode === 'bw') fx += '<a:biLevel thresh="50000"/>';
        if (im.mode === 'wash') fx += '<a:lum bright="70000" contrast="-70000"/>';
        else if (im.bright || im.contrast) fx += `<a:lum bright="${Math.round((im.bright || 0) * 100000)}" contrast="${Math.round((im.contrast || 0) * 100000)}"/>`;
        if (svgRid) fx += `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${svgRid}"/></a:ext></a:extLst>`;
        const blip = fx ? `<a:blip${blipAttr}>${fx}</a:blip>` : `<a:blip${blipAttr}/>`;
        return `<p:pic><p:nvPicPr>${cNvPr(id, sh, ctx)}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>${phXML(sh.ph)}</p:nvPicPr><p:blipFill>${blip}${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(sh)}${geomXML(sh)}${sh.line && sh.line.t !== 'none' && sh.line.c ? lineXML(sh.line, design) : ''}${shadowXML(sh.shadow, design)}</p:spPr></p:pic>`;
      }
      case 'table': return tableXML(sh, id, ctx, design);
      case 'chart': return chartFrameXML(sh, id, ctx, design);
      case 'wordart': return wordartXML(sh, id, ctx, design);
      default: {
        const isTx = sh.type === 'text' && !sh.ph;
        const spLocks = sh.ph ? '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' : isTx ? '<p:cNvSpPr txBox="1"/>' : '<p:cNvSpPr/>';
        const fill = sh.fill ? fillXML(sh.fill, ctx, design) : '';
        const ln = sh.line ? lineXML(sh.line, design) : '';
        const geom = !sh.ph || (sh.geom && sh.geom !== 'rect') || (sh.fill && sh.fill.t !== 'none') ? geomXML(sh) : '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
        const tx = sh.tx ? txBodyXML(sh.tx, ctx, design) : '';
        return `<p:sp${sh.fill && sh.fill.t === 'bg' ? ' useBgFill="1"' : ''}><p:nvSpPr>${cNvPr(id, sh, ctx)}${spLocks}${phXML(sh.ph)}</p:nvSpPr><p:spPr>${xfrm(sh)}${geom}${fill}${ln}${shadowXML(sh.shadow, design)}</p:spPr>${tx}</p:sp>`;
      }
    }
  }
  function tableXML(sh, id, ctx, design) {
    const t = sh.tbl;
    const grid = t.cols.map((w) => `<a:gridCol w="${emu(w)}"/>`).join('');
    const bd = (b, tag) => lineXML(b && b.c && b.t !== 'none' ? b : { t: 'none', w: 1 }, design, tag);
    const rows = t.rows.map((r) => `<a:tr h="${emu(r.h)}">` + r.cells.map((c) => {
      const a = [];
      if (c.gs > 1) a.push(`gridSpan="${c.gs}"`);
      if (c.rs > 1) a.push(`rowSpan="${c.rs}"`);
      if (c.hm) a.push('hMerge="1"');
      if (c.vm) a.push('vMerge="1"');
      const tx = c.tx || L.txt.body();
      const ins = tx.ins || [7.2, 3.6, 7.2, 3.6];
      const ps = tx.ps && tx.ps.length ? tx.ps : [L.txt.para('')];
      const tcPr = `<a:tcPr marL="${emu(ins[0])}" marR="${emu(ins[2])}" marT="${emu(ins[1])}" marB="${emu(ins[3])}" anchor="${tx.anchor || 't'}">${bd(c.bd && c.bd.l, 'a:lnL')}${bd(c.bd && c.bd.r, 'a:lnR')}${bd(c.bd && c.bd.t, 'a:lnT')}${bd(c.bd && c.bd.b, 'a:lnB')}${fillXML(c.fill || { t: 'none' }, ctx, design) || '<a:noFill/>'}</a:tcPr>`;
      return `<a:tc${a.length ? ' ' + a.join(' ') : ''}><a:txBody><a:bodyPr/>${lstStyleXML(null)}${parasXML(ps, ctx, design)}</a:txBody>${tcPr}</a:tc>`;
    }).join('') + '</a:tr>').join('');
    return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(id, sh, ctx)}<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>${phXML(sh.ph)}</p:nvGraphicFramePr>${xfrm(sh, 'p:xfrm')}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="${t.firstRow ? 1 : 0}" bandRow="${t.bandRow ? 1 : 0}"/><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
  }
  function wordartXML(sh, id, ctx, design) {
    const wa = sh.wa || {};
    const lines = String(wa.text || '').split('\n');
    const fs = Math.max(8, Math.round(((sh.h / Math.max(lines.length, 1)) * 0.75) * 2) / 2);
    const rp = { sz: fs, b: !!wa.b, i: !!wa.i, font: wa.font || 'Arial Black', fill: wa.fill || { t: 'solid', c: 'accent1' }, ln: wa.line && wa.line.t !== 'none' ? wa.line : undefined, spc: wa.spc, shdX: wa.shadow || undefined };
    const tx = { anchor: 'ctr', wrap: false, ins: [7.2, 3.6, 7.2, 3.6], ps: lines.map((t) => ({ lvl: 0, pp: { algn: wa.algn || 'ctr' }, rs: t ? [Object.assign({ t }, rp)] : [], end: rp })) };
    const warp = `<a:prstTxWarp prst="${wa.warp || 'textPlain'}"><a:avLst/></a:prstTxWarp>`;
    const body = txBodyXML(tx, ctx, design, 'p:txBody', { warp, attrs: 'fromWordArt="1"' + (wa.vert ? '' : '') });
    const withShadow = body;
    return `<p:sp><p:nvSpPr>${cNvPr(id, sh, ctx)}<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(sh)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>${withShadow}</p:sp>`;
  }

  /* ---------- charts ---------- */
  function chartFrameXML(sh, id, ctx, design) {
    const c = sh.chart || L.chart.sample();
    if (c.srcId && c.edited) ctx.writer?.loss({ id: 'chart-edit:' + c.srcId, what: 'Editing this chart replaces its original chart-specific formatting and extensions.', where: sh.name || ctx.part, action: 'conversion' });
    /* a chart read from a file and not edited here is written back as it came, with its workbook and styles */
    const src = c.srcId && !c.edited && L.chart.src ? L.chart.src.get(c.srcId) : null;
    const source = src?.source && K.package(src.source);
    let rid, converted = false;
    if (source && ctx.writer) {
      try {
        const owner = sh.keep?.part || src.owner, rel = source.rels(owner).find(r => r.part === src.part && K.relationshipType(r.type) === RT('chart'));
        rid = ctx.writer.keepRel(ctx.part, { source: source.id, owner, ...rel });
      } catch (error) {
        converted = true;
        ctx.writer.loss({ id: 'chart:' + src.source + ':' + src.part, what: 'The chart was converted because its original dependencies are incomplete: ' + error.message, where: src.part, action: 'conversion' });
      }
    }
    if (!rid) {
      const n = src && !converted ? ctx.addChart(HEAD + String(src.xml).replace(/^﻿/, '').replace(/^<\?xml[^>]*\?>\s*/, ''), src.parts) : ctx.addChart(chartXML(c, design));
      rid = ctx.rels.add(RT('chart'), ctx.part ? K.relative(ctx.part, `ppt/charts/chart${n}.xml`) : `../charts/chart${n}.xml`);
    }
    return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(id, sh, ctx)}<p:cNvGraphicFramePr/>${phXML(sh.ph)}</p:nvGraphicFramePr>${xfrm(sh, 'p:xfrm')}<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
  }
  /** chart XML from the chart model: charts made or edited in Lectern */
  function chartXML(c, design) {
    const v2 = c.v === 2;
    const v = (x) => `<c:v>${X(x)}</c:v>`;
    const cats = c.cats || [];
    const nctx = { media: () => null };
    const strLit = (arr) => `<c:strLit><c:ptCount val="${arr.length}"/>${arr.map((x, i) => `<c:pt idx="${i}">${v(x == null ? '' : x)}</c:pt>`).join('')}</c:strLit>`;
    const numLit = (arr, fmt) => `<c:numLit><c:formatCode>${X(fmt || 'General')}</c:formatCode><c:ptCount val="${arr.length}"/>${arr.map((x, i) => (x == null || x === '' || !isFinite(+x) ? '' : `<c:pt idx="${i}">${v(+x)}</c:pt>`)).join('')}</c:numLit>`;
    const pal = L.chart.PALETTE;
    const series = c.series || [];
    const pie = c.kind === 'pie' || c.kind === 'doughnut';
    const g0 = (v2 && c.groups && c.groups[0]) || {};
    const font = c.font === '+mn' ? '+mn-lt' : c.font === '+mj' ? '+mj-lt' : c.font || (v2 ? '+mn-lt' : 'Arial');
    /* series formatting: Lectern's own charts keep the 2003 look (black outlines); imported ones keep theirs */
    const colOf = (s, i) => s.color || pal[i % pal.length];
    const spFill = (s, i) => {
      if (!v2) return `<c:spPr><a:solidFill>${clr(colOf(s, i), null, design)}</a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>`;
      const f = s.fill && s.fill.t !== 'img' ? fillXML(s.fill, nctx, design) : `<a:solidFill>${clr(colOf(s, i), null, design)}</a:solidFill>`;
      const ln = s.line ? lineXML(Object.assign({ c: colOf(s, i), w: 0.75 }, s.line), design) : '';
      return `<c:spPr>${f}${ln}</c:spPr>`;
    };
    const spLine = (s, i) => {
      if (s.line && s.line.t === 'none') return '<c:spPr><a:ln w="28575"><a:noFill/></a:ln></c:spPr>';
      const ln = Object.assign({ c: colOf(s, i), w: 2.25 }, s.line && s.line.t !== 'none' ? s.line : {});
      delete ln.t;
      return `<c:spPr>${lineXML(ln, design)}</c:spPr>`;
    };
    const markerXML = (s, i, sym) => {
      if (sym === 'none') return '<c:marker><c:symbol val="none"/></c:marker>';
      const m = s.marker || {};
      const f = m.fill && m.fill.t !== 'img' ? fillXML(m.fill, nctx, design) : `<a:solidFill>${clr(colOf(s, i), null, design)}</a:solidFill>`;
      const ln = lineXML(Object.assign({ c: colOf(s, i), w: 0.75 }, m.line && m.line.t !== 'none' ? m.line : {}), design);
      return `<c:marker><c:symbol val="${m.sym && m.sym !== 'auto' ? m.sym : sym || 'square'}"/><c:size val="${L.clamp(Math.round(m.size || (v2 ? 7 : 6)), 2, 72)}"/><c:spPr>${f}${ln}</c:spPr></c:marker>`;
    };
    const lblXML = (s) => {
      const on = v2 ? s.lbl : c.labels ? (pie ? { showPercent: true } : { showVal: true }) : null;
      if (!on) return '';
      const fmt = on.fmt || (!v2 && c.numFmt) || null;
      const b = (k) => `<c:${k} val="${on[k] ? 1 : 0}"/>`;
      /* label positions each chart type accepts */
      const POS = { col: 'ctr inEnd inBase outEnd', bar: 'ctr inEnd inBase outEnd', colStacked: 'ctr inEnd inBase', barStacked: 'ctr inEnd inBase', colPct: 'ctr inEnd inBase', barPct: 'ctr inEnd inBase', line: 't b l r ctr', lineMarkers: 't b l r ctr', scatter: 't b l r ctr', bubble: 't b l r ctr', pie: 'bestFit ctr inEnd outEnd' }[c.kind] || '';
      const pos = on.pos && POS.split(' ').includes(on.pos) ? on.pos : null;
      return `<c:dLbls>${fmt ? `<c:numFmt formatCode="${X(fmt)}" sourceLinked="0"/>` : ''}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${pos ? `<c:dLblPos val="${pos}"/>` : ''}${b('showLegendKey')}${b('showVal')}${b('showCatName')}${b('showSerName')}${b('showPercent')}${b('showBubbleSize')}</c:dLbls>`;
    };
    const head = (s, i) => `<c:idx val="${i}"/><c:order val="${i}"/><c:tx>${v(s.name || 'Series ' + (i + 1))}</c:tx>`;
    const catVal = (s) => `<c:cat>${strLit(cats)}</c:cat><c:val>${numLit(s.vals, s.fmt || (!v2 && c.numFmt))}</c:val>`;
    const xsOf = (s) => (s.xs && s.xs.length ? s.xs : cats.map((x, k) => (isFinite(parseFloat(x)) ? parseFloat(x) : k + 1)));
    const sers = (kind, keep) => series.map((s, i) => {
      if (keep && !keep(s)) return '';
      if (kind === 'overlay') return `<c:ser>${head(s, i)}${s.overlay === 'markers' ? '<c:spPr><a:ln w="28575"><a:noFill/></a:ln></c:spPr>' : spLine(s, i)}${markerXML(s, i, s.overlay === 'line' ? 'none' : 'square')}${lblXML(s)}${catVal(s)}<c:smooth val="0"/></c:ser>`;
      if (kind === 'pie') {
        const dpts = cats.map((_, k) => {
          const pt = s.pts && s.pts[k];
          const f = pt && pt.fill && pt.fill.t !== 'img' ? fillXML(pt.fill, nctx, design) : `<a:solidFill>${clr((c.pieColors && c.pieColors[k]) || pal[k % pal.length], null, design)}</a:solidFill>`;
          const ln = v2 ? (s.line ? lineXML(Object.assign({ c: '#FFFFFF', w: 0.75 }, s.line), design) : '') : '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>';
          return `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/>${pt && pt.expl ? `<c:explosion val="${Math.round(pt.expl)}"/>` : ''}<c:spPr>${f}${ln}</c:spPr></c:dPt>`;
        }).join('');
        return `<c:ser>${head(s, i)}${s.expl ? `<c:explosion val="${Math.round(s.expl)}"/>` : ''}${dpts}${lblXML(s)}${catVal(s)}</c:ser>`;
      }
      if (kind === 'line') return `<c:ser>${head(s, i)}${spLine(s, i)}${markerXML(s, i, c.kind === 'lineMarkers' ? (v2 ? 'circle' : 'square') : 'none')}${lblXML(s)}${catVal(s)}<c:smooth val="${s.smooth ? 1 : 0}"/></c:ser>`;
      if (kind === 'radar') return `<c:ser>${head(s, i)}${c.kind === 'radarFilled' ? spFill(s, i) : spLine(s, i)}${markerXML(s, i, c.kind === 'radarFilled' ? 'none' : 'circle')}${lblXML(s)}${catVal(s)}</c:ser>`;
      if (kind === 'scatter') return `<c:ser>${head(s, i)}${spLine(s, i)}${markerXML(s, i, 'square')}${lblXML(s)}<c:xVal>${numLit(xsOf(s))}</c:xVal><c:yVal>${numLit(s.vals, s.fmt)}</c:yVal><c:smooth val="${s.smooth ? 1 : 0}"/></c:ser>`;
      if (kind === 'bubble') return `<c:ser>${head(s, i)}${spFill(s, i)}<c:invertIfNegative val="0"/>${lblXML(s)}<c:xVal>${numLit(xsOf(s))}</c:xVal><c:yVal>${numLit(s.vals, s.fmt)}</c:yVal><c:bubbleSize>${numLit(s.sizes || s.vals.map(() => 1))}</c:bubbleSize><c:bubble3D val="0"/></c:ser>`;
      if (kind === 'area') return `<c:ser>${head(s, i)}${spFill(s, i)}${lblXML(s)}${catVal(s)}</c:ser>`;
      return `<c:ser>${head(s, i)}${spFill(s, i)}<c:invertIfNegative val="0"/>${lblXML(s)}${catVal(s)}</c:ser>`;
    }).join('');
    const AX1 = 50010001, AX2 = 50010002;
    const axIds = `<c:axId val="${AX1}"/><c:axId val="${AX2}"/>`;
    const ax = (v2 && c.ax) || {};
    const gridXML = (gl) => (gl && (gl.c || gl.w) ? `<c:majorGridlines><c:spPr>${lineXML(Object.assign({ c: '#D9D9D9', w: 0.75 }, gl), design)}</c:spPr></c:majorGridlines>` : '<c:majorGridlines/>');
    const grid = c.gridY !== false ? gridXML(ax.v && ax.v.grid) : '';
    const txPrXML = (o, rot) => {
      if (!o && rot == null) return '';
      o = o || {};
      const attrs = [];
      if (o.sz) attrs.push(`sz="${Math.round(o.sz * 100)}"`);
      if (o.b != null) attrs.push(`b="${o.b ? 1 : 0}"`);
      if (o.i != null) attrs.push(`i="${o.i ? 1 : 0}"`);
      const inner = o.color ? `<a:solidFill>${clr(o.color, null, design)}</a:solidFill>` : '';
      return `<c:txPr><a:bodyPr${rot != null ? ` rot="${Math.round(rot * 60000)}" vert="horz"` : ''}/><a:lstStyle/><a:p><a:pPr><a:defRPr${attrs.length ? ' ' + attrs.join(' ') : ''}${inner ? `>${inner}</a:defRPr>` : '/>'}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
    };
    const axLine = (a) => (v2 && a && a.line ? `<c:spPr>${lineXML(Object.assign({ c: '#868686', w: 0.75 }, a.line), design)}</c:spPr>` : '');
    const catAx = (pos, rev) => { const a = ax.c || {}; return `<c:catAx><c:axId val="${AX1}"/><c:scaling><c:orientation val="${rev ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="${a.del ? 1 : 0}"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="${a.tick || 'out'}"/><c:minorTickMark val="none"/><c:tickLblPos val="${a.lblPos || 'nextTo'}"/>${axLine(a)}${txPrXML(v2 ? a.tx : null, c.catRot ? c.catRot : null)}<c:crossAx val="${AX2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`; };
    const valAx = (pos, id, cross, g, fmtDef, a, crossBetween) => {
      a = a || {};
      const sc = (a.max != null ? `<c:max val="${a.max}"/>` : '') + (a.min != null ? `<c:min val="${a.min}"/>` : '');
      return `<c:valAx><c:axId val="${id}"/><c:scaling>${a.log ? `<c:logBase val="${a.log}"/>` : ''}<c:orientation val="${a.rev ? 'maxMin' : 'minMax'}"/>${sc}</c:scaling><c:delete val="${a.del ? 1 : 0}"/><c:axPos val="${pos}"/>${g ? grid : ''}<c:numFmt formatCode="${X(c.axisFmt || fmtDef || 'General')}" sourceLinked="0"/><c:majorTickMark val="${a.tick || 'out'}"/><c:minorTickMark val="none"/><c:tickLblPos val="${a.lblPos || 'nextTo'}"/>${axLine(a)}${txPrXML(v2 ? a.tx : null)}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="${crossBetween || 'between'}"/>${a.major ? `<c:majorUnit val="${a.major}"/>` : ''}</c:valAx>`;
    };
    const gap = g0.type === 'bar' && g0.gap != null ? Math.round(g0.gap) : 150;
    const ovl = (st) => (st ? 100 : g0.type === 'bar' && g0.overlap ? Math.round(g0.overlap) : 0);
    const firstPie = sers('pie').split('</c:ser>')[0] + (series.length ? '</c:ser>' : '');
    let plot;
    switch (c.kind) {
      case 'bar': case 'barStacked': case 'barPct': {
        const grp = c.kind === 'barPct' ? 'percentStacked' : c.kind === 'barStacked' ? 'stacked' : 'clustered';
        plot = `<c:barChart><c:barDir val="bar"/><c:grouping val="${grp}"/><c:varyColors val="0"/>${sers('bar', (x) => !x.overlay)}<c:gapWidth val="${gap}"/>${grp !== 'clustered' ? '<c:overlap val="100"/>' : ovl(false) ? `<c:overlap val="${ovl(false)}"/>` : ''}${axIds}</c:barChart>${catAx('l', v2 ? !!(ax.c && ax.c.rev) : true)}${valAx('b', AX2, AX1, true, c.kind === 'barPct' ? '0%' : null, ax.v)}`;
        break;
      }
      case 'line': case 'lineMarkers':
        plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers('line')}<c:marker val="1"/>${axIds}</c:lineChart>${catAx('b', ax.c && ax.c.rev)}${valAx('l', AX2, AX1, true, null, ax.v)}`;
        break;
      case 'area': case 'areaStacked': case 'areaPct':
        plot = `<c:areaChart><c:grouping val="${c.kind === 'areaPct' ? 'percentStacked' : c.kind === 'areaStacked' ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${sers('area')}${axIds}</c:areaChart>${catAx('b', ax.c && ax.c.rev)}${valAx('l', AX2, AX1, true, c.kind === 'areaPct' ? '0%' : null, ax.v, 'midCat')}`;
        break;
      case 'pie':
        plot = `<c:pieChart><c:varyColors val="1"/>${firstPie}<c:firstSliceAng val="${Math.round(g0.firstAng || 0)}"/></c:pieChart>`;
        break;
      case 'doughnut':
        plot = `<c:doughnutChart><c:varyColors val="1"/>${firstPie}<c:firstSliceAng val="${Math.round(g0.firstAng || 0)}"/><c:holeSize val="${L.clamp(Math.round(g0.hole || 50), 10, 90)}"/></c:doughnutChart>`;
        break;
      case 'radar': case 'radarFilled':
        plot = `<c:radarChart><c:radarStyle val="${c.kind === 'radarFilled' ? 'filled' : 'marker'}"/><c:varyColors val="0"/>${sers('radar')}${axIds}</c:radarChart>${catAx('b')}${valAx('l', AX2, AX1, true, null, ax.v, 'between')}`;
        break;
      case 'bubble':
        plot = `<c:bubbleChart><c:varyColors val="0"/>${sers('bubble')}<c:bubbleScale val="${Math.round(g0.bubbleScale || 100)}"/><c:showNegBubbles val="0"/>${axIds}</c:bubbleChart>${valAx('b', AX1, AX2, false, null, ax.c, 'midCat')}${valAx('l', AX2, AX1, true, null, ax.v, 'midCat')}`;
        break;
      case 'scatter':
        plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers('scatter')}${axIds}</c:scatterChart>${valAx('b', AX1, AX2, false, null, ax.c, 'midCat')}${valAx('l', AX2, AX1, true, null, ax.v, 'midCat')}`;
        break;
      default: {
        const grp = c.kind === 'colStacked' ? 'stacked' : c.kind === 'colPct' ? 'percentStacked' : 'clustered';
        plot = `<c:barChart><c:barDir val="col"/><c:grouping val="${grp}"/><c:varyColors val="0"/>${sers('bar', (x) => !x.overlay)}<c:gapWidth val="${gap}"/>${grp !== 'clustered' ? '<c:overlap val="100"/>' : ovl(false) ? `<c:overlap val="${ovl(false)}"/>` : ''}${axIds}</c:barChart>${series.some((x) => x.overlay) ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers('overlay', (x) => x.overlay)}<c:marker val="1"/>${axIds}</c:lineChart>` : ''}${catAx('b', ax.c && ax.c.rev)}${valAx('l', AX2, AX1, true, c.kind === 'colPct' ? '0%' : null, ax.v)}`;
      }
    }
    const tsz = v2 ? (c.titleTx && c.titleTx.sz) || (c.fsz || 10) * 1.2 : (c.fsz || 12) * 1.4;
    const tb = v2 && c.titleTx && c.titleTx.b === false ? 0 : 1;
    const tcol = v2 && c.titleTx && c.titleTx.color ? `<a:solidFill>${clr(c.titleTx.color, null, design)}</a:solidFill>` : '';
    const titleLines = String(c.title || '').split('\n');
    const title = c.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>${titleLines.map((t) => `<a:p><a:pPr><a:defRPr sz="${Math.round(tsz * 100)}" b="${tb}"/></a:pPr><a:r><a:rPr lang="en-US" sz="${Math.round(tsz * 100)}" b="${tb}"${tcol ? `>${tcol}</a:rPr>` : '/>'}<a:t>${X(t)}</a:t></a:r></a:p>`).join('')}</c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    const legend = c.legend && c.legend !== 'none' ? `<c:legend><c:legendPos val="${c.legend}"/><c:overlay val="0"/>${v2 ? txPrXML(c.legendTx) : ''}</c:legend>` : '';
    const area = v2 ? `<c:spPr>${c.bg ? fillXML(c.bg, nctx, design) : '<a:noFill/>'}${c.border ? lineXML(Object.assign({ c: '#868686', w: 0.75 }, c.border), design) : '<a:ln><a:noFill/></a:ln>'}</c:spPr>` : '';
    const tcolor = v2 && c.textColor ? `<a:solidFill>${clr(c.textColor, null, design)}</a:solidFill>` : '';
    return HEAD + `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>${area}<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round((c.fsz || 12) * 100)}">${tcolor}<a:latin typeface="${X(font)}"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:chartSpace>`;
  }

  /* ---------- animation timing ---------- */
  function timingXML(slide, ctx) {
    const anims = (slide.anims || []).filter((a) => ctx.idOf(a.sid));
    if (!anims.length) return '';
    let cid = 2;
    const nid = ctx.nextTimingId || (() => ++cid);
    /* expand by-paragraph builds */
    const items = [];
    for (const a of anims) {
      const sh = ctx.findShape(a.sid);
      if (a.by === 'para' && sh && sh.tx && sh.tx.ps.length > 1) {
        let first = true;
        sh.tx.ps.forEach((p, i) => { if (L.txt.paraText(p).trim() || i === 0) { items.push(Object.assign({}, a, { para: i, start: first ? a.start : (p.lvl || 0) > 0 ? 'with' : a.start, delay: !first && (p.lvl || 0) > 0 ? 0 : a.delay })); first = false; } });
      } else items.push(a);
    }
    /* group: click groups → sequential sub-groups (after) → parallel effects (with) */
    const groups = [];
    for (const a of items) {
      if (a.start === 'click' || !groups.length) groups.push({ click: a.start === 'click', subs: [[a]] });
      else if (a.start === 'after') groups[groups.length - 1].subs.push([a]);
      else groups[groups.length - 1].subs[groups[groups.length - 1].subs.length - 1].push(a);
    }
    const bld = new Map();
    let grp = 0;
    const clickPars = groups.map((g) => {
      let t = 0;
      const subPars = g.subs.map((sub) => {
        const start = t;
        let end = start;
        const effs = sub.map((a) => {
          const total = (a.delay || 0) + (a.dur || 500) * (a.cls === 'emph' ? Math.max(1, a.repeat || 1) : 1);
          end = Math.max(end, start + total);
          if (!bld.has(a.sid + ':' + a.id)) bld.set(a.sid + ':' + a.id, { sid: a.sid, grpId: ctx.nextBuildId ? ctx.nextBuildId() : grp++, para: a.by === 'para' });
          const gi = bld.get(a.sid + ':' + a.id).grpId;
          return effectXML(a, ctx, nid, gi, a === sub[0] && sub === g.subs[0] ? (g.click ? 'clickEffect' : a.start === 'with' ? 'withEffect' : 'afterEffect') : a.start === 'after' ? 'afterEffect' : 'withEffect');
        }).join('');
        t = end;
        return `<p:par><p:cTn id="${nid()}" fill="hold"><p:stCondLst><p:cond delay="${Math.round(start)}"/></p:stCondLst><p:childTnLst>${effs}</p:childTnLst></p:cTn></p:par>`;
      }).join('');
      const cond = g.click ? '<p:cond delay="indefinite"/>' : '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="' + (ctx.timingMainId || 2) + '"/></p:cond>';
      return `<p:par><p:cTn id="${nid()}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${subPars}</p:childTnLst></p:cTn></p:par>`;
    }).join('');
    const bldXML = Array.from(bld.values()).map((b) => {
      const sh = ctx.findShape(b.sid);
      const isText = sh && sh.tx && !L.txt.isEmpty(sh.tx);
      if (!isText) return '';
      return `<p:bldP spid="${ctx.idOf(b.sid)}" grpId="${b.grpId}"${b.para ? ' build="p"' : ''}${sh.fill && sh.fill.t !== 'none' ? ' animBg="1"' : ''}/>`;
    }).join('');
    return `<p:timing><p:tnLst><p:par><p:cTn id="${ctx.timingRootId || 1}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="${ctx.timingMainId || 2}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clickPars}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>${bldXML ? `<p:bldLst>${bldXML}</p:bldLst>` : ''}</p:timing>`;
  }
  function effectXML(a, ctx, nid, grpId, nodeType) {
    const spid = ctx.idOf(a.sid);
    const tgt = a.para != null ? `<p:tgtEl><p:spTgt spid="${spid}"><p:txEl><p:pRg st="${a.para}" end="${a.para}"/></p:txEl></p:spTgt></p:tgtEl>` : `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;
    const dur = Math.max(1, Math.round(a.dur || 500));
    const info = (L.anim && L.anim.info(a.cls, a.eff)) || { id: 10, filter: 'fade' };
    const sub = L.anim ? L.anim.subtype(a) : 0;
    const vis = (val, delay) => `<p:set><p:cBhvr><p:cTn id="${nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay || 0}"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`;
    const fx = (trans, filter) => `<p:animEffect transition="${trans}" filter="${filter}"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}"/>${tgt}</p:cBhvr></p:animEffect>`;
    const anim = (attr, from, to, d) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${nid()}" dur="${d || dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
    let kids = '';
    const filt = L.anim ? L.anim.filter(a) : 'fade';
    const off = L.anim ? L.anim.flyFrom(a.dir) : { x: '0-#ppt_w/2', y: '#ppt_y' };
    if (a.cls === 'entr') {
      kids = vis('visible');
      switch (a.eff) {
        case 'appear': break;
        case 'flyIn': kids += anim('ppt_x', off.x, '#ppt_x') + anim('ppt_y', off.y, '#ppt_y'); break;
        case 'zoom': kids += anim('ppt_w', '0', '#ppt_w') + anim('ppt_h', '0', '#ppt_h'); break;
        case 'ascend': kids += fx('in', 'fade') + anim('ppt_y', '#ppt_y+.1', '#ppt_y'); break;
        case 'descend': kids += fx('in', 'fade') + anim('ppt_y', '#ppt_y-.1', '#ppt_y'); break;
        case 'expand': kids += fx('in', 'fade') + anim('ppt_w', '0.7*#ppt_w', '#ppt_w'); break;
        case 'spinner': kids += fx('in', 'fade') + anim('ppt_w', '0', '#ppt_w') + anim('ppt_h', '0', '#ppt_h') + `<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>style.rotation</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:fltVal val="360"/></p:val></p:tav><p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav></p:tavLst></p:anim>`; break;
        default: kids += fx('in', filt);
      }
    } else if (a.cls === 'exit') {
      switch (a.eff) {
        case 'disappear': kids = vis('hidden'); break;
        case 'flyOut': kids = anim('ppt_x', '#ppt_x', off.x.replace('0-', '0-').replace('1+', '1+')) + anim('ppt_y', '#ppt_y', off.y) + vis('hidden', dur - 1); break;
        case 'zoom': kids = anim('ppt_w', '#ppt_w', '0') + anim('ppt_h', '#ppt_h', '0') + vis('hidden', dur - 1); break;
        default: kids = fx('out', filt) + vis('hidden', dur - 1);
      }
    } else if (a.cls === 'emph') {
      const rep = a.repeat > 1 ? ` repeatCount="${a.repeat * 1000}"` : '';
      switch (a.eff) {
        case 'spin': kids = `<p:animRot by="${Math.round((a.amount || 360) * 60000)}"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"${rep}/>${tgt}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`; break;
        case 'transparency': kids = `<p:set><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"${rep}/>${tgt}<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="0.5"/></p:to></p:set>`; break;
        case 'teeter': {
          const seg = Math.max(1, Math.round(dur / 5));
          kids = [120000, -240000, 240000, -240000, 120000].map((by, i) => `<p:animRot by="${by}"><p:cBhvr><p:cTn id="${nid()}" dur="${seg}" fill="hold"><p:stCondLst><p:cond delay="${i * seg}"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`).join('');
          break;
        }
        case 'blink': kids = `<p:anim calcmode="discrete" valueType="str"><p:cBhvr override="childStyle"><p:cTn id="${nid()}" dur="${dur}" fill="hold"${rep}/>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="hidden"/></p:val></p:tav><p:tav tm="50000"><p:val><p:strVal val="visible"/></p:val></p:tav></p:tavLst></p:anim>`; break;
        default: kids = `<p:animScale><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"${rep}/>${tgt}</p:cBhvr><p:by x="${Math.round((a.amount || 150) * 1000)}" y="${Math.round((a.amount || 150) * 1000)}"/></p:animScale>`;
      }
    } else if (a.cls === 'path') {
      const p = L.anim ? L.anim.pathString(a) : 'M 0 0 L 0.25 0 E';
      kids = `<p:animMotion origin="layout" path="${p}" pathEditMode="relative" ptsTypes=""><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`;
    }
    const cls = a.cls === 'path' ? 'path' : a.cls;
    return `<p:par><p:cTn id="${nid()}" presetID="${info.id}" presetClass="${cls}" presetSubtype="${sub}" fill="hold"${a.cls !== 'path' ? ` grpId="${grpId}"` : ''} nodeType="${nodeType}"><p:stCondLst><p:cond delay="${Math.round(a.delay || 0)}"/></p:stCondLst><p:childTnLst>${kids}</p:childTnLst></p:cTn></p:par>`;
  }

  /* ---------- transitions ---------- */
  function transitionXML(tr) {
    if (!tr || (tr.type === 'none' && tr.click !== false && !tr.after)) return '';
    const a = [`spd="${tr.spd || 'fast'}"`];
    if (tr.click === false) a.push('advClick="0"');
    if (tr.after != null && tr.after !== '') a.push(`advTm="${Math.round(tr.after)}"`);
    let el = '';
    switch (tr.type) {
      case 'none': el = ''; break;
      case 'blinds': case 'checker': case 'randomBar': case 'comb': el = `<p:${tr.type} dir="${tr.dir || 'horz'}"/>`; break;
      case 'cover': case 'pull': el = `<p:${tr.type} dir="${tr.dir || 'l'}"/>`; break;
      case 'push': case 'wipe': el = `<p:${tr.type} dir="${tr.dir || 'l'}"/>`; break;
      case 'split': el = `<p:split orient="${tr.orient || 'horz'}" dir="${tr.dir || 'out'}"/>`; break;
      case 'strips': el = `<p:strips dir="${tr.dir || 'lu'}"/>`; break;
      case 'wheel': el = `<p:wheel spokes="${tr.spokes || 4}"/>`; break;
      case 'zoom': el = `<p:zoom dir="${tr.dir || 'in'}"/>`; break;
      case 'cut': el = tr.thruBlk ? '<p:cut thruBlk="1"/>' : '<p:cut/>'; break;
      case 'fade': el = tr.thruBlk ? '<p:fade thruBlk="1"/>' : '<p:fade/>'; break;
      case 'circle': case 'diamond': case 'dissolve': case 'newsflash': case 'plus': case 'random': case 'wedge': el = `<p:${tr.type}/>`; break;
      default: el = '';
    }
    return `<p:transition ${a.join(' ')}>${el}</p:transition>`;
  }

  /* ---------- masters, layouts, theme ---------- */
  function themeXML(d, name) {
    const c = d.colors;
    const sc = (k) => `<a:${k}><a:srgbClr val="${(c[k] || '#000000').replace('#', '').toUpperCase()}"/></a:${k}>`;
    const fnt = (f) => `<a:latin typeface="${X(f)}"/><a:ea typeface=""/><a:cs typeface=""/>`;
    const ph = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    const ln = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>`;
    return HEAD + `<a:theme xmlns:a="${NS_A}" name="${X(name || d.name)}"><a:themeElements><a:clrScheme name="${X(d.name)}">${['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(sc).join('')}</a:clrScheme>` +
      `<a:fontScheme name="${X(d.name)}"><a:majorFont>${fnt(d.fonts.major || 'Arial')}</a:majorFont><a:minorFont>${fnt(d.fonts.minor || 'Arial')}</a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="Office"><a:fillStyleLst>${ph}${ph}${ph}</a:fillStyleLst><a:lnStyleLst>${ln(9525)}${ln(25400)}${ln(38100)}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${ph}${ph}${ph}</a:bgFillStyleLst></a:fmtScheme>` +
      `</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
  }
  const CLRMAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  const GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  function bgXML(fill, ctx, d) {
    if (!fill) return '';
    return `<p:bg><p:bgPr>${fillXML(fill, ctx, d)}<a:effectLst/></p:bgPr></p:bg>`;
  }
  function levelsXML(levels, ctx, d) {
    return levels.slice(0, 9).map((lv, i) => {
      const l2 = Object.assign({}, lv, { rPr: Object.assign({ font: '+mn' }, lv.rPr || {}) });
      return pPrXML(l2, `a:lvl${i + 1}pPr`, 0, ctx, d, true) || `<a:lvl${i + 1}pPr/>`;
    }).join('');
  }
  function placeholderSp(id, name, ph, r, tx, ctx, d) {
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${X(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>${phXML(ph)}</p:nvSpPr><p:spPr>${xfrm(r)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>${txBodyXML(tx, ctx, d)}</p:sp>`;
  }
  function masterXML(d, ctx, layoutIds) {
    let id = 1;
    const nid = ctx.nextId || (() => ++id);
    ctx.nextId ||= nid;
    const deco = (d.deco || []).map((s) => shapeXML(s, ctx, d)).join('');
    const body = (t) => L.txt.body([L.txt.para(t)], { anchor: 't' });
    const ph = d.ph;
    const tp = placeholderSp(nid(), 'Title Placeholder 1', { type: 'title' }, ph.title, Object.assign(body('Click to edit Master title style'), { anchor: 'ctr' }), ctx, d);
    const bp = placeholderSp(nid(), 'Text Placeholder 2', { type: 'body', idx: 1 }, ph.body, L.txt.body(['Click to edit Master text styles', 'Second level', 'Third level', 'Fourth level', 'Fifth level'].map((t, i) => L.txt.para(t, null, null, i)), { anchor: 't' }), ctx, d);
    const footLst = (algn) => ({ 0: { algn, rPr: Object.assign({}, d.footer && d.footer.rPr) } });
    const dt = placeholderSp(nid(), 'Date Placeholder 3', { type: 'dt', idx: 2 }, ph.dt, Object.assign(L.txt.body([{ lvl: 0, pp: {}, rs: [] }]), { lst: footLst('l') }), ctx, d);
    const ft = placeholderSp(nid(), 'Footer Placeholder 4', { type: 'ftr', idx: 3 }, ph.ftr, Object.assign(L.txt.body([{ lvl: 0, pp: {}, rs: [] }]), { lst: footLst('ctr') }), ctx, d);
    const sn = placeholderSp(nid(), 'Slide Number Placeholder 5', { type: 'sldNum', idx: 4 }, ph.sldNum, Object.assign(L.txt.body([{ lvl: 0, pp: {}, rs: [{ t: '‹#›', fld: 'slidenum' }] }]), { lst: footLst('r') }), ctx, d);
    const tStyle = levelsXML([L.deepMerge(L.model.DEFAULT_TX.title[0], d.tx.title[0])], ctx, d);
    const bStyle = levelsXML(L.model.DEFAULT_TX.body.map((lv, i) => L.deepMerge(lv, d.tx.body[i])), ctx, d);
    const oStyle = levelsXML(d.tx.other || L.model.DEFAULT_TX.other, ctx, d);
    const lids = layoutIds.map((l) => `<p:sldLayoutId id="${l.id}" r:id="${l.rid}"/>`).join('');
    return HEAD + `<p:sldMaster ${NSDECL}><p:cSld>${bgXML(d.bg || { t: 'solid', c: 'bg1' }, ctx, d)}<p:spTree>${GRP}${deco}${tp}${bp}${dt}${ft}${sn}</p:spTree></p:cSld>${CLRMAP}<p:sldLayoutIdLst>${lids}</p:sldLayoutIdLst><p:txStyles><p:titleStyle>${tStyle}</p:titleStyle><p:bodyStyle>${bStyle}</p:bodyStyle><p:otherStyle>${oStyle}</p:otherStyle></p:txStyles></p:sldMaster>`;
  }
  function layoutXML(d, key, lkey, ctx) {
    const info = L.model.layoutInfo(key);
    let id = 1;
    const nid = ctx.nextId || (() => ++id);
    ctx.nextId ||= nid;
    let deco = '', showMaster = true, bg = '';
    if (key === 'title' && d.titleDeco) { deco = d.titleDeco.map((s) => shapeXML(s, ctx, d)).join(''); showMaster = false; }
    else if (lkey && d.layoutDecos && d.layoutDecos[lkey]) { deco = d.layoutDecos[lkey].map((s) => shapeXML(s, ctx, d)).join(''); showMaster = !!d.layoutShowMaster?.[lkey]; }
    if (key === 'title' && d.titleBg) bg = bgXML(d.titleBg, ctx, d);
    else if (lkey && d.layoutBgs && d.layoutBgs[lkey]) bg = bgXML(d.layoutBgs[lkey], ctx, d);
    const frames = L.model.layoutFrames(key, d);
    const phs = frames.map((fr, i) => {
      const sh = L.model.makePlaceholder(fr, d);
      sh.tx.ps = [L.txt.para(L.model.PROMPTS[fr.type] || 'Click to add text')];
      return placeholderSp(nid(), sh.name || 'Placeholder ' + (i + 1), sh.ph, fr, sh.tx, ctx, d);
    }).join('');
    const footer = ['dt', 'ftr', 'sldNum'].map((t, i) => placeholderSp(nid(), ['Date Placeholder', 'Footer Placeholder', 'Slide Number Placeholder'][i] + ' ' + (frames.length + i + 1), { type: t, idx: 10 + i }, d.ph[t], L.txt.body([{ lvl: 0, pp: {}, rs: t === 'sldNum' ? [{ t: '‹#›', fld: 'slidenum' }] : [] }]), ctx, d)).join('');
    return HEAD + `<p:sldLayout ${NSDECL} type="${info.ox}" preserve="1"${showMaster ? '' : ' showMasterSp="0"'}><p:cSld name="${X(lkey && d.layoutNames && d.layoutNames[lkey] ? d.layoutNames[lkey] : info.name)}">${bg}<p:spTree>${GRP}${deco}${phs}${footer}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  }

  /* ---------- notes ---------- */
  function notesMasterXML(pres) {
    const W = 6858000, H = 9144000;
    const img = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1143000" y="685800"/><a:ext cx="4572000" cy="${Math.round((4572000 * pres.H) / pres.W)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>`;
    const body = `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>`;
    const lvl = Array.from({ length: 9 }, (_, i) => `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`).join('');
    void W; void H;
    return HEAD + `<p:notesMaster ${NSDECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRP}${img}${body}</p:spTree></p:cSld>${CLRMAP}<p:notesStyle>${lvl}</p:notesStyle></p:notesMaster>`;
  }
  function notesXML(pres, text) {
    const paras = String(text).split('\n').map((t) => (t ? `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${X(t)}</a:t></a:r></a:p>` : '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>')).join('');
    return HEAD + `<p:notes ${NSDECL}><p:cSld><p:spTree>${GRP}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras || '<a:p><a:endParaRPr lang="en-US"/></a:p>'}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
  }

  /* ---------- media ---------- */
  async function mediaBytes(id) {
    const m = L.media.get(id);
    if (!m) return null;
    let blob = m.blob, type = blob.type || m.type || 'image/png';
    const svg = /svg/.test(type) ? new Uint8Array(await L.readAsArrayBuffer(blob)) : null;
    if (/svg|webp|avif/.test(type)) {
      try {
        const img = await L.loadImage(m.url);
        const w = Math.max(1, Math.min(2048, img.naturalWidth || 512)), hh = Math.max(1, Math.min(2048, img.naturalHeight || 512));
        const k = /svg/.test(type) ? Math.max(1, 1024 / Math.max(w, hh)) : 1;
        const c = document.createElement('canvas');
        c.width = Math.round(w * k); c.height = Math.round(hh * k);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        blob = await new Promise((res) => c.toBlob(res, 'image/png'));
        type = 'image/png';
      } catch (e) { /* keep original */ }
    }
    return { bytes: new Uint8Array(await L.readAsArrayBuffer(blob)), ext: L.mimeToExt(type), type, svg };
  }

  /* ---------- main ---------- */
  async function write(pres, opts) {
    opts = opts || {};
    const format = K.format(pres, opts.format, 'pptx'), pack = L.preserve.begin(pres, format), writer = pack.writer;
    const freshName = (dir, base, ext) => pack.use(writer.name(dir, base, ext));
    const files = [];
    const add = (name, data) => files.push({ name, data });
    const overrides = [];
    const defaults = new Map([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]);
    const mainCT = format.contentType;
    const presRels = pack.rels('ppt/presentation.xml');
    const frames = L.frames.prepare(pres, writer);
    const mediaMap = new Map(); /* media id → {name} */
    let chartN = 0;
    const charts = [];

    const allSlides = pres.slides;
    // Regenerated fields have no opaque referrers. Allocate deterministic UUIDs
    // per part so repeated saves and undo do not churn otherwise identical XML.
    const fieldIds = part => {
      let index = 0;
      return () => {
        const seed = (pres.pkg?.id || pres.props?.created || '') + ':' + part + ':' + ++index;
        const bytes = L.sha.sha1(new TextEncoder().encode(seed)).slice(0, 16);
        bytes[6] = (bytes[6] & 15) | 80; bytes[8] = (bytes[8] & 63) | 128;
        const hex = Array.from(bytes, b => b.toString(16).padStart(2, '0')).join('').toUpperCase();
        return '{' + [hex.slice(0, 8), hex.slice(8, 12), hex.slice(12, 16), hex.slice(16, 20), hex.slice(20)].join('-') + '}';
      };
    };
    const slideParts = allSlides.map(s => {
      const original = pres.pkg && s.keep?.source === pres.pkg.id && !s.keep.copy ? s.keep.part : null;
      if (original) { pack.bind(original, original); return original; }
      return freshName('ppt/slides', 'slide', 'xml');
    });
    const slideIndex = (id) => allSlides.findIndex((s) => s.id === id);
    /* Retain the imported design library, including masters with no slides. */
    const designIds = Object.keys(pres.designs);
    for (const s of allSlides) if (!designIds.includes(s.design)) designIds.push(s.design);
    if (!designIds.length) designIds.push(Object.keys(pres.designs)[0]);

    /* media must be resolved synchronously during XML generation, so preload */
    const usedMedia = new Set();
    const scanFill = (f) => { if (f && f.t === 'img' && f.media) usedMedia.add(f.media); };
    const scanShape = (s) => {
      if (s.type === 'image' && s.media) usedMedia.add(s.media);
      if (s.type === 'image' && s.svgMedia) usedMedia.add(s.svgMedia);
      scanFill(s.fill);
      if (s.tbl) s.tbl.rows.forEach((r) => r.cells.forEach((c) => scanFill(c.fill)));
      if (s.kids) s.kids.forEach(scanShape);
    };
    for (const s of allSlides) { scanFill(s.bg); s.shapes.forEach(scanShape); }
    for (const id of designIds) { const d = pres.designs[id]; if (!d) continue; scanFill(d.bg); scanFill(d.titleBg); (d.deco || []).forEach(scanShape); (d.titleDeco || []).forEach(scanShape); if (d.layoutDecos) Object.values(d.layoutDecos).forEach((a) => a.forEach(scanShape)); if (d.layoutBgs) Object.values(d.layoutBgs).forEach(scanFill); }
    const preloaded = new Map();
    for (const id of usedMedia) preloaded.set(id, await mediaBytes(id));
    const mediaSync = (rels, owner = rels.base || 'ppt/slides/slide1.xml') => (id) => {
      if (!id) return null;
      if (!mediaMap.has(id)) {
        const mb = preloaded.get(id);
        if (!mb) { mediaMap.set(id, null); return null; }
        const name = freshName('ppt/media', 'image', mb.ext).split('/').pop();
        const origin = pres.keep?.media?.[id], bytes = origin && pres.pkg?.bytes(origin);
        if (bytes && bytes.length === mb.bytes.length && bytes.every((b, i) => b === mb.bytes[i])) pack.bind('ppt/media/' + name, origin);
        add('ppt/media/' + name, mb.bytes);
        defaults.set(mb.ext, mb.type);
        mediaMap.set(id, { name });
      }
      const m = mediaMap.get(id);
      return m ? rels.add(RT('image'), K.relative(owner, 'ppt/media/' + m.name)) : null;
    };
    /* SVG pictures: the PNG copy goes in r:embed, the vector original in the svgBlip extension */
    const svgMap = new Map();
    const svgSync = (rels, owner = rels.base || 'ppt/slides/slide1.xml') => (id) => {
      const mb = id ? preloaded.get(id) : null;
      if (!mb || !mb.svg) return null;
      if (!svgMap.has(id)) {
        const name = freshName('ppt/media', 'vector', 'svg').split('/').pop();
        add('ppt/media/' + name, mb.svg);
        defaults.set('svg', 'image/svg+xml');
        svgMap.set(id, name);
      }
      return rels.add(RT('image'), K.relative(owner, 'ppt/media/' + svgMap.get(id)));
    };

    /* masters & layouts */
    const masterIdList = [];
    const layoutPartOf = new Map(); /* designId|layoutKey|lkey -> layout file number */
    let layoutN = 0;
    const masterN = new Map();
    let firstTheme = null;
    designIds.forEach((did, mi) => {
      const d = pres.designs[did] || L.model.buildDesign('default', pres.W, pres.H);
      const mNum = mi + 1;
      if (d.keep?.master) {
        try {
          const kept = L.designs.write(d, pres, pack, {
            name: freshName, master: masterXML, layout: layoutXML, theme: themeXML, shape: shapeXML,
            context: (part, shapes, owner) => {
              pack.use(part);
              // Generated masters/layouts use 1 for their structural spTree.
              // Reserve it before allocating placeholders in a new layout.
              owner.ids.reserve(part, 'shape', '1');
              const rels = pack.rels(part), ids = new Map(), byId = new Map();
              L.model.walk(shapes, s => { byId.set(s.id, s); return true; });
              return { writer: owner, frames, part, rels, media: mediaSync(rels, part), svg: svgSync(rels, part), slideIndex,
                fieldId: fieldIds(part), slideTarget: i => K.relative(part, slideParts[i]),
                nextId: mid => {
                  if (ids.has(mid)) return ids.get(mid);
                  const f = byId.get(mid)?.keep?.identity, ref = f?.ids.find(r => r.kind === 'shape' && r.definition);
                  const id = ref ? owner.ids.resolve(ref.source, ref.scope, ref.kind, ref.id, { primary: pres.pkg?.id, scope: part }) : owner.ids.fresh(part, 'shape');
                  if (mid) ids.set(mid, id); return id;
                }, idOf: mid => ids.get(mid), findShape: mid => byId.get(mid),
                addChart: (xml, parts) => { chartN++; charts.push({ n: chartN, xml, parts }); return chartN; } };
            },
          });
          if (kept) {
            if (!mi) firstTheme = kept.theme;
            masterIdList.push({ id: kept.id, rid: writer.rels(presRels.owner).add(RT('slideMaster'), K.relative(presRels.owner, kept.part), false, kept.rid) });
            for (const [key, part] of kept.layouts) layoutPartOf.set(did + '|' + key, { part });
            masterN.set(did, mNum); return;
          }
        } catch (error) {
          if (error.code !== 'OOXML_DESIGN_DEPENDENCY') throw error;
          writer.loss({ id: 'design:' + d.keep.part, what: 'The design was converted because its original dependencies are incomplete: ' + error.message, where: d.keep.part, action: 'conversion' });
        }
      }
      const original = !d.keep?.master && d.keep?.source === pres.pkg?.id && d.keep?.values === L.preserve.designValues(d);
      const used = allSlides.filter(s => s.design === did).map(s => s.layout + (s.lkey ? '|' + s.lkey : ''));
      if (original && used.every(key => d.keep.layouts[key])) {
        try {
          const part = writer.carry(pres.pkg, d.keep.part);
          pack.bind(`ppt/slideMasters/slideMaster${mNum}.xml`, d.keep.part, 'opaque');
          pack.bind(`ppt/theme/theme${mNum}.xml`, d.keep.theme, 'opaque');
          const dep = pres.pkg.rels(pres.pkg.main).find(r => r.part === d.keep.part && K.relationshipType(r.type) === RT('slideMaster'));
          const id = Array.from(pres.pkg.xml(pres.pkg.main).getElementsByTagName('*')).find(e => e.localName === 'sldMasterId' && Array.from(e.attributes).some(a => a.localName === 'id' && a.namespaceURI?.includes('relationships') && a.value === dep?.id))?.getAttribute('id');
          const rid = writer.rels(presRels.owner).add(dep?.type || RT('slideMaster'), K.relative(presRels.owner, part), false, dep?.id);
          masterIdList.push({ id: id || writer.ids.fresh('presentation', 'sldMasterId'), rid });
          for (const key of used) layoutPartOf.set(did + '|' + key, { original: d.keep.layouts[key] });
          masterN.set(did, mNum);
          return;
        } catch (error) {
          writer.loss({ id: 'design:' + d.keep.part, what: 'The design was converted because its original dependencies are incomplete: ' + error.message, where: d.keep.part, action: 'conversion' });
        }
      }
      pack.bind(`ppt/slideMasters/slideMaster${mNum}.xml`, d.keep?.part);
      pack.bind(`ppt/theme/theme${mNum}.xml`, d.keep?.theme, d.keep?.themeValues === JSON.stringify([d.colors, d.fonts]) ? 'opaque' : 'regenerated');
      const mRels = pack.rels(`ppt/slideMasters/slideMaster${mNum}.xml`);
      masterN.set(did, mNum);
      const mid = writer.ids.fresh('presentation', 'sldMasterId');
      const keys = new Set(['title', 'text', 'blank', 'titleOnly', 'content']);
      for (const s of allSlides) if (s.design === did) keys.add(s.layout + (s.lkey ? '|' + s.lkey : ''));
      const layoutIds = [];
      for (const k of keys) {
        const [key, lkey] = k.split('|');
        const n = ++layoutN;
        layoutPartOf.set(did + '|' + k, n);
        pack.bind(`ppt/slideLayouts/slideLayout${n}.xml`, d.keep?.layouts?.[k]);
        const lRels = pack.rels(`ppt/slideLayouts/slideLayout${n}.xml`);
        lRels.add(RT('slideMaster'), `../slideMasters/slideMaster${mNum}.xml`);
        const lctx = { rels: lRels, media: mediaSync(lRels), slideIndex, nextId: null, idOf: () => null, fieldId: fieldIds(`ppt/slideLayouts/slideLayout${n}.xml`) };
        add(`ppt/slideLayouts/slideLayout${n}.xml`, layoutXML(d, key, lkey, lctx));
        add(`ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`, lRels.xml());
        overrides.push([`/ppt/slideLayouts/slideLayout${n}.xml`, CT.layout]);
        const rid = mRels.add(RT('slideLayout'), `../slideLayouts/slideLayout${n}.xml`);
        layoutIds.push({ id: writer.ids.fresh('presentation', 'sldLayoutId'), rid });
      }
      const themeN = mNum;
      mRels.add(RT('theme'), `../theme/theme${themeN}.xml`);
      add(`ppt/theme/theme${themeN}.xml`, themeXML(d, d.name));
      overrides.push([`/ppt/theme/theme${themeN}.xml`, CT.theme]);
      const mctx = { rels: mRels, media: mediaSync(mRels), slideIndex, nextId: null, idOf: () => null, fieldId: fieldIds(`ppt/slideMasters/slideMaster${mNum}.xml`) };
      add(`ppt/slideMasters/slideMaster${mNum}.xml`, masterXML(d, mctx, layoutIds));
      add(`ppt/slideMasters/_rels/slideMaster${mNum}.xml.rels`, mRels.xml());
      overrides.push([`/ppt/slideMasters/slideMaster${mNum}.xml`, CT.master]);
      masterIdList.push({ id: mid, rid: presRels.add(RT('slideMaster'), `slideMasters/slideMaster${mNum}.xml`) });
    });

    /* notes master */
    const notesThemeN = designIds.length + 1;
    const sourceNotesMaster = pack.originals.get('ppt/notesMasters/notesMaster1.xml');
    let keptNotesMaster = false;
    if (sourceNotesMaster) {
      try { writer.carry(pres.pkg, sourceNotesMaster); keptNotesMaster = true; }
      catch (_) { /* The normal generated fallback below reports the conversion. */ }
    }
    if (!keptNotesMaster) {
    add(`ppt/theme/theme${notesThemeN}.xml`, themeXML(pres.designs[designIds[0]] || L.model.buildDesign('default', pres.W, pres.H), 'Notes Theme'));
    overrides.push([`/ppt/theme/theme${notesThemeN}.xml`, CT.theme]);
    add('ppt/notesMasters/notesMaster1.xml', notesMasterXML(pres));
    const nmRels = pack.rels('ppt/notesMasters/notesMaster1.xml');
    nmRels.add(RT('theme'), `../theme/theme${notesThemeN}.xml`);
    add('ppt/notesMasters/_rels/notesMaster1.xml.rels', nmRels.xml());
    overrides.push(['/ppt/notesMasters/notesMaster1.xml', CT.notesMaster]);
    }
    const notesMasterRid = presRels.add(RT('notesMaster'), 'notesMasters/notesMaster1.xml');

    /* slides */
    const sldIds = []; writer.slideIds = new Map();
    const commentState = L.comments.begin(pres, pack);
    allSlides.forEach((s, i) => {
      const n = i + 1;
      const part = slideParts[i];
      const originalSlide = s.keep?.part === part && s.keep?.source === pres.pkg?.id;
      const slideId = originalSlide && s.keep.sldId ? s.keep.sldId : writer.ids.fresh('presentation', 'sldId');
      writer.slideIds.set(s.id, slideId);
      const d = pres.designs[s.design] || pres.designs[designIds[0]];
      const sRels = pack.rels(part);
      const lk = s.design + '|' + s.layout + (s.lkey ? '|' + s.lkey : '');
      const ln = layoutPartOf.get(lk) || layoutPartOf.get(designIds[0] + '|text') || 1;
      if (ln.part) {
        const dep = pres.pkg?.rels(s.keep?.part || '').find(r => r.part === ln.part && K.relationshipType(r.type) === RT('slideLayout'));
        writer.rels(sRels.owner).add(dep?.type || RT('slideLayout'), K.relative(sRels.owner, ln.part), false, dep?.id);
      } else if (ln.original) {
        const dep = pres.pkg.rels(s.keep?.part || '').find(r => r.part === ln.original && K.relationshipType(r.type) === RT('slideLayout'));
        writer.rels(sRels.owner).add(dep?.type || RT('slideLayout'), K.relative(sRels.owner, writer.target(pres.pkg, ln.original)), false, dep?.id);
      } else sRels.add(RT('slideLayout'), K.relative(part, `ppt/slideLayouts/slideLayout${ln}.xml`));
      let rootId = s.keep?.rootId || '1'; writer.ids.reserve(part, 'shape', rootId);
      // Some source decks reuse the structural spTree ID for an actual shape.
      // Keep the visible shape's identity (timing may target it) and repair the root.
      let rootClash = false;
      L.model.walk(s.shapes, shape => { if (shape.keep?.identity?.ids.some(ref => ref.kind === 'shape' && ref.definition && String(ref.id) === String(rootId) && ref.scope === part)) rootClash = true; return true; });
      if (rootClash) rootId = writer.ids.fresh(part, 'shape');
      const idMap = new Map();
      const ctx = {
        writer, frames, part, place: 'Slide ' + n, rels: sRels, media: mediaSync(sRels, part), svg: svgSync(sRels, part), slideIndex, fieldId: fieldIds(part),
        slideTarget: index => K.relative(part, slideParts[index]),
        nextId: (mid) => {
          if (idMap.has(mid)) return idMap.get(mid);
          const f = mid && L.model.shapeById(s, mid)?.keep?.identity, ref = f?.ids.find(i => i.kind === 'shape' && i.definition);
          if (f?.copy) writer.ids.copy(f.copy, f.copyDefinitions || f.ids.filter(i => i.definition));
          const v = ref ? writer.ids.resolve(ref.source, ref.scope, ref.kind, ref.id, { primary: pres.pkg?.id, scope: part, copy: f.copy }) : writer.ids.fresh(part, 'shape');
          if (mid) idMap.set(mid, v); return v;
        },
        idOf: (mid) => idMap.get(mid),
        findShape: (mid) => L.model.shapeById(s, mid),
        addChart: (xml, parts) => { chartN++; charts.push({ n: chartN, xml, parts }); return chartN; },
      };
      L.model.walk(s.shapes, shape => { ctx.nextId(shape.id); return true; });
      ctx.liveShapeIds = new Set(Array.from(idMap.values(), String));
      const comments = L.comments.slide(commentState, s, ctx);
      let tree = s.shapes.map((sh) => shapeXML(sh, ctx, d)).join('');
      /* header & footer placeholders */
      const hf = Object.assign({}, pres.hf || {}, s.hf || {});
      if (!(hf.notOnTitle && s.layout === 'title')) {
        const has = (t) => s.shapes.some((x) => x.ph && x.ph.type === t);
        const foot = (type, run, name) => {
          const r = d.ph[type];
          if (!r) return '';
          const id = ctx.nextId();
          const tx = L.txt.body([{ lvl: 0, pp: {}, rs: [run] }]);
          return placeholderSp(id, name + ' ' + id, { type, idx: { dt: 10, ftr: 11, sldNum: 12 }[type] }, r, tx, ctx, d);
        };
        if (hf.dt && !has('dt')) tree += foot('dt', hf.dtAuto !== false ? { t: L.fmtDate(new Date(), hf.dtFmt || 'datetime1'), fld: hf.dtFmt || 'datetime1' } : { t: hf.dtText || '' }, 'Date Placeholder');
        if (hf.ftr && hf.ftrText && !has('ftr')) tree += foot('ftr', { t: hf.ftrText }, 'Footer Placeholder');
        if (hf.num && !has('sldNum')) tree += foot('sldNum', { t: String(i + (pres.firstNum || 1)), fld: 'slidenum' }, 'Slide Number Placeholder');
      }
      const timing = L.preserve.timing(s, ctx, () => timingXML(s, ctx));
      const transition = L.properties.transition(s, ctx, transitionXML(s.trans));
      const attrs = [];
      if (s.hidden) attrs.push('show="0"');
      if (s.hideMaster) attrs.push('showMasterSp="0"');
      const xml = L.properties.finishSlide(s, ctx, HEAD + `<p:sld ${NSDECL}${attrs.length ? ' ' + attrs.join(' ') : ''}><p:cSld>${s.bg ? bgXML(s.bg, ctx, d) : ''}<p:spTree>${GRP.replace('id="1"', 'id="' + rootId + '"')}${tree}</p:spTree>${comments.creation}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transition}${timing}${comments.ext}</p:sld>`);
      if (s.notes && s.notes.trim() || s.keep?.notes) {
        if (s.keep?.copy && s.keep.notes?.text === s.notes && pres.pkg?.has(s.keep.notes.part)) {
          const target = pack.part(`ppt/notesSlides/notesSlide${n}.xml`);
          writer.copyPart(pres.pkg, s.keep.notes.part, target, { [s.keep.part]: part });
          sRels.add(RT('notesSlide'), K.relative(part, `ppt/notesSlides/notesSlide${n}.xml`));
        } else {
        if (s.keep?.notes && s.keep.notes.text !== s.notes) writer.loss({ id: 'notes-edit:' + s.id, what: 'Editing these notes replaces their original notes-page formatting.', where: s.keep.notes.part, action: 'conversion' });
        pack.bind(`ppt/notesSlides/notesSlide${n}.xml`, s.keep?.notes?.part, s.keep?.notes?.text === s.notes ? 'opaque' : 'regenerated');
        const nRels = pack.rels(`ppt/notesSlides/notesSlide${n}.xml`);
        nRels.add(RT('notesMaster'), '../notesMasters/notesMaster1.xml');
        nRels.add(RT('slide'), K.relative(`ppt/notesSlides/notesSlide${n}.xml`, part));
        add(`ppt/notesSlides/notesSlide${n}.xml`, notesXML(pres, s.notes));
        add(`ppt/notesSlides/_rels/notesSlide${n}.xml.rels`, nRels.xml());
        overrides.push([`/ppt/notesSlides/notesSlide${n}.xml`, CT.notes]);
        sRels.add(RT('notesSlide'), K.relative(part, `ppt/notesSlides/notesSlide${n}.xml`));
        }
      }
      add(part, xml);
      add(K.relsPath(part), sRels.xml());
      overrides.push(['/' + part, CT.slide]);
      sldIds.push({ id: slideId, rid: presRels.add(RT('slide'), K.relative('ppt/presentation.xml', part)) });
    });
    L.comments.finish(commentState);
    /* charts, and the parts an unedited imported chart brings with it (workbook, style, colours, theme override) */
    let embN = 0;
    const PART_CT = { xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12', xlsb: 'application/vnd.ms-excel.sheet.binary.macroEnabled.main', xls: 'application/vnd.ms-excel', bin: 'application/vnd.openxmlformats-officedocument.oleObject', xml: 'application/xml', png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', emf: 'image/x-emf', wmf: 'image/x-wmf' };
    for (const c of charts) {
      add(`ppt/charts/chart${c.n}.xml`, c.xml);
      overrides.push([`/ppt/charts/chart${c.n}.xml`, CT.chart]);
      if (!c.parts || !c.parts.length) continue;
      const rels = [];
      c.parts.forEach((p, k) => {
        if (p.external) { rels.push(`<Relationship Id="${X(p.id)}" Type="${X(p.fullType)}" Target="${X(p.target)}" TargetMode="External"/>`); return; }
        const ext = (String(p.name || '').split('.').pop() || 'bin').toLowerCase().replace(/[^a-z0-9]/g, '') || 'bin';
        let part;
        if (p.type === 'package' || p.type === 'oleObject') part = `ppt/embeddings/${p.type === 'package' ? 'Microsoft_Excel_Worksheet' : 'oleObject'}${++embN}.${ext}`;
        else if (p.type === 'chartStyle') part = `ppt/charts/style${c.n}.xml`;
        else if (p.type === 'chartColorStyle') part = `ppt/charts/colors${c.n}.xml`;
        else if (p.type === 'themeOverride') part = `ppt/theme/themeOverride${c.n}.xml`;
        else if (p.type === 'image') part = `ppt/media/chart${c.n}_${k + 1}.${ext}`;
        else part = `ppt/charts/chart${c.n}_part${k + 1}.${ext}`;
        add(part, p.bytes);
        /* binary parts are typed by extension (as PowerPoint writes them), XML parts by name */
        const ct = p.ct || PART_CT[ext] || 'application/octet-stream';
        if (ext !== 'xml' && (!defaults.has(ext) || defaults.get(ext) === ct)) defaults.set(ext, ct);
        else overrides.push(['/' + part, ct]);
        const target = part.startsWith('ppt/charts/') ? part.slice('ppt/charts/'.length) : '../' + part.slice('ppt/'.length);
        rels.push(`<Relationship Id="${X(p.id)}" Type="${X(p.fullType)}" Target="${X(target)}"/>`);
      });
      add(`ppt/charts/_rels/chart${c.n}.xml.rels`, HEAD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
    }

    /* presentation-level parts */
    presRels.add(RT('presProps'), 'presProps.xml');
    presRels.add(RT('viewProps'), 'viewProps.xml');
    if (firstTheme) writer.rels(presRels.owner).add(RT('theme'), K.relative(presRels.owner, firstTheme));
    else presRels.add(RT('theme'), 'theme/theme1.xml');
    presRels.add(RT('tableStyles'), 'tableStyles.xml');
    const sizeType = { '720x540': 'screen4x3', '720x405': 'screen16x9', '720x450': 'screen16x10', '780x540': 'A4', '810x540': '35mm', '576x72': 'banner' }[`${Math.round(pres.W)}x${Math.round(pres.H)}`] || 'custom';
    const defStyle = levelsXML(L.model.DEFAULT_TX.other, { rels: presRels, media: () => null }, pres.designs[designIds[0]]);
    const presXML = HEAD + `<p:presentation ${NSDECL} saveSubsetFonts="1"${pres.firstNum && pres.firstNum !== 1 ? ` firstSlideNum="${pres.firstNum}"` : ''}><p:sldMasterIdLst>${masterIdList.map((m) => `<p:sldMasterId id="${m.id}" r:id="${m.rid}"/>`).join('')}</p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRid}"/></p:notesMasterIdLst>${sldIds.length ? `<p:sldIdLst>${sldIds.map((s) => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('')}</p:sldIdLst>` : ''}<p:sldSz cx="${emu(pres.W)}" cy="${emu(pres.H)}"${sizeType !== 'custom' ? ` type="${sizeType}"` : ''}/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${defStyle}</p:defaultTextStyle></p:presentation>`;
    add('ppt/presentation.xml', presXML);
    add('ppt/_rels/presentation.xml.rels', presRels.xml());
    overrides.unshift(['/ppt/presentation.xml', mainCT]);
    const sh = pres.show || {};
    add('ppt/presProps.xml', HEAD + `<p:presentationPr ${NSDECL}><p:showPr${sh.loop ? ' loop="1"' : ''} showNarration="1"${sh.noAnim ? ' showAnimation="0"' : ''}${sh.useTimings === false ? ' useTimings="0"' : ''}>${sh.kiosk ? '<p:kiosk/>' : '<p:present/>'}${sh.from && sh.to ? `<p:sldRg st="${sh.from}" end="${sh.to}"/>` : '<p:sldAll/>'}<p:penClr>${clr(sh.penColor || '#FF0000')}</p:penClr></p:showPr></p:presentationPr>`);
    overrides.push(['/ppt/presProps.xml', CT.presProps]);
    add('ppt/viewProps.xml', HEAD + `<p:viewPr ${NSDECL}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`);
    overrides.push(['/ppt/viewProps.xml', CT.viewProps]);
    add('ppt/tableStyles.xml', HEAD + `<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
    overrides.push(['/ppt/tableStyles.xml', CT.tableStyles]);

    /* document properties */
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const p = pres.props || {};
    add('docProps/core.xml', HEAD + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${X(p.title || pres.title || '')}</dc:title><dc:subject>${X(p.subject || '')}</dc:subject><dc:creator>${X(p.author || '')}</dc:creator><cp:keywords>${X(p.keywords || '')}</cp:keywords><dc:description>${X(p.comments || '')}</dc:description><cp:lastModifiedBy>${X(p.author || '')}</cp:lastModifiedBy><cp:revision>${(p.revision || 0) + 1}</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">${X((p.created || now).replace(/\.\d+Z$/, 'Z'))}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>${p.category ? `<cp:category>${X(p.category)}</cp:category>` : ''}</cp:coreProperties>`);
    overrides.push(['/docProps/core.xml', CT.core]);
    const words = allSlides.reduce((a, s) => a + (L.model.allText(s).match(/\S+/g) || []).length, 0);
    const fmtName = (L.model.SLIDE_SIZES.find((z) => z.w === pres.W && z.h === pres.H) || { name: 'Custom' }).name;
    add('docProps/app.xml', HEAD + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Words>${words}</Words><Application>Lectern 2003 Web Edition</Application><PresentationFormat>${X(fmtName)}</PresentationFormat><Paragraphs>0</Paragraphs><Slides>${allSlides.length}</Slides><Notes>${allSlides.filter((s) => s.notes && s.notes.trim()).length}</Notes><HiddenSlides>${allSlides.filter((s) => s.hidden).length}</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><Company>${X(p.company || '')}</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>11.0000</AppVersion></Properties>`);
    overrides.push(['/docProps/app.xml', CT.app]);

    const rootRels = pack.rels('');
    rootRels.add(RT('officeDocument'), 'ppt/presentation.xml');
    rootRels.add('http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml');
    rootRels.add(RT('extended-properties'), 'docProps/app.xml');
    add('_rels/.rels', rootRels.xml());

    const ct = HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      Array.from(defaults.entries()).map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join('') +
      overrides.map(([pn, t]) => `<Override PartName="${pn}" ContentType="${t}"/>`).join('') + '</Types>';
    files.unshift({ name: '[Content_Types].xml', data: ct });
    const mime = format.mime;
    const types = new Map(overrides.map(([name, type]) => [name.slice(1), type]));
    for (const file of files) pack.put(file.name, file.data, types.get(file.name) || defaults.get(file.name.split('.').pop()));
    const result = pack.finish();
    const blob = await L.zip.write(result.files, mime);
    blob.dropped = result.dropped;
    /* a password to open: the package is encrypted (AES-256) inside a compound file, as PowerPoint 2013 and later do */
    if (pres.password && L.officeCrypto) {
      const enc = await L.officeCrypto.encrypt(new Uint8Array(await blob.arrayBuffer()), pres.password);
      const encrypted = new Blob([enc], { type: mime }); encrypted.dropped = result.dropped; return encrypted;
    }
    return blob;
  }

  L.pptx = L.pptx || {};
  function tileXML(o) {
    o = o || {};
    const ALG = ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'], FL = ['none', 'x', 'y', 'xy'];
    return `<a:tile tx="${Math.round((o.tx || 0) * 12700)}" ty="${Math.round((o.ty || 0) * 12700)}" sx="${Math.round((o.sx || 1) * 100000)}" sy="${Math.round((o.sy || 1) * 100000)}" flip="${FL.includes(o.flip) ? o.flip : 'none'}" algn="${ALG.includes(o.algn) ? o.algn : 'tl'}"/>`;
  }
  L.pptx.write = write;
  L.pptx._internal = { chartXML, themeXML, transitionXML };
})();
