/* Quire — WordprocessingML writer. Produces a package that follows the OOXML schema element order,
 * readable by Word 2007+, LibreOffice and other consumers.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D;
  const X = L.xesc;
  const W = (L.docx = L.docx || {});
  const emu = L.pt2emu;
  // Preserved paragraph properties can have namespace/compatibility attributes.
  // Insert generated content after them, including in paragraphs with IDs.
  const prependParagraph = (xml, contents) => xml.replace(/<w:p\b[^>]*>(\s*<w:pPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:pPr>))?/, m => m + contents);
  const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const NS = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture', wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup', mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    v: 'urn:schemas-microsoft-com:vml', o: 'urn:schemas-microsoft-com:office:office', w10: 'urn:schemas-microsoft-com:office:word',
    m: 'http://schemas.openxmlformats.org/officeDocument/2006/math', c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    w14: 'http://schemas.microsoft.com/office/word/2010/wordml', w15: 'http://schemas.microsoft.com/office/word/2012/wordml', wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  };
  const NSDECL = `xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="${NS.mc}" xmlns:o="${NS.o}" xmlns:r="${NS.r}" xmlns:m="${NS.m}" xmlns:v="${NS.v}" xmlns:wp14="${NS.wp14}" xmlns:wp="${NS.wp}" xmlns:w10="${NS.w10}" xmlns:w="${NS.w}" xmlns:w14="${NS.w14}" xmlns:w15="${NS.w15}" xmlns:wpg="${NS.wpg}" xmlns:wps="${NS.wps}" mc:Ignorable="w14 w15 wp14"`;
  const RT = (n) => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + n;
  const CT = {
    doc: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    tmpl: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
    styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
    numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
    settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    fontTable: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
    webSettings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml',
    header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
    footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
    endnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
    comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
    commentsEx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
    theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
    core: 'application/vnd.openxmlformats-package.core-properties+xml',
    app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    custom: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
    chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    chartStyle: 'application/vnd.ms-office.chartstyle+xml',
    chartColors: 'application/vnd.ms-office.chartcolorstyle+xml',
  };
  const tw = (pt) => Math.round((pt || 0) * 20);
  const hp = (pt) => Math.round((pt || 0) * 2);
  const ON = (b) => (b === false ? ' w:val="0"' : '');

  const DT = (d) => (d ? `w:date="${X(d)}"` : '');
  const hex = (c) => (!c || c === 'auto' ? 'auto' : String(c).replace('#', '').toUpperCase());

  /* ================= properties ================= */
  function borderEl(tag, b) {
    if (!b) return '';
    if (!b.val || b.val === 'nil' || b.val === 'none') return `<w:${tag} w:val="nil"/>`;
    return `<w:${tag} w:val="${b.val}" w:sz="${Math.max(2, Math.round((b.sz || 0.5) * 8))}" w:space="${Math.round(b.space || 0)}" w:color="${hex(b.color)}"${b.shadow ? ' w:shadow="1"' : ''}/>`;
  }
  function bordersEl(tag, bd, names) {
    if (!bd) return '';
    const inner = names.map((n) => borderEl(n, bd[n])).join('');
    return inner ? `<w:${tag}>${inner}</w:${tag}>` : '';
  }
  function shdEl(s) {
    if (!s) return '';
    return `<w:shd w:val="${s.val || 'clear'}" w:color="${hex(s.color || 'auto')}" w:fill="${hex(s.fill || 'auto')}"/>`;
  }
  /** run properties; mark: paragraph-mark revision info */
  function rPrXML(r, mark, ctx) {
    r = r || {};
    let x = '';
    if (mark) {
      if (mark.ins) x += `<w:ins w:id="${ctx.rev()}" w:author="${X(mark.ins.author)}" ${DT(mark.ins.date)}/>`;
      if (mark.del) x += `<w:del w:id="${ctx.rev()}" w:author="${X(mark.del.author)}" ${DT(mark.del.date)}/>`;
    }
    if (r.style) x += `<w:rStyle w:val="${X(r.style)}"/>`;
    if (r.font || r.fontEA || r.fontCS) {
      const a = [];
      if (r.font) a.push(`w:ascii="${X(r.font)}" w:hAnsi="${X(r.font)}"`);
      if (r.fontEA) a.push(`w:eastAsia="${X(r.fontEA)}"`);
      if (r.fontCS) a.push(`w:cs="${X(r.fontCS)}"`);
      if (r.hint) a.push(`w:hint="${X(r.hint)}"`);
      x += `<w:rFonts ${a.join(' ')}/>`;
    }
    const tog = (k, tag) => { if (r[k] != null) x += `<w:${tag}${ON(r[k])}/>`; };
    tog('b', 'b'); tog('bCs', 'bCs'); tog('i', 'i'); tog('iCs', 'iCs'); tog('caps', 'caps'); tog('smallCaps', 'smallCaps'); tog('strike', 'strike'); tog('dstrike', 'dstrike');
    tog('outline', 'outline'); tog('shadow', 'shadow'); tog('emboss', 'emboss'); tog('imprint', 'imprint'); tog('noProof', 'noProof'); tog('hidden', 'vanish'); tog('webHidden', 'webHidden');
    if (r.color) x += `<w:color w:val="${hex(r.color)}"/>`;
    if (r.spacing) x += `<w:spacing w:val="${tw(r.spacing)}"/>`;
    if (r.w && r.w !== 100) x += `<w:w w:val="${Math.round(r.w)}"/>`;
    if (r.kern) x += `<w:kern w:val="${hp(r.kern)}"/>`;
    if (r.pos) x += `<w:position w:val="${hp(r.pos)}"/>`;
    if (r.sz) x += `<w:sz w:val="${hp(r.sz)}"/>`;
    if (r.szCs || r.sz) x += `<w:szCs w:val="${hp(r.szCs || r.sz)}"/>`;
    if (r.hl) x += `<w:highlight w:val="${r.hl}"/>`;
    if (r.u) x += `<w:u w:val="${r.u}"${r.uColor ? ` w:color="${hex(r.uColor)}"` : ''}/>`;
    if (r.effect) x += `<w:effect w:val="${r.effect}"/>`;
    if (r.border && r.border.val) x += borderEl('bdr', r.border);
    if (r.shd) x += shdEl(r.shd);
    if (r.vert) x += `<w:vertAlign w:val="${r.vert}"/>`;
    tog('rtl', 'rtl'); tog('cs', 'cs');
    if (r.em) x += `<w:em w:val="${r.em}"/>`;
    if (r.lang || r.langEA) x += `<w:lang${r.lang ? ` w:val="${X(r.lang)}"` : ''}${r.langEA ? ` w:eastAsia="${X(r.langEA)}"` : ''}/>`;
    tog('specVanish', 'specVanish');
    if (r.chg && ctx) { const old = { ...r.chg.old }; delete old.chg; x += `<w:rPrChange w:id="${ctx.rev()}" w:author="${X(r.chg.author || 'Unknown')}" ${DT(r.chg.date)}>${rPrXML(old, null, ctx) || '<w:rPr/>'}</w:rPrChange>`; }
    return L.preserve.propertyXML('rPr', r, x, ctx);
  }
  function pPrXML(p, ctx, opts) {
    p = p || {};
    let x = '';
    if (p.style && p.style !== 'Normal') x += `<w:pStyle w:val="${X(p.style)}"/>`;
    else if (p.style === 'Normal' && opts && opts.inStyle) x += '';
    if (p.keepNext != null) x += `<w:keepNext${ON(p.keepNext)}/>`;
    if (p.keepLines != null) x += `<w:keepLines${ON(p.keepLines)}/>`;
    if (p.pageBreakBefore != null) x += `<w:pageBreakBefore${ON(p.pageBreakBefore)}/>`;
    if (p.frame || p.dropCap) {
      if (p.dropCap) x += `<w:framePr w:dropCap="${p.dropCap.type || 'drop'}" w:lines="${p.dropCap.lines || 3}" w:hSpace="${tw(p.dropCap.dist || 0)}" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>`;
      else {
        const f = p.frame, a = [];
        if (f.w) a.push(`w:w="${tw(f.w)}"`);
        if (f.h) a.push(`w:h="${tw(f.h)}"`);
        if (f.hRule) a.push(`w:hRule="${f.hRule}"`);
        if (f.hSpace != null) a.push(`w:hSpace="${tw(f.hSpace)}"`);
        if (f.vSpace != null) a.push(`w:vSpace="${tw(f.vSpace)}"`);
        if (f.wrap) a.push(`w:wrap="${f.wrap}"`);
        if (f.vAnchor) a.push(`w:vAnchor="${f.vAnchor}"`);
        if (f.hAnchor) a.push(`w:hAnchor="${f.hAnchor}"`);
        if (typeof f.x === 'number') a.push(`w:x="${tw(f.x)}"`);
        if (f.xAlign) a.push(`w:xAlign="${f.xAlign}"`);
        if (typeof f.y === 'number') a.push(`w:y="${tw(f.y)}"`);
        if (f.yAlign) a.push(`w:yAlign="${f.yAlign}"`);
        x += `<w:framePr ${a.join(' ')}/>`;
      }
    }
    if (p.widow != null) x += `<w:widowControl${ON(p.widow)}/>`;
    if (p.num && !(p.num.id != null && p.num.id !== '0' && D.doc && !D.doc.numbering.nums[p.num.id])) {
      const id = p.num.id != null ? (ctx && ctx.numMap ? ctx.numMap(p.num.id) : p.num.id) : null;
      x += `<w:numPr>${p.num.lvl != null ? `<w:ilvl w:val="${p.num.lvl || 0}"/>` : ''}${id != null ? `<w:numId w:val="${id}"/>` : ''}</w:numPr>`;
    }
    if (p.noLineNum) x += '<w:suppressLineNumbers/>';
    x += bordersEl('pBdr', p.borders, ['top', 'left', 'bottom', 'right', 'between', 'bar']);
    if (p.shd) x += shdEl(p.shd);
    if (p.tabs && p.tabs.length) x += '<w:tabs>' + p.tabs.slice().sort((a, b) => a.pos - b.pos).map((t) => `<w:tab w:val="${t.al === 'num' ? 'num' : t.al || 'left'}"${t.leader && t.leader !== 'none' ? ` w:leader="${t.leader}"` : ''} w:pos="${tw(t.pos)}"/>`).join('') + '</w:tabs>';
    if (p.noHyphen != null) x += `<w:suppressAutoHyphens${ON(p.noHyphen)}/>`;
    if (p.bidi) x += '<w:bidi/>';
    if (p.noSnap) x += '<w:snapToGrid w:val="0"/>';
    if (p.sp) {
      const s = p.sp, a = [];
      if (s.b != null) a.push(`w:before="${Math.max(0, tw(s.b))}"`);
      if (s.bAuto != null) a.push(`w:beforeAutospacing="${s.bAuto ? 1 : 0}"`);
      if (s.a != null) a.push(`w:after="${Math.max(0, tw(s.a))}"`);
      if (s.aAuto != null) a.push(`w:afterAutospacing="${s.aAuto ? 1 : 0}"`);
      if (s.line != null) { const rule = s.rule || 'auto'; a.push(`w:line="${rule === 'auto' ? Math.round(s.line * 240) : tw(s.line)}"`, `w:lineRule="${rule}"`); }
      if (a.length) x += `<w:spacing ${a.join(' ')}/>`;
    }
    if (p.ind) {
      const i = p.ind, a = [];
      if (i.l != null) a.push(`w:left="${tw(i.l)}"`);
      if (i.r != null) a.push(`w:right="${tw(i.r)}"`);
      if (i.fl != null) a.push(i.fl < 0 ? `w:hanging="${tw(-i.fl)}"` : `w:firstLine="${tw(i.fl)}"`);
      if (a.length) x += `<w:ind ${a.join(' ')}/>`;
    }
    if (p.contextual != null) x += `<w:contextualSpacing${ON(p.contextual)}/>`;
    if (p.mirrorInd) x += '<w:mirrorIndents/>';
    if (p.jc) x += `<w:jc w:val="${p.jc}"/>`;
    if (p.textDir) x += `<w:textDirection w:val="${p.textDir}"/>`;
    if (p.textAlign) x += `<w:textAlignment w:val="${p.textAlign}"/>`;
    if (p.outline != null && p.outline < 9) x += `<w:outlineLvl w:val="${p.outline}"/>`;
    if (opts && opts.rPr != null) x += opts.rPr;
    if (opts && opts.sect) x += opts.sect;
    if (p.chg && ctx) { const o = Object.assign({}, p.chg.old || {}); delete o.chg; x += `<w:pPrChange w:id="${ctx.rev()}" w:author="${X(p.chg.author || 'Unknown')}" ${DT(p.chg.date)}>${pPrXML(o, ctx, { noChg: true }) || '<w:pPr/>'}</w:pPrChange>`; }
    return L.preserve.propertyXML('pPr', p, x, ctx);
  }
  function sectXML(s, ctx) {
    let x = '';
    const refs = s.refs || { hdr: {}, ftr: {} };
    for (const ty of ['default', 'first', 'even']) if (refs.hdr && refs.hdr[ty] && ctx.hfRid[refs.hdr[ty]]) x += `<w:headerReference w:type="${ty}" r:id="${ctx.hfRid[refs.hdr[ty]]}"/>`;
    for (const ty of ['default', 'first', 'even']) if (refs.ftr && refs.ftr[ty] && ctx.hfRid[refs.ftr[ty]]) x += `<w:footerReference w:type="${ty}" r:id="${ctx.hfRid[refs.ftr[ty]]}"/>`;
    if (s.fnPr && Object.keys(s.fnPr).length) x += notePrXML('footnotePr', s.fnPr, false);
    if (s.enPr && Object.keys(s.enPr).length) x += notePrXML('endnotePr', s.enPr, false);
    if (s.type && s.type !== 'nextPage') x += `<w:type w:val="${s.type}"/>`;
    x += `<w:pgSz w:w="${tw(s.pgW)}" w:h="${tw(s.pgH)}"${s.orient === 'landscape' ? ' w:orient="landscape"' : ''}${s.paperCode ? ` w:code="${s.paperCode}"` : ''}/>`;
    x += `<w:pgMar w:top="${tw(s.mt)}" w:right="${tw(s.mr)}" w:bottom="${tw(s.mb)}" w:left="${tw(s.ml)}" w:header="${tw(s.hdr)}" w:footer="${tw(s.ftr)}" w:gutter="${tw(s.gutter || 0)}"/>`;
    if (s.paperSrc) x += `<w:paperSrc w:first="${s.paperSrc.first || 0}" w:other="${s.paperSrc.other || 0}"/>`;
    if (s.borders && ['top', 'left', 'bottom', 'right'].some((k) => s.borders[k])) {
      const b = s.borders;
      x += `<w:pgBorders w:offsetFrom="${b.offsetFrom === 'page' ? 'page' : 'text'}"${b.display && b.display !== 'allPages' ? ` w:display="${b.display}"` : ''}>${['top', 'left', 'bottom', 'right'].map((k) => (b[k] && b[k].val !== 'nil' ? borderEl(k, Object.assign({ space: b.offsetFrom === 'page' ? 24 : 4 }, b[k])) : '')).join('')}</w:pgBorders>`;
    }
    if (s.lnNum) x += `<w:lnNumType w:countBy="${s.lnNum.countBy || 1}"${s.lnNum.start ? ` w:start="${s.lnNum.start}"` : ''}${s.lnNum.distance ? ` w:distance="${tw(s.lnNum.distance)}"` : ''} w:restart="${s.lnNum.restart || 'newPage'}"/>`;
    if (s.pgNum && (s.pgNum.start != null || s.pgNum.chapStyle || s.pgNum.chapSep || (s.pgNum.fmt && s.pgNum.fmt !== 'decimal'))) x += `<w:pgNumType${s.pgNum.fmt && s.pgNum.fmt !== 'decimal' ? ` w:fmt="${X(s.pgNum.fmt)}"` : ''}${s.pgNum.start != null ? ` w:start="${s.pgNum.start}"` : ''}${s.pgNum.chapStyle ? ` w:chapStyle="${X(s.pgNum.chapStyle)}"` : ''}${s.pgNum.chapSep ? ` w:chapSep="${X(s.pgNum.chapSep)}"` : ''}/>`;
    const c = s.cols || { n: 1 };
    if (c.n > 1 && c.eq === false && c.w && c.w.length) x += `<w:cols w:num="${c.n}" w:equalWidth="0"${c.sep ? ' w:sep="1"' : ''}>${c.w.map((cc) => `<w:col w:w="${tw(cc.w)}" w:space="${tw(cc.space || 0)}"/>`).join('')}</w:cols>`;
    else x += `<w:cols w:space="${tw(c.space != null ? c.space : 36)}"${c.n > 1 ? ` w:num="${c.n}"` : ''}${c.sep ? ' w:sep="1"' : ''}/>`;
    if (s.formProt) x += '<w:formProt/>';
    if (s.vAlign && s.vAlign !== 'top') x += `<w:vAlign w:val="${s.vAlign}"/>`;
    if (s.titlePg) x += '<w:titlePg/>';
    if (s.textDir) x += `<w:textDirection w:val="${s.textDir}"/>`;
    if (s.bidi) x += '<w:bidi/>';
    x += s.docGrid && (s.docGrid.type || s.docGrid.linePitch || s.docGrid.charSpace) ? `<w:docGrid${s.docGrid.type ? ` w:type="${X(s.docGrid.type)}"` : ''}${s.docGrid.linePitch ? ` w:linePitch="${s.docGrid.linePitch}"` : ''}${s.docGrid.charSpace ? ` w:charSpace="${s.docGrid.charSpace}"` : ''}/>` : '<w:docGrid w:linePitch="360"/>';
    return L.preserve.propertyXML('sectPr', s, x, ctx);
  }
  function notePrXML(tag, pr, inSettings) {
    let x = '';
    if (pr.pos) x += `<w:pos w:val="${pr.pos}"/>`;
    if (pr.fmt) x += `<w:numFmt w:val="${pr.fmt}"/>`;
    if (pr.start != null && pr.start !== 1) x += `<w:numStart w:val="${pr.start}"/>`;
    if (pr.restart && pr.restart !== 'continuous') x += `<w:numRestart w:val="${pr.restart}"/>`;
    if (inSettings) x += tag === 'footnotePr' ? '<w:footnote w:id="-1"/><w:footnote w:id="0"/>' : '<w:endnote w:id="-1"/><w:endnote w:id="0"/>';
    return x ? `<w:${tag}>${x}</w:${tag}>` : '';
  }
  function tblPrXML(t, ctx) {
    const p = t || {};
    let x = '';
    if (p.style) x += `<w:tblStyle w:val="${X(p.style)}"/>`;
    if (p.float && ctx && ctx.full) {
      const f = p.float, a = [];
      if (f.l != null) a.push(`w:leftFromText="${tw(f.l)}"`);
      if (f.r != null) a.push(`w:rightFromText="${tw(f.r)}"`);
      if (f.t != null) a.push(`w:topFromText="${tw(f.t)}"`);
      if (f.b != null) a.push(`w:bottomFromText="${tw(f.b)}"`);
      if (f.vAnchor) a.push(`w:vertAnchor="${f.vAnchor}"`);
      if (f.hAnchor) a.push(`w:horzAnchor="${f.hAnchor}"`);
      if (f.xAlign) a.push(`w:tblpXSpec="${f.xAlign}"`); else if (f.x != null) a.push(`w:tblpX="${tw(f.x)}"`);
      if (f.yAlign) a.push(`w:tblpYSpec="${f.yAlign}"`); else if (f.y != null) a.push(`w:tblpY="${tw(f.y)}"`);
      x += `<w:tblpPr ${a.join(' ')}/>`;
    }
    if (p.bidi) x += '<w:bidiVisual/>';
    if (p.rowBand) x += `<w:tblStyleRowBandSize w:val="${p.rowBand}"/>`;
    if (p.colBand) x += `<w:tblStyleColBandSize w:val="${p.colBand}"/>`;
    if (p.w) x += p.w.type === 'pct' ? `<w:tblW w:w="${Math.round(p.w.v)}" w:type="pct"/>` : p.w.type === 'dxa' ? `<w:tblW w:w="${tw(p.w.v)}" w:type="dxa"/>` : '<w:tblW w:w="0" w:type="auto"/>';
    if (p.jc) x += `<w:jc w:val="${p.jc === 'both' ? 'left' : p.jc}"/>`;
    if (p.spacing) x += `<w:tblCellSpacing w:w="${tw(p.spacing)}" w:type="dxa"/>`;
    if (p.ind != null) x += `<w:tblInd w:w="${tw(p.ind)}" w:type="dxa"/>`;
    x += bordersEl('tblBorders', p.borders, ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']);
    if (p.shd) x += shdEl(p.shd);
    if (p.layout) x += `<w:tblLayout w:type="${p.layout === 'fixed' ? 'fixed' : 'autofit'}"/>`;
    if (p.cellMar) { const m = p.cellMar; x += `<w:tblCellMar>${m.t != null ? `<w:top w:w="${tw(m.t)}" w:type="dxa"/>` : ''}${m.l != null ? `<w:left w:w="${tw(m.l)}" w:type="dxa"/>` : ''}${m.b != null ? `<w:bottom w:w="${tw(m.b)}" w:type="dxa"/>` : ''}${m.r != null ? `<w:right w:w="${tw(m.r)}" w:type="dxa"/>` : ''}</w:tblCellMar>`; }
    if (p.look) { const l = p.look; const v = (l.firstRow ? 0x20 : 0) | (l.lastRow ? 0x40 : 0) | (l.firstCol ? 0x80 : 0) | (l.lastCol ? 0x100 : 0) | (l.noHBand ? 0x200 : 0) | (l.noVBand ? 0x400 : 0); x += `<w:tblLook w:val="${v.toString(16).padStart(4, '0').toUpperCase()}" w:firstRow="${l.firstRow ? 1 : 0}" w:lastRow="${l.lastRow ? 1 : 0}" w:firstColumn="${l.firstCol ? 1 : 0}" w:lastColumn="${l.lastCol ? 1 : 0}" w:noHBand="${l.noHBand ? 1 : 0}" w:noVBand="${l.noVBand ? 1 : 0}"/>`; }
    if (ctx && ctx.full) { if (p.caption) x += `<w:tblCaption w:val="${X(p.caption)}"/>`; if (p.desc) x += `<w:tblDescription w:val="${X(p.desc)}"/>`; }
    return x;
  }
  function tcPrXML(c, w) {
    c = c || {};
    let x = '';
    if (w != null) x += `<w:tcW w:w="${tw(w)}" w:type="dxa"/>`;
    if (c.span > 1) x += `<w:gridSpan w:val="${c.span}"/>`;
    if (c.vMerge) x += c.vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>';
    x += bordersEl('tcBorders', c.borders, ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl']);
    if (c.shd) x += shdEl(c.shd);
    if (c.noWrap) x += '<w:noWrap/>';
    if (c.mar) { const m = c.mar; x += `<w:tcMar>${m.t != null ? `<w:top w:w="${tw(m.t)}" w:type="dxa"/>` : ''}${m.l != null ? `<w:left w:w="${tw(m.l)}" w:type="dxa"/>` : ''}${m.b != null ? `<w:bottom w:w="${tw(m.b)}" w:type="dxa"/>` : ''}${m.r != null ? `<w:right w:w="${tw(m.r)}" w:type="dxa"/>` : ''}</w:tcMar>`; }
    if (c.textDir) x += `<w:textDirection w:val="${c.textDir}"/>`;
    if (c.fit) x += '<w:tcFitText/>';
    if (c.vAlign && c.vAlign !== 'top') x += `<w:vAlign w:val="${c.vAlign}"/>`;
    if (c.hideMark) x += '<w:hideMark/>';
    return x;
  }
  function trPrXML(t, ctx) {
    t = t || {};
    let x = '';
    if (t.gridBefore) x += `<w:gridBefore w:val="${t.gridBefore}"/>`;
    if (t.gridAfter) x += `<w:gridAfter w:val="${t.gridAfter}"/>`;
    if (t.cantSplit) x += '<w:cantSplit/>';
    if (t.h) x += `<w:trHeight w:val="${tw(t.h)}"${t.hRule && t.hRule !== 'atLeast' ? ` w:hRule="${t.hRule}"` : ''}/>`;
    if (t.header) x += '<w:tblHeader/>';
    if (t.jc) x += `<w:jc w:val="${t.jc}"/>`;
    if (t.hidden) x += '<w:hidden/>';
    if (t.ins && ctx) x += `<w:ins w:id="${ctx.rev()}" w:author="${X(t.ins.author)}" ${DT(t.ins.date)}/>`;
    if (t.del && ctx) x += `<w:del w:id="${ctx.rev()}" w:author="${X(t.del.author)}" ${DT(t.del.date)}/>`;
    return x;
  }

  /* ================= content ================= */
  const sameLink = (a, b) => (a === b) || (!!a && !!b && a.url === b.url && a.anchor === b.anchor && a.tip === b.tip);
  const sameRev = (a, b) => (!a && !b) || (!!a && !!b && a.author === b.author && a.date === b.date);
  function textXML(s, del) {
    /* split at characters that need their own elements */
    let out = '';
    let buf = '';
    const tag = del ? 'w:delText' : 'w:t';
    const flush = () => { if (buf) { out += `<${tag} xml:space="preserve">${X(buf)}</${tag}>`; buf = ''; } };
    for (const ch of s) {
      if (ch === '‑') { flush(); out += '<w:noBreakHyphen/>'; }
      else if (ch === '­') { flush(); out += '<w:softHyphen/>'; }
      else if (ch === '\n') { flush(); out += '<w:br/>'; }
      else if (ch === '\t') { flush(); out += '<w:tab/>'; }
      else if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/.test(ch)) continue;
      else buf += ch;
    }
    flush();
    return out;
  }
  /** XML for a paragraph's runs */
  function wrapRunGroup(inner, r, ctx) {
    if (!inner) return '';
    const { link, ins, del } = r;
    /* inserted, then deleted: Word nests the deletion inside the insertion */
    if (del) inner = `<w:del w:id="${ctx.rev()}" w:author="${X(del.author)}" ${DT(del.date)}>${inner}</w:del>`;
    if (ins) inner = `<w:ins w:id="${ctx.rev()}" w:author="${X(ins.author)}" ${DT(ins.date)}>${inner}</w:ins>`;
    if (link) {
      const a = [];
      if (link.url) a.push(`r:id="${ctx.rels.add(RT('hyperlink'), link.url, true)}"`);
      if (link.anchor) a.push(`w:anchor="${X(link.anchor)}"`);
      if (link.tip) a.push(`w:tooltip="${X(link.tip)}"`);
      a.push('w:history="1"'); inner = `<w:hyperlink ${a.join(' ')}>${inner}</w:hyperlink>`;
    }
    return inner;
  }
  function runsXML(p, ctx) {
    const items = p.runs;
    let out = '';
    /* group by hyperlink & revision wrappers */
    let i = 0;
    while (i < items.length) {
      const it = items[i];
      const original = L.preserve.objectXML(it, 'ac', ctx);
      if (original != null) {
        if (original) out += it.keep.ac.scope === 'run' ? wrapRunGroup(runWrap(it.rPr, original, ctx), it.rPr || {}, ctx) : wrapRunGroup(original, it.keep.ac.wrapper || {}, ctx);
        i++; continue;
      }
      if (it.t === 'sdts') {
        const end = items.findIndex((n, at) => at > i && n.t === 'sdte' && n.key === it.control.key);
        if (end >= 0) { out += L.preserve.controlXML(it.control, runsXML({ ...p, runs: items.slice(i + 1, end) }, ctx), ctx); i = end + 1; continue; }
      }
      if (it.t === 'sdts' || it.t === 'sdte') { i++; continue; }
      const r = it.rPr || {};
      const link = D.ilen(it) ? r.link : null;
      const ins = D.ilen(it) || it.t === 'fb' || it.t === 'fs' || it.t === 'fe' ? r.ins : null;
      const del = D.ilen(it) || it.t === 'fb' || it.t === 'fs' || it.t === 'fe' ? r.del : null;
      let j = i + 1;
      if (link || ins || del) {
        while (j < items.length) {
          const n = items[j], nr = n.rPr || {};
          if (n.keep?.ac) break;
          if (D.isMarker(n) && n.t !== 'fb' && n.t !== 'fs' && n.t !== 'fe') break;
          if (!sameLink(link, D.ilen(n) ? nr.link : null) || !sameRev(ins, nr.ins) || !sameRev(del, nr.del)) break;
          j++;
        }
      }
      let inner = '';
      for (let k = i; k < j; k++) inner += itemXML(items[k], p, ctx, !!del);
      /* revision marks sit inside hyperlinks (w:hyperlink may contain w:ins/w:del, not the reverse) */
      out += wrapRunGroup(inner, { link, ins, del }, ctx);
      i = j;
    }
    return out;
  }
  function cleanRPr(r) {
    if (!r) return {};
    const o = Object.assign({}, r);
    delete o.link; delete o.ins; delete o.del;
    return o;
  }
  function runWrap(r, content, ctx) {
    const rp = rPrXML(cleanRPr(r), null, ctx);
    return `<w:r>${rp}${content}</w:r>`;
  }
  function itemXML(it, p, ctx, del) {
    const r = it.rPr || {};
    const original = L.preserve.objectXML(it, 'opaque', ctx);
    if (original != null) return original ? runWrap(r, original, ctx) : '';
    switch (it.t) {
      case 'text': return runWrap(r, textXML(it.text, del), ctx);
      case 'tab': return runWrap(r, '<w:tab/>', ctx);
      case 'ptab': return runWrap(r, `<w:ptab w:relativeTo="${it.rel || 'margin'}" w:alignment="${it.al || 'left'}" w:leader="${it.leader || 'none'}"/>`, ctx);
      case 'br':
        if (it.recoveredPlacement) ctx.writer.loss({ id: 'break-placement:' + ctx.part, where: ctx.part, action: 'conversion',
          what: 'Page or line breaks outside their required text runs were moved into standard paragraphs and runs.' });
        return runWrap(r, it.type === 'page' ? '<w:br w:type="page"/>' : it.type === 'column' ? '<w:br w:type="column"/>' : `<w:br${it.clear ? ` w:clear="${it.clear}"` : ''}/>`, ctx);
      case 'ruby': {
        const pr = it.pr || {};
        const prx = ['rubyAlign', 'hps', 'hpsRaise', 'hpsBaseText', 'lid'].filter((k) => pr[k] != null).map((k) => `<w:${k} w:val="${X(String(pr[k]))}"/>`).join('');
        const sub = (rp, t) => `<w:r>${rPrXML(rp || {}, null, ctx)}${textXML(t || '')}</w:r>`;
        return runWrap(r, `<w:ruby><w:rubyPr>${prx || '<w:rubyAlign w:val="distributeSpace"/><w:hps w:val="10"/><w:hpsRaise w:val="20"/><w:hpsBaseText w:val="21"/><w:lid w:val="ja-JP"/>'}</w:rubyPr><w:rt>${sub(it.rtRPr, it.rt)}</w:rt><w:rubyBase>${sub(it.baseRPr, it.base)}</w:rubyBase></w:ruby>`, ctx);
      }
      case 'sym': return runWrap(r, `<w:sym w:font="${X(it.font || 'Symbol')}" w:char="${it.code || it.char.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}"/>`, ctx);
      case 'fb': {
        let ff = '';
        if (it.ff) {
          const f = it.ff;
          let inner = `<w:name w:val="${X(f.name || '')}"/><w:enabled/><w:calcOnExit w:val="0"/>`;
          if (f.help) inner += `<w:helpText w:type="text" w:val="${X(f.help)}"/>`;
          if (f.status) inner += `<w:statusText w:type="text" w:val="${X(f.status)}"/>`;
          if (f.type === 'checkbox') inner += `<w:checkBox>${f.size ? `<w:size w:val="${hp(f.size)}"/>` : '<w:sizeAuto/>'}<w:default w:val="${f.def ? 1 : 0}"/>${f.checked !== f.def ? `<w:checked w:val="${f.checked ? 1 : 0}"/>` : ''}</w:checkBox>`;
          else if (f.type === 'text') inner += `<w:textInput>${f.textType && f.textType !== 'regular' ? `<w:type w:val="${f.textType}"/>` : ''}${f.def ? `<w:default w:val="${X(f.def)}"/>` : ''}${f.maxLength ? `<w:maxLength w:val="${f.maxLength}"/>` : ''}${f.format ? `<w:format w:val="${X(f.format)}"/>` : ''}</w:textInput>`;
          else if (f.type === 'dropdown') inner += `<w:ddList><w:result w:val="${f.result || 0}"/>${(f.entries || []).map((e) => `<w:listEntry w:val="${X(e)}"/>`).join('')}</w:ddList>`;
          ff = `<w:ffData>${inner}</w:ffData>`;
        }
        const begin = runWrap(r, `<w:fldChar w:fldCharType="begin"${it.lock ? ' w:fldLock="1"' : ''}${it.dirty ? ' w:dirty="1"' : ''}>${ff}${ff ? '</w:fldChar>' : '</w:fldChar>'}`.replace('></w:fldChar>', ff ? '></w:fldChar>' : '/>'), ctx);
        return begin + runWrap(r, del ? `<w:delInstrText xml:space="preserve">${X(it.instr || '')}</w:delInstrText>` : `<w:instrText xml:space="preserve">${X(it.instr || '')}</w:instrText>`, ctx);
      }
      case 'fs': return runWrap(r, '<w:fldChar w:fldCharType="separate"/>', ctx);
      case 'fe': return runWrap(r, '<w:fldChar w:fldCharType="end"/>', ctx);
      case 'bs': { const k = 's' + it.id; if (ctx.bmSeen.has(k)) return ''; ctx.bmSeen.add(k); return L.preserve.markerXML(it, ctx) ?? `<w:bookmarkStart w:id="${ctx.bmId(it.id)}" w:name="${X(it.name || '_bm' + it.id)}"/>`; }
      case 'be': { const k = 'e' + it.id; if (ctx.bmSeen.has(k)) return ''; ctx.bmSeen.add(k); return L.preserve.markerXML(it, ctx) ?? `<w:bookmarkEnd w:id="${ctx.bmId(it.id)}"/>`; }
      case 'perm': return L.preserve.markerXML(it, ctx) || '';
      case 'cs': return ctx.cmtOk(it.id) && !it.point ? `<w:commentRangeStart w:id="${ctx.cmtId(it.id)}"/>` : '';
      case 'ce': return ctx.cmtOk(it.id) ? (it.point ? '' : `<w:commentRangeEnd w:id="${ctx.cmtId(it.id)}"/>`) +
        (ctx.cmtReference(it.id) ? runWrap({ style: 'CommentReference' }, `<w:commentReference w:id="${ctx.cmtId(it.id)}"/>`, ctx) : '') : '';
      case 'fn': case 'en': {
        const kind = it.t === 'fn' ? 'footnote' : 'endnote';
        if (it.self) return runWrap(Object.assign({ style: it.t === 'fn' ? 'FootnoteReference' : 'EndnoteReference' }, cleanRPr(r)), `<w:${kind}Ref/>`, ctx);
        const nid = ctx.noteId(it.t, it.id);
        if (nid == null) return '';
        const rr = Object.assign({}, r);
        if (!rr.style) rr.style = it.t === 'fn' ? 'FootnoteReference' : 'EndnoteReference';
        return runWrap(rr, `<w:${kind}Reference w:id="${nid}"${it.custom ? ' w:customMarkFollows="1"' : ''}/>`, ctx) + (it.custom ? runWrap(rr, textXML(it.custom), ctx) : '');
      }
      case 'sep': return runWrap(r, '<w:separator/>', ctx);
      case 'csep': return runWrap(r, '<w:continuationSeparator/>', ctx);
      case 'img': case 'shape': case 'group': case 'chart': return runWrap(r, drawingXML(it, ctx), ctx);
      case 'raw': return it.xml && it.math ? (it.para ? it.xml.replace(/^<m:oMathPara/, '<m:oMathPara') : it.xml) : it.text ? runWrap(r, textXML(it.text), ctx) : '';
      default: return '';
    }
  }

  /* ================= drawings ================= */
  function clr(c, alpha) {
    const h = String(c || '#000000').replace('#', '').toUpperCase();
    const a = alpha != null && alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : '';
    return a ? `<a:srgbClr val="${/^[0-9A-F]{6}$/.test(h) ? h : '000000'}">${a}</a:srgbClr>` : `<a:srgbClr val="${/^[0-9A-F]{6}$/.test(h) ? h : '000000'}"/>`;
  }
  function fillXML(f, ctx) {
    if (!f) return '';
    switch (f.t) {
      case 'none': return '<a:noFill/>';
      case 'solid': return `<a:solidFill>${clr(f.c, f.a)}</a:solidFill>`;
      case 'grad': {
        let stops = (f.stops || []).slice().sort((a, b) => a.p - b.p);
        if (stops.length < 2) stops = [stops[0] || { p: 0, c: '#FFFFFF' }, Object.assign({}, stops[0] || { c: '#000000' }, { p: 1 })];
        const gs = stops.map((s) => `<a:gs pos="${Math.round(L.clamp(s.p, 0, 1) * 100000)}">${clr(s.c, s.a)}</a:gs>`).join('');
        const shade = f.path && f.path !== 'lin' ? `<a:path path="${f.path === 'rect' ? 'rect' : f.path === 'shape' ? 'shape' : 'circle'}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>` : `<a:lin ang="${Math.round((((f.ang || 0) % 360) + 360) % 360 * 60000)}" scaled="1"/>`;
        return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst>${shade}</a:gradFill>`;
      }
      case 'patt': return `<a:pattFill prst="${f.prst || 'pct50'}"><a:fgClr>${clr(f.fg || '#000000')}</a:fgClr><a:bgClr>${clr(f.bg || '#FFFFFF')}</a:bgClr></a:pattFill>`;
      case 'img': { const rid = ctx.media(f.media); if (!rid) return '<a:noFill/>'; return `<a:blipFill rotWithShape="1"><a:blip r:embed="${rid}"/>${f.tile ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>' : '<a:stretch><a:fillRect/></a:stretch>'}</a:blipFill>`; }
      default: return '';
    }
  }
  function lineXML(ln) {
    if (!ln) return '';
    if (ln.t === 'none' || !ln.c) return '<a:ln><a:noFill/></a:ln>';
    let x = `<a:ln w="${emu(ln.w == null ? 0.75 : ln.w)}"${ln.cmpd && ln.cmpd !== 'sng' ? ` cmpd="${ln.cmpd}"` : ''}><a:solidFill>${clr(ln.c, ln.a)}</a:solidFill>`;
    if (ln.dash && ln.dash !== 'solid') x += `<a:prstDash val="${ln.dash}"/>`;
    const end = (e, n) => (e && e.type && e.type !== 'none' ? `<a:${n} type="${e.type}" w="${e.w || 'med'}" len="${e.len || 'med'}"/>` : '');
    return x + end(ln.head, 'headEnd') + end(ln.tail, 'tailEnd') + '</a:ln>';
  }
  function geomXML(sh) {
    if (sh.geom === 'custom' && sh.path) {
      const paths = sh.path.paths.map((p) => {
        const pt = (x, y) => `<a:pt x="${Math.round(x * L.EMU)}" y="${Math.round(y * L.EMU)}"/>`;
        const cmds = p.cmds.map((c) => {
          switch (c[0]) {
            case 'M': return `<a:moveTo>${pt(c[1], c[2])}</a:moveTo>`;
            case 'L': return `<a:lnTo>${pt(c[1], c[2])}</a:lnTo>`;
            case 'C': return `<a:cubicBezTo>${pt(c[1], c[2])}${pt(c[3], c[4])}${pt(c[5], c[6])}</a:cubicBezTo>`;
            case 'Q': return `<a:quadBezTo>${pt(c[1], c[2])}${pt(c[3], c[4])}</a:quadBezTo>`;
            case 'A': return `<a:arcTo wR="${Math.round(c[1] * L.EMU)}" hR="${Math.round(c[2] * L.EMU)}" stAng="${Math.round(c[3] * 60000)}" swAng="${Math.round(c[4] * 60000)}"/>`;
            case 'Z': return '<a:close/>';
            default: return '';
          }
        }).join('');
        return `<a:path w="${Math.round(p.w * L.EMU)}" h="${Math.round(p.h * L.EMU)}"${p.fill === 'none' ? ' fill="none"' : ''}${p.stroke === false ? ' stroke="0"' : ''}>${cmds}</a:path>`;
      }).join('');
      return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`;
    }
    const g = sh.geom === 'line' ? 'line' : sh.geom || 'rect';
    const av = L.opc.presetAdjust(g, sh.adj);
    return `<a:prstGeom prst="${g}"><a:avLst>${av}</a:avLst></a:prstGeom>`;
  }
  function xfrmXML(sh, x, y, tag, extra) {
    const a = [];
    if (sh.rot) a.push(`rot="${Math.round((((sh.rot % 360) + 360) % 360) * 60000)}"`);
    if (sh.flipH) a.push('flipH="1"');
    if (sh.flipV) a.push('flipV="1"');
    return `<${tag || 'a:xfrm'}${a.length ? ' ' + a.join(' ') : ''}><a:off x="${emu(x || 0)}" y="${emu(y || 0)}"/><a:ext cx="${Math.max(0, emu(sh.w))}" cy="${Math.max(0, emu(sh.h))}"/>${extra || ''}</${tag || 'a:xfrm'}>`;
  }
  function picXML(it, ctx, x, y, id) {
    const rid = ctx.media(it.media);
    const c = it.crop;
    const src = c && (c.l || c.t || c.r || c.b) ? `<a:srcRect l="${Math.round((c.l || 0) * 100000)}" t="${Math.round((c.t || 0) * 100000)}" r="${Math.round((c.r || 0) * 100000)}" b="${Math.round((c.b || 0) * 100000)}"/>` : '<a:srcRect/>';
    let eff = '';
    if (it.gray) eff += '<a:grayscl/>';
    if (it.bw) eff += '<a:biLevel thresh="50000"/>';
    const bright = it.washout ? 0.7 : it.bright, contrast = it.washout ? -0.75 : it.contrast;
    if (bright || contrast) eff += `<a:lum${bright ? ` bright="${Math.round(bright * 100000)}"` : ''}${contrast ? ` contrast="${Math.round(contrast * 100000)}"` : ''}/>`;
    if (it.alpha != null && it.alpha < 1) eff += `<a:alphaModFix amt="${Math.round(it.alpha * 100000)}"/>`;
    const blip = rid ? `<a:blip r:embed="${rid}"${eff ? `>${eff}</a:blip>` : '/>'}` : '<a:blip/>';
    const ln = it.border && it.border.val && it.border.val !== 'nil' ? `<a:ln w="${emu(it.border.sz || 0.75)}"><a:solidFill>${clr('#' + hex(it.border.color === 'auto' ? '000000' : it.border.color))}</a:solidFill></a:ln>` : '';
    const geom = it.geom && it.geom.geom ? geomXML(it.geom) : '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    return `<pic:pic xmlns:pic="${NS.pic}"><pic:nvPicPr><pic:cNvPr id="${id}" name="${X(it.name || 'Picture ' + id)}"${it.alt ? ` descr="${X(it.alt)}"` : ''}/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill>${blip}${src}<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto">${xfrmXML(it, x, y)}${geom}${ln}</pic:spPr></pic:pic>`;
  }
  function shapeXML(it, ctx, x, y, id) {
    const isLine = it.geom === 'line' || L.geom.isLine(it.geom);
    const ins = it.ins || { l: 7.2, t: 3.6, r: 7.2, b: 3.6 };
    let txbx = '';
    if (it.tb && it.tb.blocks) {
      const was = ctx.inTxbx;
      ctx.inTxbx = true;
      try { txbx = `<wps:txbx${it.tbId != null ? ` id="${X(String(it.tbId))}"` : ''}><w:txbxContent>${blocksXML(it.tb.blocks, ctx)}</w:txbxContent></wps:txbx>`; } finally { ctx.inTxbx = was; }
    }
    else if (it.tbLink) txbx = `<wps:linkedTxbx id="${X(String(it.tbLink.id))}" seq="${it.tbLink.seq | 0}"/>`;
    let shadow = '';
    if (it.shadow) { const sh = it.shadow; const dx = sh.dx == null ? 3 : sh.dx, dy = sh.dy == null ? 3 : sh.dy; shadow = `<a:effectLst><a:outerShdw blurRad="${emu(sh.blur || 0)}" dist="${emu(Math.hypot(dx, dy))}" dir="${Math.round(((((Math.atan2(dy, dx) * 180) / Math.PI) + 360) % 360) * 60000)}" algn="tl" rotWithShape="0">${clr(sh.c || '#000000', sh.a == null ? 0.5 : sh.a)}</a:outerShdw></a:effectLst>`; }
    const warp = it.wordart || it.warp ? `<a:prstTxWarp prst="${it.warp || (it.wordart && it.wordart.warp) || 'textPlain'}"><a:avLst/></a:prstTxWarp>` : '';
    let wa = '';
    const isWA = !!(it.wordart && !it.tb);
    if (isWA) {
      /* WordArt (Word 2010+ form): the letters carry the fill and outline; the shape itself is transparent */
      const w = it.wordart;
      const f = it.fill || { t: 'solid', c: '#3366CC' };
      const w14fill = (fl) => {
        if (!fl || fl.t === 'none') return '<w14:noFill/>';
        if (fl.t === 'grad' && fl.stops && fl.stops.length) return `<w14:gradFill><w14:gsLst>${fl.stops.map((g) => `<w14:gs w14:pos="${Math.round((g.p || 0) * 100000)}"><w14:srgbClr w14:val="${hex(g.c)}"/></w14:gs>`).join('')}</w14:gsLst><w14:lin w14:ang="${Math.round(((fl.ang || 0) % 360) * 60000)}" w14:scaled="0"/></w14:gradFill>`;
        const c = fl.c || fl.fg || '#3366CC';
        return `<w14:solidFill><w14:srgbClr w14:val="${hex(c)}">${fl.a != null && fl.a < 1 ? `<w14:alpha w14:val="${Math.round((1 - fl.a) * 100000)}"/>` : ''}</w14:srgbClr></w14:solidFill>`;
      };
      const ln = it.line && it.line.t !== 'none' && it.line.c ? `<w14:textOutline w14:w="${emu(it.line.w || 0.75)}" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr"><w14:solidFill><w14:srgbClr w14:val="${hex(it.line.c)}"/></w14:solidFill><w14:prstDash w14:val="solid"/><w14:round/></w14:textOutline>` : '';
      const baseColor = f.t === 'solid' ? hex(f.c) : f.t === 'grad' && f.stops ? hex(f.stops[0].c) : '3366CC';
      const rp = rPrXML({ font: w.font || 'Arial Black', sz: w.sz || 36, b: w.b || undefined, i: w.i || undefined, color: baseColor }, null, ctx);
      const lines = String(w.text || '').split('\n');
      const effects = rp.replace('</w:rPr>', ln + '<w14:textFill>' + w14fill(f) + '</w14:textFill></w:rPr>');   // the fill only counts inside textFill
      wa = `<wps:txbx><w:txbxContent>${lines.map((t) => `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${effects}${textXML(t)}</w:r></w:p>`).join('')}</w:txbxContent></wps:txbx>`;
    }
    const shapeFill = isWA ? '<a:noFill/>' : isLine ? '<a:noFill/>' : fillXML(it.fill || { t: 'none' }, ctx);
    const shapeLine = isWA ? '<a:ln><a:noFill/></a:ln>' : lineXML(it.line || { t: 'none' });
    return `<wps:wsp><wps:cNvPr id="${id}" name="${X(it.name || (it.tb ? 'Text Box ' : isWA ? 'LettersArt ' : 'Shape ') + id)}"/><wps:cNvSpPr${it.tb || it.tbLink || isWA ? ' txBox="1"' : ''}/><wps:spPr>${xfrmXML(it, x, y)}${geomXML(it)}${shapeFill}${shapeLine}${isWA ? '' : shadow}</wps:spPr>${txbx || wa}<wps:bodyPr rot="0" vert="${it.vert || 'horz'}" wrap="${it.noWrap || isWA ? 'none' : 'square'}" lIns="${emu(isWA ? 0 : ins.l)}" tIns="${emu(isWA ? 0 : ins.t)}" rIns="${emu(isWA ? 0 : ins.r)}" bIns="${emu(isWA ? 0 : ins.b)}" anchor="${it.anchor || (isWA ? 'ctr' : 't')}" anchorCtr="0"${isWA ? ' fromWordArt="1"' : ''}>${warp}${isWA ? '<a:normAutofit/>' : it.autofit ? '<a:spAutoFit/>' : '<a:noAutofit/>'}</wps:bodyPr></wps:wsp>`;
  }
  function groupXML(g, ctx, x, y, id, nested) {
    let kids = '';
    for (const k of g.kids || []) {
      const kid = ctx.docPrId();
      if (k.t === 'img') kids += picXML(k, ctx, k.x, k.y, kid);
      else if (k.t === 'group') kids += groupXML(k, ctx, k.x, k.y, kid, true);
      else kids += shapeXML(k, ctx, k.x, k.y, kid);
    }
    const tag = nested ? 'wpg:grpSp' : 'wpg:wgp';
    const head = nested ? `<wpg:cNvPr id="${id}" name="Group ${id}"/><wpg:cNvGrpSpPr/>` : '<wpg:cNvGrpSpPr/>';
    return `<${tag}>${head}<wpg:grpSpPr>${xfrmXML(g, x, y, 'a:xfrm', `<a:chOff x="${emu(x || 0)}" y="${emu(y || 0)}"/><a:chExt cx="${emu(g.w)}" cy="${emu(g.h)}"/>`)}</wpg:grpSpPr>${kids.replace(/<pic:pic /g, '<pic:pic ')}</${tag}>`.replace(/<wps:wsp>/g, '<wps:wsp>');
  }
  /* Shapes inside a text box: Word keeps them as a locked canvas (a text box cannot hold wps/wpg shapes) */
  function dmlParasXML(blocks) {
    const out = [];
    for (const b of blocks || []) {
      if (b.t !== 'p') continue;
      const algn = { center: 'ctr', right: 'r', end: 'r', both: 'just' }[(b.pPr && b.pPr.jc) || ''] || 'l';
      let runs = '';
      for (const it of b.runs || []) {
        if (it.t === 'text' && it.text) {
          const r = it.rPr || {};
          const col = r.color && r.color !== 'auto' ? `<a:solidFill><a:srgbClr val="${hex(r.color)}"/></a:solidFill>` : '';
          runs += `<a:r><a:rPr lang="en-US"${r.sz ? ` sz="${Math.round(r.sz * 100)}"` : ''}${r.b ? ' b="1"' : ''}${r.i ? ' i="1"' : ''}${r.u && r.u !== 'none' ? ' u="sng"' : ''}>${col}${r.font ? `<a:latin typeface="${X(r.font)}"/>` : ''}</a:rPr><a:t>${X(it.text)}</a:t></a:r>`;
        } else if (it.t === 'tab') runs += '<a:r><a:rPr lang="en-US"/><a:t>\t</a:t></a:r>';
        else if (it.t === 'br') runs += '<a:br><a:rPr lang="en-US"/></a:br>';
      }
      out.push(`<a:p><a:pPr algn="${algn}"/>${runs}<a:endParaRPr lang="en-US"/></a:p>`);
    }
    return out.join('') || '<a:p><a:endParaRPr lang="en-US"/></a:p>';
  }
  function canvasKidXML(k, ctx, id) {
    if (k.t === 'img') return picXML(k, ctx, k.x, k.y, id).replace(/<(\/?)pic:/g, '<$1a:').replace(/ xmlns:pic="[^"]*"/, '');
    if (k.t === 'group') {
      let kids = '';
      for (const c of k.kids || []) kids += canvasKidXML(c, ctx, ctx.docPrId());
      return `<a:grpSp><a:nvGrpSpPr><a:cNvPr id="${id}" name="Group ${id}"/><a:cNvGrpSpPr/></a:nvGrpSpPr><a:grpSpPr>${xfrmXML(k, k.x, k.y, 'a:xfrm', `<a:chOff x="${emu(k.x || 0)}" y="${emu(k.y || 0)}"/><a:chExt cx="${emu(k.w)}" cy="${emu(k.h)}"/>`)}</a:grpSpPr>${kids}</a:grpSp>`;
    }
    const ins = k.ins || { l: 7.2, t: 3.6, r: 7.2, b: 3.6 };
    const isLine = k.geom === 'line' || L.geom.isLine(k.geom);
    const text = k.tb && k.tb.blocks ? `<a:txSp><a:txBody><a:bodyPr wrap="square" lIns="${emu(ins.l)}" tIns="${emu(ins.t)}" rIns="${emu(ins.r)}" bIns="${emu(ins.b)}" anchor="${k.anchor || 't'}"/><a:lstStyle/>${dmlParasXML(k.tb.blocks)}</a:txBody><a:useSpRect/></a:txSp>` : '';
    return `<a:sp><a:nvSpPr><a:cNvPr id="${id}" name="${X(k.name || 'Shape ' + id)}"/><a:cNvSpPr${k.tb ? ' txBox="1"' : ''}/></a:nvSpPr><a:spPr>${xfrmXML(k, k.x, k.y)}${geomXML(k)}${isLine ? '<a:noFill/>' : fillXML(k.fill || { t: 'none' }, ctx)}${lineXML(k.line || { t: 'none' })}</a:spPr>${text}</a:sp>`;
  }
  function lockedCanvasXML(it, ctx) {
    const kids = it.t === 'group' ? (it.kids || []).map((k) => canvasKidXML(k, ctx, ctx.docPrId())).join('') : canvasKidXML(Object.assign({}, it, { x: 0, y: 0 }), ctx, ctx.docPrId());
    const ox = it.t === 'group' ? 0 : 0;
    return `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"><lc:lockedCanvas xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"><a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr><a:grpSpPr>${xfrmXML(it, ox, ox, 'a:xfrm', `<a:chOff x="0" y="0"/><a:chExt cx="${emu(it.w)}" cy="${emu(it.h)}"/>`)}</a:grpSpPr>${kids}</lc:lockedCanvas></a:graphicData>`;
  }
  function drawingXML(it, ctx) {
    const id = ctx.docPrId();
    let graphic;
    if (ctx.inTxbx && (it.t === 'group' || it.t === 'shape')) graphic = lockedCanvasXML(it, ctx);
    else if (it.t === 'img') graphic = `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">${picXML(it, ctx, 0, 0, id)}</a:graphicData>`;
    else if (it.t === 'shape') graphic = `<a:graphicData uri="${NS.wps}">${shapeXML(it, ctx, 0, 0, id)}</a:graphicData>`;
    else if (it.t === 'group') graphic = `<a:graphicData uri="${NS.wpg}">${groupXML(it, ctx, 0, 0, id, false)}</a:graphicData>`;
    else if (it.t === 'chart') { const rid = ctx.chart(it); graphic = `<a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" r:id="${rid}"/></a:graphicData>`; }
    const ext = `<wp:extent cx="${Math.max(0, emu(it.w))}" cy="${Math.max(0, emu(it.h))}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`;
    const docPr = `<wp:docPr id="${id}" name="${X(it.name || (it.t === 'img' ? 'Picture ' : it.t === 'chart' ? 'Chart ' : 'Shape ') + id)}"${it.alt ? ` descr="${X(it.alt)}"` : ''}${it.title ? ` title="${X(it.title)}"` : ''}${it.hidden ? ' hidden="1"' : ''}>${it.link && it.link.url ? `<a:hlinkClick xmlns:a="${NS.a}" r:id="${ctx.rels.add(RT('hyperlink'), it.link.url, true)}"/>` : ''}</wp:docPr>`;
    const lock = it.t === 'img' ? `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` : '<wp:cNvGraphicFramePr/>';
    const graphicEl = `<a:graphic xmlns:a="${NS.a}">${graphic}</a:graphic>`;
    if (!it.float) return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">${ext}${docPr}${lock}${graphicEl}</wp:inline></w:drawing>`;
    const f = it.float;
    const d = f.dist || { t: 0, b: 0, l: 9, r: 9 };
    const RELH = /^(character|column|insideMargin|leftMargin|margin|outsideMargin|page|rightMargin)$/, RELV = /^(bottomMargin|insideMargin|line|margin|outsideMargin|page|paragraph|topMargin)$/;
    const relOf = (tag, r) => (r && (tag === 'H' ? RELH : RELV).test(r) ? r : tag === 'H' ? 'column' : 'paragraph');
    const pos = (tag, p) => `<wp:position${tag} relativeFrom="${relOf(tag, p.rel)}">${p.align ? `<wp:align>${p.align}</wp:align>` : `<wp:posOffset>${emu(p.off || 0)}</wp:posOffset>`}</wp:position${tag}>`;
    const side = f.side || 'bothSides';
    const poly = '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';
    const wrap = f.wrap === 'square' ? `<wp:wrapSquare wrapText="${side}"/>` : f.wrap === 'tight' ? `<wp:wrapTight wrapText="${side}">${poly}</wp:wrapTight>` : f.wrap === 'through' ? `<wp:wrapThrough wrapText="${side}">${poly}</wp:wrapThrough>` : f.wrap === 'topBottom' ? '<wp:wrapTopAndBottom/>' : '<wp:wrapNone/>';
    const behind = f.wrap === 'behind' || (f.behind && f.wrap !== 'none');
    return `<w:drawing><wp:anchor distT="${emu(d.t || 0)}" distB="${emu(d.b || 0)}" distL="${emu(d.l == null ? 9 : d.l)}" distR="${emu(d.r == null ? 9 : d.r)}" simplePos="0" relativeHeight="${Math.max(0, Math.min(4294967295, Math.round(f.z ?? 251659264)))}" behindDoc="${behind ? 1 : 0}" locked="${f.locked ? 1 : 0}" layoutInCell="${f.layoutInCell === false ? 0 : 1}" allowOverlap="${f.allowOverlap === false ? 0 : 1}"><wp:simplePos x="0" y="0"/>${pos('H', f.posH)}${pos('V', f.posV)}${ext}${wrap}${docPr}${lock}${graphicEl}</wp:anchor></w:drawing>`;
  }

  /* ================= blocks ================= */
  function paraXML(p, ctx) {
    const markR = rPrXML(p.rPr || {}, p.mark, ctx);
    const pPr = pPrXML(p.pPr, ctx, { rPr: markR, sect: p.sect && ctx.main ? sectXML(p.sect, ctx) : '' });
    /* display math paragraphs */
    const runs = runsXML(p, ctx);
    return `<w:p>${pPr}${runs}</w:p>`;
  }
  function tableXML(t, ctx) {
    if (t.keep?.gridRepair) ctx.writer.loss({ id: 'table-grid:' + ctx.part, what: 'Incomplete or unusable table column widths were reconstructed. Table layout may change.', where: ctx.part, action: 'conversion' });
    const map = L.R.tblMap(t);
    const pr = L.preserve.propertyXML('tblPr', t.tblPr, tblPrXML(Object.assign({}, t.tblPr, { w: t.tblPr.w || { type: 'auto', v: 0 } }), { full: true }), ctx);
    let x = `<w:tbl>${pr}<w:tblGrid>${t.grid.map((g) => `<w:gridCol w:w="${tw(g)}"/>`).join('')}</w:tblGrid>`;
    x += L.preserve.controlSequence(t.rows, row => row.trPr, row => L.preserve.tableMemberXML(row, ctx, () => {
      const ri = t.rows.indexOf(row);
      const trp = trPrXML(row.trPr, ctx);
      let rowXML = `<w:tr>${trp ? `<w:trPr>${trp}</w:trPr>` : ''}`;
      rowXML += L.preserve.controlSequence(row.cells, cell => cell.tcPr, cell => L.preserve.tableMemberXML(cell, ctx, () => {
        const ci = row.cells.indexOf(cell);
        const m = map[ri][ci];
        const w = cell.tcPr.w != null ? cell.tcPr.w : t.grid.slice(m.c0, m.c0 + m.span).reduce((a, b) => a + b, 0);
        const blocks = cell.blocks.length ? cell.blocks : [D.para()];
        let inner = blocksXML(blocks, ctx);
        if (!blocks.length || blocks[blocks.length - 1].t !== 'p') inner += '<w:p/>';
        return `<w:tc><w:tcPr>${tcPrXML(cell.tcPr, w)}</w:tcPr>${inner}</w:tc>`;
      }), ctx);
      return rowXML + '</w:tr>';
    }), ctx);
    return x + '</w:tbl>';
  }
  function blocksXML(blocks, ctx) {
    return L.preserve.controlSequence(blocks, b => b.t === 'p' ? b.pPr : b.tblPr, b => L.preserve.objectXML(b, 'ac', ctx) ?? L.preserve.objectXML(b, 'opaque', ctx) ?? (b.t === 'p' ? paraXML(b, ctx) : b.t === 'tbl' ? tableXML(b, ctx) : ''), ctx);
  }

  /* ================= styles, numbering, settings ================= */
  function stylesXML(doc, ctx) {
    const dd = doc.defaults || {};
    let x = HEAD + `<w:styles xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:mc="${NS.mc}" xmlns:w14="${NS.w14}" mc:Ignorable="w14">`;
    x += `<w:docDefaults><w:rPrDefault>${rPrXML(Object.assign({ font: 'Times New Roman', sz: 12, lang: 'en-US' }, dd.rPr || {}), null, ctx)}</w:rPrDefault><w:pPrDefault>${pPrXML(dd.pPr, ctx)}</w:pPrDefault></w:docDefaults>`;
    const order = Object.values(doc.styles).sort((a, b) => (a.id === 'Normal' ? -1 : b.id === 'Normal' ? 1 : 0));
    for (const s of order) {
      if (!s || !s.id) continue;
      if (s.keep?.defaultStyleRepair) ctx.writer.loss({ id: 'default-styles:' + ctx.part, what: 'Conflicting default styles were normalized. Text layout may change.', where: ctx.part, action: 'conversion' });
      const type = s.type || 'paragraph';
      let st = `<w:style w:type="${type}"${s.isDefault ? ' w:default="1"' : ''}${s.custom ? ' w:customStyle="1"' : ''} w:styleId="${X(s.id)}"><w:name w:val="${X(s.name || s.id)}"/>`;
      if (s.basedOn && doc.styles[s.basedOn]) st += `<w:basedOn w:val="${X(s.basedOn)}"/>`;
      if (s.next && doc.styles[s.next]) st += `<w:next w:val="${X(s.next)}"/>`;
      if (s.link && doc.styles[s.link]) st += `<w:link w:val="${X(s.link)}"/>`;
      if (s.autoRedefine) st += '<w:autoRedefine/>';
      if (s.hidden) st += '<w:hidden/>';
      if (s.prio != null) st += `<w:uiPriority w:val="${s.prio}"/>`;
      if (s.semiHidden) st += '<w:semiHidden/>';
      if (s.q) st += '<w:qFormat/>';
      if (type === 'paragraph' || type === 'table' || type === 'numbering') {
        const pp = Object.assign({}, s.pPr || {});
        delete pp.style;
        st += pPrXML(pp, ctx);
      }
      if (type !== 'numbering') st += rPrXML(s.rPr || {}, null, ctx);
      if (type === 'table') {
        const tp = tblPrXML(s.tblPr || {}, null);
        if (tp) st += `<w:tblPr>${tp}</w:tblPr>`;
        if (s.tcPr) { const tc = tcPrXML(s.tcPr, null); if (tc) st += `<w:tcPr>${tc}</w:tcPr>`; }
        const CONDS = ['wholeTable', 'band1Vert', 'band2Vert', 'band1Horz', 'band2Horz', 'firstRow', 'lastRow', 'firstCol', 'lastCol', 'neCell', 'nwCell', 'seCell', 'swCell'];
        for (const ty of CONDS) {
          const c = s.cond && s.cond[ty];
          if (!c) continue;
          let cx = '';
          if (c.pPr) cx += pPrXML(c.pPr, ctx);
          if (c.rPr) cx += rPrXML(c.rPr, null, ctx);
          cx += `<w:tblPr>${c.tblPr ? tblPrXML(c.tblPr, null) : ''}</w:tblPr>`;
          if (c.tcPr) { const v = tcPrXML(c.tcPr, null); if (v) cx += `<w:tcPr>${v}</w:tcPr>`; }
          st += `<w:tblStylePr w:type="${ty}">${cx}</w:tblStylePr>`;
        }
      }
      x += st + '</w:style>';
    }
    return x + '</w:styles>';
  }
  function lvlXML(lv, i, ctx) {
    let x = `<w:start w:val="${lv.start != null ? lv.start : 1}"/>`;
    x += lv.custFmt ? `<w:numFmt w:val="custom" w:format="${X(lv.custFmt)}"/>` : `<w:numFmt w:val="${lv.fmt || 'decimal'}"/>`;
    if (lv.restart != null) x += `<w:lvlRestart w:val="${lv.restart + 1}"/>`;
    if (lv.pStyle) x += `<w:pStyle w:val="${X(lv.pStyle)}"/>`;
    if (lv.isLgl) x += '<w:isLgl/>';
    if (lv.suff && lv.suff !== 'tab') x += `<w:suff w:val="${lv.suff}"/>`;
    x += `<w:lvlText w:val="${X(lv.text == null ? '' : lv.text)}"/>`;
    const pictureId = ctx.pictureBullet(lv);
    if (pictureId != null) x += `<w:lvlPicBulletId w:val="${X(pictureId)}"/>`;
    if (lv.legacy) x += '<w:legacy w:legacy="1"/>';
    x += `<w:lvlJc w:val="${lv.jc || 'left'}"/>`;
    const pp = { ...lv.pPr };
    if (!lv.pPrValues || lv.tabPos !== lv.pPrValues.tabPos) {
      if (lv.tabPos != null) pp.tabs = [{ pos: lv.tabPos, al: 'num' }]; else delete pp.tabs;
    }
    if (!lv.pPrValues || JSON.stringify(lv.ind) !== JSON.stringify(lv.pPrValues.ind)) {
      if (lv.ind) pp.ind = { l: lv.ind.l || 0, fl: lv.ind.fl || 0 }; else delete pp.ind;
    }
    x += pPrXML(pp, ctx);
    x += rPrXML(lv.rPr || {}, null, ctx);
    return L.preserve.propertyXML('lvl', lv, x, ctx, { ilvl: i });
  }
  function numberingXML(doc, ctx) {
    const nb = doc.numbering;
    let x = '';
    const pictures = new Map(); let pictureXML = '';
    const picture = record => {
      const f = record.fragment, key = JSON.stringify([f.source, f.part, record.id, f.copy || '']);
      if (!pictures.has(key)) {
        try {
          pictureXML += ctx.writer.emit(f, ctx.part);
          pictures.set(key, ctx.writer.ids.resolve(f.source, 'numbering', 'numPicBulletId', record.id, { primary: ctx.writer.pkg?.id, copy: f.copy }));
        } catch (error) {
          pictures.set(key, null);
          ctx.writer.loss({ id: 'picture-bullet:' + key, what: 'A picture bullet could not be retained: ' + error.message, where: ctx.part, action: 'conversion' });
        }
      }
      return pictures.get(key);
    };
    // Definitions belong to the numbering model, including unused originals.
    // Imported levels bring their own definition and dependency references.
    for (const record of nb.keep?.pictures || []) picture(record);
    ctx = Object.assign({}, ctx, { pictureBullet: lv => {
      if (!lv.pictureKeep) return null;
      if (L.preserve.pictureBulletIntact(lv)) return picture(lv.pictureKeep);
      ctx.writer.loss({ id: 'picture-bullet-edit:' + lv.pictureKeep.fragment.source + ':' + lv.pictureKeep.id, what: 'Editing this list symbol replaced its original picture bullet.', where: ctx.part, action: 'conversion' });
      return null;
    } });
    // Collect imported pictures before writing any abstract/concrete definitions.
    for (const a of Object.values(nb.abs)) for (const lv of a.levels) if (lv?.pictureKeep && L.preserve.pictureBulletIntact(lv)) picture(lv.pictureKeep);
    for (const n of Object.values(nb.nums)) for (const o of Object.values(n.ov || {})) if (o.lvl?.pictureKeep && L.preserve.pictureBulletIntact(o.lvl)) picture(o.lvl.pictureKeep);
    x += pictureXML;
    const absIds = Object.keys(nb.abs);
    const absMap = new Map();
    absIds.forEach(k => absMap.set(String(k), String(k)));
    for (const k of absIds) {
      const a = nb.abs[k];
      let ax = `<w:multiLevelType w:val="${a.multi || 'hybridMultilevel'}"/>`;
      if (a.name) ax += `<w:name w:val="${X(a.name)}"/>`;
      if (a.styleDef) ax += `<w:styleLink w:val="${X(a.styleDef)}"/>`;
      if (a.styleLink && !a.styleDef) ax += `<w:numStyleLink w:val="${X(a.styleLink)}"/>`;
      if (!a.styleLink || a.styleDef) {
        for (let i = 0; i < 9; i++) if (L.preserve.writeNumberingLevel(a, i)) ax += lvlXML(a.levels[i], i, ctx);
      } else if (a.x?.ownLevels) {
        for (let i = 0; i < 9; i++) if (!Object.hasOwn(a.x.absentLevels || {}, i)) ax += lvlXML(a.x.ownLevels[i], i, ctx);
      }
      if (a.x?.materialized) ctx.writer.loss({ id: 'numbering-style:' + a.x.key, what: 'Editing this list replaced its numbering-style link with direct list formatting.', where: ctx.part, action: 'conversion' });
      x += L.preserve.propertyXML('abstractNum', a, ax, ctx, { abstractNumId: absMap.get(String(k)) });
    }
    for (const id of Object.keys(nb.nums)) {
      const n = nb.nums[id];
      if (!absMap.has(String(n.abs))) continue;
      let nx = `<w:abstractNumId w:val="${absMap.get(String(n.abs))}"/>`;
      for (const l of Object.keys(n.ov || {})) {
        const o = n.ov[l];
        nx += L.preserve.propertyXML('lvlOverride', o, `${o.start != null ? `<w:startOverride w:val="${o.start}"/>` : ''}${o.lvl ? lvlXML(o.lvl, +l, ctx) : ''}`, ctx, { ilvl: l });
      }
      x += L.preserve.propertyXML('num', n, nx, ctx, { numId: ctx.numMap(id) });
    }
    return HEAD + L.preserve.propertyXML('numbering', nb, x, ctx);
  }
  function settingsXML(doc, opts) {
    const s = doc.settings;
    let x = HEAD + `<w:settings xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:m="${NS.m}" xmlns:v="${NS.v}" xmlns:o="${NS.o}">`;
    x += `<w:zoom w:percent="${Math.round(s.zoom || 100)}"/>`;
    if (doc.bg) x += '<w:displayBackgroundShape/>';
    if (s.mirror) x += '<w:mirrorMargins/>';
    if (s.gutterTop) x += '<w:gutterAtTop/>';
    if (s.track) x += '<w:trackRevisions/>';
    if (s.protect && s.protect.kind) {
      const p = s.protect;
      const ed = p.kind === 'tracked' ? 'trackedChanges' : p.kind;
      const raw = p.raw || {};
      const keep = ['w:cryptProviderType', 'w:cryptAlgorithmClass', 'w:cryptAlgorithmType', 'w:cryptAlgorithmSid', 'w:cryptSpinCount', 'w:hash', 'w:salt', 'w:algorithmName', 'w:hashValue', 'w:saltValue', 'w:spinCount'];
      const extra = p.hashed ? Object.entries(p.hashed).map(([k, v]) => ` w:${k}="${X(v)}"`).join('') : keep.filter((k) => raw[k] != null).map((k) => ` ${k}="${X(raw[k])}"`).join('');
      x += `<w:documentProtection w:edit="${ed}" w:enforcement="${p.on ? 1 : 0}"${extra}/>`;
    }
    x += `<w:defaultTabStop w:val="${tw(s.defTab || 36)}"/>`;
    if (s.autoHyphen) x += '<w:autoHyphenation/>';
    if (s.evenOdd) x += '<w:evenAndOddHeaders/>';
    x += '<w:characterSpacingControl w:val="doNotCompress"/>';
    if (s.updateFields || (opts && opts.updateFields)) x += '<w:updateFields w:val="true"/>';
    if (opts && opts.hasFn) x += notePrXML('footnotePr', s.fnPr || {}, true);
    if (opts && opts.hasEn) x += notePrXML('endnotePr', s.enPr || {}, true);
    const legacy = (s.compatXML || '').replace(/<w:compatSetting [^>]*\/>/g, (m) => (/compatibilityMode/.test(m) ? '' : m));
    /* legacy layout flags Word 2003 understands; newer flags outside the transitional schema are dropped */
    const COMPAT_OK = /^(useSingleBorderforContiguousCells|wpJustification|noTabHangInd|noLeading|spaceForUL|noColumnBalance|balanceSingleByteDoubleByteWidth|noExtraLineSpacing|doNotLeaveBackslashAlone|ulTrailSpace|doNotExpandShiftReturn|spacingInWholePoints|lineWrapLikeWord6|printBodyTextBeforeHeader|printColBlack|wpSpaceWidth|showBreaksInFrames|subFontBySize|suppressBottomSpacing|suppressTopSpacing|suppressSpacingAtTopOfPage|suppressTopSpacingWP|suppressSpBfAfterPgBrk|swapBordersFacingPages|convMailMergeEsc|truncateFontHeightsLikeWP6|mwSmallCaps|usePrinterMetrics|doNotSuppressParagraphBorders|wrapTrailSpaces|footnoteLayoutLikeWW8|shapeLayoutLikeWW8|alignTablesRowByRow|forgetLastTabAlignment|adjustLineHeightInTable|autoSpaceLikeWord95|noSpaceRaiseLower|doNotUseHTMLParagraphAutoSpacing|layoutRawTableWidth|layoutTableRowsApart|useWord97LineBreakRules|doNotBreakWrappedTables|doNotSnapToGridInCell|selectFldWithFirstOrLastChar|applyBreakingRules|doNotWrapTextWithPunct|doNotUseEastAsianBreakRules|useWord2002TableStyleRules|growAutofit|useFELayout|useNormalStyleForList|doNotUseIndentAsNumberingTabStop|useAltKinsokuLineBreakRules|allowSpaceOfSameStyleInTable|doNotSuppressIndentation|doNotAutofitConstrainedTables|autofitToFirstFixedWidthCell|underlineTabInNumList|displayHangulFixedWidth|doNotVertAlignCellWithSp|doNotBreakConstrainedForcedTable|doNotVertAlignInTxbx|useAnsiKerningPairs|cachedColBalance)$/;
    const legacyFlags = legacy.replace(/<w:compatSetting [^>]*\/>/g, '').replace(/<w:(\w+)[^>]*\/>/g, (m, n) => (COMPAT_OK.test(n) ? m : '')).replace(/<(?!w:)[\w.-]+:[^>]*\/>|<(?!\/?w:)[\w.-]+:[^>]*>[^<]*<\/[\w.-]+:[^>]*>/g, '');
    const legacySettings = (legacy.match(/<w:compatSetting [^>]*\/>/g) || []).join('');
    x += `<w:compat>${legacyFlags}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${s.compat || 12}"/>${legacySettings}</w:compat>`;
    if (s.docVars && s.docVars.length) x += `<w:docVars>${s.docVars.map((v) => `<w:docVar w:name="${X(v.name)}" w:val="${X(v.val)}"/>`).join('')}</w:docVars>`;
    x += '<w:themeFontLang w:val="en-US"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>';
    x += '<w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>';
    return x + '</w:settings>';
  }
  function themeXML(doc) {
    const t = doc.theme || {};
    const c = Object.assign({}, L.dml.DEFAULT_THEME, t.colors || {});
    const sc = (k) => (k === 'dk1' || k === 'lt1' ? `<a:${k}><a:sysClr val="${k === 'dk1' ? 'windowText' : 'window'}" lastClr="${(c[k] || '#000000').replace('#', '').toUpperCase()}"/></a:${k}>` : `<a:${k}><a:srgbClr val="${(c[k] || '#000000').replace('#', '').toUpperCase()}"/></a:${k}>`);
    const fnt = (f) => `<a:latin typeface="${X(f)}"/><a:ea typeface=""/><a:cs typeface=""/>`;
    const ph = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    const ln = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>`;
    return HEAD + `<a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements><a:clrScheme name="Office">${['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(sc).join('')}</a:clrScheme>` +
      `<a:fontScheme name="Office"><a:majorFont>${fnt(t.major || 'Cambria')}</a:majorFont><a:minorFont>${fnt(t.minor || 'Calibri')}</a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="Office"><a:fillStyleLst>${ph}${ph}${ph}</a:fillStyleLst><a:lnStyleLst>${ln(9525)}${ln(25400)}${ln(38100)}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${ph}${ph}${ph}</a:bgFillStyleLst></a:fmtScheme>` +
      '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
  }
  function fontTableXML(doc) {
    const fonts = new Set(['Times New Roman', 'Arial', 'Symbol', 'Courier New', 'Wingdings']);
    const add = (r) => { if (r && r.font) fonts.add(r.font); };
    for (const k in doc.styles) add(doc.styles[k].rPr);
    add(doc.defaults.rPr);
    for (const st of D.stories(doc)) D.walk(st, (b) => { if (b.t === 'p') { add(b.rPr); for (const it of b.runs) add(it.rPr); } });
    return HEAD + `<w:fonts xmlns:w="${NS.w}" xmlns:r="${NS.r}">` + Array.from(fonts).map((f) => `<w:font w:name="${X(f)}">${/symbol|wingdings/i.test(f) ? '<w:charset w:val="02"/>' : ''}<w:family w:val="${/courier|console|mono/i.test(f) ? 'modern' : /times|georgia|garamond|cambria|book|palatino|serif/i.test(f) ? 'roman' : /symbol|wingdings/i.test(f) ? 'auto' : 'swiss'}"/><w:pitch w:val="${/courier|console|mono/i.test(f) ? 'fixed' : 'variable'}"/></w:font>`).join('') + '</w:fonts>';
  }
  const iso = (d) => { try { return new Date(d || Date.now()).toISOString().replace(/\.\d+Z$/, 'Z'); } catch (e) { return new Date().toISOString().replace(/\.\d+Z$/, 'Z'); } };
  function coreXML(doc) {
    const p = doc.props;
    return HEAD + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      (p.title ? `<dc:title>${X(p.title)}</dc:title>` : '') + (p.subject ? `<dc:subject>${X(p.subject)}</dc:subject>` : '') + `<dc:creator>${X(p.creator || '')}</dc:creator>` +
      (p.keywords ? `<cp:keywords>${X(p.keywords)}</cp:keywords>` : '') + (p.description ? `<dc:description>${X(p.description)}</dc:description>` : '') +
      `<cp:lastModifiedBy>${X(p.lastModifiedBy || L.store.get('userName', 'Quire User'))}</cp:lastModifiedBy><cp:revision>${Math.max(1, (p.revision | 0))}</cp:revision>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${iso(p.created)}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${iso(p.modified)}</dcterms:modified>` +
      (p.category ? `<cp:category>${X(p.category)}</cp:category>` : '') +
      (p.lastPrinted ? `<cp:lastPrinted>${X(p.lastPrinted)}</cp:lastPrinted>` : '') + (p.contentStatus ? `<cp:contentStatus>${X(p.contentStatus)}</cp:contentStatus>` : '') +
      (p.language ? `<dc:language>${X(p.language)}</dc:language>` : '') + (p.identifier ? `<dc:identifier>${X(p.identifier)}</dc:identifier>` : '') + (p.version ? `<cp:version>${X(p.version)}</cp:version>` : '') + '</cp:coreProperties>';
  }
  function appXML(doc, stats) {
    const p = doc.props;
    return HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      `<Template>${X(p.template || 'Normal.dotm')}</Template><TotalTime>0</TotalTime><Pages>${stats.pages}</Pages><Words>${stats.words}</Words><Characters>${stats.charsNoSp}</Characters><Application>Quire 2003</Application><DocSecurity>0</DocSecurity><Lines>${stats.lines}</Lines><Paragraphs>${stats.paras}</Paragraphs><ScaleCrop>false</ScaleCrop>` +
      (p.manager ? `<Manager>${X(p.manager)}</Manager>` : '') + `<Company>${X(p.company || '')}</Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>${stats.chars}</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>11.0000</AppVersion></Properties>`;
  }
  function customXML(doc) {
    const keys = Object.keys(doc.custom || {});
    return HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      keys.map((k, i) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${X(k)}"><vt:lpwstr>${X(doc.custom[k])}</vt:lpwstr></property>`).join('') + '</Properties>';
  }
  /** VML text/picture watermark for headers */
  function watermarkXML(wm, ctx) {
    if (!wm) return '';
    const typeId = '_x0000_t' + ctx.writer.ids.fresh(ctx.part, 'vmlType:_x0000_t');
    const shapeId = '_x0000_s' + ctx.writer.ids.fresh(ctx.part, 'vml:_x0000_s');
    if (wm.type === 'text' && wm.text) {
      const sz = wm.size && wm.size > 1 ? wm.size : 1;
      const w = Math.max(100, Math.min(500, (wm.text.length * (wm.size || 40)) * 0.55)), h = Math.max(40, (wm.size || 60) * 1.1);
      return `<w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="${typeId}" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype><v:shape id="PowerPlusWaterMarkObject${ctx.docPrId()}" o:spid="${shapeId}" type="#${typeId}" style="position:absolute;margin-left:0;margin-top:0;width:${L.round(w, 1)}pt;height:${L.round(h, 1)}pt;${wm.layout === 'horizontal' ? '' : 'rotation:315;'}z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="#${wm.color || 'C0C0C0'}" stroked="f"><v:fill opacity="${wm.semi === false ? '1' : '.5'}"/><v:textpath style="font-family:&quot;${X(wm.font || 'Times New Roman')}&quot;;font-size:${Math.round(sz)}pt" string="${X(wm.text)}"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>`;
    }
    if (wm.type === 'picture' && wm.media) {
      const rid = ctx.media(wm.media);
      if (!rid) return '';
      const m = L.media.get(wm.media);
      const w = wm.w || 300, h = wm.h || 200;
      void m;
      return `<w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="${typeId}" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="WordPictureWatermark${ctx.docPrId()}" o:spid="${shapeId}" type="#${typeId}" style="position:absolute;margin-left:0;margin-top:0;width:${L.round(w, 1)}pt;height:${L.round(h, 1)}pt;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f"><v:imagedata r:id="${rid}" o:title="" gain="19661f" blacklevel="22938f"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>`;
    }
    return '';
  }

  /* ================= media ================= */
  async function mediaBytes(id) {
    const m = L.media.get(id);
    if (!m) return null;
    let blob = m.blob, type = blob.type || m.type || 'image/png';
    if (/webp|avif|heic/.test(type)) {
      try {
        const img = await L.loadImage(m.url);
        const c = document.createElement('canvas');
        c.width = img.naturalWidth || 512; c.height = img.naturalHeight || 512;
        c.getContext('2d').drawImage(img, 0, 0);
        blob = await new Promise((res) => c.toBlob(res, 'image/png'));
        type = 'image/png';
      } catch (e) { /* keep */ }
    }
    let ext = L.mimeToExt(type);
    if (ext === 'png' && m.name && /\.(emf|wmf|tiff?|svg)$/i.test(m.name)) ext = m.name.split('.').pop().toLowerCase();
    return { bytes: new Uint8Array(await L.readAsArrayBuffer(blob)), ext };
  }

  /* ================= package ================= */
  W.write = async function (doc, opts) {
    opts = opts || {};
    const K = L.opc, format = K.format(doc, opts.format || (opts.template ? 'dotx' : null), 'docx');
    const pack = L.preserve.begin(doc, format);
    const files = [];
    const add = (name, data) => files.push({ name, data });
    const overrides = [];
    const defaults = new Map([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]);
    const docRels = pack.rels('word/document.xml');
    D.reindex(doc);
    const ranges = L.preserve.prepareRanges(doc);
    const objects = L.preserve.prepareObjects(doc, pack.writer);
    L.preserve.prepareNumbering(doc.numbering, pack.writer);
    /* --- context shared by every story --- */
    const bmIds = new Map();
    const mediaMap = new Map();
    let mediaN = 0;
    const pendingMedia = [];
    const numIds = new Map();
    const numMap = (id) => { if (!numIds.has(String(id))) numIds.set(String(id), String(id)); return numIds.get(String(id)); };
    const cmtIds = new Map();
    const cmtReferences = new Set();
    const fnIds = new Map(), enIds = new Map();
    const highNote = kind => Math.max(0, ...Object.keys(kind === 'fn' ? doc.fn : doc.en).map(x => +x || 0), ...(doc.keep?.notes?.[kind] || []).map(n => +n.id || 0)) + 1;
    let fnN = highNote('fn'), enN = highNote('en');
    const charts = [];
    const usedPaths = new Set();
    const mkCtx = (rels, main) => ({
      rels, main, doc, ranges, objects, writer: pack.writer, part: rels.owner, mainPart: pack.part('word/document.xml'),
      drawingXML,
      rev: () => pack.writer.ids.fresh('document', 'revision'),
      docPrId: () => {
        // Generated frames use one value for docPr and cNvPr. Original shapes
        // have an independent, part-local ID space, which must also be reserved.
        const ids = pack.writer.ids, shapes = ids.space(rels.owner, 'shape');
        let id;
        do { id = ids.fresh('document', 'docPr'); } while (shapes.used.has(ids.canonical(id)));
        ids.reserve(rels.owner, 'shape', id);
        return id;
      },
      bmSeen: new Set(),
      bmId: (id) => { const k = rels.owner + ':' + id; if (!bmIds.has(k)) bmIds.set(k, pack.writer.ids.fresh(rels.owner, 'bookmark')); return bmIds.get(k); },
      cmtOk: (id) => !!doc.comments[id],
      cmtReference: (id) => {
        const key = String(id);
        if (cmtReferences.has(key)) {
          pack.writer.loss({ id: 'comment-reference:' + key, what: 'A duplicate reference to the same comment was removed; the comment and its first reference are kept.', where: rels.owner, action: 'conversion' });
          return false;
        }
        cmtReferences.add(key); return true;
      },
      cmtId: (id) => { const k = String(id); if (!cmtIds.has(k)) cmtIds.set(k, cmtIds.size); return cmtIds.get(k); },
      noteId: (kind, id) => {
        const store = kind === 'fn' ? doc.fn : doc.en;
        if (!store[id]) return null;
        const m = kind === 'fn' ? fnIds : enIds;
        if (!m.has(String(id))) m.set(String(id), /^\d+$/.test(String(id)) ? +id : kind === 'fn' ? fnN++ : enN++);
        return m.get(String(id));
      },
      numMap,
      media: (mid) => {
        if (!mid || !L.media.has(mid)) return null;
        let e = mediaMap.get(mid);
        if (!e) { e = { name: null, pending: true }; mediaMap.set(mid, e); pendingMedia.push(mid); }
        return rels.add(RT('image'), () => 'media/' + mediaMap.get(mid).name);
      },
      chart: (it) => {
        const src = it.src && L.chart.src.get(it.src), source = src?.source && K.package(src.source);
        let conversion = false;
        if (src && it.dirtyChart) pack.writer.loss({ id: 'chart-edit:' + it.src, what: 'Editing this chart replaces its original chart-specific formatting and extensions.', where: src.path || rels.owner, action: 'conversion' });
        if (source && !it.dirtyChart) {
          try {
            const target = pack.writer.carry(source, src.path);
            return pack.writer.rels(rels.owner).add(RT('chart'), K.relative(rels.owner, target));
          } catch (error) {
            conversion = true;
            pack.writer.loss({ id: 'chart:' + src.source + ':' + src.path, what: 'The chart was converted because its original dependencies are incomplete: ' + error.message, where: src.path, action: 'conversion' });
          }
        }
        const n = charts.length + 1;
        let path = 'word/charts/chart' + n + '.xml';
        while (usedPaths.has(path)) path = path.replace(/chart(\d+)\.xml$/, (m, k) => 'chart' + (+k + 1) + '.xml');
        usedPaths.add(path);
        charts.push({ it, path, conversion });
        return rels.add(RT('chart'), path.replace(/^word\//, ''));
      },
    });
    /* resolve deferred media targets after names are assigned */
    const finishRels = (rels) => { for (const r of rels.list) if (typeof r.target === 'function') r.target = r.target(); };

    /* --- headers & footers --- */
    const hfRid = {};
    let hN = 0, fN = 0;
    const hfParts = [];
    /* the watermark lives in the default header of every section */
    const wmHeaders = new Set();
    if (doc.watermark && !L.preserve.keepWatermark(doc) && (doc.watermark.text || doc.watermark.media)) {
      const secs = D.sections(doc);
      for (const s of secs) {
        if (!s.sect.refs) s.sect.refs = { hdr: {}, ftr: {} };
        for (const ty of ['default', 'first', 'even']) {
          if (ty !== 'default' && !(ty === 'first' && s.sect.titlePg) && !(ty === 'even' && doc.settings.evenOdd)) continue;
          let id = s.sect.refs.hdr[ty];
          if (!id) { id = 'wmh' + ty + secs.indexOf(s); if (!doc.hf[id]) doc.hf[id] = { kind: 'hdr', id, blocks: [D.para([], { style: doc.styles.Header ? 'Header' : undefined })], _temp: true }; s.sect.refs.hdr[ty] = id; s._tempRefs = (s._tempRefs || []).concat([[ty, id]]); }
          wmHeaders.add(id);
        }
      }
    }
    for (const id of Object.keys(doc.hf)) {
      const st = doc.hf[id];
      const isH = st.kind === 'hdr';
      const name = isH ? `header${++hN}.xml` : `footer${++fN}.xml`;
      pack.bind('word/' + name, st.keep?.part);
      const rels = pack.rels('word/' + name);
      const ctx = mkCtx(rels, false);
      let body = blocksXML(st.blocks.length ? st.blocks : [D.para()], ctx);
      if (!st.blocks.length || st.blocks[st.blocks.length - 1].t !== 'p') body += '<w:p/>';
      if (wmHeaders.has(id)) body = prependParagraph(body, watermarkXML(doc.watermark, ctx));
      hfParts.push({ name, rels, isH, body });
      hfRid[id] = docRels.add(RT(isH ? 'header' : 'footer'), name);
    }
    /* --- body --- */
    const mainCtx = mkCtx(docRels, true);
    mainCtx.hfRid = hfRid;
    for (const p of hfParts) p.ctxHf = hfRid;
    let body = blocksXML(doc.main.blocks, Object.assign(mainCtx, { hfRid }));
    body += sectXML(doc.sect, mainCtx);
    /* undo temporary watermark header refs */
    for (const s of D.sections(doc)) if (s._tempRefs) for (const [ty] of s._tempRefs) delete s.sect.refs.hdr[ty];
    for (const k of Object.keys(doc.hf)) if (doc.hf[k]._temp) delete doc.hf[k];
    const bg = doc.keep?.background && doc.bg === doc.keep.background.value ? pack.writer.emit(doc.keep.background.fragment, mainCtx.part) : doc.bg ? `<w:background w:color="${hex(doc.bg)}"/>` : '';
    const documentXML = HEAD + `<w:document ${NSDECL}>${bg}<w:body>${body}</w:body></w:document>`;
    /* --- notes --- */
    const noteXML = (kind) => {
      const tag = kind === 'fn' ? 'footnote' : 'endnote';
      const store = kind === 'fn' ? doc.fn : doc.en;
      const ids = kind === 'fn' ? fnIds : enIds;
      const rels = pack.rels('word/' + tag + 's.xml');
      const ctx = mkCtx(rels, false);
      let x = `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}><w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`;
      if (doc.keep?.notes?.[kind]) x = doc.keep.notes[kind].map(n => pack.writer.emit(n.fragment, rels.owner)).join('');
      for (const [orig, nid] of ids) {
        const st = store[orig];
        if (!st) continue;
        let inner = blocksXML(st.blocks, ctx);
        if (!/<w:(footnote|endnote)Ref\/>/.test(inner)) inner = prependParagraph(inner, `<w:r><w:rPr><w:rStyle w:val="${kind === 'fn' ? 'FootnoteReference' : 'EndnoteReference'}"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>`);
        x += `<w:${tag} w:id="${nid}">${inner}</w:${tag}>`;
      }
      return { xml: HEAD + `<w:${tag}s ${NSDECL}>${x}</w:${tag}s>`, rels };
    };
    let fnPart = null, enPart = null;
    if (fnIds.size || pack.originals.has('word/footnotes.xml')) { fnPart = noteXML('fn'); }
    if (enIds.size || pack.originals.has('word/endnotes.xml')) { enPart = noteXML('en'); }
    /* --- comments --- */
    let cmPart = null, cmExPart = null;
    if (cmtIds.size) {
      const rels = pack.rels('word/comments.xml');
      const ctx = mkCtx(rels, false);
      let x = '', ex = '';
      const reservedParaIds = new Set(Object.values(doc.comments).map(c => c.keep?.paraId).filter(Boolean));
      let paraN = Math.max(0x10000000, ...Array.from(reservedParaIds, p => parseInt(p, 16) || 0));
      const paraIds = new Map();
      for (const [orig, nid] of cmtIds) {
        const c = doc.comments[orig];
        if (!c) continue;
        let inner = blocksXML(c.blocks, ctx);
        if (!/annotationRef/.test(inner)) inner = prependParagraph(inner, '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>');
        const pid = c.keep?.paraId || (++paraN).toString(16).toUpperCase().padStart(8, '0');
        paraIds.set(String(orig), pid);
        /* tag the last paragraph with a paraId for threading / done state */
        const lastIdx = inner.lastIndexOf('<w:p>');
        if (lastIdx >= 0) inner = inner.slice(0, lastIdx) + `<w:p w14:paraId="${pid}" w14:textId="77777777">` + inner.slice(lastIdx + 5);
        x += `<w:comment w:id="${nid}" w:author="${X(c.author || 'Unknown')}" w:date="${X(c.date || iso())}" w:initials="${X(c.initials || '')}">${inner}</w:comment>`;
      }
      for (const [orig] of cmtIds) {
        const c = doc.comments[orig];
        if (!c) continue;
        const pid = paraIds.get(String(orig));
        const parent = c.parent != null ? paraIds.get(String(c.parent)) : null;
        ex += `<w15:commentEx w15:paraId="${pid}"${parent ? ` w15:paraIdParent="${parent}"` : ''} w15:done="${c.done ? 1 : 0}"/>`;
      }
      cmPart = { xml: HEAD + `<w:comments ${NSDECL}>${x}</w:comments>`, rels };
      const extendedNS = doc.keep?.commentExtension?.namespace || NS.w15;
      cmExPart = HEAD + `<w15:commentsEx xmlns:mc="${NS.mc}" xmlns:w15="${X(extendedNS)}" mc:Ignorable="w15">${ex}</w15:commentsEx>`;
    }
    /* --- media names --- */
    for (const mid of pendingMedia) {
      const b = await mediaBytes(mid);
      const e = mediaMap.get(mid);
      if (!b) { e.name = 'missing.png'; continue; }
      e.name = `image${++mediaN}.${b.ext}`;
      const source = doc.keep?.media?.[mid], original = source && doc.pkg?.bytes(source);
      if (original && original.length === b.bytes.length && original.every((value, i) => value === b.bytes[i])) pack.bind('word/media/' + e.name, source);
      add('word/media/' + e.name, b.bytes);
      if (!defaults.has(b.ext)) defaults.set(b.ext, L.extToMime(b.ext) === 'application/octet-stream' ? 'image/' + b.ext : L.extToMime(b.ext));
    }
    if (pendingMedia.length) { /* media referenced after the loop (notes/comments) were added above too */ }
    /* --- charts --- */
    for (const ch of charts) {
      const src = typeof ch.it.src === 'string' ? L.chart.src.get(ch.it.src) : ch.it.src;
      if (src && src.files && src.files.length && !ch.it.dirtyChart && !ch.conversion) {
        /* copy the original chart part and everything it references */
        const base = src.path;
        const relp = L.dml.relsPath(base);
        for (const f of src.files) {
          let name = f.path === base ? ch.path : f.path === relp ? L.dml.relsPath(ch.path) : f.path;
          if (!files.some((x) => x.name === name)) add(name, f.data);
          if (f.path === base) overrides.push([ch.path, CT.chart]);
          else if (/\/style\d*\.xml$/.test(f.path)) overrides.push([name, CT.chartStyle]);
          else if (/\/colors\d*\.xml$/.test(f.path)) overrides.push([name, CT.chartColors]);
          const ext = (name.split('.').pop() || '').toLowerCase();
          if (ext === 'xlsx' && !defaults.has('xlsx')) defaults.set('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          if (ext === 'bin' && !defaults.has('bin')) defaults.set('bin', 'application/vnd.openxmlformats-officedocument.oleObject');
        }
      } else {
        add(ch.path, W.chartXML(ch.it.chart || (L.chart && L.chart.sample()) || { kind: 'col', cats: [], series: [] }));
        overrides.push([ch.path, CT.chart]);
      }
    }
    /* --- assemble --- */
    const flatRels = (rels) => { finishRels(rels); return rels; };
    add('word/document.xml', documentXML);
    overrides.push(['word/document.xml', format.contentType]);
    for (const p of hfParts) {
      add('word/' + p.name, HEAD + `<w:${p.isH ? 'hdr' : 'ftr'} ${NSDECL}>${p.body}</w:${p.isH ? 'hdr' : 'ftr'}>`);
      overrides.push(['word/' + p.name, p.isH ? CT.header : CT.footer]);
      if (p.rels.list.length) add('word/_rels/' + p.name + '.rels', flatRels(p.rels).xml());
    }
    const stylesCtx = mkCtx(pack.rels('word/styles.xml'), false);
    add('word/styles.xml', stylesXML(doc, stylesCtx));
    docRels.add(RT('styles'), 'styles.xml');
    overrides.push(['word/styles.xml', CT.styles]);
    const usesNumbering = Object.keys(doc.numbering.nums).length > 0 || Object.keys(doc.numbering.abs).length > 0 || doc.numbering.keep?.pictures?.length || doc.numbering.x;
    if (usesNumbering) {
      add('word/numbering.xml', numberingXML(doc, mkCtx(pack.rels('word/numbering.xml'), false)));
      docRels.add(RT('numbering'), 'numbering.xml');
      overrides.push(['word/numbering.xml', CT.numbering]);
    }
    add('word/settings.xml', settingsXML(doc, { hasFn: !!fnPart, hasEn: !!enPart }));
    docRels.add(RT('settings'), 'settings.xml');
    overrides.push(['word/settings.xml', CT.settings]);
    add('word/fontTable.xml', fontTableXML(doc));
    docRels.add(RT('fontTable'), 'fontTable.xml');
    overrides.push(['word/fontTable.xml', CT.fontTable]);
    add('word/webSettings.xml', HEAD + `<w:webSettings xmlns:w="${NS.w}"><w:optimizeForBrowser/><w:allowPNG/></w:webSettings>`);
    docRels.add(RT('webSettings'), 'webSettings.xml');
    overrides.push(['word/webSettings.xml', CT.webSettings]);
    add('word/theme/theme1.xml', themeXML(doc));
    docRels.add(RT('theme'), 'theme/theme1.xml');
    overrides.push(['word/theme/theme1.xml', CT.theme]);
    if (fnPart) { add('word/footnotes.xml', fnPart.xml); docRels.add(RT('footnotes'), 'footnotes.xml'); overrides.push(['word/footnotes.xml', CT.footnotes]); if (fnPart.rels.list.length) add('word/_rels/footnotes.xml.rels', flatRels(fnPart.rels).xml()); }
    if (enPart) { add('word/endnotes.xml', enPart.xml); docRels.add(RT('endnotes'), 'endnotes.xml'); overrides.push(['word/endnotes.xml', CT.endnotes]); if (enPart.rels.list.length) add('word/_rels/endnotes.xml.rels', flatRels(enPart.rels).xml()); }
    if (cmPart) {
      add('word/comments.xml', cmPart.xml); docRels.add(RT('comments'), 'comments.xml'); overrides.push(['word/comments.xml', CT.comments]);
      if (cmPart.rels.list.length) add('word/_rels/comments.xml.rels', flatRels(cmPart.rels).xml());
      add('word/commentsExtended.xml', cmExPart); docRels.add('http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'commentsExtended.xml'); overrides.push(['word/commentsExtended.xml', CT.commentsEx]);
    }
    add('word/_rels/document.xml.rels', flatRels(docRels).xml());
    /* package-level parts */
    const stats = opts.stats || { pages: 1, words: 0, chars: 0, charsNoSp: 0, paras: 0, lines: 0 };
    add('docProps/core.xml', coreXML(doc));
    overrides.push(['docProps/core.xml', CT.core]);
    add('docProps/app.xml', appXML(doc, stats));
    overrides.push(['docProps/app.xml', CT.app]);
    const hasCustom = Object.keys(doc.custom || {}).length > 0 || pack.originals.has('docProps/custom.xml');
    if (hasCustom) { add('docProps/custom.xml', customXML(doc)); overrides.push(['docProps/custom.xml', CT.custom]); }
    const rootRels = pack.rels('');
    rootRels.add(RT('officeDocument'), 'word/document.xml');
    rootRels.add('http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml');
    rootRels.add(RT('extended-properties'), 'docProps/app.xml');
    if (hasCustom) rootRels.add(RT('custom-properties'), 'docProps/custom.xml');
    const ct = HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      Array.from(defaults).map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join('') +
      overrides.map(([p, t]) => `<Override PartName="/${p}" ContentType="${t}"/>`).join('') + '</Types>';
    files.unshift({ name: '[Content_Types].xml', data: ct });
    const types = new Map(overrides);
    for (const file of files) pack.put(file.name, file.data, types.get(file.name) || defaults.get(file.name.split('.').pop()));
    L.preserve.writeBindings(doc, pack.writer);
    const result = pack.finish();
    const blob = await L.zip.write(result.files, format.mime); blob.dropped = result.dropped;
    return blob;
  };

  /* ---------- chart XML from the chart model (charts created in Quire) ---------- */
  W.chartXML = function (c) {
    const NS_C = NS.c, NS_A = NS.a;
    const v = (x) => `<c:v>${X(x)}</c:v>`;
    const cats = c.cats || [];
    const strLit = (arr) => `<c:strLit><c:ptCount val="${arr.length}"/>${arr.map((x, i) => `<c:pt idx="${i}">${v(x)}</c:pt>`).join('')}</c:strLit>`;
    const numLit = (arr) => `<c:numLit><c:formatCode>${X(c.numFmt || 'General')}</c:formatCode><c:ptCount val="${arr.length}"/>${arr.map((x, i) => `<c:pt idx="${i}">${v(+x || 0)}</c:pt>`).join('')}</c:numLit>`;
    const sp = (col, isLine) => (isLine ? `<c:spPr><a:ln w="28575"><a:solidFill>${clr(col)}</a:solidFill></a:ln></c:spPr>` : `<c:spPr><a:solidFill>${clr(col)}</a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>`);
    const pal = (L.chart && L.chart.PALETTE) || ['#9999FF', '#993366', '#FFFFCC'];
    const series = c.series || [];
    const dl = c.labels ? `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${c.kind === 'pie' || c.kind === 'doughnut' ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${c.kind === 'pie' || c.kind === 'doughnut' ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>` : '';
    const sers = (kind) => series.map((s, i) => {
      const col = s.color || pal[i % pal.length];
      const head = `<c:idx val="${i}"/><c:order val="${i}"/><c:tx>${v(s.name || 'Series ' + (i + 1))}</c:tx>`;
      if (kind === 'pie') { const dpts = cats.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/>${sp((c.pieColors && c.pieColors[k]) || pal[k % pal.length])}</c:dPt>`).join(''); return `<c:ser>${head}${dpts}${dl}<c:cat>${strLit(cats)}</c:cat><c:val>${numLit(s.vals)}</c:val></c:ser>`; }
      if (kind === 'line') return `<c:ser>${head}${sp(col, true)}<c:marker><c:symbol val="${c.kind === 'lineMarkers' ? 'square' : 'none'}"/>${c.kind === 'lineMarkers' ? '<c:size val="6"/>' : ''}</c:marker>${dl}<c:cat>${strLit(cats)}</c:cat><c:val>${numLit(s.vals)}</c:val><c:smooth val="0"/></c:ser>`;
      if (kind === 'scatter') return `<c:ser>${head}${sp(col, true)}<c:marker><c:symbol val="square"/><c:size val="6"/></c:marker>${dl}<c:xVal>${numLit(cats.map((x) => +x || 0))}</c:xVal><c:yVal>${numLit(s.vals)}</c:yVal><c:smooth val="0"/></c:ser>`;
      if (kind === 'area') return `<c:ser>${head}${sp(col)}${dl}<c:cat>${strLit(cats)}</c:cat><c:val>${numLit(s.vals)}</c:val></c:ser>`;
      return `<c:ser>${head}${sp(col)}<c:invertIfNegative val="0"/>${dl}<c:cat>${strLit(cats)}</c:cat><c:val>${numLit(s.vals)}</c:val></c:ser>`;
    }).join('');
    const AX1 = 50010001, AX2 = 50010002;
    const axIds = `<c:axId val="${AX1}"/><c:axId val="${AX2}"/>`;
    const grid = c.gridY !== false ? '<c:majorGridlines/>' : '';
    const axTitle = (pos) => { const t = /[bt]/.test(pos) ? c.axTitleX : c.axTitleY; return t ? `<c:title><c:tx><c:rich><a:bodyPr${/[lr]/.test(pos) ? (c.axTitleYHoriz ? ' rot="0" vert="horz"' : ' rot="-5400000" vert="horz"') : ''}/><a:lstStyle/><a:p><a:pPr><a:defRPr b="1"/></a:pPr><a:r><a:rPr lang="en-US" b="1"/><a:t>${X(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : ''; };
    const catAx = (pos) => `<c:catAx><c:axId val="${AX1}"/><c:scaling><c:orientation val="${pos === 'l' ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${axTitle(pos)}<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${AX2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const valAx = (pos, id, cross, g) => `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${g ? grid : ''}${axTitle(pos)}<c:numFmt formatCode="${c.kind === 'colPct' ? '0%' : 'General'}" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
    let plot;
    switch (c.kind) {
      case 'bar': case 'barStacked': plot = `<c:barChart><c:barDir val="bar"/><c:grouping val="${c.kind === 'barStacked' ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>${sers('bar')}<c:gapWidth val="150"/>${c.kind === 'barStacked' ? '<c:overlap val="100"/>' : ''}${axIds}</c:barChart>${catAx('l')}${valAx('b', AX2, AX1, true)}`; break;
      case 'line': case 'lineMarkers': plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers('line')}<c:marker val="1"/>${axIds}</c:lineChart>${catAx('b')}${valAx('l', AX2, AX1, true)}`; break;
      case 'area': case 'areaStacked': plot = `<c:areaChart><c:grouping val="${c.kind === 'areaStacked' ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${sers('area')}${axIds}</c:areaChart>${catAx('b')}${valAx('l', AX2, AX1, true)}`; break;
      case 'pie': plot = `<c:pieChart><c:varyColors val="1"/>${sers('pie').split('</c:ser>')[0] + (series.length ? '</c:ser>' : '')}<c:firstSliceAng val="0"/></c:pieChart>`; break;
      case 'doughnut': plot = `<c:doughnutChart><c:varyColors val="1"/>${sers('pie').split('</c:ser>')[0] + (series.length ? '</c:ser>' : '')}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`; break;
      case 'scatter': plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers('scatter')}${axIds}</c:scatterChart>${valAx('b', AX1, AX2, false)}${valAx('l', AX2, AX1, true)}`; break;
      default: plot = `<c:barChart><c:barDir val="col"/><c:grouping val="${c.kind === 'colStacked' ? 'stacked' : c.kind === 'colPct' ? 'percentStacked' : 'clustered'}"/><c:varyColors val="0"/>${sers('bar')}<c:gapWidth val="150"/>${c.kind === 'colStacked' || c.kind === 'colPct' ? '<c:overlap val="100"/>' : ''}${axIds}</c:barChart>${catAx('b')}${valAx('l', AX2, AX1, true)}`;
    }
    const title = c.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round((c.fsz || 10) * 140)}" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="${Math.round((c.fsz || 10) * 140)}" b="1"/><a:t>${X(c.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    const legend = c.legend && c.legend !== 'none' ? `<c:legend><c:legendPos val="${c.legend}"/><c:overlay val="0"/></c:legend>` : '';
    return HEAD + `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS.r}"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round((c.fsz || 10) * 100)}"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:chartSpace>`;
  };
})();
