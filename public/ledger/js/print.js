/* Ledger — printing: pagination (automatic and manual page breaks, fit to pages, print titles),
 * page rendering with headers / footers, Print Preview, and PDF output. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, LY = L.layout, F = L.formula, C = L.calc, NF = L.numfmt;
  const { h } = L;
  const ui = L.ui;
  const P = (L.print = {});
  const G = () => L.grid;
  /* paper sizes (inches) by SpreadsheetML paperSize code */
  P.PAPERS = { 1: ['Letter', 8.5, 11], 3: ['Tabloid', 11, 17], 5: ['Legal', 8.5, 14], 7: ['Executive', 7.25, 10.5], 8: ['A3', 11.69, 16.54], 9: ['A4', 8.27, 11.69], 11: ['A5', 5.83, 8.27], 13: ['B5 (JIS)', 7.17, 10.12], 20: ['Envelope #10', 4.125, 9.5], 27: ['Envelope DL', 4.33, 8.66], 34: ['B5 (ISO)', 6.93, 9.84] };
  P.paperSize = (p) => { const x = P.PAPERS[p.paper] || P.PAPERS[1]; return p.orientation === 'landscape' ? { w: x[2] * 72, h: x[1] * 72, name: x[0] } : { w: x[1] * 72, h: x[2] * 72, name: x[0] }; };

  /* ------------------------------------------------------------ print area */
  const nameRange = (wb, sh, key) => {
    const idx = wb.sheets.indexOf(sh);
    const n = wb.names.find((x) => x.name.toLowerCase() === key && x.scope === idx);
    if (!n) return null;
    const out = [];
    for (const part of String(n.ref).split(',')) {
      try {
        const ast = F.parse(part.trim().replace(/^=/, ''));
        const v = C.ev(ast, C.ctx(wb, sh, 0, 0));
        const list = v instanceof C.RefList ? v.list : v instanceof C.Ref ? [v] : [];
        for (const a of list) out.push({ r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 });
      } catch (e) { /* ignore */ }
    }
    return out.length ? out : null;
  };
  /** the cells that print: the print area, else the used range (values, formats that show, objects) */
  P.area = function (sh) {
    const wb = sh.wb;
    const pa = nameRange(wb, sh, '_xlnm.print_area');
    if (pa) return pa[0];
    let r2 = -1, c2 = -1;
    sh.rows.forEach((row, r) => {
      if (!row) return;
      row.cells.forEach((cell, c) => {
        if (!cell) return;
        let shows = cell.v != null && cell.v !== '';
        if (!shows && cell.s) { const st = wb.styles.get(cell.s); shows = !!(st.border || (st.fill && st.fill.pattern && st.fill.pattern !== 'none')); }
        if (shows) { if (r > r2) r2 = r; if (c > c2) c2 = c; }
      });
    });
    for (const m of sh.merges) if (m.r1 <= r2 && m.c1 <= c2) { r2 = Math.max(r2, m.r2); c2 = Math.max(c2, m.c2); }
    for (const d of sh.drawings) { if (d.noPrint || !d.anchor) continue; const a = d.anchor; if (a.to) { r2 = Math.max(r2, a.to.r); c2 = Math.max(c2, a.to.c); } else if (a.from) { r2 = Math.max(r2, a.from.r + 5); c2 = Math.max(c2, a.from.c + 3); } }
    if (r2 < 0) return null;
    return { r1: 0, c1: 0, r2, c2 };
  };
  P.titles = function (sh) {
    const t = nameRange(sh.wb, sh, '_xlnm.print_titles');
    const out = { rows: null, cols: null };
    if (!t) return out;
    for (const a of t) { if (a.c1 === 0 && a.c2 >= M.MAXC - 1) out.rows = { r1: a.r1, r2: a.r2 }; else if (a.r1 === 0 && a.r2 >= M.MAXR - 1) out.cols = { c1: a.c1, c2: a.c2 }; }
    return out;
  };

  /* ------------------------------------------------------------ pagination */
  /** split [a, b] by sizes (pt) into page spans of at most `avail`, honouring manual breaks */
  function split(a, b, size, avail, manual, titleSize) {
    const spans = [];
    let start = a, acc = 0, first = true;
    for (let i = a; i <= b; i++) {
      const s = size(i);
      const room = avail - (first ? 0 : titleSize);
      if ((manual.has(i) && i > start) || (acc + s > room && i > start)) { spans.push([start, i - 1]); start = i; acc = 0; first = false; }
      acc += s;
    }
    spans.push([start, b]);
    return spans;
  }
  /** pages of a sheet: {pages: [{r1,r2,c1,c2}], scale, rows/cols (break positions for the grid), area} */
  P.paginate = function (sh) {
    const p = sh.print;
    const area = P.area(sh);
    if (!area) return { pages: [], scale: 1, rows: [], cols: [], area: null };
    const g = LY.geo(sh);
    const paper = P.paperSize(p);
    const m = p.margins || { l: 0.75, r: 0.75, t: 1, b: 1, header: 0.5, footer: 0.5 };
    const availW = paper.w - (m.l + m.r) * 72, availH = paper.h - (m.t + m.b) * 72;
    const colPt = (c) => g.cols.size(c) * 0.75, rowPt = (r) => g.rows.size(r) * 0.75;
    const titles = P.titles(sh);
    const tRowH = titles.rows ? Array.from({ length: titles.rows.r2 - titles.rows.r1 + 1 }, (_, i) => rowPt(titles.rows.r1 + i)).reduce((x, y) => x + y, 0) : 0;
    const tColW = titles.cols ? Array.from({ length: titles.cols.c2 - titles.cols.c1 + 1 }, (_, i) => colPt(titles.cols.c1 + i)).reduce((x, y) => x + y, 0) : 0;
    const headW = p.headings ? 30 : 0, headH = p.headings ? 13 : 0;
    let scale = (p.scale || 100) / 100;
    if (p.fit) {
      let totW = 0, totH = 0;
      for (let c = area.c1; c <= area.c2; c++) totW += colPt(c);
      for (let r = area.r1; r <= Math.min(area.r2, area.r1 + 200000); r++) totH += rowPt(r);
      const sw = p.fitW ? (availW * p.fitW) / (totW + headW * p.fitW) : Infinity;
      const shh = p.fitH ? (availH * p.fitH) / (totH + headH * p.fitH) : Infinity;
      scale = Math.min(1, sw, shh);
      if (!isFinite(scale)) scale = 1;
      scale = Math.max(0.1, Math.floor(scale * 100) / 100);
    }
    const mr = new Set((p.rowBreaks || []).filter((b) => b > area.r1 && b <= area.r2));
    const mc = new Set((p.colBreaks || []).filter((b) => b > area.c1 && b <= area.c2));
    const rowSpans = split(area.r1, area.r2, (r) => rowPt(r) * scale, availH - headH * scale, p.fit && p.fitH ? new Set() : mr, tRowH * scale);
    const colSpans = split(area.c1, area.c2, (c) => colPt(c) * scale, availW - headW * scale, p.fit && p.fitW ? new Set() : mc, tColW * scale);
    const pages = [];
    if (p.pageOrder === 'overThenDown') { for (const rs of rowSpans) for (const cs of colSpans) pages.push({ r1: rs[0], r2: rs[1], c1: cs[0], c2: cs[1] }); }
    else for (const cs of colSpans) for (const rs of rowSpans) pages.push({ r1: rs[0], r2: rs[1], c1: cs[0], c2: cs[1] });
    return {
      pages, scale, area, titles, paper, margins: m,
      rows: rowSpans.slice(1).map((s) => ({ r: s[0], manual: mr.has(s[0]) })).concat([{ r: area.r2 + 1, end: true }]),
      cols: colSpans.slice(1).map((s) => ({ c: s[0], manual: mc.has(s[0]) })).concat([{ c: area.c2 + 1, end: true }]),
    };
  };
  /** remember the page breaks so the grid can draw them */
  P.computeBreaks = function (sh) {
    try { sh._pages = P.paginate(sh); } catch (e) { console.error(e); sh._pages = null; }
    if (G() && G().paint) G().paint();
    return sh._pages;
  };

  /* ------------------------------------------------------------ header / footer codes */
  /** "&LLeft&CCenter&RRight" → {l, c, r} of [{text, b, i, u, sz, font}] */
  P.parseHF = function (code, info) {
    const out = { l: [], c: [], r: [] };
    let sec = 'c', cur = { text: '' }, i = 0;
    const s = String(code || '');
    const push = () => { if (cur.text) out[sec].push(cur); cur = Object.assign({}, cur, { text: '' }); };
    while (i < s.length) {
      const ch = s[i];
      if (ch !== '&') { cur.text += ch; i++; continue; }
      const n = s[i + 1];
      if (n == null) { i++; continue; }
      i += 2;
      switch (n) {
        case 'L': push(); sec = 'l'; cur = { text: '' }; break;
        case 'C': push(); sec = 'c'; cur = { text: '' }; break;
        case 'R': push(); sec = 'r'; cur = { text: '' }; break;
        case 'P': cur.text += String(info.page || 1); break;
        case 'N': cur.text += String(info.pages || 1); break;
        case 'D': cur.text += new Date().toLocaleDateString('en-US'); break;
        case 'T': cur.text += new Date().toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' }); break;
        case 'F': cur.text += info.file || ''; break;
        case 'Z': cur.text += ''; break;
        case 'A': cur.text += info.sheet || ''; break;
        case '&': cur.text += '&'; break;
        case 'B': push(); cur.b = !cur.b; break;
        case 'I': push(); cur.i = !cur.i; break;
        case 'U': case 'E': push(); cur.u = !cur.u; break;
        case 'S': push(); cur.s = !cur.s; break;
        case 'X': case 'Y': break;
        case 'G': break;
        case 'K': i += 6; break;
        case '"': {
          const e = s.indexOf('"', i);
          const spec = s.slice(i, e < 0 ? s.length : e);
          i = e < 0 ? s.length : e + 1;
          push();
          const [font, style] = spec.split(',');
          if (font && font !== '-') cur.font = font;
          if (style) { cur.b = /bold/i.test(style); cur.i = /italic/i.test(style); }
          break;
        }
        default:
          if (/\d/.test(n)) { let num = n; while (/\d/.test(s[i])) num += s[i++]; push(); cur.sz = +num; }
          else cur.text += n;
      }
    }
    push();
    return out;
  };

  /* ------------------------------------------------------------ page rendering */
  const imgCache = new Map();
  function loadImg(src) {
    if (imgCache.has(src)) return imgCache.get(src);
    const pr = new Promise((res) => { const im = new Image(); im.onload = () => res(im); im.onerror = () => res(null); im.src = src; });
    imgCache.set(src, pr);
    return pr;
  }
  /** draw a block of cells with the grid renderer onto ctx at (x, y) */
  function drawCells(ctx, sh, rg, x, y, zoom, opts) {
    const g = LY.geo(sh);
    const w = Math.ceil((g.cols.pos(rg.c2 + 1) - g.cols.pos(rg.c1)) * zoom), hh = Math.ceil((g.rows.pos(rg.r2 + 1) - g.rows.pos(rg.r1)) * zoom);
    if (w <= 0 || hh <= 0) return { w: 0, h: 0 };
    const cv = document.createElement('canvas');
    const saved = { freeze: sh.view.freeze, grid: sh.view.grid, headings: sh.view.headings, top: sh.view.top };
    sh.view.freeze = null; sh.view.grid = !!opts.grid; sh.view.headings = !!opts.headings; sh.view.top = { r: 0, c: 0 };
    try {
      L.render.draw(cv, { sh, zoom, scrollR: rg.r1, scrollC: rg.c1, w: w + (opts.headings ? 0 : 0), h: hh, dpr: 1, headers: !!opts.headings }, { sel: null, hideSel: true, pages: null, print: true, bw: opts.bw });
    } finally { Object.assign(sh.view, saved); }
    ctx.drawImage(cv, x, y);
    return { w, h: hh };
  }
  /** a page as a canvas (pxPerPt pixels per point) */
  P.renderPage = async function (sh, pg, info, pxPerPt) {
    const pag = info.pag;
    const k = pxPerPt;
    const paper = pag.paper, m = pag.margins, p = sh.print;
    const cv = document.createElement('canvas');
    cv.width = Math.round(paper.w * k); cv.height = Math.round(paper.h * k);
    const ctx = cv.getContext('2d');
    ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, cv.width, cv.height);
    const zoom = pag.scale * k * 0.75; /* CSS px → canvas px */
    const g = LY.geo(sh);
    const titles = pag.titles || {};
    const repeatRows = titles.rows && pg.r1 > titles.rows.r2 ? { r1: titles.rows.r1, r2: titles.rows.r2 } : null;
    const repeatCols = titles.cols && pg.c1 > titles.cols.c2 ? { c1: titles.cols.c1, c2: titles.cols.c2 } : null;
    const tW = repeatCols ? (g.cols.pos(repeatCols.c2 + 1) - g.cols.pos(repeatCols.c1)) * zoom : 0;
    const tH = repeatRows ? (g.rows.pos(repeatRows.r2 + 1) - g.rows.pos(repeatRows.r1)) * zoom : 0;
    const bodyW = (g.cols.pos(pg.c2 + 1) - g.cols.pos(pg.c1)) * zoom + tW;
    const bodyH = (g.rows.pos(pg.r2 + 1) - g.rows.pos(pg.r1)) * zoom + tH;
    let x0 = m.l * 72 * k, y0 = m.t * 72 * k;
    if (p.hCenter) x0 = (cv.width - bodyW) / 2;
    if (p.vCenter) y0 = (cv.height - bodyH) / 2;
    const opts = { grid: p.gridLines, headings: false, bw: p.bw };
    if (repeatRows && repeatCols) drawCells(ctx, sh, { r1: repeatRows.r1, r2: repeatRows.r2, c1: repeatCols.c1, c2: repeatCols.c2 }, x0, y0, zoom, opts);
    if (repeatRows) drawCells(ctx, sh, { r1: repeatRows.r1, r2: repeatRows.r2, c1: pg.c1, c2: pg.c2 }, x0 + tW, y0, zoom, opts);
    if (repeatCols) drawCells(ctx, sh, { r1: pg.r1, r2: pg.r2, c1: repeatCols.c1, c2: repeatCols.c2 }, x0, y0 + tH, zoom, opts);
    drawCells(ctx, sh, pg, x0 + tW, y0 + tH, zoom, opts);
    /* pictures and charts on this page */
    for (const d of sh.drawings) {
      if (d.noPrint || d.hidden || !d.anchor || d.anchor.type === 'abs') continue;
      const a = d.anchor;
      const ax = g.cols.pos(a.from.c) + (a.from.cOff || 0) / 0.75, ay = g.rows.pos(a.from.r) + (a.from.rOff || 0) / 0.75;
      let aw, ah;
      if (a.to) { aw = g.cols.pos(a.to.c) + (a.to.cOff || 0) / 0.75 - ax; ah = g.rows.pos(a.to.r) + (a.to.rOff || 0) / 0.75 - ay; } else { aw = (a.w || 72) / 0.75; ah = (a.h || 72) / 0.75; }
      const px0 = g.cols.pos(pg.c1), py0 = g.rows.pos(pg.r1), px1 = g.cols.pos(pg.c2 + 1), py1 = g.rows.pos(pg.r2 + 1);
      if (ax + aw <= px0 || ax >= px1 || ay + ah <= py0 || ay >= py1) continue;
      const dx = x0 + tW + (ax - px0) * zoom, dy = y0 + tH + (ay - py0) * zoom, dw = aw * zoom, dh = ah * zoom;
      ctx.save();
      ctx.beginPath(); ctx.rect(x0 + tW, y0 + tH, bodyW - tW, bodyH - tH); ctx.clip();
      try {
        if (d.kind === 'image') {
          const mm = sh.wb.media.get(d.media);
          if (mm && !/emf|wmf/i.test(mm.ext)) { const url = URL.createObjectURL(new Blob([mm.bytes], { type: mm.type || 'image/png' })); const im = await loadImg(url); if (im) ctx.drawImage(im, dx, dy, dw, dh); }
        } else if (d.kind === 'chart' && d.chart) {
          const svg = L.xchart.render(d.chart, sh.wb, sh, aw, ah, 1);
          svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
          const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(new XMLSerializer().serializeToString(svg));
          const im = await loadImg(url);
          if (im) { ctx.drawImage(im, dx, dy, dw, dh); ctx.strokeStyle = '#000'; ctx.lineWidth = Math.max(1, zoom * 0.75); ctx.strokeRect(dx, dy, dw, dh); }
        } else if (d.kind === 'shape') {
          const el = L.drawing.element(d, sh);
          L.drawing.update(el, d, sh, { x: 0, y: 0, w: aw, h: ah }, false);
          const svg = el.querySelector('svg');
          if (svg) { svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg'); const im = await loadImg('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(new XMLSerializer().serializeToString(svg))); if (im) ctx.drawImage(im, dx, dy, dw, dh); }
          const t = d.text || (d._props && d._props.text);
          if (t) { ctx.fillStyle = '#000'; ctx.font = `${10 * zoom * 96 / 72}px Arial`; ctx.textBaseline = 'top'; String(t).split('\n').forEach((ln, i) => ctx.fillText(ln, dx + 4 * zoom, dy + 4 * zoom + i * 13 * zoom)); }
        }
      } catch (e) { console.warn('print object', e); }
      ctx.restore();
    }
    /* header and footer */
    const hfInfo = { page: info.page, pages: info.pages, file: (sh.wb.fileName || '').replace(/\.[^.]+$/, '') + (sh.wb.fileName && /\./.test(sh.wb.fileName) ? sh.wb.fileName.slice(sh.wb.fileName.lastIndexOf('.')) : ''), sheet: sh.name };
    const first = info.page === 1 && p.diffFirst, even = info.page % 2 === 0 && p.diffOddEven;
    const hdr = first ? p.firstHeader : even ? p.evenHeader : p.header;
    const ftr = first ? p.firstFooter : even ? p.evenFooter : p.footer;
    const drawHF = (code, y, bottom) => {
      if (!code) return;
      const secs = P.parseHF(code, hfInfo);
      for (const key of ['l', 'c', 'r']) {
        const runs = secs[key];
        if (!runs.length) continue;
        /* lines split on newlines inside runs */
        const lines = [[]];
        for (const r of runs) { const parts = r.text.split('\n'); parts.forEach((t, i) => { if (i) lines.push([]); if (t) lines[lines.length - 1].push(Object.assign({}, r, { text: t })); }); }
        const lh = 12 * k;
        lines.forEach((ln, li) => {
          const fonts = ln.map((r) => `${r.i ? 'italic ' : ''}${r.b ? 'bold ' : ''}${(r.sz || 10) * k}px ${LY.fontStack(r.font || 'Arial')}`);
          let tw = 0;
          ln.forEach((r, i) => { ctx.font = fonts[i]; tw += ctx.measureText(r.text).width; });
          let x = key === 'l' ? m.l * 72 * k : key === 'r' ? cv.width - m.r * 72 * k - tw : (cv.width - tw) / 2;
          const yy = bottom ? y - (lines.length - 1 - li) * lh : y + li * lh;
          ln.forEach((r, i) => { ctx.font = fonts[i]; ctx.fillStyle = '#000'; ctx.textBaseline = bottom ? 'alphabetic' : 'top'; ctx.fillText(r.text, x, yy); const w = ctx.measureText(r.text).width; if (r.u) ctx.fillRect(x, yy + (bottom ? 2 : (r.sz || 10) * k), w, Math.max(1, k * 0.6)); x += w; });
        });
      }
    };
    drawHF(hdr, (m.header || 0.5) * 72 * k, false);
    drawHF(ftr, cv.height - (m.footer || 0.5) * 72 * k, true);
    return cv;
  };
  /** every page of the given sheets: [{sh, pg, pag}] */
  P.allPages = function (sheets, selRange) {
    const out = [];
    for (const sh of sheets) {
      if (sh.kind === 'chartsheet') { out.push({ sh, chartsheet: true }); continue; }
      let pag = P.paginate(sh);
      if (selRange) {
        const saved = sh.print;
        const wb = sh.wb, idx = wb.sheets.indexOf(sh);
        const names = wb.names;
        wb.names = names.filter((n) => !(n.name.toLowerCase() === '_xlnm.print_area' && n.scope === idx)).concat([{ name: '_xlnm.Print_Area', ref: F.quoteSheet(sh.name) + '!' + F.absRangeName(selRange), scope: idx }]);
        try { pag = P.paginate(sh); } finally { wb.names = names; }
        void saved;
      }
      for (const pg of pag.pages) out.push({ sh, pg, pag });
    }
    return out;
  };
  async function renderAny(item, i, n, k) {
    if (item.chartsheet) {
      const sh = item.sh;
      const paper = P.paperSize(Object.assign({}, sh.print, { orientation: sh.print.orientation || 'landscape' }));
      const cv = document.createElement('canvas');
      cv.width = Math.round(paper.w * k); cv.height = Math.round(paper.h * k);
      const ctx = cv.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height);
      const d = sh.drawings.find((x) => x.kind === 'chart' && x.chart);
      if (d) {
        const m = sh.print.margins;
        const w = paper.w - (m.l + m.r) * 72, hh = paper.h - (m.t + m.b) * 72;
        const svg = L.xchart.render(d.chart, sh.wb, sh, w / 0.75, hh / 0.75, 1);
        svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
        const im = await loadImg('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(new XMLSerializer().serializeToString(svg)));
        if (im) ctx.drawImage(im, m.l * 72 * k, m.t * 72 * k, w * k, hh * k);
      }
      return { cv, w: paper.w, h: paper.h };
    }
    const cv = await P.renderPage(item.sh, item.pg, { page: i + 1, pages: n, pag: item.pag }, k);
    return { cv, w: item.pag.paper.w, h: item.pag.paper.h };
  }

  /* ------------------------------------------------------------ PDF */
  /**
   * opts: {what: 'sheets'|'selection'|'workbook', from, to, quick}
   */
  P.exportPDF = async function (opts) {
    opts = opts || {};
    const A = L.app, wb = G().wb;
    const sheets = opts.what === 'workbook' ? wb.sheets.filter((s) => s.state === 'visible') : A.selectedSheets();
    const sel = opts.what === 'selection' ? G().range() : null;
    const items = P.allPages(sheets, sel);
    if (!items.length) { ui.msg('Ledger did not find anything to print.', { icon: 'info' }); return; }
    const from = Math.max(1, opts.from || 1), to = Math.min(items.length, opts.to || items.length);
    ui.busy(true, 'Preparing pages...');
    try {
      const pages = [];
      for (let i = from - 1; i < to; i++) {
        const r = await renderAny(items[i], i, items.length, 2);
        const blob = await new Promise((res) => r.cv.toBlob(res, 'image/jpeg', 0.92));
        pages.push({ jpeg: new Uint8Array(await blob.arrayBuffer()), iw: r.cv.width, ih: r.cv.height, pw: r.w, ph: r.h });
      }
      const pdf = L.buildPDF(pages, { title: (wb.props && wb.props.title) || (A.books[A.cur] ? A.books[A.cur].name : 'Book1'), author: (wb.props && wb.props.creator) || '' });
      ui.busy(false);
      const name = (A.books[A.cur] ? A.books[A.cur].name : 'Book1').replace(/\.[^.]+$/, '') + '.pdf';
      const res = await L.saveFile(name, pdf);
      if (res === 'saved' && !opts.quiet) ui.toast(`${pages.length} page${pages.length === 1 ? '' : 's'} saved as ${name}. Open it to print.`);
    } catch (e) { ui.busy(false); A.error(e); }
    finally { ui.busy(false); }
  };

  /* ------------------------------------------------------------ Print Preview */
  P.preview = async function (opts) {
    opts = opts || {};
    const A = L.app, wb = G().wb;
    if (L.editor.active && !L.editor.commit()) return;
    let items = P.allPages(A.selectedSheets(), null);
    const ov = h('div', { class: 'pp-overlay', role: 'dialog', 'aria-label': 'Print Preview' });
    const bar = h('div', { class: 'pp-bar toolbar' });
    const view = h('div', { class: 'pp-view' });
    const status = h('div', { class: 'pp-status' });
    ov.append(bar, view, status);
    document.body.appendChild(ov);
    let i = 0, zoomed = false, showMargins = false;
    const btn = (label, fn, tip) => { const b = h('button', { type: 'button', class: 'tb-btn with-label', html: `<span class="tb-lbl">${L.esc(label).replace(/&amp;(\w)/, '<u>$1</u>')}</span>`, 'data-tip': tip || label.replace('&', '') }); b.addEventListener('click', fn); bar.appendChild(b); return b; };
    const bNext = btn('&Next', () => go(1));
    const bPrev = btn('&Previous', () => go(-1));
    btn('&Zoom', () => { zoomed = !zoomed; draw(); });
    bar.appendChild(h('span', { class: 'tb-sep' }));
    btn('&Print...', async () => { close(); await L.dlg.print(); });
    btn('&Setup...', async () => { await L.dlg.pageSetup(0); items = P.allPages(A.selectedSheets(), null); i = Math.min(i, items.length - 1); draw(); });
    btn('&Margins', () => { showMargins = !showMargins; draw(); });
    btn('Page Break Pre&view', () => { close(); A.setPageBreakPreview(true); });
    bar.appendChild(h('span', { class: 'tb-sep' }));
    btn('&Close', () => close());
    btn('&Help', () => { close(); L.panes.task.show('help'); });
    const keys = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); } else if (e.key === 'PageDown' || e.key === 'ArrowDown' || e.key === 'ArrowRight') { e.preventDefault(); e.stopPropagation(); go(1); } else if (e.key === 'PageUp' || e.key === 'ArrowUp' || e.key === 'ArrowLeft') { e.preventDefault(); e.stopPropagation(); go(-1); } };
    document.addEventListener('keydown', keys, true);
    function close() { document.removeEventListener('keydown', keys, true); ov.remove(); for (const s of A.selectedSheets()) P.computeBreaks(s); G().focus(); }
    function go(d) { const n = i + d; if (n < 0 || n >= items.length) return; i = n; draw(); }
    let token = 0;
    async function draw() {
      const my = ++token;
      bNext.disabled = i >= items.length - 1; bPrev.disabled = i <= 0;
      if (!items.length) { view.textContent = ''; view.appendChild(h('div', { class: 'pp-empty', text: 'Ledger did not find anything to print.' })); status.textContent = ''; return; }
      status.textContent = `Preview: Page ${i + 1} of ${items.length}`;
      const r = await renderAny(items[i], i, items.length, 1.6);
      if (my !== token) return;
      view.textContent = '';
      const vw = view.clientWidth - 40, vh = view.clientHeight - 40;
      const sc = zoomed ? 1.3 : Math.min(vw / r.cv.width, vh / r.cv.height);
      const page = h('div', { class: 'pp-page' + (zoomed ? ' zoomed' : '') });
      r.cv.style.width = r.cv.width * sc + 'px'; r.cv.style.height = r.cv.height * sc + 'px';
      page.appendChild(r.cv);
      if (showMargins && !items[i].chartsheet) {
        const m = items[i].pag.margins, k = 1.6 * sc;
        const lines = [['v', m.l * 72 * k], ['v', r.cv.width * sc - m.r * 72 * k], ['h', m.t * 72 * k], ['h', r.cv.height * sc - m.b * 72 * k], ['h', (m.header || 0.5) * 72 * k], ['h', r.cv.height * sc - (m.footer || 0.5) * 72 * k]];
        for (const [o, v] of lines) page.appendChild(h('div', { class: 'pp-margin ' + o, style: o === 'v' ? `left:${v}px` : `top:${v}px` }));
      }
      page.addEventListener('click', () => { zoomed = !zoomed; draw(); });
      view.appendChild(page);
    }
    draw();
    void wb; void opts;
  };
})(typeof window !== 'undefined' ? window : globalThis);
