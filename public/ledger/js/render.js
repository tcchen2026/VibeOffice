/* Ledger — worksheet renderer (canvas): headers, gridlines, fills, text, borders, conditional formats,
 * selection, frozen panes. All geometry is in CSS pixels; the canvas is scaled by devicePixelRatio. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, LY = L.layout, NF = L.numfmt, F = L.formula;
  const RD = (L.render = {});

  const COL = {
    grid: '#C0C0C0', hdrTop: '#FFFFFF', hdrBot: '#E4ECF7', hdrLine: '#9EB6CE', hdrText: '#000000',
    hdrSelTop: '#FFE7A2', hdrSelBot: '#F8C55E', hdrSelLine: '#D7A43C', hdrAllTop: '#FFD58D', hdrAllBot: '#F5B54A',
    sel: 'rgba(49, 74, 160, 0.20)', selBorder: '#000000', freeze: '#000000', pageBreak: '#0000FF',
  };
  RD.COL = COL;
  RD.HL = ['#0000FF', '#008000', '#9900CC', '#800000', '#00CC33', '#CC6600', '#CC0099'];

  /* ------------------------------------------------------------ fills */
  const PAT = {
    darkGray: ['1110', '1011', '1110', '1011'], mediumGray: ['10', '01'], lightGray: ['1000', '0010', '1000', '0010'],
    gray125: ['10000000', '00001000', '10000000', '00001000'], gray0625: ['1000000000000000', '0000000000000000', '0000000010000000', '0000000000000000'],
    darkHorizontal: ['1111', '1111', '0000', '0000'], darkVertical: ['1100', '1100', '1100', '1100'], darkDown: ['1100', '0110', '0011', '1001'], darkUp: ['0011', '0110', '1100', '1001'],
    darkGrid: ['1111', '1111', '1100', '1100'], darkTrellis: ['1111', '0110', '1111', '1001'],
    lightHorizontal: ['1111', '0000', '0000', '0000'], lightVertical: ['1000', '1000', '1000', '1000'], lightDown: ['1000', '0100', '0010', '0001'], lightUp: ['0001', '0010', '0100', '1000'],
    lightGrid: ['1111', '1000', '1000', '1000'], lightTrellis: ['1010', '0101', '1010', '0000'],
  };
  const patCache = new Map();
  function pattern(ctx, type, fg, bg) {
    const key = type + fg + bg;
    let p = patCache.get(key);
    if (p) return p;
    const rows = PAT[type];
    if (!rows || typeof document === 'undefined') return fg;
    const c = document.createElement('canvas');
    c.width = rows[0].length; c.height = rows.length;
    const x = c.getContext('2d');
    x.fillStyle = bg; x.fillRect(0, 0, c.width, c.height);
    x.fillStyle = fg;
    rows.forEach((row, i) => { for (let j = 0; j < row.length; j++) if (row[j] === '1') x.fillRect(j, i, 1, 1); });
    p = ctx.createPattern(c, 'repeat');
    patCache.set(key, p);
    return p;
  }
  /** canvas paint for a style fill at rect, or null */
  RD.fillPaint = function (ctx, wb, fill, x, y, w, h) {
    if (!fill) return null;
    if (fill.gradient) {
      const g = fill.gradient;
      let gr;
      if (g.type === 'path') {
        const cx = x + w * (((g.left || 0) + (g.right || 0)) / 2 || 0.5), cy = y + h * (((g.top || 0) + (g.bottom || 0)) / 2 || 0.5);
        gr = ctx.createRadialGradient(cx, cy, 0, cx, cy, Math.max(w, h) * 0.75);
      } else {
        const a = ((g.degree || 0) * Math.PI) / 180;
        const dx = Math.cos(a), dy = Math.sin(a);
        const half = (Math.abs(dx) * w + Math.abs(dy) * h) / 2;
        const cx = x + w / 2, cy = y + h / 2;
        gr = ctx.createLinearGradient(cx - dx * half, cy - dy * half, cx + dx * half, cy + dy * half);
      }
      for (const s of g.stops) gr.addColorStop(Math.max(0, Math.min(1, s.pos)), M.colorHex(wb, s.color, '#FFFFFF'));
      return gr;
    }
    const pt = fill.pattern;
    if (!pt || pt === 'none') return null;
    const fg = M.colorHex(wb, fill.fg, '#000000'), bg = M.colorHex(wb, fill.bg, '#FFFFFF');
    if (pt === 'solid') return fill.fg ? fg : M.colorHex(wb, fill.bg, '#000000');
    return pattern(ctx, pt, fg, bg);
  };

  /* ------------------------------------------------------------ borders */
  const BW = { hair: 1, thin: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
  const DASH = { hair: [1, 1], dotted: [1, 2], dashed: [3, 1], dashDot: [8, 3, 3, 3], dashDotDot: [8, 3, 3, 3, 3, 3], mediumDashed: [8, 3], mediumDashDot: [8, 3, 3, 3], mediumDashDotDot: [8, 3, 3, 3, 3, 3], slantDashDot: [10, 2, 4, 2] };
  RD.borderWidth = (b) => (b ? BW[b.style] || 1 : 0);
  /** draw a horizontal (h) or vertical edge line centred on (x1,y1)-(x2,y2) */
  function edge(ctx, wb, b, x1, y1, x2, y2, horiz, z) {
    const col = M.colorHex(wb, b.color, '#000000');
    const st = b.style;
    const w = BW[st] || 1;
    ctx.fillStyle = col;
    ctx.strokeStyle = col;
    if (st === 'double') {
      if (horiz) { ctx.fillRect(x1 - 1, y1 - 2, x2 - x1 + 2, 1); ctx.fillRect(x1 - 1, y1, x2 - x1 + 2, 1); }
      else { ctx.fillRect(x1 - 2, y1 - 1, 1, y2 - y1 + 2); ctx.fillRect(x1, y1 - 1, 1, y2 - y1 + 2); }
      return;
    }
    const dash = DASH[st];
    const off = w === 2 ? -1 : w === 3 ? -2 : -1;
    if (!dash) {
      if (horiz) ctx.fillRect(x1 - (w > 1 ? 1 : 0), y1 + off + (w === 3 ? 1 : 0), x2 - x1 + (w > 1 ? 2 : 0), w);
      else ctx.fillRect(x1 + off + (w === 3 ? 1 : 0), y1 - (w > 1 ? 1 : 0), w, y2 - y1 + (w > 1 ? 2 : 0));
      return;
    }
    ctx.save();
    ctx.lineWidth = w;
    ctx.setLineDash(dash.map((d) => d * Math.max(1, z)));
    ctx.beginPath();
    if (horiz) { const yy = y1 + off + w / 2 + (w === 3 ? 1 : 0); ctx.moveTo(x1, yy); ctx.lineTo(x2, yy); }
    else { const xx = x1 + off + w / 2 + (w === 3 ? 1 : 0); ctx.moveTo(xx, y1); ctx.lineTo(xx, y2); }
    ctx.stroke();
    ctx.restore();
  }
  RD.edge = edge;

  /* ------------------------------------------------------------ icon sets */
  const ICONSETS = {
    '3Arrows': [['arrow', 'down', '#C00000'], ['arrow', 'side', '#E6B200'], ['arrow', 'up', '#00A050']],
    '3ArrowsGray': [['arrow', 'down', '#808080'], ['arrow', 'side', '#808080'], ['arrow', 'up', '#808080']],
    '3Flags': [['flag', 0, '#C00000'], ['flag', 0, '#E6B200'], ['flag', 0, '#00A050']],
    '3TrafficLights1': [['dot', 0, '#C00000'], ['dot', 0, '#E6B200'], ['dot', 0, '#00A050']],
    '3TrafficLights2': [['light', 0, '#C00000'], ['light', 0, '#E6B200'], ['light', 0, '#00A050']],
    '3Signs': [['diamond', 0, '#C00000'], ['tri', 0, '#E6B200'], ['dot', 0, '#00A050']],
    '3Symbols': [['xmark', 0, '#C00000'], ['excl', 0, '#E6B200'], ['check', 0, '#00A050']],
    '3Symbols2': [['xmark', 1, '#C00000'], ['excl', 1, '#E6B200'], ['check', 1, '#00A050']],
    '3Stars': [['star', 0, '#C9C9C9'], ['star', 0.5, '#E6B200'], ['star', 1, '#E6B200']],
    '3Triangles': [['tri', 'down', '#C00000'], ['dash', 0, '#E6B200'], ['tri', 'up', '#00A050']],
    '4Arrows': [['arrow', 'down', '#C00000'], ['arrow', 'downside', '#E6B200'], ['arrow', 'upside', '#E6B200'], ['arrow', 'up', '#00A050']],
    '4ArrowsGray': [['arrow', 'down', '#808080'], ['arrow', 'downside', '#808080'], ['arrow', 'upside', '#808080'], ['arrow', 'up', '#808080']],
    '4RedToBlack': [['dot', 0, '#404040'], ['dot', 0, '#808080'], ['dot', 0, '#E08080'], ['dot', 0, '#C00000']],
    '4Rating': [['bars', 1, '#1F5DA8'], ['bars', 2, '#1F5DA8'], ['bars', 3, '#1F5DA8'], ['bars', 4, '#1F5DA8']],
    '4TrafficLights': [['dot', 0, '#404040'], ['dot', 0, '#C00000'], ['dot', 0, '#E6B200'], ['dot', 0, '#00A050']],
    '5Arrows': [['arrow', 'down', '#C00000'], ['arrow', 'downside', '#E6B200'], ['arrow', 'side', '#E6B200'], ['arrow', 'upside', '#E6B200'], ['arrow', 'up', '#00A050']],
    '5ArrowsGray': [['arrow', 'down', '#808080'], ['arrow', 'downside', '#808080'], ['arrow', 'side', '#808080'], ['arrow', 'upside', '#808080'], ['arrow', 'up', '#808080']],
    '5Rating': [['bars', 0, '#1F5DA8'], ['bars', 1, '#1F5DA8'], ['bars', 2, '#1F5DA8'], ['bars', 3, '#1F5DA8'], ['bars', 4, '#1F5DA8']],
    '5Quarters': [['quarter', 0, '#404040'], ['quarter', 1, '#404040'], ['quarter', 2, '#404040'], ['quarter', 3, '#404040'], ['quarter', 4, '#404040']],
    '3Smilies': [['face', -1, '#C00000'], ['face', 0, '#E6B200'], ['face', 1, '#00A050']],
    '5Boxes': [['boxes', 0, '#1F5DA8'], ['boxes', 1, '#1F5DA8'], ['boxes', 2, '#1F5DA8'], ['boxes', 3, '#1F5DA8'], ['boxes', 4, '#1F5DA8']],
  };
  RD.ICONSETS = ICONSETS;
  RD.drawIcon = function (ctx, set, i, x, y, s) {
    if (set === 'NoIcons') return;
    const def = (ICONSETS[set] || ICONSETS['3TrafficLights1'])[i] || ['dot', 0, '#808080'];
    const [kind, arg, col] = def;
    const cx = x + s / 2, cy = y + s / 2, r = s / 2 - 1;
    ctx.save();
    ctx.fillStyle = col; ctx.strokeStyle = col; ctx.lineWidth = Math.max(1, s / 8);
    ctx.beginPath();
    switch (kind) {
      case 'arrow': {
        const ang = { up: -90, upside: -45, side: 0, downside: 45, down: 90 }[arg] * Math.PI / 180;
        ctx.translate(cx, cy); ctx.rotate(ang);
        ctx.moveTo(r, 0); ctx.lineTo(0, -r); ctx.lineTo(0, -r * 0.4); ctx.lineTo(-r, -r * 0.4); ctx.lineTo(-r, r * 0.4); ctx.lineTo(0, r * 0.4); ctx.lineTo(0, r); ctx.closePath(); ctx.fill();
        break;
      }
      case 'dot': case 'light': ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fill(); if (kind === 'light') { ctx.strokeStyle = '#333'; ctx.lineWidth = 1; ctx.stroke(); } break;
      case 'diamond': ctx.moveTo(cx, y + 1); ctx.lineTo(x + s - 1, cy); ctx.lineTo(cx, y + s - 1); ctx.lineTo(x + 1, cy); ctx.closePath(); ctx.fill(); break;
      case 'tri': if (arg === 'down') { ctx.moveTo(x + 1, y + 3); ctx.lineTo(x + s - 1, y + 3); ctx.lineTo(cx, y + s - 2); } else { ctx.moveTo(cx, y + 2); ctx.lineTo(x + s - 1, y + s - 3); ctx.lineTo(x + 1, y + s - 3); } ctx.closePath(); ctx.fill(); break;
      case 'dash': ctx.fillRect(x + 2, cy - s / 8, s - 4, s / 4); break;
      case 'flag': ctx.fillRect(x + 2, y + 1, 1.5, s - 2); ctx.moveTo(x + 3.5, y + 1); ctx.lineTo(x + s - 1, y + s / 4 + 1); ctx.lineTo(x + 3.5, y + s / 2 + 1); ctx.closePath(); ctx.fill(); break;
      case 'xmark': case 'excl': case 'check':
        if (!arg) { ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fill(); ctx.strokeStyle = '#fff'; }
        ctx.lineWidth = Math.max(1.5, s / 7);
        ctx.beginPath();
        if (kind === 'xmark') { ctx.moveTo(cx - r / 2, cy - r / 2); ctx.lineTo(cx + r / 2, cy + r / 2); ctx.moveTo(cx + r / 2, cy - r / 2); ctx.lineTo(cx - r / 2, cy + r / 2); }
        else if (kind === 'excl') { ctx.moveTo(cx, cy - r / 2); ctx.lineTo(cx, cy + r / 6); ctx.moveTo(cx, cy + r / 2.2); ctx.lineTo(cx, cy + r / 1.8); }
        else { ctx.moveTo(cx - r / 2, cy); ctx.lineTo(cx - r / 8, cy + r / 2.5); ctx.lineTo(cx + r / 2, cy - r / 2.5); }
        ctx.stroke();
        break;
      case 'star': {
        const pts = [];
        for (let k = 0; k < 10; k++) { const a = -Math.PI / 2 + (k * Math.PI) / 5, rr = k % 2 ? r * 0.45 : r; pts.push([cx + rr * Math.cos(a), cy + rr * Math.sin(a)]); }
        ctx.moveTo(pts[0][0], pts[0][1]); pts.forEach((p) => ctx.lineTo(p[0], p[1])); ctx.closePath();
        ctx.strokeStyle = '#B08A00'; ctx.lineWidth = 1;
        if (arg === 1) ctx.fill(); else if (arg === 0.5) { ctx.save(); ctx.clip(); ctx.fillRect(x, y, s / 2, s); ctx.restore(); }
        ctx.stroke();
        break;
      }
      case 'bars': for (let k = 0; k < 4; k++) { const hh = (s - 2) * (k + 1) / 4; ctx.fillStyle = k < arg ? col : '#C9C9C9'; ctx.fillRect(x + 1 + k * (s / 4), y + s - 1 - hh, s / 4 - 1, hh); } break;
      case 'boxes': for (let k = 0; k < 4; k++) { ctx.fillStyle = k < arg ? col : '#C9C9C9'; ctx.fillRect(x + 1 + (k % 2) * (s / 2), y + 1 + Math.floor(k / 2) * (s / 2), s / 2 - 2, s / 2 - 2); } break;
      case 'quarter': ctx.strokeStyle = col; ctx.lineWidth = 1; ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.stroke(); if (arg) { ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + (Math.PI * 2 * arg) / 4); ctx.closePath(); ctx.fill(); } break;
      case 'face': ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fill(); ctx.fillStyle = '#000'; ctx.fillRect(cx - r / 2.5, cy - r / 3, 1.5, 1.5); ctx.fillRect(cx + r / 3.5, cy - r / 3, 1.5, 1.5); ctx.strokeStyle = '#000'; ctx.lineWidth = 1; ctx.beginPath(); if (arg > 0) ctx.arc(cx, cy, r / 2, 0.2 * Math.PI, 0.8 * Math.PI); else if (arg < 0) ctx.arc(cx, cy + r / 1.4, r / 2, 1.2 * Math.PI, 1.8 * Math.PI); else { ctx.moveTo(cx - r / 2.5, cy + r / 3); ctx.lineTo(cx + r / 2.5, cy + r / 3); } ctx.stroke(); break;
      default: ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fill();
    }
    ctx.restore();
  };

  /* ------------------------------------------------------------ view geometry */
  /**
   * Compute panes for a view: {hdrW, hdrH, panes:[{r0, r1, c0, c1, x, y, w, h, kind}]}
   * view: {sh, zoom, scrollR, scrollC, w, h, headers}
   */
  RD.frame = function (view) {
    const sh = view.sh, z = view.zoom, g = LY.geo(sh);
    const headers = view.headers !== false && sh.view.headings !== false;
    const hdrFont = LY.cssFont({ name: 'Arial', sz: 10 }, z);
    let hdrW = 0, hdrH = 0;
    if (headers) {
      hdrH = Math.max(Math.round(g.defRow * z), Math.round(17 * z));
      const lastRow = Math.min(M.MAXR, g.rows.at(g.rows.pos(view.scrollR) + view.h / z) + 2);
      hdrW = Math.max(Math.round(25 * z), Math.ceil(LY.measure(hdrFont, String(lastRow)) + 10 * z));
      if (sh.outline && (sh.outline.levelRow || sh.outline.levelCol)) { /* outline gutters handled by the grid controller */ }
    }
    const fz = sh.view.freeze && (sh.view.freeze.r || sh.view.freeze.c) ? sh.view.freeze : null;
    const oR = fz && sh.view.top ? sh.view.top.r : 0, oC = fz && sh.view.top ? sh.view.top.c : 0;
    const fr = fz ? fz.r : 0, fc = fz ? fz.c : 0;
    const fzH = fr ? Math.round((g.rows.pos(oR + fr) - g.rows.pos(oR)) * z) : 0;
    const fzW = fc ? Math.round((g.cols.pos(oC + fc) - g.cols.pos(oC)) * z) : 0;
    const x0 = hdrW, y0 = hdrH;
    const W = view.w, H = view.h;
    const sR = Math.max(view.scrollR, oR + fr), sC = Math.max(view.scrollC, oC + fc);
    const panes = [];
    const mk = (r0, rMax, c0, cMax, x, y, w, h, kind) => { if (w > 0 && h > 0) panes.push({ r0, rMax, c0, cMax, x, y, w, h, kind }); };
    if (fr && fc) mk(oR, oR + fr, oC, oC + fc, x0, y0, Math.min(fzW, W - x0), Math.min(fzH, H - y0), 'corner');
    if (fr) mk(oR, oR + fr, sC, M.MAXC, x0 + fzW, y0, W - x0 - fzW, Math.min(fzH, H - y0), 'top');
    if (fc) mk(sR, M.MAXR, oC, oC + fc, x0, y0 + fzH, Math.min(fzW, W - x0), H - y0 - fzH, 'left');
    mk(sR, M.MAXR, sC, M.MAXC, x0 + fzW, y0 + fzH, W - x0 - fzW, H - y0 - fzH, 'main');
    for (const p of panes) {
      /* last visible row / column */
      p.rEnd = Math.min(p.rMax - 1, g.rows.at(g.rows.pos(p.r0) + p.h / z));
      p.cEnd = Math.min(p.cMax - 1, g.cols.at(g.cols.pos(p.c0) + p.w / z));
    }
    return { hdrW, hdrH, panes, g, z, fzW, fzH, oR, oC, fr, fc, sR, sC, hdrFont };
  };
  /** x of column c's left edge inside pane p (canvas coords) */
  const colX = (fr, p, c) => p.x + Math.round((fr.g.cols.pos(c) - fr.g.cols.pos(p.c0)) * fr.z);
  const rowY = (fr, p, r) => p.y + Math.round((fr.g.rows.pos(r) - fr.g.rows.pos(p.r0)) * fr.z);
  RD.colX = colX; RD.rowY = rowY;
  /** pane that shows cell (r, c), or null */
  RD.paneOf = function (fr, r, c) {
    for (const p of fr.panes) if (r >= p.r0 && r < p.rMax && c >= p.c0 && c < p.cMax) return p;
    return null;
  };
  /** canvas rect of a cell range in the pane that shows its top-left (clipped to that pane's region) */
  RD.rangeRect = function (fr, rg, pane) {
    const p = pane || RD.paneOf(fr, Math.max(rg.r1, fr.panes[0] ? 0 : 0), rg.c1) || fr.panes[fr.panes.length - 1];
    const x1 = colX(fr, p, rg.c1), x2 = colX(fr, p, rg.c2 + 1), y1 = rowY(fr, p, rg.r1), y2 = rowY(fr, p, rg.r2 + 1);
    return { x: x1, y: y1, w: x2 - x1, h: y2 - y1, p };
  };
  /** hit test canvas point → {area: 'cell'|'colhdr'|'rowhdr'|'corner'|'none', r, c, pane} */
  RD.hit = function (fr, x, y) {
    const g = fr.g, z = fr.z;
    if (x < fr.hdrW && y < fr.hdrH) return { area: 'corner' };
    const findPane = (px, py) => fr.panes.find((p) => px >= p.x && px < p.x + p.w && py >= p.y && py < p.y + p.h);
    if (y < fr.hdrH) {
      const p = findPane(Math.max(x, fr.hdrW), fr.hdrH + 1) || fr.panes.find((q) => x >= q.x && x < q.x + q.w);
      if (!p) return { area: 'none' };
      const c = g.cols.at(g.cols.pos(p.c0) + (x - p.x) / z);
      return { area: 'colhdr', c, pane: p };
    }
    if (x < fr.hdrW) {
      const p = findPane(fr.hdrW + 1, y) || fr.panes.find((q) => y >= q.y && y < q.y + q.h);
      if (!p) return { area: 'none' };
      const r = g.rows.at(g.rows.pos(p.r0) + (y - p.y) / z);
      return { area: 'rowhdr', r, pane: p };
    }
    const p = findPane(x, y);
    if (!p) return { area: 'none' };
    const c = g.cols.at(g.cols.pos(p.c0) + (x - p.x) / z), r = g.rows.at(g.rows.pos(p.r0) + (y - p.y) / z);
    return { area: 'cell', r: Math.min(r, p.rMax - 1), c: Math.min(c, p.cMax - 1), pane: p };
  };

  /* ------------------------------------------------------------ text */
  function fontOf(wb, st, dxf) {
    let f = st.font || {};
    /* a font without a name or size takes the Normal style's */
    if (f.name == null || f.sz == null) { const d = wb.defaultFont || {}; f = Object.assign({}, f, { name: f.name == null ? d.name : f.name, sz: f.sz == null ? d.sz : f.sz }); }
    if (dxf && dxf.font) f = Object.assign({}, f, dxf.font);
    return f;
  }
  function textColor(wb, f, fmtColor, dxf) {
    if (dxf && dxf.font && dxf.font.color) return M.colorHex(wb, dxf.font.color, '#000000');
    if (fmtColor) return fmtColor;
    return M.colorHex(wb, f.color, '#000000');
  }
  /* wrap text into lines that fit width (word wrap, breaking long words) */
  function wrapLines(font, text, width) {
    const out = [];
    for (const para of String(text).split(/\r?\n/)) {
      if (!para) { out.push(''); continue; }
      const words = para.split(/(\s+)/);
      let line = '';
      for (const w of words) {
        const cand = line + w;
        if (LY.measure(font, cand) <= width || !line) {
          if (!line && LY.measure(font, w) > width && w.trim()) {
            /* break a long word */
            let part = '';
            for (const ch of w) { if (LY.measure(font, part + ch) > width && part) { out.push(part); part = ch; } else part += ch; }
            line = part;
          } else line = cand;
        } else { out.push(line.replace(/\s+$/, '')); line = w.trim() ? w : ''; }
      }
      out.push(line.replace(/\s+$/, ''));
    }
    return out;
  }
  RD.wrapLines = wrapLines;

  /* ------------------------------------------------------------ main draw */
  /**
   * Draw the sheet. state: {sel:{ranges, r, c, active}, editing:{r,c}, marquee:{range, phase}, hl:[{range, color}],
   *   fillTo: range, dragTo: range, page: {rowBreaks, colBreaks}, focus}
   */
  RD.draw = function (canvas, view, state) {
    const dpr = view.dpr || 1;
    const W = view.w, H = view.h;
    if (canvas.width !== Math.round(W * dpr) || canvas.height !== Math.round(H * dpr)) {
      canvas.width = Math.round(W * dpr); canvas.height = Math.round(H * dpr);
      canvas.style.width = W + 'px'; canvas.style.height = H + 'px';
    }
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.textBaseline = 'alphabetic';
    const fr = RD.frame(view);
    view.frame = fr;
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, W, H);
    for (const p of fr.panes) drawPane(ctx, view, fr, p, state);
    drawSelection(ctx, view, fr, state);
    /* frozen pane dividers */
    ctx.fillStyle = COL.freeze;
    if (fr.fzH) ctx.fillRect(fr.hdrW, fr.hdrH + fr.fzH - 1, W - fr.hdrW, 1);
    if (fr.fzW) ctx.fillRect(fr.hdrW + fr.fzW - 1, fr.hdrH, 1, H - fr.hdrH);
    if (fr.hdrW) drawHeaders(ctx, view, fr, state);
    return fr;
  };

  function drawPane(ctx, view, fr, p, state) {
    const sh = view.sh, wb = sh.wb, z = fr.z, g = fr.g;
    ctx.save();
    ctx.beginPath(); ctx.rect(p.x, p.y, p.w, p.h); ctx.clip();
    const r0 = p.r0, r1 = p.rEnd, c0 = p.c0, c1 = p.cEnd;
    /* column x positions / row y positions in the pane */
    const xs = [], ys = [];
    for (let c = c0; c <= c1 + 1; c++) xs.push(colX(fr, p, c));
    for (let r = r0; r <= r1 + 1; r++) ys.push(rowY(fr, p, r));
    const X = (c) => (c >= c0 && c <= c1 + 1 ? xs[c - c0] : colX(fr, p, c));
    const Y = (r) => (r >= r0 && r <= r1 + 1 ? ys[r - r0] : rowY(fr, p, r));
    const showGrid = sh.view.grid !== false;
    /* gridlines */
    if (showGrid) {
      ctx.fillStyle = sh.view.gridColor ? M.colorHex(wb, sh.view.gridColor, COL.grid) : COL.grid;
      for (let c = c0; c <= c1; c++) { const x = xs[c - c0 + 1] - 1; if (xs[c - c0 + 1] > xs[c - c0]) ctx.fillRect(x, p.y, 1, p.h); }
      for (let r = r0; r <= r1; r++) { const y = ys[r - r0 + 1] - 1; if (ys[r - r0 + 1] > ys[r - r0]) ctx.fillRect(p.x, y, p.w, 1); }
    }
    const mi = LY.mergeIndex(sh);
    const tstyle = L.tstyle && wb.tables.length ? L.tstyle : null;
    /* collect cells: stored cells plus styled rows / columns */
    const items = [];
    const doneMerge = new Set();
    const colStyled = [];
    for (let c = c0; c <= c1; c++) { const o = sh.cols[c] || (sh.colTail && c >= sh.colTail.from ? sh.colTail.o : null); if (o && o.s) colStyled.push(c); }
    const rowRange = (r) => sh.rows[r];
    /* merged areas that start above / left of the viewport */
    for (const m of sh.merges) {
      if (m.r2 < r0 || m.r1 > r1 || m.c2 < c0 || m.c1 > c1) continue;
      if (m.r1 < r0 || m.c1 < c0) items.push({ r: m.r1, c: m.c1, cell: sh.get(m.r1, m.c1), m });
    }
    for (let r = r0; r <= r1; r++) {
      if (ys[r - r0 + 1] === ys[r - r0]) continue; /* hidden row */
      const row = rowRange(r);
      const cells = row ? row.cells : null;
      const seen = new Set();
      const consider = (c, cell) => {
        if (xs[c - c0 + 1] === xs[c - c0] && !(cell && cell.v != null)) return;
        const m = mi.byCell.get(M.key(r, c));
        if (m) { if (m.r1 === r && m.c1 === c) items.push({ r, c, cell, m }); return; }
        items.push({ r, c, cell });
      };
      if (cells) {
        if (c1 - c0 > 60) { for (const k in cells) { const c = +k; if (c >= c0 && c <= c1) { seen.add(c); consider(c, cells[c]); } } }
        else for (let c = c0; c <= c1; c++) { const cell = cells[c]; if (cell) { seen.add(c); consider(c, cell); } }
      }
      if (row && row.s) { for (let c = c0; c <= c1; c++) if (!seen.has(c)) { seen.add(c); consider(c, null); } }
      else for (const c of colStyled) if (!seen.has(c)) consider(c, null);
      if (tstyle) for (const t of wb.tables) if (t.sheet === sh && r >= t.ref.r1 && r <= t.ref.r2) for (let c = Math.max(c0, t.ref.c1); c <= Math.min(c1, t.ref.c2); c++) if (!seen.has(c)) { seen.add(c); consider(c, null); }
      if (sh.cf.length) for (const cf of sh.cf) for (const rg of cf.ranges) if (r >= rg.r1 && r <= rg.r2) for (let c = Math.max(c0, rg.c1); c <= Math.min(c1, rg.c2); c++) if (!seen.has(c)) { seen.add(c); consider(c, null); }
    }
    /* pass 1: fills (and erase gridlines under merges) */
    const texts = [];
    for (const it of items) {
      const { r, c, cell, m } = it;
      if (m) { const k = m.r1 * 16384 + m.c1; if (doneMerge.has(k)) continue; doneMerge.add(k); }
      const st = LY.styleOf(sh, r, c, cell);
      const cf = sh.cf.length ? L.cf.at(sh, r, c, cell) : null;
      const ts = tstyle ? tstyle.at(sh, r, c) : null;
      it.st = st; it.cf = cf; it.ts = ts;
      const x = m ? X(m.c1) : X(c), y = m ? Y(m.r1) : Y(r);
      const w = (m ? X(m.c2 + 1) : X(c + 1)) - x, h = (m ? Y(m.r2 + 1) : Y(r + 1)) - y;
      it.x = x; it.y = y; it.w = w; it.h = h;
      if (w <= 0 || h <= 0) continue;
      let fill = st.fill;
      if (ts && ts.fill && !fill) fill = ts.fill;
      if (cf && cf.dxf && cf.dxf.fill) fill = cf.dxf.fill;
      let paint = fill ? RD.fillPaint(ctx, wb, fill, x, y, w, h) : null;
      if (cf && cf.scale) paint = cf.scale;
      if (paint) { ctx.fillStyle = paint; ctx.fillRect(x - 1, y - 1, w + 1, h + 1); it.filled = true; }
      else if (m) { ctx.fillStyle = '#FFFFFF'; ctx.fillRect(x, y, w - 1, h - 1); }
      if (cf && cf.bar) drawBar(ctx, cf.bar, x, y, w, h, z);
      if (cell && (cell.v != null || cell.dirty)) texts.push(it);
    }
    /* gridlines over filled cells' neighbours stay hidden; redraw grid lines between unfilled neighbours is unnecessary */
    /* pass 2: text */
    const rowsWithText = new Map();
    for (const it of texts) { let l = rowsWithText.get(it.r); if (!l) { l = []; rowsWithText.set(it.r, l); } l.push(it); }
    for (const [r, list] of rowsWithText) {
      list.sort((a, b) => a.c - b.c);
      for (const it of list) drawText(ctx, view, fr, p, it, X, Y, state);
    }
    /* overflow text from cells left / right of the viewport */
    drawOffscreenOverflow(ctx, view, fr, p, X, Y, state, r0, r1, c0, c1);
    /* pass 3: borders */
    for (const it of items) {
      if (it.w <= 0 || it.h <= 0 || !it.st) continue;
      let b = it.st.border;
      if (it.ts && it.ts.border) b = Object.assign({}, it.ts.border, b || {});
      if (it.cf && it.cf.dxf && it.cf.dxf.border) b = Object.assign({}, b || {}, it.cf.dxf.border);
      if (!b) continue;
      const { x, y, w, h } = it;
      const right = x + w, bottom = y + h;
      if (b.t) edge(ctx, wb, b.t, x - 1, y, right, y, true, z);
      if (b.b) edge(ctx, wb, b.b, x - 1, bottom, right, bottom, true, z);
      if (b.l) edge(ctx, wb, b.l, x, y - 1, x, bottom, false, z);
      if (b.r) edge(ctx, wb, b.r, right, y - 1, right, bottom, false, z);
      if (b.d) {
        ctx.save();
        ctx.strokeStyle = M.colorHex(wb, b.d.color, '#000000');
        ctx.lineWidth = RD.borderWidth(b.d);
        ctx.beginPath();
        if (b.du) { ctx.moveTo(x, bottom - 1); ctx.lineTo(right - 1, y); }
        if (b.dd) { ctx.moveTo(x, y); ctx.lineTo(right - 1, bottom - 1); }
        ctx.stroke();
        ctx.restore();
      }
    }
    /* pass 3b: sparklines (Excel 2010 workbooks) */
    if (sh.sparklines && sh.sparklines.length) drawSparklines(ctx, sh, X, Y, r0, r1, c0, c1, z);
    /* pass 4: comment indicators */
    if (sh.comments.size) {
      ctx.fillStyle = '#FF0000';
      for (const cm of sh.comments.values()) {
        if (cm.r < r0 || cm.r > r1 || cm.c < c0 || cm.c > c1) continue;
        const m = mi.byCell.get(M.key(cm.r, cm.c));
        const right = m ? X(m.c2 + 1) : X(cm.c + 1), top = m ? Y(m.r1) : Y(cm.r);
        if (right <= X(cm.c)) continue;
        const s = Math.max(4, Math.round(5 * z));
        ctx.beginPath(); ctx.moveTo(right - 1 - s, top); ctx.lineTo(right - 1, top); ctx.lineTo(right - 1, top + s); ctx.closePath(); ctx.fill();
      }
    }
    /* page breaks (after print preview / page setup, or in page break preview) */
    if (state.pages) drawPageBreaks(ctx, view, fr, p, state.pages, X, Y, r0, r1, c0, c1);
    ctx.restore();
  }

  /* ------------------------------------------------------------ sparklines */
  function sparkValues(sh, g, it) {
    const XC = L.xchart;
    if (!XC || !XC.evalRef) return null;
    const res = XC.evalRef(it.f, sh.wb, sh);
    if (!res) return null;
    return res.vals.map((v) => (typeof v === 'number' ? v : typeof v === 'boolean' ? null : v == null || v === '' ? (g.displayEmptyCellsAs === 'zero' ? 0 : null) : null));
  }
  function drawSparklines(ctx, sh, X, Y, r0, r1, c0, c1, z) {
    const wb = sh.wb;
    const hex = (c, d) => (c ? M.colorHex(wb, c, d) : d);
    for (const g of sh.sparklines) {
      /* group scaling needs every member's values */
      let gMin = Infinity, gMax = -Infinity;
      const groupAxis = g.minAxisType === 'group' || g.maxAxisType === 'group';
      const members = [];
      for (const it of g.items) {
        if (!it._at) { const p = F.parseCell(String(it.sqref || '').replace(/\$/g, '')); it._at = p || { r: -1, c: -1 }; }
        const at = it._at;
        const visible = at.r >= r0 && at.r <= r1 && at.c >= c0 && at.c <= c1;
        if (!visible && !groupAxis) continue;
        const vals = sparkValues(sh, g, it);
        if (!vals) continue;
        if (g.rightToLeft) vals.reverse();
        members.push({ at, vals, visible });
        for (const v of vals) if (v != null) { if (v < gMin) gMin = v; if (v > gMax) gMax = v; }
      }
      for (const m of members) {
        if (!m.visible) continue;
        const { at, vals } = m;
        const x0 = X(at.c), y0 = Y(at.r), w = X(at.c + 1) - x0, h = Y(at.r + 1) - y0;
        if (w < 4 || h < 4) continue;
        const nums = vals.filter((v) => v != null);
        if (!nums.length) continue;
        let lo = Math.min(...nums), hi = Math.max(...nums);
        if (g.minAxisType === 'group') lo = gMin; else if (g.minAxisType === 'custom' && g.manualMin != null) lo = g.manualMin;
        if (g.maxAxisType === 'group') hi = gMax; else if (g.maxAxisType === 'custom' && g.manualMax != null) hi = g.manualMax;
        if (g.type === 'column' && g.minAxisType !== 'custom') { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
        const padX = Math.max(2, Math.round(3 * z)), padY = Math.max(2, Math.round(3 * z));
        const L0 = x0 + padX, T0 = y0 + padY, W = w - 2 * padX - 1, H = h - 2 * padY - 1;
        if (W < 2 || H < 2) continue;
        const span = hi - lo || 1;
        const yOf = (v) => T0 + H - ((Math.min(hi, Math.max(lo, v)) - lo) / span) * H;
        const n = vals.length;
        const iMax = vals.indexOf(Math.max(...nums)), iMin = vals.indexOf(Math.min(...nums));
        let iFirst = vals.findIndex((v) => v != null), iLast = n - 1; while (iLast > 0 && vals[iLast] == null) iLast--;
        const pointColor = (i, v, base) => {
          if (g.high && i === iMax && g.highColor) return hex(g.highColor, base);
          if (g.low && i === iMin && g.lowColor) return hex(g.lowColor, base);
          if (g.first && i === iFirst && g.firstColor) return hex(g.firstColor, base);
          if (g.last && i === iLast && g.lastColor) return hex(g.lastColor, base);
          if (g.negative && v < 0 && g.negColor) return hex(g.negColor, base);
          return null;
        };
        const series = hex(g.color, '#376092');
        ctx.save();
        ctx.beginPath(); ctx.rect(x0, y0, w - 1, h - 1); ctx.clip();
        if (g.displayXAxis && lo < 0 && hi > 0) { ctx.fillStyle = hex(g.axisColor, '#000000'); ctx.fillRect(L0, Math.round(yOf(0)), W, 1); }
        if (g.type === 'line') {
          const xOf = (i) => (n === 1 ? L0 + W / 2 : L0 + (i * W) / (n - 1));
          ctx.strokeStyle = series;
          ctx.lineWidth = Math.max(0.75, ((g.lineWeight || 0.75) * 96 / 72) * z);
          ctx.lineJoin = 'round';
          ctx.beginPath();
          let pen = false;
          for (let i = 0; i < n; i++) {
            const v = vals[i];
            if (v == null) { if (g.displayEmptyCellsAs !== 'span') pen = false; continue; }
            const px = xOf(i), py = yOf(v);
            if (pen) ctx.lineTo(px, py); else ctx.moveTo(px, py);
            pen = true;
          }
          ctx.stroke();
          const ms = Math.max(2, Math.round(2 * z));
          for (let i = 0; i < n; i++) {
            const v = vals[i];
            if (v == null) continue;
            const c = pointColor(i, v, null) || (g.markers ? hex(g.markerColor, series) : null);
            if (!c) continue;
            ctx.fillStyle = c;
            ctx.fillRect(Math.round(xOf(i) - ms), Math.round(yOf(v) - ms), ms * 2, ms * 2);
          }
        } else {
          const slot = W / n, bw = Math.max(1, slot * 0.8);
          for (let i = 0; i < n; i++) {
            const v = vals[i];
            if (v == null) continue;
            const bx = L0 + i * slot + (slot - bw) / 2;
            let top, bot;
            if (g.type === 'stacked') { if (v === 0) continue; const mid = T0 + H / 2; if (v > 0) { top = T0; bot = mid; } else { top = mid; bot = T0 + H; } }
            else { const base = lo > 0 ? lo : hi < 0 ? hi : 0; top = Math.min(yOf(v), yOf(base)); bot = Math.max(yOf(v), yOf(base)); if (bot - top < 1) bot = top + 1; }
            ctx.fillStyle = pointColor(i, v, null) || (v < 0 && g.type === 'stacked' ? hex(g.negColor, series) : series);
            ctx.fillRect(bx, top, bw, bot - top);
          }
        }
        ctx.restore();
      }
    }
  }
  RD.drawSparklines = drawSparklines;

  function drawBar(ctx, bar, x, y, w, h, z) {
    const pad = Math.max(1, Math.round(2 * z));
    const bh = h - pad * 2 - 1, bw = w - pad * 2 - 1;
    if (bh <= 0 || bw <= 0) return;
    let bx = x + pad, len = Math.max(0, bw * bar.frac);
    if (bar.axis != null) {
      const ax = x + pad + bw * bar.axis;
      if (bar.neg) { bx = ax - len; } else bx = ax;
      ctx.save(); ctx.setLineDash([2, 1]); ctx.strokeStyle = '#000'; ctx.beginPath(); ctx.moveTo(ax + 0.5, y + 1); ctx.lineTo(ax + 0.5, y + h - 1); ctx.stroke(); ctx.restore();
    }
    const col = bar.neg ? bar.negColor : bar.color;
    if (bar.gradient) {
      const gr = ctx.createLinearGradient(bx, 0, bx + Math.max(1, len), 0);
      if (bar.neg && bar.axis != null) { gr.addColorStop(0, '#FFFFFF'); gr.addColorStop(1, col); } else { gr.addColorStop(0, col); gr.addColorStop(1, '#FFFFFF'); }
      ctx.fillStyle = gr;
    } else ctx.fillStyle = col;
    ctx.fillRect(bx, y + pad, len, bh);
    if (bar.border || bar.gradient) { ctx.strokeStyle = bar.borderColor || col; ctx.lineWidth = 1; ctx.strokeRect(bx + 0.5, y + pad + 0.5, Math.max(0, len - 1), bh - 1); }
  }

  /** horizontal alignment of an item */
  function hAlign(st, disp, cell) {
    const a = st.align && st.align.h;
    if (a && a !== 'general') return a;
    if (!disp) return 'left';
    return disp.align === 'r' ? 'right' : disp.align === 'c' ? 'center' : 'left';
  }
  function drawText(ctx, view, fr, p, it, X, Y, state) {
    const sh = view.sh, wb = sh.wb, z = fr.z;
    const { r, c, cell, st, m } = it;
    if (state.editing && state.editing.r === r && state.editing.c === c && state.editing.sh === sh) return;
    let x = it.x, y = it.y, w = it.w, h = it.h;
    if (w <= 1 || h <= 1) return;
    const dxf = it.cf && it.cf.dxf;
    const ts = it.ts;
    let f = fontOf(wb, st, dxf);
    if (ts && ts.font && !(st.font && (st.font.b || st.font.color))) f = Object.assign({}, f, ts.font);
    const fname = LY.fontName(wb, f);
    const ff = Object.assign({}, f, { name: fname });
    const vert = f.vert === 'superscript' || f.vert === 'subscript';
    const font = LY.cssFont(ff, z, vert ? 0.7 : 1);
    const al = st.align || {};
    const indentPx = (al.indent || 0) * Math.round(9 * z);
    const pad = Math.max(2, Math.round(2 * z));
    const wrap = !!al.wrap && !it.cf?.bar;
    let stEff = st;
    if (dxf && dxf.nf) stEff = Object.assign({}, st, { nf: dxf.nf });
    const rot = al.rot || 0;
    const avail = w - pad * 2 - 1 - indentPx;
    const iconW = it.cf && it.cf.icon ? Math.round(16 * z) : 0;
    let disp = LY.display(sh, cell.dirty ? Object.assign({}, cell, { v: L.calc.cellValue(sh, r, c) }) : cell, stEff, rot || wrap || al.shrink ? null : avail - iconW, font);
    if (!disp && !(it.cf && it.cf.icon)) return;
    if (it.cf && it.cf.icon) {
      const s = Math.min(Math.round(16 * z), h - 2);
      L.render.drawIcon(ctx, it.cf.icon.set, it.cf.icon.i, x + pad, y + h - s - Math.max(1, Math.round(2 * z)), s);
      if (!it.cf.icon.showValue) return;
    }
    if (it.cf && it.cf.bar && !it.cf.bar.showValue) return;
    if (!disp) return;
    let ha = hAlign(st, disp, cell);
    const va = al.v || 'bottom';
    ctx.fillStyle = textColor(wb, f, disp.color, dxf);
    ctx.font = font;
    const fpx = ((f.sz || 11) * 96) / 72 * z;
    const lineH = Math.round(fpx * 1.2);
    let text = disp.text;
    if (ha === 'fill' && text) { const tw = LY.measure(font, text); if (tw > 0) text = text.repeat(Math.max(1, Math.floor(avail / tw))); ha = 'left'; }
    /* clip region: own cell, or overflow neighbours for unwrapped text */
    let clipL = x, clipR = x + w - 1;
    const tw = disp.segs ? LY.segWidth(disp.segs, font) : LY.measure(font, text);
    const overflow = !wrap && !rot && !m && !disp.num && !al.shrink && tw > avail;
    if (overflow) {
      const row = sh.rows[r];
      const empty = (cc) => { const cl = row && row.cells[cc]; return !(cl && cl.v != null && cl.v !== '') && !LY.mergeAt(sh, r, cc); };
      if (ha === 'left' || ha === 'general' || ha === 'center' || ha === 'centerContinuous') { let cc = c + 1, right = x + w; while (right - x - pad < tw + indentPx && cc < M.MAXC && empty(cc)) { right = X(cc + 1); cc++; if (cc - c > 200) break; } clipR = right - 1; }
      if (ha === 'right' || ha === 'center' || ha === 'centerContinuous') { let cc = c - 1, left = x; while (x + w - left - pad < tw + indentPx && cc >= 0 && empty(cc)) { left = X(cc); cc--; if (c - cc > 200) break; } clipL = left; }
      if (ha === 'center' || ha === 'centerContinuous') { const need = (tw - avail) / 2; clipL = Math.max(clipL, x - need - pad); clipR = Math.min(clipR, x + w + need + pad); }
    }
    if (ha === 'centerContinuous') {
      /* centre across following empty cells that share this alignment */
      const row = sh.rows[r];
      let cc = c + 1, right = x + w;
      while (cc < M.MAXC) { const cl = row && row.cells[cc]; if (cl && cl.v != null) break; const s2 = LY.styleOf(sh, r, cc, cl); if (!(s2.align && s2.align.h === 'centerContinuous')) break; right = X(cc + 1); cc++; }
      w = right - x; clipR = Math.max(clipR, right - 1);
    }
    ctx.save();
    ctx.beginPath(); ctx.rect(clipL, y, clipR - clipL, h - 1); ctx.clip();
    const deco = (tx, ty, width) => {
      if (f.u) { const uy = ty + Math.max(1, Math.round(fpx * 0.12)); ctx.fillRect(tx, uy, width, Math.max(1, Math.round(z))); if (/double/i.test(f.u)) ctx.fillRect(tx, uy + Math.max(2, Math.round(2 * z)), width, Math.max(1, Math.round(z))); }
      if (f.strike) ctx.fillRect(tx, ty - Math.round(fpx * 0.3), width, Math.max(1, Math.round(z)));
    };
    if (rot) {
      drawRotated(ctx, text, font, f, fpx, rot, x, y, w, h, ha, va, pad, z, deco);
    } else if (wrap || /\n/.test(text) && wrap) {
      const lines = wrapLines(font, text, Math.max(4, avail));
      const total = lines.length * lineH;
      let ty;
      if (va === 'top') ty = y + pad + fpx * 0.95;
      else if (va === 'center' || va === 'distributed' || va === 'justify') ty = y + (h - total) / 2 + fpx * 0.95;
      else ty = y + h - total - pad + fpx * 0.95;
      for (const ln of lines) {
        const lw = LY.measure(font, ln);
        let tx = ha === 'right' ? x + w - pad - 1 - lw - indentPx : ha === 'center' ? x + (w - lw) / 2 : x + pad + indentPx;
        if (ha === 'justify' && ln !== lines[lines.length - 1]) tx = x + pad;
        ctx.fillText(ln, tx, Math.round(ty));
        deco(tx, Math.round(ty), lw);
        ty += lineH;
      }
    } else {
      let scale = 1;
      let fontUse = font;
      if (al.shrink && tw > avail && tw > 0) { scale = Math.max(0.1, avail / tw); fontUse = LY.cssFont(ff, z, scale * (vert ? 0.7 : 1)); ctx.font = fontUse; }
      const width = tw * scale;
      let ty;
      const asc = fpx * scale * 0.8;
      if (va === 'top') ty = y + pad + asc;
      else if (va === 'center' || va === 'distributed' || va === 'justify') ty = y + h / 2 + asc / 2 - 1;
      else ty = y + h - Math.max(2, Math.round(fpx * scale * 0.22 + 1));
      if (f.vert === 'superscript') ty -= fpx * 0.35;
      if (f.vert === 'subscript') ty += fpx * 0.1;
      ty = Math.round(ty);
      if (disp.segs) {
        /* number formats with fill (*) and padding (_) characters, e.g. accounting */
        const fixedW = LY.segWidth(disp.segs.filter((s) => s.fill == null), fontUse);
        const fills = disp.segs.filter((s) => s.fill != null).length;
        const spare = Math.max(0, avail - fixedW);
        let tx = fills ? x + pad + indentPx : ha === 'right' ? x + w - pad - 1 - fixedW - indentPx : ha === 'center' ? x + (w - fixedW) / 2 : x + pad + indentPx;
        const start = tx;
        for (const s of disp.segs) {
          if (s.s != null) { ctx.fillText(s.s, tx, ty); tx += LY.measure(fontUse, s.s); }
          else if (s.pad != null) tx += LY.measure(fontUse, s.pad);
          else if (s.fill != null) {
            const cw = LY.measure(fontUse, s.fill) || 1;
            const n = Math.floor(spare / fills / cw);
            if (s.fill !== ' ') ctx.fillText(s.fill.repeat(n), tx, ty);
            tx += spare / fills;
          }
        }
        deco(start, ty, tx - start);
      } else {
        let tx;
        if (ha === 'right') tx = x + w - pad - 1 - width - indentPx;
        else if (ha === 'center' || ha === 'centerContinuous' || ha === 'distributed') tx = x + (w - width) / 2;
        else tx = x + pad + indentPx + (it.cf && it.cf.icon && ha !== 'right' ? iconW : 0);
        if (disp.hashes) tx = x + pad;
        /* a number that fits by Excel's (hinted) measure but not by the browser's is narrowed, never clipped */
        if (disp.num && !disp.hashes && width > avail + 1) { const room = w - 2; tx = ha === 'left' ? x + 1 : ha === 'right' ? x + w - 1 - room : x + 1; ctx.fillText(text, Math.round(tx), ty, room); deco(Math.round(tx), ty, room); }
        else { ctx.fillText(text, Math.round(tx), ty); deco(Math.round(tx), ty, width); }
      }
    }
    ctx.restore();
  }
  function drawRotated(ctx, text, font, f, fpx, rot, x, y, w, h, ha, va, pad, z, deco) {
    ctx.save();
    if (rot === 255) {
      /* stacked vertical text */
      const chars = Array.from(String(text));
      const lh = Math.round(fpx * 1.15);
      let ty = va === 'top' ? y + pad + fpx : va === 'center' ? y + (h - chars.length * lh) / 2 + fpx : y + h - chars.length * lh - pad + fpx * 0.9;
      for (const ch of chars) { const cw = LY.measure(font, ch); ctx.fillText(ch, x + (w - cw) / 2, Math.round(ty)); ty += lh; }
      ctx.restore();
      return;
    }
    const ang = rot <= 90 ? -rot : rot - 90;
    const a = (ang * Math.PI) / 180;
    const tw = LY.measure(font, text);
    const cx = x + w / 2, cy = y + h / 2;
    ctx.translate(cx, cy);
    ctx.rotate(a);
    /* place along the rotated baseline, anchored like Excel (bottom-aligned text starts at the cell bottom) */
    let tx = -tw / 2;
    const vSpan = Math.abs(Math.sin(a)) * tw;
    if (va === 'bottom' && Math.abs(ang) > 0) tx = ang < 0 ? -(h / 2 - pad) / Math.abs(Math.sin(a)) + 0 : -tw + (h / 2 - pad) / Math.abs(Math.sin(a));
    if (va === 'top' && Math.abs(ang) > 0) tx = ang < 0 ? (h / 2 - pad) / Math.abs(Math.sin(a)) - tw : -(h / 2 - pad) / Math.abs(Math.sin(a));
    if (vSpan < h - pad * 2 && va !== 'center') { /* fits: keep the chosen anchor */ } else if (va === 'center') tx = -tw / 2;
    ctx.fillText(text, tx, fpx * 0.35);
    deco(tx, fpx * 0.35, tw);
    void ha; void z;
    ctx.restore();
  }
  /** text that spills into the viewport from cells outside it */
  function drawOffscreenOverflow(ctx, view, fr, p, X, Y, state, r0, r1, c0, c1) {
    const sh = view.sh;
    for (let r = r0; r <= r1; r++) {
      const row = sh.rows[r];
      if (!row || !row.cells.length) continue;
      if (Y(r + 1) === Y(r)) continue;
      /* left side: nearest non-empty cell before c0 */
      if (c0 > 0) {
        let cc = c0 - 1, n = 0;
        while (cc >= 0 && n < 256) { const cl = row.cells[cc]; if (cl && cl.v != null && cl.v !== '') break; cc--; n++; }
        if (cc >= 0 && n < 256) {
          const cl = row.cells[cc];
          if (cl && typeof cl.v === 'string' && !LY.mergeAt(sh, r, cc)) {
            const st = LY.styleOf(sh, r, cc, cl);
            const al = st.align || {};
            if (!al.wrap && !al.rot && (!al.h || al.h === 'general' || al.h === 'left' || al.h === 'center' || al.h === 'centerContinuous')) {
              const it = { r, c: cc, cell: cl, st, x: X(cc), y: Y(r), w: X(cc + 1) - X(cc), h: Y(r + 1) - Y(r), cf: sh.cf.length ? L.cf.at(sh, r, cc, cl) : null };
              drawText(ctx, view, fr, p, it, X, Y, state);
            }
          }
        }
      }
      /* right side: right-aligned text from beyond the last visible column */
      if (c1 < M.MAXC - 1) {
        let cc = c1 + 1, n = 0;
        while (cc < M.MAXC && n < 64) { const cl = row.cells[cc]; if (cl && cl.v != null && cl.v !== '') break; cc++; n++; }
        const cl = row.cells[cc];
        if (cl && typeof cl.v === 'string' && !LY.mergeAt(sh, r, cc)) {
          const st = LY.styleOf(sh, r, cc, cl);
          const al = st.align || {};
          if (!al.wrap && (al.h === 'right' || al.h === 'center')) {
            const it = { r, c: cc, cell: cl, st, x: X(cc), y: Y(r), w: X(cc + 1) - X(cc), h: Y(r + 1) - Y(r) };
            drawText(ctx, view, fr, p, it, X, Y, state);
          }
        }
      }
    }
  }
  function drawPageBreaks(ctx, view, fr, p, pages, X, Y, r0, r1, c0, c1) {
    const sh = view.sh;
    const pbp = sh.view.pageBreakPreview;
    ctx.save();
    if (!pbp) {
      /* Normal view after printing or Page Setup: thin dashed lines */
      ctx.strokeStyle = '#000000'; ctx.lineWidth = 1; ctx.setLineDash([3, 3]);
      ctx.beginPath();
      for (const br of pages.rows || []) { if (br.end || br.r < r0 || br.r > r1 + 1) continue; const y = Y(br.r) - 0.5; ctx.moveTo(p.x, y); ctx.lineTo(p.x + p.w, y); }
      for (const br of pages.cols || []) { if (br.end || br.c < c0 || br.c > c1 + 1) continue; const x = X(br.c) - 0.5; ctx.moveTo(x, p.y); ctx.lineTo(x, p.y + p.h); }
      ctx.stroke();
      ctx.restore();
      return;
    }
    /* Page Break Preview, as Excel 2003 draws it: grey outside the print area, a page number on each page,
       a solid blue outline and manual breaks, dashed blue automatic breaks */
    const a = pages.area;
    if (!a) { ctx.restore(); return; }
    const ax1 = X(a.c1), ay1 = Y(a.r1), ax2 = X(a.c2 + 1), ay2 = Y(a.r2 + 1);
    ctx.fillStyle = 'rgba(128,128,128,0.5)';
    ctx.beginPath();
    ctx.rect(p.x, p.y, p.w, p.h);
    ctx.rect(Math.max(p.x - 1, Math.min(p.x + p.w + 1, ax1)), Math.max(p.y - 1, Math.min(p.y + p.h + 1, ay1)), Math.max(0, Math.min(p.x + p.w + 1, ax2) - Math.max(p.x - 1, ax1)), Math.max(0, Math.min(p.y + p.h + 1, ay2) - Math.max(p.y - 1, ay1)));
    ctx.fill('evenodd');
    const rs = [a.r1].concat((pages.rows || []).map((b) => b.r)), cs = [a.c1].concat((pages.cols || []).map((b) => b.c));
    const across = sh.print.pageOrder === 'overThenDown';
    let n = 0;
    const label = (i, j) => {
      const x1 = X(cs[j]), x2 = X(cs[j + 1]), y1 = Y(rs[i]), y2 = Y(rs[i + 1]);
      n++;
      if (x2 < p.x || x1 > p.x + p.w || y2 < p.y || y1 > p.y + p.h) return;
      const size = Math.max(12, Math.min(96 * fr.z * 1.6, (x2 - x1) / 4, (y2 - y1) / 3));
      ctx.font = `bold ${Math.round(size)}px Arial, Arimo, sans-serif`;
      ctx.fillStyle = 'rgba(150,150,150,0.55)';
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText('Page ' + n, (x1 + x2) / 2, (y1 + y2) / 2);
    };
    if (across) { for (let i = 0; i < rs.length - 1; i++) for (let j = 0; j < cs.length - 1; j++) label(i, j); }
    else { for (let j = 0; j < cs.length - 1; j++) for (let i = 0; i < rs.length - 1; i++) label(i, j); }
    ctx.strokeStyle = '#0000FF';
    ctx.lineWidth = 3;
    ctx.setLineDash([]);
    ctx.strokeRect(ax1 - 0.5, ay1 - 0.5, ax2 - ax1, ay2 - ay1);
    for (const manual of [true, false]) {
      ctx.setLineDash(manual ? [] : [8, 4]);
      ctx.beginPath();
      for (const br of pages.rows || []) { if (br.end || !!br.manual !== manual || br.r < r0 || br.r > r1 + 1) continue; const y = Y(br.r) - 0.5; ctx.moveTo(ax1, y); ctx.lineTo(ax2, y); }
      for (const br of pages.cols || []) { if (br.end || !!br.manual !== manual || br.c < c0 || br.c > c1 + 1) continue; const x = X(br.c) - 0.5; ctx.moveTo(x, ay1); ctx.lineTo(x, ay2); }
      ctx.stroke();
    }
    ctx.restore();
  }

  /* ------------------------------------------------------------ selection */
  function drawSelection(ctx, view, fr, state) {
    const sel = state.sel;
    const sh = view.sh;
    const z = fr.z;
    if (!sel) return;
    for (const p of fr.panes) {
      ctx.save();
      ctx.beginPath(); ctx.rect(p.x, p.y, p.w, p.h); ctx.clip();
      const rectOf = (rg) => {
        const r1 = Math.max(rg.r1, p.r0), r2 = Math.min(rg.r2, p.rEnd), c1 = Math.max(rg.c1, p.c0), c2 = Math.min(rg.c2, p.cEnd);
        if (r1 > r2 || c1 > c2) return null;
        const x1 = rg.c1 < p.c0 ? p.x - 3 : colX(fr, p, rg.c1), y1 = rg.r1 < p.r0 ? p.y - 3 : rowY(fr, p, rg.r1);
        const x2 = rg.c2 > p.cEnd ? p.x + p.w + 3 : colX(fr, p, rg.c2 + 1), y2 = rg.r2 > p.rEnd ? p.y + p.h + 3 : rowY(fr, p, rg.r2 + 1);
        return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
      };
      /* range highlight for formula references while editing */
      if (state.hl) {
        for (const h of state.hl) {
          if (h.sh && h.sh !== sh) continue;
          const q = rectOf(h.range);
          if (!q) continue;
          ctx.strokeStyle = h.color; ctx.lineWidth = 2;
          ctx.strokeRect(q.x + 0.5, q.y + 0.5, q.w - 2, q.h - 2);
          if (h.handles) { ctx.fillStyle = h.color; ctx.fillRect(q.x + q.w - 4, q.y + q.h - 4, 5, 5); }
        }
      }
      if (!state.hideSel) {
        /* shaded selection (the active cell stays white) */
        ctx.fillStyle = COL.sel;
        const ranges = sel.ranges;
        const am = LY.mergeAt(sh, sel.r, sel.c);
        const ac = am || { r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c };
        const aq = rectOf(ac);
        for (let k = 0; k < ranges.length; k++) {
          const rg = ranges[k];
          const q = rectOf(rg);
          if (!q) continue;
          if (rg.r1 === rg.r2 && rg.c1 === rg.c2) continue;
          if (aq && M.rangesOverlap(rg, ac)) {
            /* tint around the active cell */
            const ax1 = Math.max(aq.x, q.x), ay1 = Math.max(aq.y, q.y), ax2 = Math.min(aq.x + aq.w, q.x + q.w), ay2 = Math.min(aq.y + aq.h, q.y + q.h);
            if (ay1 > q.y) ctx.fillRect(q.x, q.y, q.w - 1, ay1 - q.y);
            if (q.y + q.h > ay2) ctx.fillRect(q.x, ay2, q.w - 1, q.y + q.h - ay2 - 1);
            if (ax1 > q.x) ctx.fillRect(q.x, ay1, ax1 - q.x, ay2 - ay1);
            if (q.x + q.w > ax2) ctx.fillRect(ax2, ay1, q.x + q.w - ax2 - 1, ay2 - ay1);
          } else ctx.fillRect(q.x, q.y, q.w - 1, q.h - 1);
        }
        /* border of the active range */
        const act = ranges[sel.active || 0] || ac;
        const big = act.r1 === act.r2 && act.c1 === act.c2 ? ac : act;
        const bq = rectOf(big);
        if (bq) {
          ctx.fillStyle = COL.selBorder;
          const t = Math.max(2, Math.round(2 * Math.min(z, 1.5)));
          ctx.fillRect(bq.x - 1, bq.y - 1, bq.w + 1, t);
          ctx.fillRect(bq.x - 1, bq.y + bq.h - t, bq.w + 1, t);
          ctx.fillRect(bq.x - 1, bq.y - 1, t, bq.h + 1);
          ctx.fillRect(bq.x + bq.w - t, bq.y - 1, t, bq.h + 1);
          /* fill handle */
          if (ranges.length <= 1 && !state.editing && big.r2 <= p.rEnd && big.c2 <= p.cEnd) {
            ctx.fillStyle = '#FFFFFF'; ctx.fillRect(bq.x + bq.w - 5, bq.y + bq.h - 5, 7, 7);
            ctx.fillStyle = '#000000'; ctx.fillRect(bq.x + bq.w - 4, bq.y + bq.h - 4, 5, 5);
          }
        }
        if (ranges.length > 1) {
          /* multi-area selections: thin borders around each area */
          ctx.strokeStyle = '#000'; ctx.lineWidth = 1;
          for (const rg of ranges) { const q = rectOf(rg); if (q) ctx.strokeRect(q.x + 0.5, q.y + 0.5, q.w - 2, q.h - 2); }
        }
      }
      /* fill handle drag preview / move preview */
      for (const pr of [state.fillTo, state.dragTo]) {
        if (!pr) continue;
        const q = rectOf(pr);
        if (!q) continue;
        ctx.save(); ctx.strokeStyle = '#808080'; ctx.lineWidth = 2; ctx.setLineDash([3, 2]);
        ctx.strokeRect(q.x + 0.5, q.y + 0.5, q.w - 2, q.h - 2); ctx.restore();
      }
      /* copy marquee (marching ants) */
      if (state.marquee && (!state.marquee.sh || state.marquee.sh === sh)) {
        for (const rg of state.marquee.ranges) {
          const q = rectOf(rg);
          if (!q) continue;
          ctx.save();
          ctx.lineWidth = 2;
          ctx.strokeStyle = '#000'; ctx.setLineDash([4, 4]); ctx.lineDashOffset = -(state.marquee.phase || 0);
          ctx.strokeRect(q.x + 0.5, q.y + 0.5, q.w - 2, q.h - 2);
          ctx.strokeStyle = '#fff'; ctx.lineDashOffset = -(state.marquee.phase || 0) + 4;
          ctx.strokeRect(q.x + 0.5, q.y + 0.5, q.w - 2, q.h - 2);
          ctx.restore();
        }
      }
      ctx.restore();
    }
  }

  /* ------------------------------------------------------------ headers */
  function drawHeaders(ctx, view, fr, state) {
    const sh = view.sh, z = fr.z, g = fr.g;
    const W = view.w, H = view.h;
    const sel = state.sel;
    const selRows = new Set(), selCols = new Set();
    const fullRows = new Set(), fullCols = new Set();
    if (sel) for (const rg of sel.ranges) {
      const allCols = rg.c1 === 0 && rg.c2 >= M.MAXC - 1, allRows = rg.r1 === 0 && rg.r2 >= M.MAXR - 1;
      for (const p of fr.panes) {
        for (let r = Math.max(rg.r1, p.r0); r <= Math.min(rg.r2, p.rEnd); r++) { selRows.add(r); if (allCols) fullRows.add(r); }
        for (let c = Math.max(rg.c1, p.c0); c <= Math.min(rg.c2, p.cEnd); c++) { selCols.add(c); if (allRows) fullCols.add(c); }
      }
    }
    ctx.font = fr.hdrFont;
    ctx.textBaseline = 'middle';
    /* column headers */
    const gradH = ctx.createLinearGradient(0, 0, 0, fr.hdrH);
    gradH.addColorStop(0, COL.hdrTop); gradH.addColorStop(1, COL.hdrBot);
    const gradHs = ctx.createLinearGradient(0, 0, 0, fr.hdrH);
    gradHs.addColorStop(0, COL.hdrSelTop); gradHs.addColorStop(1, COL.hdrSelBot);
    ctx.fillStyle = gradH; ctx.fillRect(fr.hdrW, 0, W - fr.hdrW, fr.hdrH);
    const r1c1 = sh.wb.r1c1;
    for (const p of fr.panes) {
      if (p.kind === 'left' || (p.kind === 'main' && fr.fr)) continue;
      ctx.save(); ctx.beginPath(); ctx.rect(p.x, 0, p.w, fr.hdrH); ctx.clip();
      for (let c = p.c0; c <= p.cEnd; c++) {
        const x1 = colX(fr, p, c), x2 = colX(fr, p, c + 1);
        if (x2 <= x1) { ctx.fillStyle = COL.hdrLine; ctx.fillRect(x1 - 2, 0, 1, fr.hdrH); continue; }
        if (selCols.has(c)) { ctx.fillStyle = fullCols.has(c) ? COL.hdrAllBot : gradHs; ctx.fillRect(x1, 0, x2 - x1, fr.hdrH); }
        ctx.fillStyle = selCols.has(c) ? COL.hdrSelLine : COL.hdrLine;
        ctx.fillRect(x2 - 1, 0, 1, fr.hdrH);
        ctx.fillStyle = COL.hdrText;
        const label = r1c1 ? String(c + 1) : L.formula.colName(c);
        const lw = LY.measure(fr.hdrFont, label);
        if (x2 - x1 > lw / 2) ctx.fillText(label, Math.round(x1 + (x2 - x1 - lw) / 2), Math.round(fr.hdrH / 2) + 1);
      }
      ctx.restore();
    }
    ctx.fillStyle = COL.hdrLine; ctx.fillRect(fr.hdrW, fr.hdrH - 1, W - fr.hdrW, 1);
    /* row headers */
    const gradV = ctx.createLinearGradient(0, 0, fr.hdrW, 0);
    gradV.addColorStop(0, COL.hdrTop); gradV.addColorStop(1, COL.hdrBot);
    const gradVs = ctx.createLinearGradient(0, 0, fr.hdrW, 0);
    gradVs.addColorStop(0, COL.hdrSelTop); gradVs.addColorStop(1, COL.hdrSelBot);
    ctx.fillStyle = gradV; ctx.fillRect(0, fr.hdrH, fr.hdrW, H - fr.hdrH);
    for (const p of fr.panes) {
      if (p.kind === 'top' || (p.kind === 'main' && fr.fc)) continue;
      ctx.save(); ctx.beginPath(); ctx.rect(0, p.y, fr.hdrW, p.h); ctx.clip();
      for (let r = p.r0; r <= p.rEnd; r++) {
        const y1 = rowY(fr, p, r), y2 = rowY(fr, p, r + 1);
        if (y2 <= y1) { ctx.fillStyle = COL.hdrLine; ctx.fillRect(0, y1 - 2, fr.hdrW, 1); continue; }
        if (selRows.has(r)) { ctx.fillStyle = fullRows.has(r) ? COL.hdrAllBot : gradVs; ctx.fillRect(0, y1, fr.hdrW, y2 - y1); }
        ctx.fillStyle = selRows.has(r) ? COL.hdrSelLine : COL.hdrLine;
        ctx.fillRect(0, y2 - 1, fr.hdrW, 1);
        if (y2 - y1 >= 6) {
          ctx.fillStyle = COL.hdrText;
          const label = String(r + 1);
          const lw = LY.measure(fr.hdrFont, label);
          ctx.fillText(label, Math.round((fr.hdrW - lw) / 2), Math.round((y1 + y2) / 2) + 1);
        }
      }
      ctx.restore();
    }
    ctx.fillStyle = COL.hdrLine; ctx.fillRect(fr.hdrW - 1, fr.hdrH, 1, H - fr.hdrH);
    /* select-all corner */
    ctx.fillStyle = gradH; ctx.fillRect(0, 0, fr.hdrW, fr.hdrH);
    ctx.fillStyle = COL.hdrLine; ctx.fillRect(fr.hdrW - 1, 0, 1, fr.hdrH); ctx.fillRect(0, fr.hdrH - 1, fr.hdrW, 1);
    ctx.fillStyle = '#B5C5DA';
    ctx.beginPath(); ctx.moveTo(fr.hdrW - 4, 3); ctx.lineTo(fr.hdrW - 4, fr.hdrH - 4); ctx.lineTo(fr.hdrW - 4 - (fr.hdrH - 7), fr.hdrH - 4); ctx.closePath(); ctx.fill();
    ctx.textBaseline = 'alphabetic';
    void g;
  }
})(typeof window !== 'undefined' ? window : globalThis);
