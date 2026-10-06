/* VibeOffice — Windows metafile (WMF / EMF) renderer, shared by Quire, Ledger and Lectern.
 * Browsers cannot display .wmf/.emf, yet older documents are full of them (clip art, logos,
 * pasted Visio/Excel graphics). This plays the GDI records onto a canvas so the picture
 * shows; the original metafile bytes are kept for saving.
 * Covered: pens, brushes (solid, hatched, DIB pattern), fonts and text, polygons, polylines,
 * Béziers, paths, rectangles, ellipses, arcs/pies/chords, DIB bitmaps, window/viewport mapping,
 * world transforms, SaveDC/RestoreDC and rectangular clipping. EMF+ comments are skipped; the
 * GDI fallback records that "EMF+ dual" files carry are drawn instead.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});

  /* ---------- shared helpers ---------- */
  const rgb = (cr) => `rgb(${cr & 255},${(cr >> 8) & 255},${(cr >> 16) & 255})`;
  const HATCH = [[[0, 4, 8, 4]], [[4, 0, 4, 8]], [[0, 0, 8, 8]], [[0, 8, 8, 0]], [[0, 4, 8, 4], [4, 0, 4, 8]], [[0, 0, 8, 8], [0, 8, 8, 0]]];
  function hatchPattern(g, style, color, bk) {
    const c = document.createElement('canvas');
    c.width = c.height = 8;
    const x = c.getContext('2d');
    if (bk) { x.fillStyle = bk; x.fillRect(0, 0, 8, 8); }
    x.strokeStyle = color; x.lineWidth = 1;
    for (const [a, b, cc, d] of HATCH[style] || HATCH[0]) { x.beginPath(); x.moveTo(a + 0.5, b + 0.5); x.lineTo(cc + 0.5, d + 0.5); x.stroke(); }
    return g.createPattern(c, 'repeat');
  }
  /** Decode a device-independent bitmap (BITMAPINFOHEADER or BITMAPCOREHEADER) into a canvas. */
  function dibCanvas(u8, dv, off, bitsOff, usage, end, extPal) {
    try {
      const hs = dv.getUint32(off, true);
      let w, h, bpp, comp = 0, clrUsed = 0, palOff, palEntry;
      if (hs === 12) { w = dv.getUint16(off + 4, true); h = dv.getInt16(off + 6, true); bpp = dv.getUint16(off + 10, true); palEntry = 3; }
      else { w = dv.getInt32(off + 4, true); h = dv.getInt32(off + 8, true); bpp = dv.getUint16(off + 14, true); comp = dv.getUint32(off + 16, true); clrUsed = dv.getUint32(off + 32, true); palEntry = 4; }
      if (!(w > 0) || !h || Math.abs(h) > 20000 || w > 20000) return null;
      if (comp !== 0 && comp !== 3) return null; /* RLE / JPEG / PNG inside a DIB: not handled */
      const nPal = bpp <= 8 ? clrUsed || 1 << bpp : 0;
      palOff = off + hs + (comp === 3 && hs === 40 ? 12 : 0);
      const ah = Math.abs(h), stride = Math.floor((w * bpp + 31) / 32) * 4;
      /* some writers omit the colour table and rely on the selected logical palette */
      const noTable = bitsOff == null && end != null && palOff + nPal * palEntry + stride * ah > end && palOff + stride * ah <= end;
      const pal = [];
      for (let i = 0; i < nPal; i++) {
        if (noTable || usage === 1) { const q = extPal && extPal[i]; pal.push(q ? q.slice(0, 3) : [i, i, i]); continue; }
        const p = palOff + i * palEntry;
        pal.push([u8[p + 2], u8[p + 1], u8[p]]);
      }
      let bits = bitsOff != null ? bitsOff : palOff + (noTable ? 0 : nPal * palEntry);
      if (bits + stride * ah > u8.length + stride) return null;
      const c = document.createElement('canvas');
      c.width = w; c.height = ah;
      const x = c.getContext('2d');
      const img = x.createImageData(w, ah);
      const d = img.data;
      for (let row = 0; row < ah; row++) {
        const sy = h > 0 ? ah - 1 - row : row;
        const rp = bits + sy * stride;
        for (let col = 0; col < w; col++) {
          let r = 0, g = 0, b = 0, a = 255;
          if (bpp === 24) { const p = rp + col * 3; b = u8[p]; g = u8[p + 1]; r = u8[p + 2]; }
          else if (bpp === 32) { const p = rp + col * 4; b = u8[p]; g = u8[p + 1]; r = u8[p + 2]; }
          else if (bpp === 16) { const v = u8[rp + col * 2] | (u8[rp + col * 2 + 1] << 8); r = ((v >> 10) & 31) << 3; g = ((v >> 5) & 31) << 3; b = (v & 31) << 3; }
          else if (bpp === 8) { const q = pal[u8[rp + col]] || [0, 0, 0]; [r, g, b] = q; }
          else if (bpp === 4) { const v = u8[rp + (col >> 1)]; const q = pal[col & 1 ? v & 15 : v >> 4] || [0, 0, 0]; [r, g, b] = q; }
          else if (bpp === 1) { const v = u8[rp + (col >> 3)]; const q = pal[(v >> (7 - (col & 7))) & 1] || [0, 0, 0]; [r, g, b] = q; }
          const o = (row * w + col) * 4;
          d[o] = r; d[o + 1] = g; d[o + 2] = b; d[o + 3] = a;
        }
      }
      x.putImageData(img, 0, 0);
      return c;
    } catch (e) { return null; }
  }

  /* ---------- a GDI device context replayed onto a canvas ---------- */
  function Player(g, map) {
    this.g = g;
    this.map = map; /* (x, y) in device units → canvas */
    this.s = this.fresh();
    this.stack = [];
    this.path = null;
  }
  Player.prototype.fresh = function () {
    return {
      pen: { style: 0, w: 1, color: 0 }, brush: { style: 0, color: 0xffffff }, font: { h: -12, w: 400, i: false, face: 'Arial', esc: 0, charset: 0 },
      text: 0, bk: 0xffffff, bkMode: 2, align: 0, fill: 1, x: 0, y: 0,
      wo: [0, 0], we: null, vo: [0, 0], ve: null, mapMode: 1, xf: [1, 0, 0, 1, 0, 0], clip: null,
    };
  };
  /* pens, brushes and fonts are never mutated after creation, so a shallow copy is a faithful DC snapshot */
  Player.prototype.save = function () {
    const s = this.s;
    this.stack.push(Object.assign({}, s, { wo: s.wo.slice(), we: s.we && s.we.slice(), vo: s.vo.slice(), ve: s.ve && s.ve.slice(), xf: s.xf.slice(), clip: s.clip && s.clip.slice() }));
  };
  Player.prototype.restore = function (n) {
    let k = n < 0 ? -n : 1;
    let st = null;
    while (k-- > 0 && this.stack.length) st = this.stack.pop();
    if (st) this.s = st;
  };
  /* logical → device → canvas */
  Player.prototype.pt = function (x, y) {
    const s = this.s;
    const X = s.xf[0] * x + s.xf[2] * y + s.xf[4], Y = s.xf[1] * x + s.xf[3] * y + s.xf[5];
    let dx = X, dy = Y;
    if (s.we && s.mapMode >= 7) { const ve = s.ve || s.we; dx = ((X - s.wo[0]) * ve[0]) / (s.we[0] || 1) + s.vo[0]; dy = ((Y - s.wo[1]) * ve[1]) / (s.we[1] || 1) + s.vo[1]; }
    else { dx = X - s.wo[0] + s.vo[0]; dy = Y - s.wo[1] + s.vo[1]; }
    return this.map(dx, dy);
  };
  Player.prototype.scale = function () {
    const a = this.pt(0, 0), b = this.pt(1000, 1000);
    return { x: Math.abs(b[0] - a[0]) / 1000, y: Math.abs(b[1] - a[1]) / 1000 };
  };
  Player.prototype.penOn = function () { const p = this.s.pen; return p && (p.style & 15) !== 5; };
  Player.prototype.brushOn = function () { const b = this.s.brush; return b && b.style !== 1; };
  Player.prototype.applyPen = function () {
    const g = this.g, p = this.s.pen, k = this.scale();
    g.strokeStyle = rgb(p.color);
    g.lineWidth = Math.max(1, (p.w || 0) * (k.x + k.y) / 2);
    const st = p.style & 15, lw = g.lineWidth;
    g.setLineDash(st === 1 ? [lw * 3, lw] : st === 2 ? [lw, lw] : st === 3 ? [lw * 3, lw, lw, lw] : st === 4 ? [lw * 3, lw, lw, lw, lw, lw] : []);
    const cap = (p.style >> 8) & 15, join = (p.style >> 12) & 15;
    g.lineCap = cap === 1 ? 'square' : cap === 2 ? 'butt' : 'round';
    g.lineJoin = join === 1 ? 'bevel' : join === 2 ? 'miter' : 'round';
  };
  Player.prototype.applyBrush = function () {
    const g = this.g, b = this.s.brush;
    if (b.style === 2) g.fillStyle = hatchPattern(g, b.hatch || 0, rgb(b.color), this.s.bkMode === 2 ? rgb(this.s.bk) : null);
    else if ((b.style === 3 || b.style === 5 || b.style === 6) && b.pat) g.fillStyle = b.pat;
    else g.fillStyle = rgb(b.color);
  };
  Player.prototype.withClip = function (fn) {
    const g = this.g, c = this.s.clip;
    if (c) { g.save(); g.beginPath(); g.rect(c[0], c[1], c[2] - c[0], c[3] - c[1]); g.clip(); }
    try { fn(); } finally { if (c) g.restore(); }
  };
  /** draw (or, inside BeginPath/EndPath, record) a shape described by a path-building callback */
  Player.prototype.shape = function (build, closed, fillIt) {
    if (this.path) { this.path.push(build); return; }
    const g = this.g;
    this.withClip(() => {
      g.beginPath();
      build(g);
      if (closed && fillIt !== false && this.brushOn()) { this.applyBrush(); g.fill(this.s.fill === 1 ? 'evenodd' : 'nonzero'); }
      if (this.penOn()) { this.applyPen(); g.stroke(); }
    });
  };
  Player.prototype.poly = function (pts, closed) {
    if (pts.length < 2) return;
    const P = pts.map(([x, y]) => this.pt(x, y));
    this.shape((g) => { g.moveTo(P[0][0], P[0][1]); for (let i = 1; i < P.length; i++) g.lineTo(P[i][0], P[i][1]); if (closed) g.closePath(); }, closed);
  };
  Player.prototype.polyPoly = function (polys, closed) {
    const PP = polys.filter((p) => p.length > 1).map((p) => p.map(([x, y]) => this.pt(x, y)));
    if (!PP.length) return;
    this.shape((g) => { for (const P of PP) { g.moveTo(P[0][0], P[0][1]); for (let i = 1; i < P.length; i++) g.lineTo(P[i][0], P[i][1]); if (closed) g.closePath(); } }, closed);
  };
  Player.prototype.bezier = function (pts, fromCurrent) {
    const all = fromCurrent ? [[this.s.x, this.s.y]].concat(pts) : pts;
    if (all.length < 4) return;
    const P = all.map(([x, y]) => this.pt(x, y));
    /* PolyBezierTo inside a path continues the current figure; anywhere else it starts at the current point */
    const cont = fromCurrent && !!this.path;
    this.shape((g) => { if (!cont) g.moveTo(P[0][0], P[0][1]); for (let i = 1; i + 2 < P.length; i += 3) g.bezierCurveTo(P[i][0], P[i][1], P[i + 1][0], P[i + 1][1], P[i + 2][0], P[i + 2][1]); }, false);
    const last = all[all.length - 1];
    this.s.x = last[0]; this.s.y = last[1];
  };
  Player.prototype.rect = function (l, t, r, b, rx, ry) {
    const p0 = this.pt(l, t), p1 = this.pt(r, b);
    const x = Math.min(p0[0], p1[0]), y = Math.min(p0[1], p1[1]), w = Math.abs(p1[0] - p0[0]), h = Math.abs(p1[1] - p0[1]);
    const k = this.scale();
    const cx = Math.min(w / 2, ((rx || 0) * k.x) / 2), cy = Math.min(h / 2, ((ry || 0) * k.y) / 2);
    this.shape((g) => {
      if (cx > 0 && cy > 0) {
        g.moveTo(x + cx, y); g.lineTo(x + w - cx, y); g.ellipse(x + w - cx, y + cy, cx, cy, 0, -Math.PI / 2, 0); g.lineTo(x + w, y + h - cy);
        g.ellipse(x + w - cx, y + h - cy, cx, cy, 0, 0, Math.PI / 2); g.lineTo(x + cx, y + h); g.ellipse(x + cx, y + h - cy, cx, cy, 0, Math.PI / 2, Math.PI);
        g.lineTo(x, y + cy); g.ellipse(x + cx, y + cy, cx, cy, 0, Math.PI, Math.PI * 1.5); g.closePath();
      } else g.rect(x, y, w, h);
    }, true);
  };
  Player.prototype.ellipse = function (l, t, r, b) {
    const p0 = this.pt(l, t), p1 = this.pt(r, b);
    const cx = (p0[0] + p1[0]) / 2, cy = (p0[1] + p1[1]) / 2, rx = Math.abs(p1[0] - p0[0]) / 2, ry = Math.abs(p1[1] - p0[1]) / 2;
    this.shape((g) => { g.moveTo(cx + rx, cy); g.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2); g.closePath(); }, true);
  };
  /** kind: 'arc' | 'pie' | 'chord'. GDI arcs run counter-clockwise from start to end radial. */
  Player.prototype.arc = function (l, t, r, b, xs, ys, xe, ye, kind) {
    const p0 = this.pt(l, t), p1 = this.pt(r, b), s0 = this.pt(xs, ys), e0 = this.pt(xe, ye);
    const cx = (p0[0] + p1[0]) / 2, cy = (p0[1] + p1[1]) / 2, rx = Math.abs(p1[0] - p0[0]) / 2, ry = Math.abs(p1[1] - p0[1]) / 2;
    if (!rx || !ry) return;
    const a0 = Math.atan2((s0[1] - cy) / ry, (s0[0] - cx) / rx), a1 = Math.atan2((e0[1] - cy) / ry, (e0[0] - cx) / rx);
    const flipped = (p1[1] - p0[1]) * (p1[0] - p0[0]) < 0;
    this.shape((g) => {
      if (kind === 'pie') g.moveTo(cx, cy);
      g.ellipse(cx, cy, rx, ry, 0, a0, a1 === a0 ? a0 + Math.PI * 2 : a1, !flipped);
      if (kind !== 'arc') g.closePath();
    }, kind !== 'arc');
  };
  Player.prototype.text = function (x, y, str, dxs) {
    if (!str) return;
    const g = this.g, s = this.s, f = s.font, k = this.scale();
    const size = Math.max(1, Math.abs(f.h || 12) * k.y * (f.h > 0 ? 0.82 : 1));
    const useCP = s.align & 1;
    const p = useCP ? this.pt(s.x, s.y) : this.pt(x, y);
    const face = f.face || 'Arial';
    this.withClip(() => {
      g.save();
      g.font = `${f.i ? 'italic ' : ''}${f.w >= 600 ? 'bold ' : ''}${size}px ${L.fontStack ? L.fontStack(face) : '"' + face + '", sans-serif'}`;
      g.fillStyle = rgb(s.text);
      const ha = s.align & 6, va = s.align & 24;
      g.textAlign = ha === 6 ? 'center' : ha === 2 ? 'right' : 'left';
      g.textBaseline = va === 24 ? 'alphabetic' : va === 8 ? 'bottom' : 'top';
      g.translate(p[0], p[1]);
      if (f.esc) g.rotate((-f.esc / 10) * Math.PI / 180);
      if (dxs && dxs.length === str.length && g.textAlign === 'left') {
        let cx = 0;
        for (let i = 0; i < str.length; i++) { g.fillText(str[i], cx, 0); cx += dxs[i] * k.x; }
      } else g.fillText(str, 0, 0);
      g.restore();
    });
    if (useCP) { const wpx = g.measureText(str).width; s.x += wpx / (k.x || 1); }
  };
  const ROP_SRCAND = 0x008800c6, ROP_SRCINVERT = 0x00660046, ROP_SRCPAINT = 0x00ee0086;
  /** crop a source rectangle out of a bitmap canvas into its own canvas */
  function cropTo(src, sx, sy, sw, sh, W, H) {
    const c = document.createElement('canvas');
    c.width = W; c.height = H;
    c.getContext('2d').drawImage(src, Math.max(0, sx), Math.max(0, sy), Math.max(1, Math.min(Math.abs(sw) || src.width, src.width)), Math.max(1, Math.min(Math.abs(sh) || src.height, src.height)), 0, 0, W, H);
    return c;
  }
  /* GDI raster ops: the AND-mask + XOR-image pair (how icons are drawn transparently) becomes one image with alpha */
  Player.prototype.blit = function (src, sx, sy, sw, sh, dx, dy, dw, dh, rop) {
    if (!src) return;
    if (rop === ROP_SRCAND) { this.mask = { src, sx, sy, sw, sh, key: [dx, dy, dw, dh].join() }; return; }
    if (rop === ROP_SRCINVERT && this.mask && this.mask.key === [dx, dy, dw, dh].join()) {
      const m = this.mask; this.mask = null;
      const W = Math.max(Math.abs(sw) || src.width, Math.abs(m.sw) || m.src.width), H = Math.max(Math.abs(sh) || src.height, Math.abs(m.sh) || m.src.height);
      const mc = cropTo(m.src, m.sx, m.sy, m.sw, m.sh, W, H), xc = cropTo(src, sx, sy, sw, sh, W, H);
      const md = mc.getContext('2d').getImageData(0, 0, W, H).data, xg = xc.getContext('2d'), xi = xg.getImageData(0, 0, W, H), xd = xi.data;
      for (let i = 0; i < xd.length; i += 4) {
        const maskDark = md[i] + md[i + 1] + md[i + 2] < 384;
        const xorBlack = xd[i] + xd[i + 1] + xd[i + 2] < 24;
        xd[i + 3] = maskDark || !xorBlack ? 255 : 0;
      }
      xg.putImageData(xi, 0, 0);
      this.bitmap(xc, 0, 0, W, H, dx, dy, dw, dh);
      return;
    }
    if (this.mask) { const m = this.mask; this.mask = null; void m; }
    this.bitmap(src, sx, sy, sw, sh, dx, dy, dw, dh, rop === ROP_SRCINVERT ? 'difference' : rop === ROP_SRCPAINT ? 'lighten' : rop === ROP_SRCAND ? 'multiply' : null);
  };
  Player.prototype.bitmap = function (src, sx, sy, sw, sh, dx, dy, dw, dh, op) {
    if (!src) return;
    const p0 = this.pt(dx, dy), p1 = this.pt(dx + dw, dy + dh);
    const x = Math.min(p0[0], p1[0]), y = Math.min(p0[1], p1[1]), w = Math.abs(p1[0] - p0[0]), h = Math.abs(p1[1] - p0[1]);
    const fx = (p1[0] < p0[0]) !== (sw < 0), fy = (p1[1] < p0[1]) !== (sh < 0);
    sw = Math.abs(sw) || src.width; sh = Math.abs(sh) || src.height;
    this.withClip(() => {
      const g = this.g;
      g.save();
      g.translate(x + (fx ? w : 0), y + (fy ? h : 0));
      g.scale(fx ? -1 : 1, fy ? -1 : 1);
      g.imageSmoothingEnabled = true;
      if (op) g.globalCompositeOperation = op;
      try { g.drawImage(src, Math.max(0, sx), Math.max(0, sy), Math.min(sw, src.width), Math.min(sh, src.height), 0, 0, w, h); } catch (e) { /* ignore */ }
      g.restore();
    });
  };
  /** pattern blits: only the ROPs that paint something meaningful (no-op markers are common in EMF+ dual files) */
  Player.prototype.patBlt = function (rop, x, y, w, h) {
    const keepPen = this.s.pen, keepBrush = this.s.brush;
    if (rop === 0x00f00021) { /* PATCOPY */ }
    else if (rop === 0x00ff0062) this.s.brush = { style: 0, color: 0xffffff };
    else if (rop === 0x00000042) this.s.brush = { style: 0, color: 0 };
    else return;
    this.s.pen = { style: 5 };
    this.rect(x, y, x + w, y + h);
    this.s.pen = keepPen; this.s.brush = keepBrush;
  };
  Player.prototype.fillPath = function (mode) {
    const builds = this.path || this.lastPath || [];
    this.path = null;
    if (!builds.length) return;
    const g = this.g;
    this.withClip(() => {
      g.beginPath();
      for (const b of builds) b(g);
      if (mode !== 'stroke' && this.brushOn()) { this.applyBrush(); g.fill(this.s.fill === 1 ? 'evenodd' : 'nonzero'); }
      if (mode !== 'fill' && this.penOn()) { this.applyPen(); g.stroke(); }
    });
  };

  /* ---------- WMF ---------- */
  function renderWMF(u8, maxPx) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let p = 0, bbox = null, inch = 1440;
    if (dv.getUint32(0, true) === 0x9ac6cdd7) {
      bbox = [dv.getInt16(6, true), dv.getInt16(8, true), dv.getInt16(10, true), dv.getInt16(12, true)];
      inch = dv.getUint16(14, true) || 1440;
      p = 22;
    }
    const hdrWords = dv.getUint16(p + 2, true);
    const nObjects = dv.getUint16(p + 10, true);
    p += hdrWords * 2;
    /* pre-scan for the window to size the picture when there is no placeable header */
    let win = null, org = [0, 0];
    for (let q = p; q + 6 <= u8.length;) {
      const size = dv.getUint32(q, true) * 2, fn = dv.getUint16(q + 4, true);
      if (size < 6 || fn === 0) break;
      if (fn === 0x020c && !win) win = [dv.getInt16(q + 8, true), dv.getInt16(q + 6, true)];
      if (fn === 0x020b && !win) org = [dv.getInt16(q + 8, true), dv.getInt16(q + 6, true)];
      q += size;
    }
    if (!bbox) bbox = win ? [org[0], org[1], org[0] + win[0], org[1] + win[1]] : [0, 0, 1000, 1000];
    const bw = Math.abs(bbox[2] - bbox[0]) || 1, bh = Math.abs(bbox[3] - bbox[1]) || 1;
    const placeable = p === 22 + hdrWords * 2;
    /* without a placeable header there is no physical size: treat logical units as pixels */
    const natW = placeable ? (bw / inch) * 96 : bw, natH = placeable ? (bh / inch) * 96 : bh;
    const k = placeable ? Math.min(maxPx / Math.max(natW, natH), 4) || 1 : maxPx / Math.max(natW, natH) || 1;
    const cw = Math.max(1, Math.round(natW * k)), ch = Math.max(1, Math.round(natH * k));
    const c = document.createElement('canvas');
    c.width = cw; c.height = ch;
    const g = c.getContext('2d');
    const pl = new Player(g, null);
    pl.wmf = true;
    /* WMF: the window maps onto the placeable bounds, which fill the canvas */
    pl.map = (x, y) => [((x - bbox[0]) * cw) / (bbox[2] - bbox[0] || 1), ((y - bbox[1]) * ch) / (bbox[3] - bbox[1] || 1)];
    pl.pt = function (x, y) {
      const s = this.s;
      let X = x, Y = y;
      if (s.we) {
        /* window → placeable bounds */
        X = bbox[0] + ((x - s.wo[0]) * (bbox[2] - bbox[0])) / (s.we[0] || 1);
        Y = bbox[1] + ((y - s.wo[1]) * (bbox[3] - bbox[1])) / (s.we[1] || 1);
      }
      return this.map(X, Y);
    };
    const objs = new Array(Math.max(nObjects, 1)).fill(null);
    const add = (o) => { let i = objs.indexOf(null); if (i < 0) { i = objs.length; objs.push(null); } objs[i] = o; };
    const td = (bytes, charset, face) => {
      let enc = { 128: 'shift_jis', 134: 'gbk', 136: 'big5', 129: 'euc-kr', 161: 'windows-1253', 162: 'windows-1254', 177: 'windows-1255', 178: 'windows-1256', 186: 'windows-1257', 204: 'windows-1251', 222: 'windows-874', 238: 'windows-1250' }[charset] || 'windows-1252';
      /* DEFAULT_CHARSET text was written in the author's system code page: infer it from an East Asian font name */
      if ((charset === 0 || charset === 1) && face && bytes.some((b) => b > 127)) {
        if (/YaHei|SimSun|SimHei|KaiTi|FangSong|DengXian|NSimSun|宋体|黑体|微软雅黑/i.test(face)) enc = 'gbk';
        else if (/MingLiU|JhengHei|細明|新細明/i.test(face)) enc = 'big5';
        else if (/MS (P)?(Gothic|Mincho)|Meiryo|Yu (Gothic|Mincho)|ＭＳ|MS UI Gothic/i.test(face)) enc = 'shift_jis';
        else if (/Gulim|Dotum|Batang|Gungsuh|Malgun/i.test(face)) enc = 'euc-kr';
      }
      try { return new TextDecoder(enc).decode(bytes); } catch (e) { return String.fromCharCode(...bytes); }
    };
    let n = 0;
    while (p + 6 <= u8.length && n++ < 200000) {
      const size = dv.getUint32(p, true) * 2, fn = dv.getUint16(p + 4, true);
      if (size < 6) break;
      const a = p + 6;
      const i16 = (k2) => dv.getInt16(a + k2 * 2, true);
      const u16 = (k2) => dv.getUint16(a + k2 * 2, true);
      const u32 = (k2) => dv.getUint32(a + k2 * 2, true);
      const s = pl.s;
      try {
        switch (fn) {
          case 0x0000: p = u8.length; continue;
          case 0x0102: s.bkMode = u16(0); break;
          case 0x0103: s.mapMode = u16(0); break;
          case 0x0106: s.fill = u16(0); break;
          case 0x012e: s.align = u16(0); break;
          case 0x0201: s.bk = u32(0); break;
          case 0x0209: s.text = u32(0); break;
          case 0x020b: s.wo = [i16(1), i16(0)]; break;
          case 0x020c: s.we = [i16(1), i16(0)]; break;
          case 0x020f: s.wo = [s.wo[0] + i16(1), s.wo[1] + i16(0)]; break;
          case 0x0410: if (s.we) s.we = [(s.we[0] * i16(3)) / (i16(2) || 1), (s.we[1] * i16(1)) / (i16(0) || 1)]; break;
          case 0x001e: pl.save(); break;
          case 0x0127: pl.restore(i16(0)); break;
          case 0x0214: s.x = i16(1); s.y = i16(0); break;
          case 0x0213: { const x = i16(1), y = i16(0); pl.poly([[s.x, s.y], [x, y]], false); s.x = x; s.y = y; break; }
          case 0x0416: { /* INTERSECTCLIPRECT: bottom, right, top, left */
            const b = i16(0), r = i16(1), tt = i16(2), ll = i16(3);
            const q0 = pl.pt(ll, tt), q1 = pl.pt(r, b);
            const nc = [Math.min(q0[0], q1[0]), Math.min(q0[1], q1[1]), Math.max(q0[0], q1[0]), Math.max(q0[1], q1[1])];
            s.clip = s.clip ? [Math.max(s.clip[0], nc[0]), Math.max(s.clip[1], nc[1]), Math.min(s.clip[2], nc[2]), Math.min(s.clip[3], nc[3])] : nc;
            break; }
          case 0x012c: s.clip = null; break;
          case 0x041b: pl.rect(i16(3), i16(2), i16(1), i16(0)); break;
          case 0x061c: pl.rect(i16(5), i16(4), i16(3), i16(2), i16(1), i16(0)); break;
          case 0x0418: pl.ellipse(i16(3), i16(2), i16(1), i16(0)); break;
          case 0x0817: pl.arc(i16(7), i16(6), i16(5), i16(4), i16(3), i16(2), i16(1), i16(0), 'arc'); break;
          case 0x081a: pl.arc(i16(7), i16(6), i16(5), i16(4), i16(3), i16(2), i16(1), i16(0), 'pie'); break;
          case 0x0830: pl.arc(i16(7), i16(6), i16(5), i16(4), i16(3), i16(2), i16(1), i16(0), 'chord'); break;
          case 0x0324: case 0x0325: { const cnt = i16(0); const pts = []; for (let i = 0; i < cnt; i++) pts.push([i16(1 + i * 2), i16(2 + i * 2)]); pl.poly(pts, fn === 0x0324); break; }
          case 0x0538: {
            const np = u16(0); const counts = []; for (let i = 0; i < np; i++) counts.push(u16(1 + i));
            let k2 = 1 + np; const polys = [];
            for (const cnt of counts) { const pts = []; for (let i = 0; i < cnt; i++) { pts.push([i16(k2), i16(k2 + 1)]); k2 += 2; } polys.push(pts); }
            pl.polyPoly(polys, true); break;
          }
          case 0x061d: { /* PATBLT: rop, h, w, y, x */ pl.patBlt(u32(0), i16(5), i16(4), i16(3), i16(2)); break; }
          case 0x02fa: add({ kind: 'pen', style: u16(0), w: i16(1), color: u32(3) }); break;
          case 0x02fc: add({ kind: 'brush', style: u16(0), color: u32(1), hatch: u16(3) }); break;
          case 0x02fb: {
            const fb = u8.subarray(a + 18, Math.min(a + 18 + 32, p + size)); const z = fb.indexOf(0);
            const face = td(z >= 0 ? fb.subarray(0, z) : fb, 0);
            add({ kind: 'font', h: i16(0), esc: i16(2), w: i16(4), i: !!u8[a + 10], charset: u8[a + 13], face });
            break;
          }
          case 0x0142: { /* DIBCREATEPATTERNBRUSH */ const cv = dibCanvas(u8, dv, a + 4, null, u16(1), p + size, s.palette); add({ kind: 'brush', style: 5, color: 0x808080, pat: cv ? g.createPattern(cv, 'repeat') : null, dib: true }); break; }
          case 0x01f9: add({ kind: 'brush', style: 0, color: 0x808080 }); break;
          case 0x00f7: { const cnt = u16(1), pal = []; for (let i = 0; i < cnt; i++) pal.push([u8[a + 4 + i * 4], u8[a + 5 + i * 4], u8[a + 6 + i * 4]]); add({ kind: 'palette', pal }); break; }
          case 0x0234: { const o = objs[u16(0)]; if (o && o.kind === 'palette') s.palette = o.pal; break; }
          case 0x06ff: add({ kind: 'region' }); break;
          case 0x012d: { const o = objs[u16(0)]; if (o) { if (o.kind === 'pen') s.pen = o; else if (o.kind === 'brush') s.brush = o; else if (o.kind === 'font') s.font = o; } break; }
          case 0x01f0: objs[u16(0)] = null; break;
          case 0x0521: { const cnt = i16(0); const str = td(u8.subarray(a + 2, a + 2 + cnt), s.font.charset, s.font.face); const off = 1 + Math.ceil(cnt / 2); pl.text(i16(off + 1), i16(off), str); break; }
          case 0x0a32: {
            const y = i16(0), x = i16(1), cnt = i16(2), opts = u16(3);
            let o = a + 8 + ((opts & 6) ? 8 : 0);
            const str = td(u8.subarray(o, o + cnt), s.font.charset, s.font.face);
            o += cnt + (cnt & 1);
            const dxs = []; if (o + cnt * 2 <= p + size) for (let i = 0; i < cnt; i++) dxs.push(dv.getInt16(o + i * 2, true));
            pl.text(x, y, str, dxs.length === cnt ? dxs : null);
            break;
          }
          case 0x0f43: { /* STRETCHDIB */ const cv = dibCanvas(u8, dv, a + 22, null, u16(2), p + size, s.palette); pl.blit(cv, i16(6), i16(5), i16(4), i16(3), i16(10), i16(9), i16(8), i16(7), u32(0)); break; }
          case 0x0b41: { /* DIBSTRETCHBLT */ const hasBmp = size > 6 + 22; if (hasBmp) { const cv = dibCanvas(u8, dv, a + 20, null, 0, p + size, s.palette); pl.blit(cv, i16(5), i16(4), i16(3), i16(2), i16(9), i16(8), i16(7), i16(6), u32(0)); } break; }
          case 0x0940: { /* DIBBITBLT: rop, ySrc, xSrc, height, width, yDest, xDest, DIB */ const hasBmp = size > 6 + 18; if (hasBmp) { const cv = dibCanvas(u8, dv, a + 16, null, 0, p + size, s.palette); pl.blit(cv, i16(3), i16(2), i16(5), i16(4), i16(7), i16(6), i16(5), i16(4), u32(0)); } break; }
          default: break;
        }
      } catch (e) { if (L.metafile && L.metafile.debug) console.log("rec err", e.message); }
      p += size;
    }
    return c;
  }

  /* ---------- EMF ---------- */
  function renderEMF(u8, maxPx) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const i32 = (o) => dv.getInt32(o, true), u32 = (o) => dv.getUint32(o, true), i16 = (o) => dv.getInt16(o, true), f32 = (o) => dv.getFloat32(o, true);
    if (u32(0) !== 1) return null;
    const bounds = [i32(8), i32(12), i32(16), i32(20)];
    const frame = [i32(24), i32(28), i32(32), i32(36)];
    const devPx = [i32(72), i32(76)], devMm = [i32(80), i32(84)];
    /* the picture frame (0.01 mm) in device pixels */
    const pxPerMmX = devMm[0] ? devPx[0] / devMm[0] : 96 / 25.4, pxPerMmY = devMm[1] ? devPx[1] / devMm[1] : 96 / 25.4;
    let fr = [(frame[0] / 100) * pxPerMmX, (frame[1] / 100) * pxPerMmY, (frame[2] / 100) * pxPerMmX, (frame[3] / 100) * pxPerMmY];
    if (!(fr[2] - fr[0] > 0) || !(fr[3] - fr[1] > 0)) fr = [bounds[0], bounds[1], bounds[2] + 1, bounds[3] + 1];
    const fw = fr[2] - fr[0], fh = fr[3] - fr[1];
    const natW = (frame[2] - frame[0]) / 100 / 25.4 * 96 || fw, natH = (frame[3] - frame[1]) / 100 / 25.4 * 96 || fh;
    const k = Math.min(maxPx / Math.max(natW, natH), 4) || 1;
    const cw = Math.max(1, Math.round(natW * k)), ch = Math.max(1, Math.round(natH * k));
    const c = document.createElement('canvas');
    c.width = cw; c.height = ch;
    const g = c.getContext('2d');
    const pl = new Player(g, (x, y) => [((x - fr[0]) * cw) / fw, ((y - fr[1]) * ch) / fh]);
    const objs = new Map();
    const STOCK = {
      0: { kind: 'brush', style: 0, color: 0xffffff }, 1: { kind: 'brush', style: 0, color: 0xc0c0c0 }, 2: { kind: 'brush', style: 0, color: 0x808080 }, 3: { kind: 'brush', style: 0, color: 0x404040 },
      4: { kind: 'brush', style: 0, color: 0 }, 5: { kind: 'brush', style: 1, color: 0 }, 6: { kind: 'pen', style: 0, w: 1, color: 0xffffff }, 7: { kind: 'pen', style: 0, w: 1, color: 0 },
      8: { kind: 'pen', style: 5, w: 1, color: 0 }, 13: { kind: 'font', h: -16, w: 700, face: 'Arial' }, 17: { kind: 'font', h: -11, w: 400, face: 'Tahoma' }, 18: { kind: 'brush', style: 0, color: 0xffffff }, 19: { kind: 'pen', style: 0, w: 1, color: 0 },
    };
    const sel = (ih) => {
      const o = ih & 0x80000000 ? STOCK[ih & 0x7fffffff] : objs.get(ih);
      if (!o) return;
      if (o.kind === 'pen') pl.s.pen = o; else if (o.kind === 'brush') pl.s.brush = o; else if (o.kind === 'font') pl.s.font = o;
    };
    const patBlt = (rop, x, y, w, h) => pl.patBlt(rop, x, y, w, h);
    const pts16 = (o, cnt) => { const a = []; for (let i = 0; i < cnt; i++) a.push([i16(o + i * 4), i16(o + i * 4 + 2)]); return a; };
    const pts32 = (o, cnt) => { const a = []; for (let i = 0; i < cnt; i++) a.push([i32(o + i * 8), i32(o + i * 8 + 4)]); return a; };
    const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
    let p = u32(4), n = 0;
    while (p + 8 <= u8.length && n++ < 400000) {
      const type = u32(p), size = u32(p + 4);
      if (size < 8 || p + size > u8.length + 4) break;
      const d = p + 8;
      const s = pl.s;
      try {
        switch (type) {
          case 14: p = u8.length; continue;
          case 9: s.we = [i32(d), i32(d + 4)]; break;
          case 10: s.wo = [i32(d), i32(d + 4)]; break;
          case 11: s.ve = [i32(d), i32(d + 4)]; break;
          case 12: s.vo = [i32(d), i32(d + 4)]; break;
          case 17: s.mapMode = u32(d); break;
          case 18: s.bkMode = u32(d); break;
          case 19: s.fill = u32(d); break;
          case 22: s.align = u32(d); break;
          case 24: s.text = u32(d); break;
          case 25: s.bk = u32(d); break;
          case 27: s.x = i32(d); s.y = i32(d + 4); if (pl.path) { const q = pl.pt(s.x, s.y); pl.path.push((gg) => gg.moveTo(q[0], q[1])); } break;
          case 54: { const x = i32(d), y = i32(d + 4); if (pl.path) { const q = pl.pt(x, y); pl.path.push((gg) => gg.lineTo(q[0], q[1])); } else pl.poly([[s.x, s.y], [x, y]], false); s.x = x; s.y = y; break; }
          case 33: pl.save(); break;
          case 34: pl.restore(i32(d)); break;
          case 35: s.xf = [f32(d), f32(d + 4), f32(d + 8), f32(d + 12), f32(d + 16), f32(d + 20)]; break;
          case 36: { const m = [f32(d), f32(d + 4), f32(d + 8), f32(d + 12), f32(d + 16), f32(d + 20)], mode = u32(d + 24); s.xf = mode === 1 ? [1, 0, 0, 1, 0, 0] : mode === 2 ? mul(m, s.xf) : mode === 3 ? mul(s.xf, m) : m; break; }
          case 30: { const q0 = pl.pt(i32(d), i32(d + 4)), q1 = pl.pt(i32(d + 8), i32(d + 12)); const nc = [Math.min(q0[0], q1[0]), Math.min(q0[1], q1[1]), Math.max(q0[0], q1[0]), Math.max(q0[1], q1[1])]; s.clip = s.clip ? [Math.max(s.clip[0], nc[0]), Math.max(s.clip[1], nc[1]), Math.min(s.clip[2], nc[2]), Math.min(s.clip[3], nc[3])] : nc; break; }
          case 75: if (u32(d) === 0 || u32(d + 4) === 5) s.clip = null; break;
          case 43: pl.rect(i32(d), i32(d + 4), i32(d + 8), i32(d + 12)); break;
          case 44: pl.rect(i32(d), i32(d + 4), i32(d + 8), i32(d + 12), i32(d + 16), i32(d + 20)); break;
          case 42: pl.ellipse(i32(d), i32(d + 4), i32(d + 8), i32(d + 12)); break;
          case 45: case 46: case 47: case 55:
            pl.arc(i32(d), i32(d + 4), i32(d + 8), i32(d + 12), i32(d + 16), i32(d + 20), i32(d + 24), i32(d + 28), type === 47 ? 'pie' : type === 46 ? 'chord' : 'arc'); break;
          case 3: case 4: pl.poly(pts32(d + 20, u32(d + 16)), type === 3); break;
          case 86: case 87: pl.poly(pts16(d + 20, u32(d + 16)), type === 86); break;
          case 2: pl.bezier(pts32(d + 20, u32(d + 16)), false); break;
          case 85: pl.bezier(pts16(d + 20, u32(d + 16)), false); break;
          case 5: pl.bezier(pts32(d + 20, u32(d + 16)), true); break;
          case 88: pl.bezier(pts16(d + 20, u32(d + 16)), true); break;
          case 6: case 89: {
            const cnt = u32(d + 16), pts = type === 6 ? pts32(d + 20, cnt) : pts16(d + 20, cnt);
            if (pl.path) { const P = pts.map(([x, y]) => pl.pt(x, y)); pl.path.push((gg) => { for (const q of P) gg.lineTo(q[0], q[1]); }); }
            else pl.poly([[s.x, s.y]].concat(pts), false);
            if (cnt) { s.x = pts[cnt - 1][0]; s.y = pts[cnt - 1][1]; }
            break;
          }
          case 7: case 8: case 90: case 91: {
            const np = u32(d + 16), counts = []; for (let i = 0; i < np; i++) counts.push(u32(d + 24 + i * 4));
            let o = d + 24 + np * 4; const polys = []; const small = type >= 90;
            for (const cnt of counts) { polys.push(small ? pts16(o, cnt) : pts32(o, cnt)); o += cnt * (small ? 4 : 8); }
            pl.polyPoly(polys, type === 8 || type === 91); break;
          }
          case 59: pl.path = []; break;
          case 60: pl.lastPath = pl.path; pl.path = null; break;
          case 61: if (pl.path) pl.path.push((gg) => gg.closePath()); break;
          case 62: pl.path = pl.lastPath; pl.fillPath('fill'); pl.lastPath = null; break;
          case 63: pl.path = pl.lastPath; pl.fillPath('both'); pl.lastPath = null; break;
          case 64: pl.path = pl.lastPath; pl.fillPath('stroke'); pl.lastPath = null; break;
          case 68: pl.path = null; pl.lastPath = null; break;
          case 38: objs.set(u32(d), { kind: 'pen', style: u32(d + 4), w: i32(d + 8), color: u32(d + 16) }); break;
          case 95: { const st = u32(d + 20), geometric = (st & 0xf0000) === 0x10000; objs.set(u32(d), { kind: 'pen', style: u32(d + 28) === 1 ? 5 : st & 0xffff, w: geometric ? u32(d + 24) : 0, color: u32(d + 32) }); break; }
          case 39: objs.set(u32(d), { kind: 'brush', style: u32(d + 4), color: u32(d + 8), hatch: u32(d + 12) }); break;
          case 94: { const cv = dibCanvas(u8, dv, p + u32(d + 8), p + u32(d + 16), u32(d + 4)); objs.set(u32(d), { kind: 'brush', style: 6, color: 0x808080, pat: cv ? g.createPattern(cv, 'repeat') : null, dib: true }); break; }
          case 93: objs.set(u32(d), { kind: 'brush', style: 0, color: 0x808080 }); break;
          case 82: {
            const lf = d + 4;
            let face = '';
            for (let i = 0; i < 32; i++) { const ch2 = dv.getUint16(lf + 28 + i * 2, true); if (!ch2) break; face += String.fromCharCode(ch2); }
            objs.set(u32(d), { kind: 'font', h: i32(lf), esc: i32(lf + 8), w: i32(lf + 16), i: !!u8[lf + 20], charset: u8[lf + 23], face: face || 'Arial' });
            break;
          }
          case 37: sel(u32(d)); break;
          case 40: objs.delete(u32(d)); break;
          case 84: case 83: {
            const em = d + 28; /* EMRTEXT after rclBounds(16) + iGraphicsMode(4) + exScale(4) + eyScale(4) */
            const x = i32(em), y = i32(em + 4), cnt = u32(em + 8), offStr = u32(em + 12), offDx = u32(em + 32);
            let str = '';
            if (type === 84) { for (let i = 0; i < cnt; i++) str += String.fromCharCode(dv.getUint16(p + offStr + i * 2, true)); }
            else str = new TextDecoder('windows-1252').decode(u8.subarray(p + offStr, p + offStr + cnt));
            const dxs = []; if (offDx && p + offDx + cnt * 4 <= p + size) for (let i = 0; i < cnt; i++) dxs.push(i32(p + offDx + i * 4));
            pl.text(x, y, str, dxs.length === cnt ? dxs : null);
            break;
          }
          case 81: { /* STRETCHDIBITS */
            const xDest = i32(d + 16), yDest = i32(d + 20), xSrc = i32(d + 24), ySrc = i32(d + 28), cxSrc = i32(d + 32), cySrc = i32(d + 36);
            const offBmi = u32(d + 40), offBits = u32(d + 48), usage = u32(d + 56), cxDest = i32(d + 64), cyDest = i32(d + 68);
            const cv = dibCanvas(u8, dv, p + offBmi, p + offBits, usage);
            pl.blit(cv, xSrc, ySrc, cxSrc, cySrc, xDest, yDest, cxDest, cyDest, u32(d + 60));
            break;
          }
          case 76: case 77: { /* BITBLT / STRETCHBLT */
            const xDest = i32(d + 16), yDest = i32(d + 20), cxDest = i32(d + 24), cyDest = i32(d + 28);
            const xSrc = i32(d + 36), ySrc = i32(d + 40), usage = u32(d + 68), offBmi = u32(d + 72), cbBmi = u32(d + 76), offBits = u32(d + 80);
            if (!cbBmi) { patBlt(u32(d + 32), xDest, yDest, cxDest, cyDest); break; }
            const cv = dibCanvas(u8, dv, p + offBmi, p + offBits, usage);
            const cxSrc = type === 77 ? i32(d + 88) : cxDest, cySrc = type === 77 ? i32(d + 92) : cyDest;
            pl.blit(cv, xSrc, ySrc, cxSrc, cySrc, xDest, yDest, cxDest, cyDest, u32(d + 32));
            break;
          }
          default: break;
        }
      } catch (e) { if (L.metafile && L.metafile.debug) console.log("rec err", e.message); }
      p += size;
    }
    return c;
  }

  /* ---------- public ---------- */
  L.metafile = {
    /** bytes: Uint8Array; kind: 'wmf' | 'emf'. Resolves to a PNG Blob, or null if it can't be drawn. */
    async toPNG(bytes, kind, maxPx) {
      try {
        const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
        const isEmf = kind === 'emf' || (u8.length > 44 && new DataView(u8.buffer, u8.byteOffset).getUint32(40, true) === 0x464d4520);
        const c = isEmf ? renderEMF(u8, maxPx || 1200) : renderWMF(u8, maxPx || 1200);
        if (!c) return null;
        return await new Promise((res) => c.toBlob((b) => res(b), 'image/png'));
      } catch (e) {
        console.warn('metafile', e);
        return null;
      }
    },
  };
})(typeof window !== 'undefined' ? window : globalThis);
