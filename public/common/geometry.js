/* VibeOffice — AutoShape geometry, shared by Quire, Ledger and Lectern.
 * Preset names follow the OOXML prstGeom vocabulary so shapes round-trip unchanged.
 * Each preset returns SVG path parts in a w×h box.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const n = (v) => (Math.abs(v) < 1e-9 ? 0 : Math.round(v * 100) / 100);
  const P = (pts) => 'M' + pts.map((p) => n(p[0]) + ',' + n(p[1])).join(' L') + ' Z';
  const ell = (cx, cy, rx, ry) => `M${n(cx - rx)},${n(cy)} A${n(rx)},${n(ry)} 0 1 1 ${n(cx + rx)},${n(cy)} A${n(rx)},${n(ry)} 0 1 1 ${n(cx - rx)},${n(cy)} Z`;
  const rrect = (x, y, w, h, r) => {
    r = Math.max(0, Math.min(r, w / 2, h / 2));
    if (!r) return P([[x, y], [x + w, y], [x + w, y + h], [x, y + h]]);
    return `M${n(x + r)},${n(y)} H${n(x + w - r)} A${n(r)},${n(r)} 0 0 1 ${n(x + w)},${n(y + r)} V${n(y + h - r)} A${n(r)},${n(r)} 0 0 1 ${n(x + w - r)},${n(y + h)} H${n(x + r)} A${n(r)},${n(r)} 0 0 1 ${n(x)},${n(y + h - r)} V${n(y + r)} A${n(r)},${n(r)} 0 0 1 ${n(x + r)},${n(y)} Z`;
  };
  const one = (d, extra) => [Object.assign({ d, fill: 'norm', stroke: true }, extra || {})];
  const ss = (w, h) => Math.min(w, h);
  /** point on the ellipse inscribed in w×h at visual angle deg (clockwise from 3 o'clock), as DrawingML computes it */
  const ePt = (w, h, deg, rx, ry) => {
    rx = rx == null ? w / 2 : rx; ry = ry == null ? h / 2 : ry;
    const t = (deg * Math.PI) / 180;
    const phi = Math.atan2(rx * Math.sin(t), ry * Math.cos(t));
    return [w / 2 + rx * Math.cos(phi), h / 2 + ry * Math.sin(phi)];
  };
  const sweepOf = (st, en) => { const s = (((en - st) % 360) + 360) % 360; return s === 0 ? 360 : s; };
  /** elliptical arc segment (as an SVG 'A' command) from angle st to en, clockwise */
  const arcSeg = (w, h, st, en, rx, ry) => {
    rx = rx == null ? w / 2 : rx; ry = ry == null ? h / 2 : ry;
    const sw = sweepOf(st, en);
    if (sw >= 359.99) { const m = ePt(w, h, st + 180, rx, ry), e = ePt(w, h, st, rx, ry); return `A${n(rx)},${n(ry)} 0 1 1 ${n(m[0])},${n(m[1])} A${n(rx)},${n(ry)} 0 1 1 ${n(e[0])},${n(e[1])}`; }
    const e = ePt(w, h, en, rx, ry);
    return `A${n(rx)},${n(ry)} 0 ${sw > 180 ? 1 : 0} 1 ${n(e[0])},${n(e[1])}`;
  };
  const arcSegCCW = (w, h, st, en, rx, ry) => {
    rx = rx == null ? w / 2 : rx; ry = ry == null ? h / 2 : ry;
    const sw = sweepOf(en, st);
    const e = ePt(w, h, en, rx, ry);
    return `A${n(rx)},${n(ry)} 0 ${sw > 180 ? 1 : 0} 0 ${n(e[0])},${n(e[1])}`;
  };
  const regPoly = (w, h, k, rot) => { const pts = []; for (let i = 0; i < k; i++) { const t = ((rot == null ? -90 : rot) + (i * 360) / k) * Math.PI / 180; pts.push([w / 2 + (w / 2) * Math.cos(t), h / 2 + (h / 2) * Math.sin(t)]); } return P(pts); };
  function star(w, h, k, inner) {
    const pts = [], cx = w / 2, cy = h / 2;
    for (let i = 0; i < k * 2; i++) {
      const a = -Math.PI / 2 + (i * Math.PI) / k;
      const r = i % 2 ? inner : 1;
      pts.push([cx + Math.cos(a) * cx * r, cy + Math.sin(a) * cy * r]);
    }
    return P(pts);
  }
  const poly21600 = (w, h, arr) => P(arr.map(([x, y]) => [(x * w) / 21600, (y * h) / 21600]));

  /* callout wedge on a rectangle-ish body */
  function wedgeRect(w, h, a, r) {
    const tx = w / 2 + (w * a.adj1) / 100000, ty = h / 2 + (h * a.adj2) / 100000;
    const dx = (tx - w / 2) / (w || 1), dy = (ty - h / 2) / (h || 1);
    const inside = tx >= 0 && tx <= w && ty >= 0 && ty <= h;
    const body = rrect(0, 0, w, h, r);
    if (inside) return body;
    let pts;
    if (Math.abs(dy) >= Math.abs(dx)) {
      const c = L.clamp(tx, w * 0.25, w * 0.75), bw = w * 0.08;
      pts = dy > 0 ? [[c - bw, h], [tx, ty], [c + bw, h]] : [[c + bw, 0], [tx, ty], [c - bw, 0]];
      const y0 = dy > 0 ? h : 0;
      if (dy > 0) return `M${n(r)},0 H${n(w - r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(w)},${n(r)}` : ''} V${n(h - r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(w - r)},${n(h)}` : ''} L${n(pts[2][0])},${n(y0)} L${n(tx)},${n(ty)} L${n(pts[0][0])},${n(y0)} H${n(r)} ${r ? `A${n(r)},${n(r)} 0 0 1 0,${n(h - r)}` : ''} V${n(r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(r)},0` : ''} Z`;
      return `M${n(r)},0 L${n(pts[2][0])},0 L${n(tx)},${n(ty)} L${n(pts[0][0])},0 H${n(w - r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(w)},${n(r)}` : ''} V${n(h - r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(w - r)},${n(h)}` : ''} H${n(r)} ${r ? `A${n(r)},${n(r)} 0 0 1 0,${n(h - r)}` : ''} V${n(r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(r)},0` : ''} Z`;
    }
    const c = L.clamp(ty, h * 0.25, h * 0.75), bh = h * 0.08;
    if (dx > 0) return `M${n(r)},0 H${n(w - r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(w)},${n(r)}` : ''} L${n(w)},${n(c - bh)} L${n(tx)},${n(ty)} L${n(w)},${n(c + bh)} V${n(h - r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(w - r)},${n(h)}` : ''} H${n(r)} ${r ? `A${n(r)},${n(r)} 0 0 1 0,${n(h - r)}` : ''} V${n(r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(r)},0` : ''} Z`;
    return `M${n(r)},0 H${n(w - r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(w)},${n(r)}` : ''} V${n(h - r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(w - r)},${n(h)}` : ''} H${n(r)} ${r ? `A${n(r)},${n(r)} 0 0 1 0,${n(h - r)}` : ''} L0,${n(c + bh)} L${n(tx)},${n(ty)} L0,${n(c - bh)} V${n(r)} ${r ? `A${n(r)},${n(r)} 0 0 1 ${n(r)},0` : ''} Z`;
  }
  function cloudPath(cx, cy, rx, ry, bumps) {
    const pts = [];
    for (let i = 0; i < bumps; i++) {
      const a = (i / bumps) * Math.PI * 2 + 0.3;
      const jitter = 1 + (i % 3 === 0 ? 0.06 : i % 3 === 1 ? -0.04 : 0.02);
      pts.push([cx + Math.cos(a) * rx * 0.86 * jitter, cy + Math.sin(a) * ry * 0.82 * jitter]);
    }
    let d = `M${n(pts[0][0])},${n(pts[0][1])}`;
    for (let i = 0; i < bumps; i++) {
      const p = pts[(i + 1) % bumps];
      const q = pts[i];
      const chord = Math.hypot(p[0] - q[0], p[1] - q[1]);
      const r = chord * 0.62;
      d += ` A${n(r)},${n(r)} 0 0 1 ${n(p[0])},${n(p[1])}`;
    }
    return d + ' Z';
  }

  /* action-button glyphs in a unit box (0..1) */
  function glyph(kind, x, y, s) {
    const g = (pts) => P(pts.map(([u, v]) => [x + u * s, y + v * s]));
    switch (kind) {
      case 'home': return g([[0.5, 0.05], [1, 0.5], [0.85, 0.5], [0.85, 0.95], [0.15, 0.95], [0.15, 0.5], [0, 0.5]]);
      case 'back': return g([[0.1, 0.5], [0.9, 0.05], [0.9, 0.95]]);
      case 'next': return g([[0.9, 0.5], [0.1, 0.05], [0.1, 0.95]]);
      case 'begin': return g([[0.25, 0.5], [0.9, 0.05], [0.9, 0.95]]) + ' ' + g([[0.1, 0.05], [0.2, 0.05], [0.2, 0.95], [0.1, 0.95]]);
      case 'end': return g([[0.75, 0.5], [0.1, 0.05], [0.1, 0.95]]) + ' ' + g([[0.8, 0.05], [0.9, 0.05], [0.9, 0.95], [0.8, 0.95]]);
      case 'info': return ell(x + s / 2, y + s / 2, s / 2, s / 2);
      case 'help': return ell(x + s / 2, y + s / 2, s / 2, s / 2);
      case 'return': return `M${n(x + 0.2 * s)},${n(y + 0.1 * s)} V${n(y + 0.65 * s)} A${n(0.3 * s)},${n(0.3 * s)} 0 0 0 ${n(x + 0.8 * s)},${n(y + 0.65 * s)} V${n(y + 0.35 * s)} H${n(x + 0.95 * s)} L${n(x + 0.7 * s)},${n(y + 0.05 * s)} L${n(x + 0.45 * s)},${n(y + 0.35 * s)} H${n(x + 0.62 * s)} V${n(y + 0.65 * s)} A${n(0.12 * s)},${n(0.12 * s)} 0 0 1 ${n(x + 0.38 * s)},${n(y + 0.65 * s)} V${n(y + 0.1 * s)} Z`;
      case 'doc': return g([[0.2, 0], [0.65, 0], [0.85, 0.2], [0.85, 1], [0.2, 1]]);
      case 'sound': return g([[0.05, 0.35], [0.3, 0.35], [0.6, 0.05], [0.6, 0.95], [0.3, 0.65], [0.05, 0.65]]);
      case 'movie': return g([[0, 0.25], [0.65, 0.25], [0.65, 0.42], [1, 0.25], [1, 0.75], [0.65, 0.58], [0.65, 0.75], [0, 0.75]]);
      default: return '';
    }
  }
  function actionButton(kind) {
    return (w, h) => {
      const s = ss(w, h) * 0.62, x = (w - s) / 2, y = (h - s) / 2;
      const parts = [{ d: P([[0, 0], [w, 0], [w, h], [0, h]]), fill: 'norm', stroke: true }];
      if (kind !== 'blank') {
        parts.push({ d: glyph(kind, x, y, s), fill: 'darken', stroke: false });
        if (kind === 'info') parts.push({ d: ell(x + s / 2, y + s * 0.25, s * 0.07, s * 0.07) + ' ' + P([[x + s * 0.43, y + s * 0.4], [x + s * 0.57, y + s * 0.4], [x + s * 0.57, y + s * 0.8], [x + s * 0.43, y + s * 0.8]]), fill: 'lighten', stroke: false });
        if (kind === 'help') parts.push({ d: `M${n(x + s * 0.35)},${n(y + s * 0.38)} A${n(s * 0.15)},${n(s * 0.15)} 0 1 1 ${n(x + s * 0.55)},${n(y + s * 0.52)} L${n(x + s * 0.55)},${n(y + s * 0.62)} L${n(x + s * 0.45)},${n(y + s * 0.62)} L${n(x + s * 0.45)},${n(y + s * 0.47)} A${n(s * 0.06)},${n(s * 0.06)} 0 1 0 ${n(x + s * 0.43)},${n(y + s * 0.38)} Z ` + ell(x + s * 0.5, y + s * 0.74, s * 0.055, s * 0.055), fill: 'lighten', stroke: false });
        if (kind === 'doc') parts.push({ d: P([[x + 0.65 * s, y], [x + 0.65 * s, y + 0.2 * s], [x + 0.85 * s, y + 0.2 * s]]), fill: 'lighten', stroke: false });
      }
      return parts;
    };
  }

  /* ---------- preset table ---------- */
  const G = {};
  const def = (name, label, fn, opts) => { G[name] = Object.assign({ name, label, fn }, opts || {}); };

  def('rect', 'Rectangle', (w, h) => one(P([[0, 0], [w, 0], [w, h], [0, h]])));
  def('roundRect', 'Rounded Rectangle', (w, h, a) => one(rrect(0, 0, w, h, (ss(w, h) * a.adj) / 100000)), {
    adj: { adj: 16667 },
    handles: [{ pos: (w, h, a) => [(ss(w, h) * a.adj) / 100000, 0], set: (x, y, w, h) => ({ adj: L.clamp((x / ss(w, h)) * 100000, 0, 50000) }) }],
  });
  def('ellipse', 'Oval', (w, h) => one(ell(w / 2, h / 2, w / 2, h / 2)), { text: (w, h) => [w * 0.146, h * 0.146, w * 0.854, h * 0.854] });
  def('triangle', 'Isosceles Triangle', (w, h, a) => one(P([[(w * a.adj) / 100000, 0], [w, h], [0, h]])), {
    adj: { adj: 50000 }, text: (w, h) => [w / 4, h / 2, (w * 3) / 4, h],
    handles: [{ pos: (w, h, a) => [(w * a.adj) / 100000, 0], set: (x, y, w) => ({ adj: L.clamp((x / w) * 100000, 0, 100000) }) }],
  });
  def('rtTriangle', 'Right Triangle', (w, h) => one(P([[0, 0], [w, h], [0, h]])), { text: (w, h) => [w / 12, h * 7 / 12, w * 7 / 12, h * 11 / 12] });
  def('parallelogram', 'Parallelogram', (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return one(P([[x, 0], [w, 0], [w - x, h], [0, h]])); }, {
    adj: { adj: 25000 },
    handles: [{ pos: (w, h, a) => [(ss(w, h) * a.adj) / 100000, 0], set: (x, y, w, h) => ({ adj: L.clamp((x / ss(w, h)) * 100000, 0, (w / ss(w, h)) * 100000) }) }],
  });
  def('trapezoid', 'Trapezoid', (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return one(P([[0, h], [x, 0], [w - x, 0], [w, h]])); }, {
    adj: { adj: 25000 },
    handles: [{ pos: (w, h, a) => [(ss(w, h) * a.adj) / 100000, 0], set: (x, y, w, h) => ({ adj: L.clamp((x / ss(w, h)) * 100000, 0, (w / 2 / ss(w, h)) * 100000) }) }],
  });
  def('diamond', 'Diamond', (w, h) => one(P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]])), { text: (w, h) => [w / 4, h / 4, (w * 3) / 4, (h * 3) / 4] });
  def('pentagon', 'Regular Pentagon', (w, h) => one(P([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]])));
  def('hexagon', 'Hexagon', (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return one(P([[x, 0], [w - x, 0], [w, h / 2], [w - x, h], [x, h], [0, h / 2]])); }, {
    adj: { adj: 25000 },
    handles: [{ pos: (w, h, a) => [(ss(w, h) * a.adj) / 100000, 0], set: (x, y, w, h) => ({ adj: L.clamp((x / ss(w, h)) * 100000, 0, (w / 2 / ss(w, h)) * 100000) }) }],
  });
  def('octagon', 'Octagon', (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return one(P([[x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x], [0, x]])); }, {
    adj: { adj: 29289 },
    handles: [{ pos: (w, h, a) => [(ss(w, h) * a.adj) / 100000, 0], set: (x, y, w, h) => ({ adj: L.clamp((x / ss(w, h)) * 100000, 0, 50000) }) }],
  });
  def('plus', 'Cross', (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return one(P([[x, 0], [w - x, 0], [w - x, x], [w, x], [w, h - x], [w - x, h - x], [w - x, h], [x, h], [x, h - x], [0, h - x], [0, x], [x, x]])); }, {
    adj: { adj: 25000 },
    handles: [{ pos: (w, h, a) => [0, (ss(w, h) * a.adj) / 100000], set: (x, y, w, h) => ({ adj: L.clamp((y / ss(w, h)) * 100000, 0, 50000) }) }],
  });
  def('can', 'Can', (w, h, a) => {
    const e = (ss(w, h) * a.adj) / 100000, r = e / 2;
    return [
      { d: `M0,${n(r)} A${n(w / 2)},${n(r)} 0 0 0 ${n(w)},${n(r)} V${n(h - r)} A${n(w / 2)},${n(r)} 0 0 1 0,${n(h - r)} Z`, fill: 'norm', stroke: true },
      { d: ell(w / 2, r, w / 2, r), fill: 'lighten', stroke: true },
    ];
  }, {
    adj: { adj: 25000 }, text: (w, h, a) => [0, (ss(w, h) * a.adj) / 100000, w, h - (ss(w, h) * a.adj) / 200000],
    handles: [{ pos: (w, h, a) => [w / 2, (ss(w, h) * a.adj) / 100000], set: (x, y, w, h) => ({ adj: L.clamp((y / ss(w, h)) * 100000, 0, (h / ss(w, h)) * 100000) }) }],
  });
  def('cube', 'Cube', (w, h, a) => {
    const d = (ss(w, h) * a.adj) / 100000;
    return [
      { d: P([[0, d], [w - d, d], [w - d, h], [0, h]]), fill: 'norm', stroke: true },
      { d: P([[0, d], [d, 0], [w, 0], [w - d, d]]), fill: 'lightenLess', stroke: true },
      { d: P([[w - d, d], [w, 0], [w, h - d], [w - d, h]]), fill: 'darken', stroke: true },
    ];
  }, {
    adj: { adj: 25000 }, text: (w, h, a) => { const d = (ss(w, h) * a.adj) / 100000; return [0, d, w - d, h]; },
    handles: [{ pos: (w, h, a) => [0, (ss(w, h) * a.adj) / 100000], set: (x, y, w, h) => ({ adj: L.clamp((y / ss(w, h)) * 100000, 0, 100000) }) }],
  });
  def('bevel', 'Bevel', (w, h, a) => {
    const d = (ss(w, h) * a.adj) / 100000;
    return [
      { d: P([[d, d], [w - d, d], [w - d, h - d], [d, h - d]]), fill: 'norm', stroke: true },
      { d: P([[0, 0], [w, 0], [w - d, d], [d, d]]), fill: 'lightenLess', stroke: true },
      { d: P([[0, 0], [d, d], [d, h - d], [0, h]]), fill: 'lighten', stroke: true },
      { d: P([[0, h], [d, h - d], [w - d, h - d], [w, h]]), fill: 'darken', stroke: true },
      { d: P([[w, 0], [w, h], [w - d, h - d], [w - d, d]]), fill: 'darkenLess', stroke: true },
    ];
  }, {
    adj: { adj: 12500 }, text: (w, h, a) => { const d = (ss(w, h) * a.adj) / 100000; return [d, d, w - d, h - d]; },
    handles: [{ pos: (w, h, a) => [(ss(w, h) * a.adj) / 100000, h / 2], set: (x, y, w, h) => ({ adj: L.clamp((x / ss(w, h)) * 100000, 0, 50000) }) }],
  });
  def('foldedCorner', 'Folded Corner', (w, h, a) => {
    const d = (ss(w, h) * a.adj) / 100000;
    return [
      { d: P([[0, 0], [w, 0], [w, h - d], [w - d, h], [0, h]]), fill: 'norm', stroke: true },
      { d: P([[w - d, h], [w - d * 0.8, h - d * 0.8], [w, h - d]]), fill: 'darkenLess', stroke: true },
    ];
  }, { adj: { adj: 16667 } });
  def('frame', 'Frame', (w, h, a) => {
    const t = (ss(w, h) * a.adj1) / 100000;
    return [{ d: P([[0, 0], [w, 0], [w, h], [0, h]]) + ' ' + P([[t, t], [t, h - t], [w - t, h - t], [w - t, t]]), fill: 'norm', stroke: true, rule: 'evenodd' }];
  }, {
    adj: { adj1: 12500 },
    handles: [{ pos: (w, h, a) => [(ss(w, h) * a.adj1) / 100000, 0], set: (x, y, w, h) => ({ adj1: L.clamp((x / ss(w, h)) * 100000, 0, 50000) }) }],
  });
  def('donut', 'Donut', (w, h, a) => {
    const t = (ss(w, h) * a.adj) / 100000;
    return [{ d: ell(w / 2, h / 2, w / 2, h / 2) + ' ' + ell(w / 2, h / 2, Math.max(0, w / 2 - t), Math.max(0, h / 2 - t)), fill: 'norm', stroke: true, rule: 'evenodd' }];
  }, {
    adj: { adj: 25000 }, text: (w, h) => [w * 0.146, h * 0.146, w * 0.854, h * 0.854],
    handles: [{ pos: (w, h, a) => [(ss(w, h) * a.adj) / 100000, h / 2], set: (x, y, w, h) => ({ adj: L.clamp((x / ss(w, h)) * 100000, 0, 50000) }) }],
  });
  def('noSmoking', '"No" Symbol', (w, h, a) => {
    const t = (ss(w, h) * a.adj) / 100000, rx = w / 2 - t, ry = h / 2 - t;
    const k = Math.SQRT1_2, ox = (t / 2) * k * 1.4;
    const band = P([[w / 2 - rx * k - ox, h / 2 - ry * k + ox], [w / 2 - rx * k + ox, h / 2 - ry * k - ox], [w / 2 + rx * k + ox, h / 2 + ry * k - ox], [w / 2 + rx * k - ox, h / 2 + ry * k + ox]]);
    return [{ d: ell(w / 2, h / 2, w / 2, h / 2) + ' ' + ell(w / 2, h / 2, rx, ry), fill: 'norm', stroke: true, rule: 'evenodd' }, { d: band, fill: 'norm', stroke: true }];
  }, { adj: { adj: 18750 } });
  def('smileyFace', 'Smiley Face', (w, h, a) => {
    const sm = (h * a.adj) / 100000;
    return [
      { d: ell(w / 2, h / 2, w / 2, h / 2), fill: 'norm', stroke: true },
      { d: ell(w * 0.34, h * 0.36, w * 0.05, h * 0.05) + ' ' + ell(w * 0.66, h * 0.36, w * 0.05, h * 0.05), fill: 'darkenLess', stroke: true },
      { d: `M${n(w * 0.27)},${n(h * 0.7)} Q${n(w / 2)},${n(h * 0.7 + sm * 4)} ${n(w * 0.73)},${n(h * 0.7)}`, fill: 'none', stroke: true },
    ];
  }, {
    adj: { adj: 4653 },
    handles: [{ pos: (w, h, a) => [w / 2, h * 0.7 + (h * a.adj * 2) / 100000], set: (x, y, w, h) => ({ adj: L.clamp(((y - h * 0.7) / 2 / h) * 100000, -4653, 4653) }) }],
  });
  def('heart', 'Heart', (w, h) => one(`M${n(w / 2)},${n(h * 0.25)} C${n(w / 2)},${n(h * 0.1)} ${n(w * 0.36)},0 ${n(w * 0.24)},0 C${n(w * 0.1)},0 0,${n(h * 0.13)} 0,${n(h * 0.3)} C0,${n(h * 0.58)} ${n(w * 0.3)},${n(h * 0.76)} ${n(w / 2)},${n(h)} C${n(w * 0.7)},${n(h * 0.76)} ${n(w)},${n(h * 0.58)} ${n(w)},${n(h * 0.3)} C${n(w)},${n(h * 0.13)} ${n(w * 0.9)},0 ${n(w * 0.76)},0 C${n(w * 0.64)},0 ${n(w / 2)},${n(h * 0.1)} ${n(w / 2)},${n(h * 0.25)} Z`));
  def('lightningBolt', 'Lightning Bolt', (w, h) => one(poly21600(w, h, [[8458, 0], [12860, 6080], [11050, 6797], [16080, 12382], [14394, 12877], [21600, 21600], [10012, 14915], [12222, 13987], [5022, 9705], [7602, 8382], [0, 3890]])));
  def('sun', 'Sun', (w, h, a) => {
    const r = 0.5 - a.adj / 100000 - 0.06;
    const parts = [{ d: ell(w / 2, h / 2, w * r, h * r), fill: 'norm', stroke: true }];
    let rays = '';
    for (let i = 0; i < 8; i++) {
      const t = (i * Math.PI) / 4, sp = 0.13;
      const tip = [w / 2 + Math.cos(t) * w / 2, h / 2 + Math.sin(t) * h / 2];
      const b1 = [w / 2 + Math.cos(t - sp) * w * (r + 0.05), h / 2 + Math.sin(t - sp) * h * (r + 0.05)];
      const b2 = [w / 2 + Math.cos(t + sp) * w * (r + 0.05), h / 2 + Math.sin(t + sp) * h * (r + 0.05)];
      rays += P([b1, tip, b2]) + ' ';
    }
    parts.push({ d: rays, fill: 'norm', stroke: true });
    return parts;
  }, { adj: { adj: 25000 } });
  def('moon', 'Moon', (w, h, a) => one(`M${n(w)},0 A${n(w)},${n(h / 2)} 0 0 0 ${n(w)},${n(h)} A${n((w * a.adj) / 100000)},${n(h / 2)} 0 0 1 ${n(w)},0 Z`), {
    adj: { adj: 50000 },
    handles: [{ pos: (w, h, a) => [w - (w * a.adj) / 100000, h / 2], set: (x, y, w) => ({ adj: L.clamp(((w - x) / w) * 100000, 0, 87500) }) }],
  });
  def('arc', 'Arc', (w, h, a) => {
    const st = a.adj1 / 60000, en = a.adj2 / 60000;
    const p0 = ePt(w, h, st), p1 = ePt(w, h, en);
    const sw = (((en - st) % 360) + 360) % 360;
    return [{ d: `M${n(p0[0])},${n(p0[1])} A${n(w / 2)},${n(h / 2)} 0 ${sw > 180 ? 1 : 0} 1 ${n(p1[0])},${n(p1[1])}`, fill: 'none', stroke: true, open: true }];
  }, { noFill: true, adj: { adj1: 16200000, adj2: 0 } });
  def('wave', 'Wave', (w, h, a) => {
    const am = (h * a.adj1) / 100000;
    return one(`M0,${n(am)} C${n(w / 3)},${n(-am)} ${n((w * 2) / 3)},${n(am * 3)} ${n(w)},${n(am)} V${n(h - am)} C${n((w * 2) / 3)},${n(h + am)} ${n(w / 3)},${n(h - am * 3)} 0,${n(h - am)} Z`);
  }, { adj: { adj1: 12500 } });

  /* block arrows */
  def('rightArrow', 'Right Arrow', (w, h, a) => {
    const t = (h * a.adj1) / 200000, hx = w - (ss(w, h) * a.adj2) / 100000;
    return one(P([[0, h / 2 - t], [hx, h / 2 - t], [hx, 0], [w, h / 2], [hx, h], [hx, h / 2 + t], [0, h / 2 + t]]));
  }, { adj: { adj1: 50000, adj2: 50000 }, handles: arrowHandles('r') });
  def('leftArrow', 'Left Arrow', (w, h, a) => {
    const t = (h * a.adj1) / 200000, hx = (ss(w, h) * a.adj2) / 100000;
    return one(P([[w, h / 2 - t], [hx, h / 2 - t], [hx, 0], [0, h / 2], [hx, h], [hx, h / 2 + t], [w, h / 2 + t]]));
  }, { adj: { adj1: 50000, adj2: 50000 }, handles: arrowHandles('l') });
  def('upArrow', 'Up Arrow', (w, h, a) => {
    const t = (w * a.adj1) / 200000, hy = (ss(w, h) * a.adj2) / 100000;
    return one(P([[w / 2 - t, h], [w / 2 - t, hy], [0, hy], [w / 2, 0], [w, hy], [w / 2 + t, hy], [w / 2 + t, h]]));
  }, { adj: { adj1: 50000, adj2: 50000 }, handles: arrowHandles('u') });
  def('downArrow', 'Down Arrow', (w, h, a) => {
    const t = (w * a.adj1) / 200000, hy = h - (ss(w, h) * a.adj2) / 100000;
    return one(P([[w / 2 - t, 0], [w / 2 - t, hy], [0, hy], [w / 2, h], [w, hy], [w / 2 + t, hy], [w / 2 + t, 0]]));
  }, { adj: { adj1: 50000, adj2: 50000 }, handles: arrowHandles('d') });
  def('leftRightArrow', 'Left-Right Arrow', (w, h, a) => {
    const t = (h * a.adj1) / 200000, hx = (ss(w, h) * a.adj2) / 100000;
    return one(P([[0, h / 2], [hx, 0], [hx, h / 2 - t], [w - hx, h / 2 - t], [w - hx, 0], [w, h / 2], [w - hx, h], [w - hx, h / 2 + t], [hx, h / 2 + t], [hx, h]]));
  }, { adj: { adj1: 50000, adj2: 50000 }, handles: arrowHandles('l') });
  def('upDownArrow', 'Up-Down Arrow', (w, h, a) => {
    const t = (w * a.adj1) / 200000, hy = (ss(w, h) * a.adj2) / 100000;
    return one(P([[w / 2, 0], [w, hy], [w / 2 + t, hy], [w / 2 + t, h - hy], [w, h - hy], [w / 2, h], [0, h - hy], [w / 2 - t, h - hy], [w / 2 - t, hy], [0, hy]]));
  }, { adj: { adj1: 50000, adj2: 50000 }, handles: arrowHandles('u') });
  def('quadArrow', 'Quad Arrow', (w, h, a) => {
    const s = ss(w, h), t = (s * a.adj1) / 200000, hw = (s * a.adj2) / 100000, hl = (s * a.adj3) / 100000;
    const cx = w / 2, cy = h / 2;
    return one(P([
      [0, cy], [hl, cy - hw], [hl, cy - t], [cx - t, cy - t], [cx - t, hl], [cx - hw, hl], [cx, 0], [cx + hw, hl], [cx + t, hl], [cx + t, cy - t],
      [w - hl, cy - t], [w - hl, cy - hw], [w, cy], [w - hl, cy + hw], [w - hl, cy + t], [cx + t, cy + t], [cx + t, h - hl], [cx + hw, h - hl], [cx, h],
      [cx - hw, h - hl], [cx - t, h - hl], [cx - t, cy + t], [hl, cy + t], [hl, cy + hw],
    ]));
  }, { adj: { adj1: 22500, adj2: 22500, adj3: 22500 } });
  def('notchedRightArrow', 'Notched Right Arrow', (w, h, a) => {
    const t = (h * a.adj1) / 200000, hx = w - (ss(w, h) * a.adj2) / 100000, nx = ((h / 2 - t) * (w - hx)) / (h / 2);
    return one(P([[0, h / 2 - t], [hx, h / 2 - t], [hx, 0], [w, h / 2], [hx, h], [hx, h / 2 + t], [0, h / 2 + t], [nx, h / 2]]));
  }, { adj: { adj1: 50000, adj2: 50000 }, handles: arrowHandles('r') });
  def('stripedRightArrow', 'Striped Right Arrow', (w, h, a) => {
    const t = (h * a.adj1) / 200000, hx = w - (ss(w, h) * a.adj2) / 100000, s = ss(w, h);
    return one(P([[0, h / 2 - t], [s / 32, h / 2 - t], [s / 32, h / 2 + t], [0, h / 2 + t]]) + ' ' + P([[s / 16, h / 2 - t], [s / 8, h / 2 - t], [s / 8, h / 2 + t], [s / 16, h / 2 + t]]) + ' ' + P([[(s * 5) / 32, h / 2 - t], [hx, h / 2 - t], [hx, 0], [w, h / 2], [hx, h], [hx, h / 2 + t], [(s * 5) / 32, h / 2 + t]]));
  }, { adj: { adj1: 50000, adj2: 50000 } });
  def('homePlate', 'Pentagon', (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return one(P([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h]])); }, {
    adj: { adj: 50000 },
    handles: [{ pos: (w, h, a) => [w - (ss(w, h) * a.adj) / 100000, 0], set: (x, y, w, h) => ({ adj: L.clamp(((w - x) / ss(w, h)) * 100000, 0, (w / ss(w, h)) * 100000) }) }],
  });
  def('chevron', 'Chevron', (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return one(P([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h], [x, h / 2]])); }, {
    adj: { adj: 50000 },
    handles: [{ pos: (w, h, a) => [w - (ss(w, h) * a.adj) / 100000, 0], set: (x, y, w, h) => ({ adj: L.clamp(((w - x) / ss(w, h)) * 100000, 0, (w / ss(w, h)) * 100000) }) }],
  });
  def('bentArrow', 'Bent Arrow', (w, h, a) => {
    const s = ss(w, h), t = (s * a.adj1) / 100000, hw = (s * a.adj2) / 100000, hl = (s * a.adj3) / 100000, r = Math.min((s * a.adj4) / 100000, h - hw - t / 2);
    const yc = hw; // centre line of the horizontal shaft
    return one(`M0,${n(h)} V${n(yc - t / 2 + r)} A${n(r)},${n(r)} 0 0 1 ${n(r)},${n(yc - t / 2)} H${n(w - hl)} V${n(yc - hw)} L${n(w)},${n(yc)} L${n(w - hl)},${n(yc + hw)} V${n(yc + t / 2)} H${n(t + Math.max(0, r - t))} ${r > t ? `A${n(r - t)},${n(r - t)} 0 0 0 ${n(t)},${n(yc + t / 2 + r - t)}` : ''} V${n(h)} Z`);
  }, { adj: { adj1: 25000, adj2: 25000, adj3: 25000, adj4: 43750 } });
  def('uturnArrow', 'U-Turn Arrow', (w, h, a) => {
    const s = ss(w, h), t = (s * a.adj1) / 100000, hw = (s * a.adj2) / 100000, hl = (s * a.adj3) / 100000;
    const r = Math.min((w - hw * 2 + t) / 2, h * 0.4), xr = w - hw - t / 2; // right shaft centre
    const top = 0, yHead = (h * a.adj5) / 100000;
    return one(`M0,${n(h)} V${n(top + r)} A${n(r)},${n(r)} 0 0 1 ${n(r)},${n(top)} H${n(xr + t / 2 - r)} A${n(r)},${n(r)} 0 0 1 ${n(xr + t / 2)},${n(top + r)} V${n(yHead - hl)} H${n(w)} L${n(xr)},${n(yHead)} L${n(xr - hw)},${n(yHead - hl)} H${n(xr - t / 2)} V${n(top + r)} A${n(Math.max(0, r - t))},${n(Math.max(0, r - t))} 0 0 0 ${n(xr + t / 2 - r)},${n(top + t)} H${n(r)} A${n(Math.max(0, r - t))},${n(Math.max(0, r - t))} 0 0 0 ${n(t)},${n(top + r)} V${n(h)} Z`);
  }, { adj: { adj1: 25000, adj2: 25000, adj3: 25000, adj4: 43750, adj5: 75000 } });
  def('circularArrow', 'Circular Arrow', (w, h, a) => {
    // a ring segment with arrow head, sweeping clockwise from adj4 to adj5 (degrees*60000)
    const t = (ss(w, h) * a.adj1) / 100000;
    const st = (a.adj4 / 60000) * Math.PI / 180, en = (a.adj5 / 60000) * Math.PI / 180;
    const cx = w / 2, cy = h / 2, rxo = w / 2 - t * 0.5, ryo = h / 2 - t * 0.5, rxi = rxo - t, ryi = ryo - t;
    const pt = (rx, ry, ang) => [cx + rx * Math.cos(ang), cy + ry * Math.sin(ang)];
    let sw = en - st; while (sw <= 0) sw += Math.PI * 2;
    const head = Math.min(0.35, sw * 0.3), body = en - head;
    const large = body - st > Math.PI ? 1 : 0;
    const o1 = pt(rxo, ryo, st), o2 = pt(rxo, ryo, body), i2 = pt(rxi, ryi, body), i1 = pt(rxi, ryi, st);
    const hm = (rxo + rxi) / 2, hmy = (ryo + ryi) / 2;
    const h1 = pt(rxo + t * 0.5, ryo + t * 0.5, body), h2 = pt(rxi - t * 0.5, ryi - t * 0.5, body), tip = pt(hm, hmy, en);
    return one(`M${n(o1[0])},${n(o1[1])} A${n(rxo)},${n(ryo)} 0 ${large} 1 ${n(o2[0])},${n(o2[1])} L${n(h1[0])},${n(h1[1])} L${n(tip[0])},${n(tip[1])} L${n(h2[0])},${n(h2[1])} L${n(i2[0])},${n(i2[1])} A${n(rxi)},${n(ryi)} 0 ${large} 0 ${n(i1[0])},${n(i1[1])} Z`);
  }, { adj: { adj1: 12500, adj2: 1142319, adj3: 20457681, adj4: 10800000, adj5: 0 } });

  /* stars & banners */
  def('star4', '4-Point Star', (w, h, a) => one(star(w, h, 4, a.adj / 50000)), { adj: { adj: 12500 } });
  def('star5', '5-Point Star', (w, h, a) => one(star(w, h, 5, a.adj / 50000)), { adj: { adj: 19098 } });
  def('star6', '6-Point Star', (w, h, a) => one(star(w, h, 6, a.adj / 50000)), { adj: { adj: 28868 } });
  def('star8', '8-Point Star', (w, h, a) => one(star(w, h, 8, a.adj / 50000)), { adj: { adj: 37500 } });
  def('star16', '16-Point Star', (w, h, a) => one(star(w, h, 16, a.adj / 50000)), { adj: { adj: 37500 } });
  def('star24', '24-Point Star', (w, h, a) => one(star(w, h, 24, a.adj / 50000)), { adj: { adj: 37500 } });
  def('star32', '32-Point Star', (w, h, a) => one(star(w, h, 32, a.adj / 50000)), { adj: { adj: 37500 } });
  def('irregularSeal1', 'Explosion 1', (w, h) => one(poly21600(w, h, [[10800, 5800], [14522, 0], [14155, 5325], [18380, 4457], [16702, 7315], [21097, 8137], [17607, 10475], [21600, 13290], [16837, 12942], [18145, 18095], [14020, 14457], [13247, 19737], [10532, 14935], [8485, 21600], [7715, 15627], [4762, 17617], [5667, 13937], [135, 14587], [3722, 11775], [0, 8615], [4627, 7617], [370, 2295], [7312, 6320], [8352, 2295]])));
  def('irregularSeal2', 'Explosion 2', (w, h) => one(poly21600(w, h, [[11462, 4342], [14790, 0], [14525, 5777], [18007, 3172], [16380, 6532], [21600, 6645], [16985, 9402], [18270, 11290], [16380, 12310], [18877, 15632], [14640, 14350], [14942, 17370], [12180, 15935], [11612, 18842], [9872, 17370], [8700, 19712], [7527, 18125], [4917, 21600], [4805, 18240], [1285, 17825], [3330, 15370], [0, 12877], [3935, 11592], [1172, 8270], [5372, 7817], [4502, 3625], [8550, 6382], [9722, 1887]])));
  def('ribbon2', 'Up Ribbon', (w, h) => {
    const e = w / 8, f = h / 3;
    return [
      { d: P([[0, h], [e * 1.5, h], [e * 1.5, f * 0.6], [0, f * 0.6], [e * 0.7, f * 1.3]]).replace('Z', '') + ' Z', fill: 'darkenLess', stroke: true },
      { d: P([[w, h], [w - e * 1.5, h], [w - e * 1.5, f * 0.6], [w, f * 0.6], [w - e * 0.7, f * 1.3]]), fill: 'darkenLess', stroke: true },
      { d: P([[e, 0], [w - e, 0], [w - e, h - f * 0.6], [e, h - f * 0.6]]), fill: 'norm', stroke: true },
    ];
  });
  def('verticalScroll', 'Vertical Scroll', (w, h) => {
    const r = Math.min(w, h) / 16;
    return [
      { d: P([[r * 2, r * 2], [w - r, r * 2], [w - r, h - r], [r * 2, h - r]]), fill: 'norm', stroke: true },
      { d: ell(w - r * 2, r * 2, r * 1.2, r * 1.2) + ' ' + ell(r * 2.2, h - r * 1.2, r * 1.2, r * 1.2), fill: 'darkenLess', stroke: true },
    ];
  });
  def('horizontalScroll', 'Horizontal Scroll', (w, h) => {
    const r = Math.min(w, h) / 16;
    return [
      { d: P([[r, r * 2], [w - r * 2, r * 2], [w - r * 2, h - r], [r, h - r]]), fill: 'norm', stroke: true },
      { d: ell(w - r * 1.6, r * 2.4, r * 1.2, r * 1.2) + ' ' + ell(r * 1.6, h - r * 1.6, r * 1.2, r * 1.2), fill: 'darkenLess', stroke: true },
    ];
  });

  /* flowchart */
  def('flowChartProcess', 'Flowchart: Process', (w, h) => one(P([[0, 0], [w, 0], [w, h], [0, h]])));
  def('flowChartAlternateProcess', 'Flowchart: Alternate Process', (w, h) => one(rrect(0, 0, w, h, ss(w, h) / 6)));
  def('flowChartDecision', 'Flowchart: Decision', (w, h) => one(P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]])), { text: (w, h) => [w / 4, h / 4, (w * 3) / 4, (h * 3) / 4] });
  def('flowChartInputOutput', 'Flowchart: Data', (w, h) => one(P([[w / 5, 0], [w, 0], [(w * 4) / 5, h], [0, h]])));
  def('flowChartPredefinedProcess', 'Flowchart: Predefined Process', (w, h) => [{ d: P([[0, 0], [w, 0], [w, h], [0, h]]), fill: 'norm', stroke: true }, { d: `M${n(w / 8)},0 V${n(h)} M${n((w * 7) / 8)},0 V${n(h)}`, fill: 'none', stroke: true, open: true }]);
  def('flowChartInternalStorage', 'Flowchart: Internal Storage', (w, h) => [{ d: P([[0, 0], [w, 0], [w, h], [0, h]]), fill: 'norm', stroke: true }, { d: `M${n(w / 8)},0 V${n(h)} M0,${n(h / 8)} H${n(w)}`, fill: 'none', stroke: true, open: true }]);
  def('flowChartDocument', 'Flowchart: Document', (w, h) => one(`M0,0 H${n(w)} V${n(h * 0.8)} C${n(w * 0.75)},${n(h * 0.62)} ${n(w * 0.5)},${n(h * 1.08)} 0,${n(h * 0.9)} Z`));
  def('flowChartMultidocument', 'Flowchart: Multidocument', (w, h) => [
    { d: P([[w * 0.1, 0], [w, 0], [w, h * 0.7], [w * 0.9, h * 0.7], [w * 0.9, h * 0.1], [w * 0.1, h * 0.1]]), fill: 'norm', stroke: true },
    { d: P([[w * 0.05, h * 0.05], [w * 0.95, h * 0.05], [w * 0.95, h * 0.75], [w * 0.85, h * 0.75], [w * 0.85, h * 0.15], [w * 0.05, h * 0.15]]), fill: 'norm', stroke: true },
    { d: `M0,${n(h * 0.15)} H${n(w * 0.85)} V${n(h * 0.82)} C${n(w * 0.65)},${n(h * 0.7)} ${n(w * 0.42)},${n(h * 1.06)} 0,${n(h * 0.9)} Z`, fill: 'norm', stroke: true },
  ]);
  def('flowChartTerminator', 'Flowchart: Terminator', (w, h) => { const r = Math.min(h / 2, w / 2); return one(`M${n(r)},0 H${n(w - r)} A${n(r)},${n(h / 2)} 0 0 1 ${n(w - r)},${n(h)} H${n(r)} A${n(r)},${n(h / 2)} 0 0 1 ${n(r)},0 Z`); });
  def('flowChartPreparation', 'Flowchart: Preparation', (w, h) => one(P([[w / 5, 0], [(w * 4) / 5, 0], [w, h / 2], [(w * 4) / 5, h], [w / 5, h], [0, h / 2]])));
  def('flowChartManualInput', 'Flowchart: Manual Input', (w, h) => one(P([[0, h / 5], [w, 0], [w, h], [0, h]])));
  def('flowChartManualOperation', 'Flowchart: Manual Operation', (w, h) => one(P([[0, 0], [w, 0], [(w * 4) / 5, h], [w / 5, h]])));
  def('flowChartConnector', 'Flowchart: Connector', (w, h) => one(ell(w / 2, h / 2, w / 2, h / 2)), { text: (w, h) => [w * 0.146, h * 0.146, w * 0.854, h * 0.854] });
  def('flowChartOffpageConnector', 'Flowchart: Off-page Connector', (w, h) => one(P([[0, 0], [w, 0], [w, h * 0.8], [w / 2, h], [0, h * 0.8]])));
  def('flowChartPunchedCard', 'Flowchart: Card', (w, h) => one(P([[w / 5, 0], [w, 0], [w, h], [0, h], [0, h / 5]])));
  def('flowChartPunchedTape', 'Flowchart: Punched Tape', (w, h) => one(`M0,${n(h * 0.1)} C${n(w * 0.25)},${n(h * 0.3)} ${n(w * 0.25)},${n(-h * 0.1)} ${n(w / 2)},${n(h * 0.1)} C${n(w * 0.75)},${n(h * 0.3)} ${n(w * 0.75)},${n(-h * 0.1)} ${n(w)},${n(h * 0.1)} V${n(h * 0.9)} C${n(w * 0.75)},${n(h * 0.7)} ${n(w * 0.75)},${n(h * 1.1)} ${n(w / 2)},${n(h * 0.9)} C${n(w * 0.25)},${n(h * 0.7)} ${n(w * 0.25)},${n(h * 1.1)} 0,${n(h * 0.9)} Z`));
  def('flowChartSummingJunction', 'Flowchart: Summing Junction', (w, h) => { const k = Math.SQRT1_2; return [{ d: ell(w / 2, h / 2, w / 2, h / 2), fill: 'norm', stroke: true }, { d: `M${n(w / 2 - (w / 2) * k)},${n(h / 2 - (h / 2) * k)} L${n(w / 2 + (w / 2) * k)},${n(h / 2 + (h / 2) * k)} M${n(w / 2 + (w / 2) * k)},${n(h / 2 - (h / 2) * k)} L${n(w / 2 - (w / 2) * k)},${n(h / 2 + (h / 2) * k)}`, fill: 'none', stroke: true, open: true }]; });
  def('flowChartOr', 'Flowchart: Or', (w, h) => [{ d: ell(w / 2, h / 2, w / 2, h / 2), fill: 'norm', stroke: true }, { d: `M${n(w / 2)},0 V${n(h)} M0,${n(h / 2)} H${n(w)}`, fill: 'none', stroke: true, open: true }]);
  def('flowChartCollate', 'Flowchart: Collate', (w, h) => one(P([[0, 0], [w, 0], [0, h], [w, h]])));
  def('flowChartSort', 'Flowchart: Sort', (w, h) => [{ d: P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]), fill: 'norm', stroke: true }, { d: `M0,${n(h / 2)} H${n(w)}`, fill: 'none', stroke: true, open: true }]);
  def('flowChartExtract', 'Flowchart: Extract', (w, h) => one(P([[w / 2, 0], [w, h], [0, h]])));
  def('flowChartMerge', 'Flowchart: Merge', (w, h) => one(P([[0, 0], [w, 0], [w / 2, h]])));
  def('flowChartDelay', 'Flowchart: Delay', (w, h) => one(`M0,0 H${n(w / 2)} A${n(w / 2)},${n(h / 2)} 0 0 1 ${n(w / 2)},${n(h)} H0 Z`));
  def('flowChartMagneticDisk', 'Flowchart: Magnetic Disk', (w, h) => { const r = h / 6; return [{ d: `M0,${n(r)} A${n(w / 2)},${n(r)} 0 0 0 ${n(w)},${n(r)} V${n(h - r)} A${n(w / 2)},${n(r)} 0 0 1 0,${n(h - r)} Z`, fill: 'norm', stroke: true }, { d: ell(w / 2, r, w / 2, r), fill: 'norm', stroke: true }]; });
  def('flowChartDisplay', 'Flowchart: Display', (w, h) => one(`M0,${n(h / 2)} L${n(w / 6)},0 H${n((w * 5) / 6)} A${n(w / 6)},${n(h / 2)} 0 0 1 ${n((w * 5) / 6)},${n(h)} H${n(w / 6)} Z`));

  /* callouts */
  const calloutHandles = [{ pos: (w, h, a) => [w / 2 + (w * a.adj1) / 100000, h / 2 + (h * a.adj2) / 100000], set: (x, y, w, h) => ({ adj1: ((x - w / 2) / w) * 100000, adj2: ((y - h / 2) / h) * 100000 }) }];
  def('wedgeRectCallout', 'Rectangular Callout', (w, h, a) => one(wedgeRect(w, h, a, 0)), { adj: { adj1: -20833, adj2: 62500 }, handles: calloutHandles });
  def('wedgeRoundRectCallout', 'Rounded Rectangular Callout', (w, h, a) => one(wedgeRect(w, h, a, ss(w, h) / 6)), { adj: { adj1: -20833, adj2: 62500 }, handles: calloutHandles });
  def('wedgeEllipseCallout', 'Oval Callout', (w, h, a) => {
    const tx = w / 2 + (w * a.adj1) / 100000, ty = h / 2 + (h * a.adj2) / 100000;
    const ang = Math.atan2((ty - h / 2) / h, (tx - w / 2) / w), sp = 0.18;
    const p1 = [w / 2 + (Math.cos(ang - sp) * w) / 2, h / 2 + (Math.sin(ang - sp) * h) / 2];
    const p2 = [w / 2 + (Math.cos(ang + sp) * w) / 2, h / 2 + (Math.sin(ang + sp) * h) / 2];
    return one(`M${n(p2[0])},${n(p2[1])} A${n(w / 2)},${n(h / 2)} 0 1 1 ${n(p1[0])},${n(p1[1])} L${n(tx)},${n(ty)} Z`);
  }, { adj: { adj1: -20833, adj2: 62500 }, handles: calloutHandles, text: (w, h) => [w * 0.146, h * 0.146, w * 0.854, h * 0.854] });
  def('cloudCallout', 'Cloud Callout', (w, h, a) => {
    const tx = w / 2 + (w * a.adj1) / 100000, ty = h / 2 + (h * a.adj2) / 100000;
    const parts = [{ d: cloudPath(w / 2, h / 2, w / 2, h / 2, 11), fill: 'norm', stroke: true }];
    const dx = tx - w / 2, dy = ty - h / 2, len = Math.hypot(dx / w, dy / h);
    if (len > 0.5) {
      const ex = w / 2 + dx * (0.5 / len), ey = h / 2 + dy * (0.5 / len);
      const s = ss(w, h);
      let d = '';
      [[0.15, 0.11], [0.5, 0.07], [0.88, 0.04]].forEach(([t, r]) => { d += ell(ex + (tx - ex) * t, ey + (ty - ey) * t, s * r, s * r * 0.8) + ' '; });
      d += ell(tx, ty, s * 0.02, s * 0.02);
      parts.push({ d, fill: 'norm', stroke: true });
    }
    return parts;
  }, { adj: { adj1: -20833, adj2: 62500 }, handles: calloutHandles, text: (w, h) => [w * 0.15, h * 0.15, w * 0.85, h * 0.85] });
  def('cloud', 'Cloud', (w, h) => one(cloudPath(w / 2, h / 2, w / 2, h / 2, 11)), { text: (w, h) => [w * 0.15, h * 0.15, w * 0.85, h * 0.85] });

  /* brackets */
  def('leftBracket', 'Left Bracket', (w, h, a) => { const r = Math.min((ss(w, h) * a.adj) / 100000, h / 2); const d = `M${n(w)},0 A${n(w)},${n(r)} 0 0 0 0,${n(r)} V${n(h - r)} A${n(w)},${n(r)} 0 0 0 ${n(w)},${n(h)}`; return [{ d: d + ' Z', fill: 'norm', stroke: false }, { d, fill: 'none', stroke: true, open: true }]; }, { adj: { adj: 8333 }, line: true });
  def('rightBracket', 'Right Bracket', (w, h, a) => { const r = Math.min((ss(w, h) * a.adj) / 100000, h / 2); const d = `M0,0 A${n(w)},${n(r)} 0 0 1 ${n(w)},${n(r)} V${n(h - r)} A${n(w)},${n(r)} 0 0 1 0,${n(h)}`; return [{ d: d + ' Z', fill: 'norm', stroke: false }, { d, fill: 'none', stroke: true, open: true }]; }, { adj: { adj: 8333 }, line: true });
  def('leftBrace', 'Left Brace', (w, h, a) => { const r = Math.min((ss(w, h) * a.adj1) / 100000, h / 4); const d = `M${n(w)},0 A${n(w / 2)},${n(r)} 0 0 0 ${n(w / 2)},${n(r)} V${n(h / 2 - r)} A${n(w / 2)},${n(r)} 0 0 1 0,${n(h / 2)} A${n(w / 2)},${n(r)} 0 0 1 ${n(w / 2)},${n(h / 2 + r)} V${n(h - r)} A${n(w / 2)},${n(r)} 0 0 0 ${n(w)},${n(h)}`; return [{ d: d + ' Z', fill: 'norm', stroke: false }, { d, fill: 'none', stroke: true, open: true }]; }, { adj: { adj1: 8333, adj2: 50000 }, line: true });
  def('rightBrace', 'Right Brace', (w, h, a) => { const r = Math.min((ss(w, h) * a.adj1) / 100000, h / 4); const d = `M0,0 A${n(w / 2)},${n(r)} 0 0 1 ${n(w / 2)},${n(r)} V${n(h / 2 - r)} A${n(w / 2)},${n(r)} 0 0 0 ${n(w)},${n(h / 2)} A${n(w / 2)},${n(r)} 0 0 0 ${n(w / 2)},${n(h / 2 + r)} V${n(h - r)} A${n(w / 2)},${n(r)} 0 0 1 0,${n(h)}`; return [{ d: d + ' Z', fill: 'norm', stroke: false }, { d, fill: 'none', stroke: true, open: true }]; }, { adj: { adj1: 8333, adj2: 50000 }, line: true });
  def('bracketPair', 'Double Bracket', (w, h, a) => { const r = (ss(w, h) * a.adj) / 100000; const l = `M${n(r)},${n(h)} A${n(r)},${n(r)} 0 0 1 0,${n(h - r)} V${n(r)} A${n(r)},${n(r)} 0 0 1 ${n(r)},0`, rr = `M${n(w - r)},0 A${n(r)},${n(r)} 0 0 1 ${n(w)},${n(r)} V${n(h - r)} A${n(r)},${n(r)} 0 0 1 ${n(w - r)},${n(h)}`; return [{ d: rrect(0, 0, w, h, r), fill: 'norm', stroke: false }, { d: l + ' ' + rr, fill: 'none', stroke: true, open: true }]; }, { adj: { adj: 16667 }, line: true });
  def('bracePair', 'Double Brace', (w, h, a) => { const r = (ss(w, h) * a.adj) / 100000; const l = `M${n(r * 2)},${n(h)} A${n(r)},${n(r)} 0 0 1 ${n(r)},${n(h - r)} V${n(h / 2 + r)} A${n(r)},${n(r)} 0 0 0 0,${n(h / 2)} A${n(r)},${n(r)} 0 0 0 ${n(r)},${n(h / 2 - r)} V${n(r)} A${n(r)},${n(r)} 0 0 1 ${n(r * 2)},0`; const rr = `M${n(w - r * 2)},0 A${n(r)},${n(r)} 0 0 1 ${n(w - r)},${n(r)} V${n(h / 2 - r)} A${n(r)},${n(r)} 0 0 0 ${n(w)},${n(h / 2)} A${n(r)},${n(r)} 0 0 0 ${n(w - r)},${n(h / 2 + r)} V${n(h - r)} A${n(r)},${n(r)} 0 0 1 ${n(w - r * 2)},${n(h)}`; return [{ d: P([[r, 0], [w - r, 0], [w - r, h], [r, h]]), fill: 'norm', stroke: false }, { d: l + ' ' + rr, fill: 'none', stroke: true, open: true }]; }, { adj: { adj: 8333 }, line: true });

  /* action buttons */
  [['actionButtonBlank', 'Action Button: Custom', 'blank'], ['actionButtonHome', 'Action Button: Home', 'home'], ['actionButtonHelp', 'Action Button: Help', 'help'],
    ['actionButtonInformation', 'Action Button: Information', 'info'], ['actionButtonBackPrevious', 'Action Button: Back or Previous', 'back'],
    ['actionButtonForwardNext', 'Action Button: Forward or Next', 'next'], ['actionButtonBeginning', 'Action Button: Beginning', 'begin'],
    ['actionButtonEnd', 'Action Button: End', 'end'], ['actionButtonReturn', 'Action Button: Return', 'return'], ['actionButtonDocument', 'Action Button: Document', 'doc'],
    ['actionButtonSound', 'Action Button: Sound', 'sound'], ['actionButtonMovie', 'Action Button: Movie', 'movie']].forEach(([k, l, g]) => def(k, l, actionButton(g), { action: g }));

  /* lines & connectors — drawn from (0,0) to (w,h) honoring flips */
  def('line', 'Line', null, { isLine: true });
  def('straightConnector1', 'Straight Connector', null, { isLine: true });
  def('bentConnector3', 'Elbow Connector', null, { isLine: true });
  def('bentConnector2', 'Elbow Connector', null, { isLine: true });
  def('bentConnector4', 'Elbow Connector', null, { isLine: true });
  def('curvedConnector2', 'Curved Connector', null, { isLine: true });
  def('curvedConnector3', 'Curved Connector', null, { isLine: true });

  function arrowHandles(dir) {
    if (dir === 'r') return [
      { pos: (w, h, a) => [w - (ss(w, h) * a.adj2) / 100000, h / 2 - (h * a.adj1) / 200000], set: (x, y, w, h) => ({ adj1: L.clamp(((h / 2 - y) * 200000) / h, 0, 100000), adj2: L.clamp(((w - x) / ss(w, h)) * 100000, 0, (w / ss(w, h)) * 100000) }) }];
    if (dir === 'l') return [
      { pos: (w, h, a) => [(ss(w, h) * a.adj2) / 100000, h / 2 - (h * a.adj1) / 200000], set: (x, y, w, h) => ({ adj1: L.clamp(((h / 2 - y) * 200000) / h, 0, 100000), adj2: L.clamp((x / ss(w, h)) * 100000, 0, (w / ss(w, h)) * 100000) }) }];
    if (dir === 'u') return [
      { pos: (w, h, a) => [w / 2 - (w * a.adj1) / 200000, (ss(w, h) * a.adj2) / 100000], set: (x, y, w, h) => ({ adj1: L.clamp(((w / 2 - x) * 200000) / w, 0, 100000), adj2: L.clamp((y / ss(w, h)) * 100000, 0, (h / ss(w, h)) * 100000) }) }];
    return [
      { pos: (w, h, a) => [w / 2 - (w * a.adj1) / 200000, h - (ss(w, h) * a.adj2) / 100000], set: (x, y, w, h) => ({ adj1: L.clamp(((w / 2 - x) * 200000) / w, 0, 100000), adj2: L.clamp(((h - y) / ss(w, h)) * 100000, 0, (h / ss(w, h)) * 100000) }) }];
  }

  /* ---------- custom geometry (freeforms, imported custGeom) ----------
   * path = { paths: [{ w, h, cmds: [['M',x,y], ['L',x,y], ['C',...6], ['Q',...4], ['A', wR, hR, stDeg, swDeg], ['Z']], fill, stroke }] }
   */
  function customParts(path, W, H) {
    const out = [];
    for (const p of path.paths || []) {
      const sx = p.w ? W / p.w : 1, sy = p.h ? H / p.h : 1;
      let d = '', cx = 0, cy = 0, sxp = 0, syp = 0;
      for (const c of p.cmds) {
        switch (c[0]) {
          case 'M': cx = c[1]; cy = c[2]; sxp = cx; syp = cy; d += `M${n(cx * sx)},${n(cy * sy)} `; break;
          case 'L': cx = c[1]; cy = c[2]; d += `L${n(cx * sx)},${n(cy * sy)} `; break;
          case 'C': d += `C${n(c[1] * sx)},${n(c[2] * sy)} ${n(c[3] * sx)},${n(c[4] * sy)} ${n(c[5] * sx)},${n(c[6] * sy)} `; cx = c[5]; cy = c[6]; break;
          case 'Q': d += `Q${n(c[1] * sx)},${n(c[2] * sy)} ${n(c[3] * sx)},${n(c[4] * sy)} `; cx = c[3]; cy = c[4]; break;
          case 'A': {
            /* OOXML arcTo: radii + start/sweep angle, starting at the current point */
            const wR = c[1] * sx, hR = c[2] * sy;
            const st = (c[3] * Math.PI) / 180, sw = (c[4] * Math.PI) / 180;
            const pcx = cx * sx - wR * Math.cos(st), pcy = cy * sy - hR * Math.sin(st);
            const segs = Math.max(1, Math.min(4, Math.ceil(Math.abs(sw) / Math.PI - 1e-6)));
            let ex = cx * sx, ey = cy * sy;
            for (let i = 1; i <= segs; i++) {
              const tt = st + (sw * i) / segs;
              ex = pcx + wR * Math.cos(tt); ey = pcy + hR * Math.sin(tt);
              if (wR > 0 && hR > 0) d += `A${n(wR)},${n(hR)} 0 0 ${sw > 0 ? 1 : 0} ${n(ex)},${n(ey)} `;
              else d += `L${n(ex)},${n(ey)} `;
            }
            cx = ex / (sx || 1); cy = ey / (sy || 1);
            break;
          }
          case 'Z': d += 'Z '; cx = sxp; cy = syp; break;
          default: break;
        }
      }
      out.push({ d: d.trim(), fill: p.fill === 'none' ? 'none' : p.fill || 'norm', stroke: p.stroke !== false, open: !/Z\s*$/.test(d.trim()) });
    }
    return out;
  }

  /* ---------- Office 2007+ presets that real-world decks use ---------- */
  const deg = (v) => v / 60000;
  def('pie', 'Pie', (w, h, a) => { const p0 = ePt(w, h, deg(a.adj1)); return one(`M${n(w / 2)},${n(h / 2)} L${n(p0[0])},${n(p0[1])} ${arcSeg(w, h, deg(a.adj1), deg(a.adj2))} Z`); }, { adj: { adj1: 0, adj2: 16200000 } });
  def('pieWedge', 'Pie Wedge', (w, h) => one(`M0,${n(h)} A${n(w)},${n(h)} 0 0 1 ${n(w)},0 L${n(w)},${n(h)} Z`));
  def('chord', 'Chord', (w, h, a) => { const p0 = ePt(w, h, deg(a.adj1)); return one(`M${n(p0[0])},${n(p0[1])} ${arcSeg(w, h, deg(a.adj1), deg(a.adj2))} Z`); }, { adj: { adj1: 2700000, adj2: 16200000 } });
  def('blockArc', 'Block Arc', (w, h, a) => {
    const st = deg(a.adj1), en = deg(a.adj2), t = (ss(w, h) * a.adj3) / 100000;
    const rx = w / 2, ry = h / 2, irx = Math.max(0, rx - t), iry = Math.max(0, ry - t);
    const o0 = ePt(w, h, st), i1 = ePt(w, h, en, irx, iry);
    return one(`M${n(o0[0])},${n(o0[1])} ${arcSeg(w, h, st, en)} L${n(i1[0])},${n(i1[1])} ${arcSegCCW(w, h, en, st, irx, iry)} Z`);
  }, { adj: { adj1: 10800000, adj2: 0, adj3: 25000 } });
  def('teardrop', 'Teardrop', (w, h, a) => {
    const k = Math.min(2, a.adj / 100000);
    const tx = w / 2 + (w / 2) * k * Math.SQRT1_2 * 1.414, ty = h / 2 - (h / 2) * k * Math.SQRT1_2 * 1.414;
    return one(`M0,${n(h / 2)} A${n(w / 2)},${n(h / 2)} 0 0 1 ${n(w / 2)},0 Q${n((w / 2 + tx) / 2 + w * 0.05)},${n(ty / 2 - h * 0.02)} ${n(tx)},${n(ty)} Q${n(w + w * 0.02)},${n((ty + h / 2) / 2 - h * 0.05)} ${n(w)},${n(h / 2)} A${n(w / 2)},${n(h / 2)} 0 0 1 ${n(w / 2)},${n(h)} A${n(w / 2)},${n(h / 2)} 0 0 1 0,${n(h / 2)} Z`);
  }, { adj: { adj: 100000 } });
  def('corner', 'L-Shape', (w, h, a) => { const y = (ss(w, h) * a.adj1) / 100000, x = (ss(w, h) * a.adj2) / 100000; return one(P([[0, 0], [x, 0], [x, h - y], [w, h - y], [w, h], [0, h]])); }, { adj: { adj1: 50000, adj2: 50000 } });
  def('halfFrame', 'Half Frame', (w, h, a) => { const t1 = (ss(w, h) * a.adj1) / 100000, t2 = (ss(w, h) * a.adj2) / 100000; return one(P([[0, 0], [w, 0], [Math.max(t2, w - (t1 * w) / h), t1], [t2, t1], [t2, Math.max(t1, h - (t2 * h) / w)], [0, h]])); }, { adj: { adj1: 33333, adj2: 33333 } });
  def('diagStripe', 'Diagonal Stripe', (w, h, a) => { const k = a.adj / 100000; return one(P([[0, h * k], [w * k, 0], [w, 0], [0, h]])); }, { adj: { adj: 50000 } });
  def('plaque', 'Plaque', (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return one(`M0,${n(x)} A${n(x)},${n(x)} 0 0 0 ${n(x)},0 L${n(w - x)},0 A${n(x)},${n(x)} 0 0 0 ${n(w)},${n(x)} L${n(w)},${n(h - x)} A${n(x)},${n(x)} 0 0 0 ${n(w - x)},${n(h)} L${n(x)},${n(h)} A${n(x)},${n(x)} 0 0 0 0,${n(h - x)} Z`); }, { adj: { adj: 16667 } });
  def('bentUpArrow', 'Bent-Up Arrow', (w, h, a) => {
    const th = (ss(w, h) * a.adj1) / 100000, hw = (ss(w, h) * a.adj2) / 100000 * 2, hl = (ss(w, h) * a.adj3) / 100000;
    const cx = w - hw / 2;
    return one(P([[0, h - th], [cx - th / 2, h - th], [cx - th / 2, hl], [w - hw, hl], [cx, 0], [w, hl], [cx + th / 2, hl], [cx + th / 2, h], [0, h]]));
  }, { adj: { adj1: 25000, adj2: 25000, adj3: 25000 } });
  def('leftUpArrow', 'Left-Up Arrow', (w, h, a) => {
    const th = (ss(w, h) * a.adj1) / 100000, hw = (ss(w, h) * a.adj2) / 100000 * 2, hl = (ss(w, h) * a.adj3) / 100000;
    const cx = w - hw / 2, cy = h - hw / 2;
    return one(P([[0, cy], [hl, h - hw], [hl, cy - th / 2], [cx - th / 2, cy - th / 2], [cx - th / 2, hl], [w - hw, hl], [cx, 0], [w, hl], [cx + th / 2, hl], [cx + th / 2, cy + th / 2], [hl, cy + th / 2], [hl, h]]));
  }, { adj: { adj1: 25000, adj2: 25000, adj3: 25000 } });
  const rcorners = (w, h, tl, tr, br, bl, snip) => {
    const c = (r, x, y, ex, ey, isSnip) => (r <= 0 ? `L${n(x)},${n(y)}` : isSnip ? `L${n(ex)},${n(ey)}` : `A${n(r)},${n(r)} 0 0 1 ${n(ex)},${n(ey)}`);
    const S = snip || {};
    return `M${n(tl)},0 L${n(w - tr)},0 ${c(tr, w, 0, w, tr, S.tr)} L${n(w)},${n(h - br)} ${c(br, w, h, w - br, h, S.br)} L${n(bl)},${n(h)} ${c(bl, 0, h, 0, h - bl, S.bl)} L0,${n(tl)} ${c(tl, 0, 0, tl, 0, S.tl)} Z`;
  };
  const sz = (w, h, v) => (ss(w, h) * v) / 100000;
  def('snip1Rect', 'Snip Single Corner Rectangle', (w, h, a) => one(rcorners(w, h, 0, sz(w, h, a.adj), 0, 0, { tr: 1 })), { adj: { adj: 16667 } });
  def('snip2SameRect', 'Snip Same Side Corner Rectangle', (w, h, a) => { const t = sz(w, h, a.adj1), b = sz(w, h, a.adj2); return one(rcorners(w, h, t, t, b, b, { tl: 1, tr: 1, br: 1, bl: 1 })); }, { adj: { adj1: 16667, adj2: 0 } });
  def('snip2DiagRect', 'Snip Diagonal Corner Rectangle', (w, h, a) => { const d1 = sz(w, h, a.adj1), d2 = sz(w, h, a.adj2); return one(rcorners(w, h, d1, d2, d1, d2, { tl: 1, tr: 1, br: 1, bl: 1 })); }, { adj: { adj1: 0, adj2: 16667 } });
  def('snipRoundRect', 'Snip and Round Single Corner Rectangle', (w, h, a) => one(rcorners(w, h, sz(w, h, a.adj1), sz(w, h, a.adj2), 0, 0, { tr: 1 })), { adj: { adj1: 16667, adj2: 16667 } });
  def('round1Rect', 'Round Single Corner Rectangle', (w, h, a) => one(rcorners(w, h, 0, sz(w, h, a.adj), 0, 0)), { adj: { adj: 16667 } });
  def('round2SameRect', 'Round Same Side Corner Rectangle', (w, h, a) => { const t = sz(w, h, a.adj1), b = sz(w, h, a.adj2); return one(rcorners(w, h, t, t, b, b)); }, { adj: { adj1: 16667, adj2: 0 } });
  def('round2DiagRect', 'Round Diagonal Corner Rectangle', (w, h, a) => { const d1 = sz(w, h, a.adj1), d2 = sz(w, h, a.adj2); return one(rcorners(w, h, d1, d2, d1, d2)); }, { adj: { adj1: 16667, adj2: 0 } });
  def('heptagon', 'Heptagon', (w, h) => one(regPoly(w, h, 7)));
  def('decagon', 'Decagon', (w, h) => one(regPoly(w, h, 10, 0)));
  def('dodecagon', 'Dodecagon', (w, h) => one(regPoly(w, h, 12, 15)));
  def('star7', '7-Point Star', (w, h, a) => one(star(w, h, 7, a.adj / 50000)), { adj: { adj: 34601 } });
  def('star10', '10-Point Star', (w, h, a) => one(star(w, h, 10, a.adj / 50000)), { adj: { adj: 42533 } });
  def('star12', '12-Point Star', (w, h, a) => one(star(w, h, 12, a.adj / 50000)), { adj: { adj: 37500 } });
  def('flowChartOnlineStorage', 'Flowchart: Stored Data', (w, h) => one(`M${n(w / 6)},0 L${n(w)},0 A${n(w / 6)},${n(h / 2)} 0 0 0 ${n(w)},${n(h)} L${n(w / 6)},${n(h)} A${n(w / 6)},${n(h / 2)} 0 0 1 ${n(w / 6)},0 Z`));
  def('flowChartOfflineStorage', 'Flowchart: Offline Storage', (w, h) => one(P([[0, 0], [w, 0], [w / 2, h]])));
  def('flowChartMagneticTape', 'Flowchart: Sequential Access Storage', (w, h) => one(`M${n(w / 2)},${n(h)} A${n(w / 2)},${n(h / 2)} 0 1 1 ${n(w * 0.854)},${n(h * 0.854)} L${n(w)},${n(h * 0.854)} L${n(w)},${n(h)} Z`));
  def('mathPlus', 'Plus', (w, h, a) => { const t = sz(w, h, a.adj1) / 2, m = 0.1; const cx = w / 2, cy = h / 2; return one(P([[w * m, cy - t], [cx - t, cy - t], [cx - t, h * m], [cx + t, h * m], [cx + t, cy - t], [w * (1 - m), cy - t], [w * (1 - m), cy + t], [cx + t, cy + t], [cx + t, h * (1 - m)], [cx - t, h * (1 - m)], [cx - t, cy + t], [w * m, cy + t]])); }, { adj: { adj1: 23520 } });
  def('mathMinus', 'Minus', (w, h, a) => { const t = (h * a.adj1) / 200000; return one(P([[w * 0.1, h / 2 - t], [w * 0.9, h / 2 - t], [w * 0.9, h / 2 + t], [w * 0.1, h / 2 + t]])); }, { adj: { adj1: 23520 } });
  def('mathEqual', 'Equal', (w, h, a) => { const t = (h * a.adj1) / 100000, g = (h * a.adj2) / 200000; return [...one(P([[w * 0.1, h / 2 - g - t], [w * 0.9, h / 2 - g - t], [w * 0.9, h / 2 - g], [w * 0.1, h / 2 - g]])), ...one(P([[w * 0.1, h / 2 + g], [w * 0.9, h / 2 + g], [w * 0.9, h / 2 + g + t], [w * 0.1, h / 2 + g + t]]))]; }, { adj: { adj1: 23520, adj2: 11760 } });
  const bar = (cx, cy, r, t, angDeg) => { const a = (angDeg * Math.PI) / 180, u = [Math.cos(a), Math.sin(a)], v = [-Math.sin(a), Math.cos(a)]; return P([[cx - u[0] * r - v[0] * t, cy - u[1] * r - v[1] * t], [cx + u[0] * r - v[0] * t, cy + u[1] * r - v[1] * t], [cx + u[0] * r + v[0] * t, cy + u[1] * r + v[1] * t], [cx - u[0] * r + v[0] * t, cy - u[1] * r + v[1] * t]]); };
  def('mathMultiply', 'Multiply', (w, h, a) => { const t = sz(w, h, a.adj1) / 2, r = ss(w, h) * 0.42; return [{ d: bar(w / 2, h / 2, r, t, 45) + ' ' + bar(w / 2, h / 2, r, t, -45), fill: 'norm', stroke: true }]; }, { adj: { adj1: 23520 } });
  def('mathDivide', 'Division', (w, h, a) => { const t = (h * a.adj1) / 200000, r = (h * a.adj3) / 100000 || h * 0.1, g = (h * a.adj2) / 100000; return [...one(P([[w * 0.1, h / 2 - t], [w * 0.9, h / 2 - t], [w * 0.9, h / 2 + t], [w * 0.1, h / 2 + t]])), ...one(ell(w / 2, h / 2 - t - g - r, r, r)), ...one(ell(w / 2, h / 2 + t + g + r, r, r))]; }, { adj: { adj1: 23520, adj2: 5880, adj3: 11760 } });

  /* ---------- text rectangles (after presetShapeDefinitions, simplified) ---------- */
  const inset = (fx, fy, fr, fb) => (w, h) => [w * fx, h * fy, w * fr, h * fb];
  const arrowH = (w, h, a, left, right) => {
    const dy = (h * a.adj1) / 200000, dx = (ss(w, h) * a.adj2) / 100000;
    const k = dx * (dy / (h / 2 || 1));
    return [left ? dx - k : 0, h / 2 - dy, right ? w - dx + k : w, h / 2 + dy];
  };
  const arrowV = (w, h, a, top, bottom) => {
    const dx = (w * a.adj1) / 200000, dy = (ss(w, h) * a.adj2) / 100000;
    const k = dy * (dx / (w / 2 || 1));
    return [w / 2 - dx, top ? dy - k : 0, w / 2 + dx, bottom ? h - dy + k : h];
  };
  const TR = {
    roundRect: (w, h, a) => { const i = ((ss(w, h) * a.adj) / 100000) * 0.29289; return [i, i, w - i, h - i]; },
    wedgeRoundRectCallout: (w, h) => { const i = ss(w, h) * 0.16667 * 0.29289; return [i, i, w - i, h - i]; },
    flowChartAlternateProcess: (w, h) => { const i = ss(w, h) * 0.16667 * 0.29289; return [i, i, w - i, h - i]; },
    parallelogram: (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return [x / 2, h * 0.08, w - x / 2, h * 0.92]; },
    trapezoid: (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return [x * 0.66, h * 0.1, w - x * 0.66, h]; },
    pentagon: inset(0.18, 0.26, 0.82, 1),
    hexagon: (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return [x / 2, h * 0.08, w - x / 2, h * 0.92]; },
    octagon: (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000 / 2; return [x, x, w - x, h - x]; },
    plus: (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return [0, x, w, h - x]; },
    heart: inset(0.2, 0.25, 0.8, 0.7),
    smileyFace: inset(0.146, 0.146, 0.854, 0.854),
    sun: inset(0.3, 0.3, 0.7, 0.7),
    moon: inset(0.25, 0.25, 0.65, 0.75),
    noSmoking: inset(0.146, 0.146, 0.854, 0.854),
    donut: inset(0.146, 0.146, 0.854, 0.854),
    wave: inset(0, 0.2, 1, 0.8),
    lightningBolt: inset(0.3, 0.3, 0.7, 0.7),
    rightArrow: (w, h, a) => arrowH(w, h, a, false, true),
    leftArrow: (w, h, a) => arrowH(w, h, a, true, false),
    leftRightArrow: (w, h, a) => arrowH(w, h, a, true, true),
    notchedRightArrow: (w, h, a) => { const r = arrowH(w, h, a, false, true); r[0] = ((h / 2 - (h * a.adj1) / 200000) * w) / h * 0.5; return r; },
    stripedRightArrow: (w, h, a) => { const r = arrowH(w, h, a, false, true); r[0] = ss(w, h) * 0.156; return r; },
    upArrow: (w, h, a) => arrowV(w, h, a, true, false),
    downArrow: (w, h, a) => arrowV(w, h, a, false, true),
    upDownArrow: (w, h, a) => arrowV(w, h, a, true, true),
    quadArrow: inset(0.3, 0.3, 0.7, 0.7),
    homePlate: (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return [0, 0, Math.max(w / 2, w - x / 2), h]; },
    chevron: (w, h, a) => { const x = (ss(w, h) * a.adj) / 100000; return w - 2 * x > w * 0.2 ? [x, 0, w - x, h] : [w * 0.3, 0, w * 0.7, h]; },
    star4: inset(0.33, 0.33, 0.67, 0.67), star5: inset(0.31, 0.39, 0.69, 0.75), star6: inset(0.25, 0.25, 0.75, 0.75), star8: inset(0.2, 0.2, 0.8, 0.8),
    star16: inset(0.18, 0.18, 0.82, 0.82), star24: inset(0.16, 0.16, 0.84, 0.84), star32: inset(0.15, 0.15, 0.85, 0.85),
    irregularSeal1: inset(0.25, 0.3, 0.72, 0.68), irregularSeal2: inset(0.25, 0.3, 0.72, 0.7),
    ribbon2: inset(0.25, 0.1, 0.75, 0.75),
    verticalScroll: (w, h) => { const x = ss(w, h) * 0.125; return [x, x, w - x, h - x / 2]; },
    horizontalScroll: (w, h) => { const x = ss(w, h) * 0.125; return [x / 2, x, w - x, h - x]; },
    flowChartInputOutput: inset(0.2, 0, 0.8, 1), flowChartPredefinedProcess: inset(0.125, 0, 0.875, 1), flowChartInternalStorage: inset(0.125, 0.125, 1, 1),
    flowChartDocument: inset(0, 0, 1, 0.802), flowChartMultidocument: inset(0, 0.167, 0.861, 0.802), flowChartTerminator: inset(0.047, 0.146, 0.953, 0.854),
    flowChartPreparation: inset(0.2, 0, 0.8, 1), flowChartManualInput: inset(0, 0.2, 1, 1), flowChartManualOperation: inset(0.2, 0, 0.8, 1),
    flowChartOffpageConnector: inset(0, 0, 1, 0.8), flowChartPunchedCard: inset(0, 0.2, 1, 1), flowChartPunchedTape: inset(0, 0.2, 1, 0.8),
    flowChartSummingJunction: inset(0.146, 0.146, 0.854, 0.854), flowChartOr: inset(0.146, 0.146, 0.854, 0.854),
    flowChartCollate: inset(0.25, 0.25, 0.75, 0.75), flowChartSort: inset(0.25, 0.25, 0.75, 0.75), flowChartExtract: inset(0.25, 0.5, 0.75, 1), flowChartMerge: inset(0.25, 0, 0.75, 0.5),
    flowChartDelay: inset(0, 0.146, 0.854, 0.854), flowChartMagneticDisk: inset(0, 0.333, 1, 0.833), flowChartDisplay: inset(0.167, 0, 0.833, 1),
  };
  for (const k in TR) if (G[k] && !G[k].text) G[k].text = TR[k];

  /* ---------- public API ---------- */
  L.geom = {
    presets: G,
    get: (name) => G[name] || G.rect,
    isLine: (name) => !!(G[name] && G[name].isLine),
    adj(sh) { const g = G[sh.geom] || G.rect; return Object.assign({}, g.adj || {}, sh.adj || {}); },
    parts(sh, w, h) {
      w = Math.max(w, 0.01); h = Math.max(h, 0.01);
      if (sh.geom === 'custom' && sh.path) return customParts(sh.path, w, h);
      const g = G[sh.geom] || G.rect;
      if (!g.fn) return [];
      return g.fn(w, h, L.geom.adj(sh));
    },
    textRect(sh) {
      const g = G[sh.geom];
      if (g && g.text) return g.text(sh.w, sh.h, L.geom.adj(sh));
      return [0, 0, sh.w, sh.h];
    },
    handles(sh) { const g = G[sh.geom]; return (g && g.handles) || []; },
    linePath(sh) {
      const w = sh.w, h = sh.h;
      const x1 = sh.flipH ? w : 0, y1 = sh.flipV ? h : 0, x2 = sh.flipH ? 0 : w, y2 = sh.flipV ? 0 : h;
      if (sh.geom === 'bentConnector2') return `M${n(x1)},${n(y1)} H${n(x2)} V${n(y2)}`;
      if (sh.geom === 'bentConnector3' || sh.geom === 'bentConnector4') { const a = sh.adj && isFinite(sh.adj.adj1) ? sh.adj.adj1 / 100000 : 0.5; const mx = x1 + (x2 - x1) * a; return `M${n(x1)},${n(y1)} H${n(mx)} V${n(y2)} H${n(x2)}`; }
      if (sh.geom === 'curvedConnector2') return `M${n(x1)},${n(y1)} Q${n(x2)},${n(y1)} ${n(x2)},${n(y2)}`;
      if (sh.geom === 'curvedConnector3') { const mx = (x1 + x2) / 2; return `M${n(x1)},${n(y1)} C${n(mx)},${n(y1)} ${n(mx)},${n(y2)} ${n(x2)},${n(y2)}`; }
      return `M${n(x1)},${n(y1)} L${n(x2)},${n(y2)}`;
    },
    lineEnds(sh) {
      return [[sh.x + (sh.flipH ? sh.w : 0), sh.y + (sh.flipV ? sh.h : 0)], [sh.x + (sh.flipH ? 0 : sh.w), sh.y + (sh.flipV ? 0 : sh.h)]];
    },
    setLineEnds(sh, p1, p2) {
      sh.x = Math.min(p1[0], p2[0]); sh.y = Math.min(p1[1], p2[1]);
      sh.w = Math.abs(p2[0] - p1[0]); sh.h = Math.abs(p2[1] - p1[1]);
      sh.flipH = p1[0] > p2[0]; sh.flipV = p1[1] > p2[1];
    },
    /* AutoShapes menu groups (2003 order) */
    MENU: {
      'Basic Shapes': ['rect', 'parallelogram', 'trapezoid', 'diamond', 'roundRect', 'octagon', 'triangle', 'rtTriangle', 'ellipse', 'hexagon', 'plus', 'pentagon', 'can', 'cube', 'bevel', 'foldedCorner', 'smileyFace', 'donut', 'noSmoking', 'heart', 'lightningBolt', 'sun', 'moon', 'arc', 'bracketPair', 'bracePair', 'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'frame', 'cloud'],
      'Block Arrows': ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'quadArrow', 'bentArrow', 'uturnArrow', 'circularArrow', 'stripedRightArrow', 'notchedRightArrow', 'homePlate', 'chevron'],
      'Flowchart': ['flowChartProcess', 'flowChartAlternateProcess', 'flowChartDecision', 'flowChartInputOutput', 'flowChartPredefinedProcess', 'flowChartInternalStorage', 'flowChartDocument', 'flowChartMultidocument', 'flowChartTerminator', 'flowChartPreparation', 'flowChartManualInput', 'flowChartManualOperation', 'flowChartConnector', 'flowChartOffpageConnector', 'flowChartPunchedCard', 'flowChartPunchedTape', 'flowChartSummingJunction', 'flowChartOr', 'flowChartCollate', 'flowChartSort', 'flowChartExtract', 'flowChartMerge', 'flowChartDelay', 'flowChartMagneticDisk', 'flowChartDisplay'],
      'Stars and Banners': ['irregularSeal1', 'irregularSeal2', 'star4', 'star5', 'star6', 'star8', 'star16', 'star24', 'star32', 'ribbon2', 'verticalScroll', 'horizontalScroll', 'wave'],
      'Callouts': ['wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout'],
      'Action Buttons': ['actionButtonBlank', 'actionButtonHome', 'actionButtonHelp', 'actionButtonInformation', 'actionButtonBackPrevious', 'actionButtonForwardNext', 'actionButtonBeginning', 'actionButtonEnd', 'actionButtonReturn', 'actionButtonDocument', 'actionButtonSound', 'actionButtonMovie'],
    },
    /* small SVG preview of a preset (for menus) */
    preview(name, size) {
      size = size || 18;
      const s = 14, o = (size - s) / 2;
      if (G[name] && G[name].isLine) {
        const d = name === 'bentConnector3' ? `M${o},${o} H${size / 2} V${o + s} H${o + s}` : name === 'curvedConnector3' ? `M${o},${o} C${size / 2},${o} ${size / 2},${o + s} ${o + s},${o + s}` : `M${o},${o + s} L${o + s},${o}`;
        return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><path d="${d}" fill="none" stroke="#1f3f7a" stroke-width="1.2"/></svg>`;
      }
      const parts = L.geom.parts({ geom: name }, s, s);
      const body = parts.map((p) => `<path d="${p.d}" transform="translate(${o},${o})" fill="${p.fill === 'none' ? 'none' : p.fill.startsWith('darken') ? '#5b7fc0' : p.fill.startsWith('lighten') ? '#dfe9fb' : '#a9c3ef'}" ${p.rule ? `fill-rule="${p.rule}"` : ''} stroke="${p.stroke ? '#1f3f7a' : 'none'}" stroke-width="0.8"/>`).join('');
      return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">${body}</svg>`;
    },
  };
})(typeof window !== 'undefined' ? window : globalThis);
