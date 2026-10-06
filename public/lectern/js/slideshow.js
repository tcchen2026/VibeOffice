/* Lectern — slide show: transitions, custom animation, pen, rehearsal. */
(function () {
  'use strict';
  const L = window.L;
  const { h } = L;

  /* ================= effect catalogue (shared with UI and .pptx I/O) ================= */
  const DIR8 = [['b', 'From Bottom', 4], ['l', 'From Left', 8], ['r', 'From Right', 2], ['t', 'From Top', 1], ['bl', 'From Bottom-Left', 12], ['br', 'From Bottom-Right', 6], ['tl', 'From Top-Left', 9], ['tr', 'From Top-Right', 3]];
  const DIR4 = [['b', 'From Bottom', 4], ['l', 'From Left', 8], ['r', 'From Right', 2], ['t', 'From Top', 1]];
  const HV = [['horz', 'Horizontal', 10], ['vert', 'Vertical', 5]];
  const INOUT = [['in', 'In', 16], ['out', 'Out', 32]];
  const ACROSS = [['across', 'Across', 10], ['down', 'Down', 5]];
  const SPLIT = [['horzIn', 'Horizontal In', 26], ['horzOut', 'Horizontal Out', 42], ['vertIn', 'Vertical In', 21], ['vertOut', 'Vertical Out', 37]];
  const DIAG = [['dl', 'Left Down', 12], ['ul', 'Left Up', 9], ['dr', 'Right Down', 6], ['ur', 'Right Up', 3]];
  const SPOKES = [['1', '1 Spoke', 1], ['2', '2 Spokes', 2], ['3', '3 Spokes', 3], ['4', '4 Spokes', 4], ['8', '8 Spokes', 8]];
  const E_ENTR = [
    ['appear', 'Appear', 1], ['blinds', 'Blinds', 3, HV], ['box', 'Box', 4, INOUT], ['checkerboard', 'Checkerboard', 5, ACROSS], ['circle', 'Circle', 6, INOUT],
    ['diamond', 'Diamond', 8, INOUT], ['dissolve', 'Dissolve In', 9], ['fade', 'Fade', 10], ['flyIn', 'Fly In', 2, DIR8], ['peekIn', 'Peek In', 12, DIR4],
    ['plus', 'Plus', 13, INOUT], ['randomBars', 'Random Bars', 14, HV], ['split', 'Split', 16, SPLIT], ['strips', 'Strips', 18, DIAG], ['wedge', 'Wedge', 20],
    ['wheel', 'Wheel', 21, SPOKES], ['wipe', 'Wipe', 22, DIR4], ['zoom', 'Zoom', 23], ['ascend', 'Ascend', 42], ['descend', 'Descend', 47], ['expand', 'Expand', 55], ['spinner', 'Spinner', 49],
  ];
  const E_EMPH = [['growShrink', 'Grow/Shrink', 6], ['spin', 'Spin', 8], ['transparency', 'Transparency', 9], ['teeter', 'Teeter', 32], ['blink', 'Blink', 35]];
  const E_EXIT = [
    ['disappear', 'Disappear', 1], ['blinds', 'Blinds', 3, HV], ['box', 'Box', 4, INOUT], ['checkerboard', 'Checkerboard', 5, ACROSS], ['circle', 'Circle', 6, INOUT],
    ['diamond', 'Diamond', 8, INOUT], ['dissolve', 'Dissolve Out', 9], ['fade', 'Fade', 10], ['flyOut', 'Fly Out', 2, DIR8], ['peekIn', 'Peek Out', 12, DIR4],
    ['plus', 'Plus', 13, INOUT], ['randomBars', 'Random Bars', 14, HV], ['split', 'Split', 16, SPLIT], ['strips', 'Strips', 18, DIAG], ['wedge', 'Wedge', 20],
    ['wheel', 'Wheel', 21, SPOKES], ['wipe', 'Wipe', 22, DIR4], ['zoom', 'Zoom', 23], ['descend', 'Sink Down', 47],
  ];
  const E_PATH = [['right', 'Right', 0], ['left', 'Left', 0], ['up', 'Up', 0], ['down', 'Down', 0], ['diagDR', 'Diagonal Down Right', 0], ['diagUR', 'Diagonal Up Right', 0], ['circle', 'Circle', 0], ['bounce', 'Bounce Right', 0]];
  const A = (L.anim = {
    ENTR: E_ENTR, EMPH: E_EMPH, EXIT: E_EXIT, PATH: E_PATH,
    list: (cls) => ({ entr: E_ENTR, emph: E_EMPH, exit: E_EXIT, path: E_PATH }[cls] || E_ENTR),
    info(cls, eff) { const e = A.list(cls).find((x) => x[0] === eff); return e ? { id: e[2], name: e[1], dirs: e[3] || null } : null; },
    name(a) { const i = A.info(a.cls, a.eff); return i ? i.name : a.eff; },
    subtype(a) {
      const i = A.info(a.cls, a.eff);
      if (!i || !i.dirs) return 0;
      const d = i.dirs.find((x) => x[0] === String(a.dir)) || i.dirs[0];
      return d[2];
    },
    dirFromSubtype(cls, eff, st) { const i = A.info(cls, eff); if (!i || !i.dirs) return undefined; const d = i.dirs.find((x) => x[2] === st); return d ? d[0] : i.dirs[0][0]; },
    defaultDir(cls, eff) { const i = A.info(cls, eff); return i && i.dirs ? i.dirs[0][0] : undefined; },
    filter(a) {
      const d = a.dir;
      switch (a.eff) {
        case 'blinds': return `blinds(${d === 'vert' ? 'vertical' : 'horizontal'})`;
        case 'box': return `box(${d === 'out' ? 'out' : 'in'})`;
        case 'checkerboard': return `checkerboard(${d === 'down' ? 'down' : 'across'})`;
        case 'circle': return `circle(${d === 'out' ? 'out' : 'in'})`;
        case 'diamond': return `diamond(${d === 'out' ? 'out' : 'in'})`;
        case 'plus': return `plus(${d === 'out' ? 'out' : 'in'})`;
        case 'dissolve': return 'dissolve';
        case 'randomBars': return `randombar(${d === 'vert' ? 'vertical' : 'horizontal'})`;
        case 'split': return { horzIn: 'barn(inHorizontal)', horzOut: 'barn(outHorizontal)', vertIn: 'barn(inVertical)', vertOut: 'barn(outVertical)' }[d] || 'barn(inVertical)';
        case 'strips': return { dl: 'strips(downLeft)', ul: 'strips(upLeft)', dr: 'strips(downRight)', ur: 'strips(upRight)' }[d] || 'strips(downLeft)';
        case 'wedge': return 'wedge';
        case 'wheel': return `wheel(${d || 1})`;
        case 'wipe': return { b: 'wipe(up)', t: 'wipe(down)', l: 'wipe(right)', r: 'wipe(left)' }[d] || 'wipe(up)';
        case 'peekIn': return { b: 'slide(fromBottom)', t: 'slide(fromTop)', l: 'slide(fromLeft)', r: 'slide(fromRight)' }[d] || 'slide(fromBottom)';
        default: return 'fade';
      }
    },
    fromFilter(f) {
      const m = /^(\w+)(?:\(([^)]*)\))?/.exec(f || '');
      if (!m) return { eff: 'fade' };
      const arg = m[2] || '';
      switch (m[1]) {
        case 'blinds': return { eff: 'blinds', dir: arg === 'vertical' ? 'vert' : 'horz' };
        case 'box': return { eff: 'box', dir: arg === 'out' ? 'out' : 'in' };
        case 'checkerboard': return { eff: 'checkerboard', dir: arg === 'down' ? 'down' : 'across' };
        case 'circle': return { eff: 'circle', dir: arg === 'out' ? 'out' : 'in' };
        case 'diamond': return { eff: 'diamond', dir: arg === 'out' ? 'out' : 'in' };
        case 'plus': return { eff: 'plus', dir: arg === 'out' ? 'out' : 'in' };
        case 'dissolve': return { eff: 'dissolve' };
        case 'randombar': return { eff: 'randomBars', dir: arg === 'vertical' ? 'vert' : 'horz' };
        case 'barn': return { eff: 'split', dir: { inHorizontal: 'horzIn', outHorizontal: 'horzOut', inVertical: 'vertIn', outVertical: 'vertOut' }[arg] || 'vertIn' };
        case 'strips': return { eff: 'strips', dir: { downLeft: 'dl', upLeft: 'ul', downRight: 'dr', upRight: 'ur' }[arg] || 'dl' };
        case 'wedge': return { eff: 'wedge' };
        case 'wheel': return { eff: 'wheel', dir: arg || '1' };
        case 'wipe': return { eff: 'wipe', dir: { up: 'b', down: 't', right: 'l', left: 'r' }[arg] || 'b' };
        case 'slide': return { eff: 'peekIn', dir: { fromBottom: 'b', fromTop: 't', fromLeft: 'l', fromRight: 'r' }[arg] || 'b' };
        default: return { eff: 'fade' };
      }
    },
    flyFrom(d) {
      const x = { l: '0-#ppt_w/2', r: '1+#ppt_w/2', bl: '0-#ppt_w/2', tl: '0-#ppt_w/2', br: '1+#ppt_w/2', tr: '1+#ppt_w/2' }[d] || '#ppt_x';
      const y = { t: '0-#ppt_h/2', b: '1+#ppt_h/2', tl: '0-#ppt_h/2', tr: '0-#ppt_h/2', bl: '1+#ppt_h/2', br: '1+#ppt_h/2' }[d] || '#ppt_y';
      return { x, y: d === 'l' || d === 'r' ? '#ppt_y' : y };
    },
    pathPoints(a) {
      const k = a.eff;
      if (a.pts && a.pts.length) return a.pts;
      if (k === 'circle') return Array.from({ length: 25 }, (_, i) => { const t = (i / 24) * Math.PI * 2; return [0.12 * Math.sin(t), -0.12 * (1 - Math.cos(t)) * 1.333]; });
      if (k === 'bounce') return Array.from({ length: 21 }, (_, i) => { const t = i / 20; return [t * 0.3, -Math.abs(Math.sin(t * Math.PI * 2)) * 0.15 * (1 - t * 0.5)]; });
      const v = { right: [0.25, 0], left: [-0.25, 0], up: [0, -0.25], down: [0, 0.25], diagDR: [0.2, 0.2], diagUR: [0.2, -0.2] }[k] || [0.25, 0];
      return [[0, 0], v];
    },
    pathString(a) { const p = A.pathPoints(a); return p.map((q, i) => `${i ? 'L' : 'M'} ${L.round(q[0], 4)} ${L.round(q[1], 4)}`).join(' ') + ' E'; },
    SPEEDS: [['Very Slow', 5000], ['Slow', 3000], ['Medium', 2000], ['Fast', 1000], ['Very Fast', 500]],
  });

  /* transitions (2003 list) */
  const TR = [];
  const tr = (name, type, extra) => TR.push(Object.assign({ name, type }, extra || {}));
  tr('No Transition', 'none');
  tr('Blinds Horizontal', 'blinds', { dir: 'horz' }); tr('Blinds Vertical', 'blinds', { dir: 'vert' });
  tr('Box In', 'zoom', { dir: 'in' }); tr('Box Out', 'zoom', { dir: 'out' });
  tr('Checkerboard Across', 'checker', { dir: 'horz' }); tr('Checkerboard Down', 'checker', { dir: 'vert' });
  tr('Comb Horizontal', 'comb', { dir: 'horz' }); tr('Comb Vertical', 'comb', { dir: 'vert' });
  tr('Cover Down', 'cover', { dir: 'd' }); tr('Cover Left', 'cover', { dir: 'l' }); tr('Cover Right', 'cover', { dir: 'r' }); tr('Cover Up', 'cover', { dir: 'u' });
  tr('Cover Left-Down', 'cover', { dir: 'ld' }); tr('Cover Left-Up', 'cover', { dir: 'lu' }); tr('Cover Right-Down', 'cover', { dir: 'rd' }); tr('Cover Right-Up', 'cover', { dir: 'ru' });
  tr('Cut', 'cut'); tr('Cut Through Black', 'cut', { thruBlk: true });
  tr('Dissolve', 'dissolve'); tr('Fade Smoothly', 'fade'); tr('Fade Through Black', 'fade', { thruBlk: true });
  tr('Newsflash', 'newsflash');
  tr('Push Down', 'push', { dir: 'd' }); tr('Push Left', 'push', { dir: 'l' }); tr('Push Right', 'push', { dir: 'r' }); tr('Push Up', 'push', { dir: 'u' });
  tr('Random Bars Horizontal', 'randomBar', { dir: 'horz' }); tr('Random Bars Vertical', 'randomBar', { dir: 'vert' });
  tr('Shape Circle', 'circle'); tr('Shape Diamond', 'diamond'); tr('Shape Plus', 'plus');
  tr('Split Horizontal In', 'split', { orient: 'horz', dir: 'in' }); tr('Split Horizontal Out', 'split', { orient: 'horz', dir: 'out' });
  tr('Split Vertical In', 'split', { orient: 'vert', dir: 'in' }); tr('Split Vertical Out', 'split', { orient: 'vert', dir: 'out' });
  tr('Strips Left-Down', 'strips', { dir: 'ld' }); tr('Strips Left-Up', 'strips', { dir: 'lu' }); tr('Strips Right-Down', 'strips', { dir: 'rd' }); tr('Strips Right-Up', 'strips', { dir: 'ru' });
  tr('Uncover Down', 'pull', { dir: 'd' }); tr('Uncover Left', 'pull', { dir: 'l' }); tr('Uncover Right', 'pull', { dir: 'r' }); tr('Uncover Up', 'pull', { dir: 'u' });
  tr('Uncover Left-Down', 'pull', { dir: 'ld' }); tr('Uncover Left-Up', 'pull', { dir: 'lu' }); tr('Uncover Right-Down', 'pull', { dir: 'rd' }); tr('Uncover Right-Up', 'pull', { dir: 'ru' });
  tr('Wedge', 'wedge');
  tr('Wheel Clockwise, 1 Spoke', 'wheel', { spokes: 1 }); tr('Wheel Clockwise, 2 Spokes', 'wheel', { spokes: 2 }); tr('Wheel Clockwise, 3 Spokes', 'wheel', { spokes: 3 });
  tr('Wheel Clockwise, 4 Spokes', 'wheel', { spokes: 4 }); tr('Wheel Clockwise, 8 Spokes', 'wheel', { spokes: 8 });
  tr('Wipe Down', 'wipe', { dir: 'd' }); tr('Wipe Left', 'wipe', { dir: 'l' }); tr('Wipe Right', 'wipe', { dir: 'r' }); tr('Wipe Up', 'wipe', { dir: 'u' });
  tr('Random Transition', 'random');
  A.TRANSITIONS = TR;
  A.transName = (t) => {
    if (!t || t.type === 'none') return 'No Transition';
    const hit = TR.find((x) => x.type === t.type && (x.dir || null) === (t.dir || null) && (!!x.thruBlk === !!t.thruBlk) && (x.spokes || null) === (t.spokes || null) && (x.orient || null) === (t.orient || null));
    return hit ? hit.name : TR.find((x) => x.type === t.type) ? TR.find((x) => x.type === t.type).name : t.type;
  };
  A.transDur = (t) => (t && t.dur) || { slow: 1000, med: 750, fast: 500 }[(t && t.spd) || 'fast'] || 500;

  /* ================= mask helpers ================= */
  function canvasMask(w, h) {
    const c = document.createElement('canvas');
    c.width = w; c.height = h;
    return c;
  }
  const RANDOM_CACHE = {};
  function randomOrder(n, key) {
    if (RANDOM_CACHE[key] && RANDOM_CACHE[key].length === n) return RANDOM_CACHE[key];
    const a = Array.from({ length: n }, () => Math.random());
    RANDOM_CACHE[key] = a;
    return a;
  }
  function setMask(el, val, size, pos) {
    el.style.webkitMaskImage = val; el.style.maskImage = val;
    el.style.webkitMaskSize = size || '100% 100%'; el.style.maskSize = size || '100% 100%';
    el.style.webkitMaskPosition = pos || '0 0'; el.style.maskPosition = pos || '0 0';
    el.style.webkitMaskRepeat = 'no-repeat'; el.style.maskRepeat = 'no-repeat';
  }
  function clearFx(el) {
    if (!el) return;
    const st = el.style;
    st.opacity = ''; st.transform = ''; st.translate = ''; st.clipPath = ''; st.webkitClipPath = '';
    st.webkitMaskImage = ''; st.maskImage = ''; st.filter = ''; st.transformOrigin = '';
  }
  /**
   * Apply a reveal of fraction t (0 hidden → 1 fully shown) for a named wipe-style effect
   * inside box b={x,y,w,h} (in el's local coordinates; ew/eh are el's own size).
   */
  function reveal(el, kind, dir, t, b, ew, eh, key) {
    const px = (v) => L.round(v, 2) + 'px';
    const ins = (top, right, bottom, left) => { const v = `inset(${px(top)} ${px(right)} ${px(bottom)} ${px(left)})`; el.style.clipPath = v; el.style.webkitClipPath = v; };
    const poly = (pts, evenodd) => { const v = `polygon(${evenodd ? 'evenodd, ' : ''}${pts.map((p) => px(p[0]) + ' ' + px(p[1])).join(', ')})`; el.style.clipPath = v; el.style.webkitClipPath = v; };
    const R = { x0: b.x, y0: b.y, x1: b.x + b.w, y1: b.y + b.h };
    const outer = [[0, 0], [ew, 0], [ew, eh], [0, eh], [0, 0]];
    const outsideB = (r) => [[R.x0, R.y0], [R.x1, R.y0], [R.x1, R.y1], [R.x0, R.y1], [R.x0, R.y0], [r[0], r[1]], [r[0], r[3]], [r[2], r[3]], [r[2], r[1]], [r[0], r[1]]];
    const cx = b.x + b.w / 2, cy = b.y + b.h / 2;
    const mpos = `${px(b.x)} ${px(b.y)}`, msize = `${px(b.w)} ${px(b.h)}`;
    void outer;
    switch (kind) {
      case 'wipe': {
        if (dir === 'b') ins(R.y0 + (1 - t) * b.h, ew - R.x1, eh - R.y1, R.x0);
        else if (dir === 't') ins(R.y0, ew - R.x1, eh - R.y1 + (1 - t) * b.h, R.x0);
        else if (dir === 'l') ins(R.y0, ew - R.x1 + (1 - t) * b.w, eh - R.y1, R.x0);
        else ins(R.y0, ew - R.x1, eh - R.y1, R.x0 + (1 - t) * b.w);
        return;
      }
      case 'box': {
        if (dir === 'out') ins(cy - (t * b.h) / 2, ew - (cx + (t * b.w) / 2), eh - (cy + (t * b.h) / 2), cx - (t * b.w) / 2);
        else { const k = 1 - t; poly(outsideB([cx - (k * b.w) / 2, cy - (k * b.h) / 2, cx + (k * b.w) / 2, cy + (k * b.h) / 2]), true); }
        return;
      }
      case 'split': {
        if (dir === 'horzOut') ins(cy - (t * b.h) / 2, ew - R.x1, eh - (cy + (t * b.h) / 2), R.x0);
        else if (dir === 'vertOut') ins(R.y0, ew - (cx + (t * b.w) / 2), eh - R.y1, cx - (t * b.w) / 2);
        else if (dir === 'horzIn') { const k = (1 - t) * b.h / 2; poly(outsideB([R.x0, cy - k, R.x1, cy + k]), true); }
        else { const k = (1 - t) * b.w / 2; poly(outsideB([cx - k, R.y0, cx + k, R.y1]), true); }
        return;
      }
      case 'circle': {
        const r = Math.hypot(b.w, b.h) / 2;
        if (dir === 'in') setMask(el, `radial-gradient(circle closest-corner at 50% 50%, transparent ${L.round((1 - t) * 100, 2)}%, #000 ${L.round((1 - t) * 100 + 0.5, 2)}%)`, msize, mpos);
        else { const v = `circle(${px(t * r)} at ${px(cx)} ${px(cy)})`; el.style.clipPath = v; el.style.webkitClipPath = v; }
        if (dir === 'in') setMask(el, `radial-gradient(${px(r)} ${px(r)} at 50% 50%, transparent ${px((1 - t) * r)}, #000 ${px((1 - t) * r + 0.5)})`, msize, mpos);
        return;
      }
      case 'diamond': {
        if (dir === 'out') poly([[cx, cy - t * b.h], [cx + t * b.w, cy], [cx, cy + t * b.h], [cx - t * b.w, cy]]);
        else { const k = 1 - t; poly([[R.x0, R.y0], [R.x1, R.y0], [R.x1, R.y1], [R.x0, R.y1], [R.x0, R.y0], [cx, cy - k * b.h], [cx - k * b.w, cy], [cx, cy + k * b.h], [cx + k * b.w, cy], [cx, cy - k * b.h]], true); }
        return;
      }
      case 'plus': {
        const a = (t * Math.max(b.w, b.h)) / 2, hx = (t * b.w) / 2 + 0.01, hy = (t * b.h) / 2 + 0.01;
        if (dir === 'in') { const k = 1 - t, ax = (k * b.w) / 2, ay = (k * b.h) / 2; poly(outsideB([cx - ax, cy - ay, cx + ax, cy + ay]), true); return; }
        poly([[cx - hx, R.y0], [cx + hx, R.y0], [cx + hx, cy - hy], [R.x1, cy - hy], [R.x1, cy + hy], [cx + hx, cy + hy], [cx + hx, R.y1], [cx - hx, R.y1], [cx - hx, cy + hy], [R.x0, cy + hy], [R.x0, cy - hy], [cx - hx, cy - hy]]);
        void a;
        return;
      }
      case 'blinds': {
        const p = L.round(t * 100, 2);
        const g = dir === 'vert' ? `repeating-linear-gradient(90deg, #000 0, #000 ${p / 6}%, transparent ${p / 6}%, transparent ${100 / 6}%)` : `repeating-linear-gradient(180deg, #000 0, #000 ${p / 6}%, transparent ${p / 6}%, transparent ${100 / 6}%)`;
        setMask(el, g, msize, mpos);
        return;
      }
      case 'wedge': {
        const a = L.round(t * 180, 2);
        setMask(el, `conic-gradient(from 0deg at 50% 50%, #000 0deg ${a}deg, transparent ${a}deg ${360 - a}deg, #000 ${360 - a}deg)`, msize, mpos);
        return;
      }
      case 'wheel': {
        const n = Math.max(1, +dir || 1), seg = 360 / n, a = L.round(t * seg, 2);
        setMask(el, `repeating-conic-gradient(from 0deg at 50% 50%, #000 0deg ${a}deg, transparent ${a}deg ${seg}deg)`, msize, mpos);
        return;
      }
      case 'strips': {
        const ang = { dl: 225, ul: 315, dr: 135, ur: 45, ld: 225, lu: 315, rd: 135, ru: 45 }[dir] || 135;
        const p = t * 130 - 15;
        setMask(el, `linear-gradient(${ang}deg, #000 ${L.round(p - 15, 2)}%, transparent ${L.round(p + 15, 2)}%)`, msize, mpos);
        return;
      }
      case 'dissolve': case 'randomBars': case 'checkerboard': {
        const cols = kind === 'dissolve' ? 36 : kind === 'checkerboard' ? 16 : dir === 'vert' ? 96 : 1;
        const rows = kind === 'dissolve' ? 27 : kind === 'checkerboard' ? 12 : dir === 'vert' ? 1 : 72;
        const c = canvasMask(kind === 'checkerboard' ? 160 : cols, kind === 'checkerboard' ? 120 : rows);
        const g = c.getContext('2d');
        g.fillStyle = '#000';
        if (kind === 'checkerboard') {
          const cw = c.width / 8, ch = c.height / 6;
          for (let j = 0; j < 6; j++) for (let i = 0; i < 8; i++) {
            const tt = L.clamp(t * 2 - (((dir === 'down' ? i : j) % 2) ? 0.5 : 0), 0, 1);
            if (dir === 'down') g.fillRect(i * cw, j * ch, cw, ch * tt); else g.fillRect(i * cw, j * ch, cw * tt, ch);
          }
        } else {
          const ord = randomOrder(cols * rows, kind + key + cols + 'x' + rows);
          for (let j = 0; j < rows; j++) for (let i = 0; i < cols; i++) if (ord[j * cols + i] < t) g.fillRect(i, j, 1, 1);
        }
        setMask(el, `url(${c.toDataURL()})`, msize, mpos);
        el.style.imageRendering = 'pixelated';
        return;
      }
      default: return;
    }
  }
  A.reveal = reveal;

  /* ================= element effects ================= */
  const easeOut = (t) => 1 - Math.pow(1 - t, 3);
  const easeInOut = (t) => (t < 0.5 ? 2 * t * t : 1 - Math.pow(-2 * t + 2, 2) / 2);
  /** Draw one frame of an animation. t in [0,1]. tgt: {el, local:{x,y,w,h}, ew, eh, slideBox, W, H} */
  function frame(a, tgt, t) {
    const el = tgt.el;
    const { local: b, W, H } = tgt;
    clearFx(el);
    const sb = tgt.slideBox;
    const orig = `${L.round(b.x + b.w / 2, 2)}px ${L.round(b.y + b.h / 2, 2)}px`;
    if (a.cls === 'path') {
      const pts = A.pathPoints(a);
      const total = pts.length - 1;
      const f = easeInOut(t) * total;
      const i = Math.min(total - 1, Math.floor(f));
      const k = f - i;
      const p0 = pts[Math.max(0, i)], p1 = pts[Math.min(total, i + 1)];
      const x = (p0[0] + (p1[0] - p0[0]) * k) * W, y = (p0[1] + (p1[1] - p0[1]) * k) * H;
      el.style.translate = `${L.round(x + (tgt.pathOff ? tgt.pathOff[0] : 0), 2)}px ${L.round(y + (tgt.pathOff ? tgt.pathOff[1] : 0), 2)}px`;
      return;
    }
    if (a.cls === 'emph') {
      el.style.transformOrigin = orig;
      const rep = Math.max(1, a.repeat || 1);
      const tt = t >= 1 ? 1 : (t * rep) % 1;
      switch (a.eff) {
        case 'spin': el.style.transform = `rotate(${L.round((t >= 1 ? 1 : tt) * (a.amount || 360), 2)}deg)`; break;
        case 'transparency': el.style.opacity = t >= 1 ? '0.5' : String(1 - 0.5 * Math.min(1, tt * 4)); break;
        case 'teeter': el.style.transform = `rotate(${L.round(Math.sin(tt * Math.PI * 4) * 6 * (1 - tt), 2)}deg)`; break;
        case 'blink': el.style.visibility = t >= 1 ? '' : tt % 0.5 < 0.25 ? 'hidden' : 'visible'; break;
        default: { const s = 1 + ((a.amount || 150) / 100 - 1) * (t >= 1 ? 1 : easeInOut(tt)); el.style.transform = `scale(${L.round(s, 4)})`; }
      }
      return;
    }
    const entering = a.cls === 'entr';
    const p = entering ? t : 1 - t;
    if (p >= 1 && entering) return;
    const eff = a.eff === 'flyOut' ? 'flyIn' : a.eff === 'disappear' ? 'appear' : a.eff;
    switch (eff) {
      case 'appear': if (!entering && t >= 1) el.style.visibility = 'hidden'; return;
      case 'fade': el.style.opacity = String(L.round(p, 4)); return;
      case 'flyIn': {
        const d = a.dir || 'b';
        const e = entering ? easeOut(t) : 1 - (1 - (1 - t)) ** 1; void e;
        const k = entering ? 1 - easeOut(t) : t * t;
        let dx = 0, dy = 0;
        if (d.includes('l')) dx = -(sb.x + sb.w);
        if (d.includes('r')) dx = W - sb.x;
        if (d.includes('t')) dy = -(sb.y + sb.h);
        if (d.includes('b')) dy = H - sb.y;
        const sc = tgt.scaleToLocal || 1;
        el.style.translate = `${L.round(dx * k * sc, 2)}px ${L.round(dy * k * sc, 2)}px`;
        return;
      }
      case 'peekIn': {
        const d = a.dir || 'b';
        reveal(el, 'wipe', d === 'b' ? 'b' : d === 't' ? 't' : d === 'l' ? 'l' : 'r', 1, b, tgt.ew, tgt.eh);
        const off = 1 - (entering ? easeOut(t) : 1 - t);
        const inner = tgt.inner || el;
        if (inner !== el) clearFx(inner);
        const dx = d === 'l' ? -b.w : d === 'r' ? b.w : 0, dy = d === 't' ? -b.h : d === 'b' ? b.h : 0;
        inner.style.translate = `${L.round(dx * off, 2)}px ${L.round(dy * off, 2)}px`;
        return;
      }
      case 'zoom': el.style.transformOrigin = orig; el.style.transform = `scale(${L.round(Math.max(0.0001, entering ? easeOut(p) : p), 4)})`; el.style.opacity = String(Math.min(1, p * 2)); return;
      case 'ascend': el.style.translate = `0 ${L.round((1 - easeOut(p)) * Math.min(b.h, 60) * 0.8, 2)}px`; el.style.opacity = String(p); return;
      case 'descend': el.style.translate = `0 ${L.round(-(1 - easeOut(p)) * Math.min(b.h, 60) * 0.8 * (entering ? 1 : -1), 2)}px`; el.style.opacity = String(p); return;
      case 'expand': el.style.transformOrigin = orig; el.style.transform = `scaleX(${L.round(0.7 + 0.3 * easeOut(p), 4)})`; el.style.opacity = String(p); return;
      case 'spinner': el.style.transformOrigin = orig; el.style.transform = `rotate(${L.round((1 - easeOut(p)) * 360, 2)}deg) scale(${L.round(Math.max(0.0001, easeOut(p)), 4)})`; el.style.opacity = String(p); return;
      default: reveal(el, eff, a.dir || A.defaultDir(a.cls, a.eff), p, b, tgt.ew, tgt.eh, a.id); return;
    }
  }
  A.frame = frame;
  A.clearFx = clearFx;

  /* ================= slide show controller ================= */
  const S = (L.show = { active: false });
  let root, stage, layer, inkCanvas, ink, popbar, raf = 0;
  let st = null; /* runtime state */

  function slideOrder() {
    const pres = L.pres;
    const n = pres.slides.length;
    const sh = pres.show || {};
    let from = 0, to = n - 1;
    if (sh.from && sh.to) { from = L.clamp(sh.from - 1, 0, n - 1); to = L.clamp(sh.to - 1, from, n - 1); }
    const out = [];
    for (let i = from; i <= to; i++) out.push(i);
    return out;
  }

  /** Start the slide show. opts: {from: slideIndex, rehearse, preview(slideIndex for single-slide preview)} */
  S.start = function (opts) {
    opts = opts || {};
    if (S.active) return;
    const pres = L.pres;
    if (!pres.slides.length) { L.ui.msg('There are no slides to show.', { icon: 'info' }); return; }
    if (L.te && L.te.active()) L.te.end();
    S.active = true;
    L.ui.closeMenus();
    root = h('div', { class: 'show', tabindex: '0', role: 'application', 'aria-label': 'Slide Show' });
    stage = h('div', { class: 'stage' });
    inkCanvas = h('canvas', { class: 'ink' });
    popbar = h('div', { class: 'popbar' },
      h('button', { type: 'button', 'aria-label': 'Previous', html: L.icons.get('showPrev'), onclick: (e) => { e.stopPropagation(); S.prev(); } }),
      h('button', { type: 'button', 'aria-label': 'Pointer options', html: L.icons.get('showPen'), onclick: (e) => { e.stopPropagation(); const r = e.currentTarget.getBoundingClientRect(); L.ui.openMenu(pointerMenu(), { x: r.left, y: r.top - 230 }); } }),
      h('button', { type: 'button', 'aria-label': 'Slide show menu', html: L.icons.get('showMenu'), onclick: (e) => { e.stopPropagation(); const r = e.currentTarget.getBoundingClientRect(); L.ui.openMenu(showMenu(), { x: r.left, y: r.top - 300 }); } }),
      h('button', { type: 'button', 'aria-label': 'Next', html: L.icons.get('showNext'), onclick: (e) => { e.stopPropagation(); S.next(); } }));
    root.append(stage, inkCanvas, popbar);
    document.body.appendChild(root);
    const sh = pres.show || {};
    st = {
      order: opts.preview != null ? [opts.preview] : slideOrder(),
      pos: 0, step: 0, steps: [], running: null, timer: 0, ink: {}, pen: null, penColor: sh.penColor || '#FF0000', history: [],
      noAnim: !!sh.noAnim, loop: !!sh.loop || !!sh.kiosk, kiosk: !!sh.kiosk, useTimings: sh.useTimings !== false, rehearse: !!opts.rehearse, preview: opts.preview != null,
      slideStart: performance.now(), times: {}, paused: false, ended: false, gotoBuf: '', blank: null, idle: 0,
    };
    let startIdx = opts.from != null ? st.order.indexOf(opts.from) : 0;
    if (startIdx < 0) startIdx = 0;
    if (opts.from == null) while (startIdx < st.order.length - 1 && pres.slides[st.order[startIdx]].hidden) startIdx++;
    if (st.rehearse) buildRehearse();
    layout();
    window.addEventListener('resize', layout);
    root.addEventListener('pointerdown', onPointerDown);
    root.addEventListener('pointermove', onPointerMove);
    root.addEventListener('pointerup', onPointerUp);
    root.addEventListener('contextmenu', onContext);
    root.addEventListener('wheel', onWheel, { passive: true });
    document.addEventListener('keydown', onKey, true);
    if (!st.preview && root.requestFullscreen) root.requestFullscreen().catch(() => { /* fullscreen is optional */ });
    try { if (navigator.wakeLock) navigator.wakeLock.request('screen').then((l) => { st.wake = l; }).catch(() => {}); } catch (e) { /* ignore */ }
    root.focus();
    showSlide(startIdx, { transition: true });
    idleCursor();
  };

  function layout() {
    if (!root) return;
    const W = L.pres.W, H = L.pres.H;
    const vw = root.clientWidth || window.innerWidth, vh = root.clientHeight || window.innerHeight;
    const k = Math.min(vw / W, vh / H);
    st.k = k;
    const ox = (vw - W * k) / 2, oy = (vh - H * k) / 2;
    stage.style.cssText = `width:${W}px;height:${H}px;transform:translate(${ox}px,${oy}px) scale(${k})`;
    const dpr = window.devicePixelRatio || 1;
    inkCanvas.width = Math.round(W * k * dpr); inkCanvas.height = Math.round(H * k * dpr);
    inkCanvas.style.cssText = `left:${ox}px;top:${oy}px;width:${W * k}px;height:${H * k}px`;
    redrawInk();
  }

  /* ---- build steps ---- */
  function compile(slide, slideEl, noAnim) {
    const anims = noAnim ? [] : (slide.anims || []).filter((a) => L.model.shapeById(slide, a.sid));
    const items = [];
    for (const a of anims) {
      const sh = L.model.shapeById(slide, a.sid);
      if (a.by === 'para' && sh && sh.tx && sh.tx.ps.length > 1 && a.cls !== 'path') {
        /* "by 1st level paragraphs": deeper levels build together with their parent */
        sh.tx.ps.forEach((p, i) => {
          if (!L.txt.paraText(p).trim()) return;
          const prior = items.some((x) => x.id === a.id);
          items.push(Object.assign({}, a, { para: i, start: !prior ? a.start : (p.lvl || 0) > 0 ? 'with' : a.start, delay: prior && (p.lvl || 0) > 0 ? 0 : a.delay }));
        });
      } else items.push(a);
    }
    const steps = [];
    let cur = null;
    for (const a of items) {
      if (!cur || a.start === 'click') { cur = { auto: a.start !== 'click', items: [], t: 0, lastStart: 0, lastEnd: 0 }; steps.push(cur); }
      const dur = a.cls === 'emph' ? (a.dur || 2000) * Math.max(1, a.repeat || 1) : a.dur || 500;
      const start = a.start === 'after' ? cur.lastEnd + (a.delay || 0) : a.start === 'with' ? cur.lastStart + (a.delay || 0) : (a.delay || 0);
      const end = start + (a.cls === 'entr' && a.eff === 'appear' ? 0 : a.cls === 'exit' && a.eff === 'disappear' ? 0 : dur);
      cur.items.push({ a, start, end, dur });
      cur.lastStart = start;
      cur.lastEnd = Math.max(cur.lastEnd, end);
      if (a.start === 'click') { cur.lastStart = start; }
    }
    /* first step that starts without a click runs on slide entry */
    for (const s of steps) {
      for (const it of s.items) it.tgt = target(slideEl, slide, it.a);
      s.total = Math.max(0, ...s.items.map((x) => x.end));
    }
    return steps;
  }
  function target(slideEl, slide, a) {
    const aw = slideEl.querySelector(`.aw[data-sid="${a.sid}"]`);
    if (!aw) return null;
    const sh = L.model.shapeById(slide, a.sid);
    const W = L.pres.W, H = L.pres.H;
    if (a.para != null) {
      const p = aw.querySelectorAll('p.pa')[a.para];
      if (!p) return null;
      const sr = slideEl.getBoundingClientRect();
      const pr = p.getBoundingClientRect();
      const k = sr.width / W;
      const sb = { x: (pr.left - sr.left) / k, y: (pr.top - sr.top) / k, w: pr.width / k, h: pr.height / k };
      return { el: p, local: { x: 0, y: 0, w: p.offsetWidth, h: p.offsetHeight }, ew: p.offsetWidth, eh: p.offsetHeight, slideBox: sb, W, H, scaleToLocal: p.offsetWidth ? sb.w / p.offsetWidth === 0 ? 1 : p.offsetWidth / sb.w : 1, para: true };
    }
    const b = sh.type === 'group' ? L.model.groupBounds(sh) : L.rotBounds(sh);
    const pad = 4;
    const box = { x: b.x - pad, y: b.y - pad, w: b.w + pad * 2, h: b.h + pad * 2 };
    return { el: aw, inner: aw.firstElementChild, local: box, ew: W, eh: H, slideBox: b, W, H };
  }
  /** set the state of all animated targets as of `done` completed steps (instant) */
  function applyState(steps, done) {
    /* initial: entrance targets hidden, others shown */
    const seen = new Set();
    for (const s of steps) for (const it of s.items) {
      if (!it.tgt) continue;
      clearFx(it.tgt.el);
      if (it.tgt.inner) clearFx(it.tgt.inner);
      const key = it.tgt.el;
      if (!seen.has(key)) { seen.add(key); it.tgt.el.style.visibility = it.a.cls === 'entr' ? 'hidden' : ''; }
    }
    for (let i = 0; i < done && i < steps.length; i++) for (const it of steps[i].items) finish(it);
  }
  function finish(it) {
    if (!it.tgt) return;
    const el = it.tgt.el;
    const a = it.a;
    if (a.cls === 'entr') { clearFx(el); if (it.tgt.inner) clearFx(it.tgt.inner); el.style.visibility = 'visible'; }
    else if (a.cls === 'exit') { clearFx(el); el.style.visibility = 'hidden'; }
    else if (a.cls === 'emph') { frame(a, it.tgt, 1); el.style.visibility = el.style.visibility === 'hidden' ? 'hidden' : 'visible'; }
    else if (a.cls === 'path') { const pts = A.pathPoints(a); const last = pts[pts.length - 1]; const prev = it.tgt.pathOff || [0, 0]; it.tgt.pathOff = [prev[0] + last[0] * it.tgt.W, prev[1] + last[1] * it.tgt.H]; el.style.translate = `${it.tgt.pathOff[0]}px ${it.tgt.pathOff[1]}px`; }
  }
  /** play one build step; returns a controller */
  function playStep(step, onDone) {
    let rafId = 0, done = false;
    const t0 = performance.now();
    const started = new Set();
    const pathBase = new Map();
    const complete = () => { if (done) return; done = true; for (const it of step.items) it.done = false; if (onDone) onDone(); };
    const ctl = {
      finishNow() { cancelAnimationFrame(rafId); for (const it of step.items) { if (it.a.cls === 'path' && pathBase.has(it)) it.tgt.pathOff = pathBase.get(it); finish(it); } complete(); },
      cancel() { cancelAnimationFrame(rafId); done = true; },
    };
    const tick = (now) => {
      if (done) return;
      const el = now - t0;
      let all = true;
      for (const it of step.items) {
        if (!it.tgt) continue;
        if (el < it.start) { all = false; continue; }
        if (!started.has(it)) {
          started.add(it);
          if (it.a.cls === 'entr') it.tgt.el.style.visibility = 'visible';
          if (it.a.cls === 'path') pathBase.set(it, it.tgt.pathOff || [0, 0]);
        }
        const dur = it.end - it.start;
        const t = dur <= 0 ? 1 : L.clamp((el - it.start) / dur, 0, 1);
        if (t < 1) { all = false; if (it.a.cls === 'path') it.tgt.pathOff = pathBase.get(it); frame(it.a, it.tgt, t); }
        else if (!it.done) { it.done = true; if (it.a.cls === 'path') it.tgt.pathOff = pathBase.get(it); finish(it); }
      }
      if (all) { complete(); return; }
      rafId = requestAnimationFrame(tick);
    };
    rafId = requestAnimationFrame(tick);
    return ctl;
  }
  function runStep(idx, onDone) {
    const step = st.steps[idx];
    if (!step) { if (onDone) onDone(); return; }
    const ctl = playStep(step, () => { if (!st) return; st.running = null; st.step = idx + 1; if (onDone) onDone(); });
    st.running = { idx, step, finishNow: () => ctl.finishNow(), cancel: () => ctl.cancel() };
  }
  /** play a slide's animations inside an arbitrary host (task-pane previews) */
  S.playInline = function (host, slide, opts) {
    opts = opts || {};
    const el = L.render.slide(L.pres, slide, { mode: 'show' });
    host.appendChild(el);
    requestAnimationFrame(() => {
      let steps = compile(slide, el, false);
      applyState(steps, 0);
      if (opts.only) {
        let found = null;
        steps.forEach((s, i) => s.items.forEach((it) => { if (it.a.id === opts.only && !found) found = { i, it }; }));
        if (!found) { if (opts.done) setTimeout(opts.done, 300); return; }
        applyState(steps, found.i);
        steps = [{ items: [Object.assign({}, found.it, { start: 0, end: found.it.end - found.it.start })] }];
      }
      let i = 0;
      const next = () => { if (i >= steps.length) { if (opts.done) setTimeout(opts.done, 700); return; } const s = steps[i++]; playStep(s, () => setTimeout(next, 250)); };
      next();
    });
  };

  /* ---- slides ---- */
  function renderSlideEl(i) {
    const slide = L.pres.slides[i];
    const el = L.render.slide(L.pres, slide, { mode: 'show', index: i });
    return el;
  }
  function showSlide(pos, opts) {
    opts = opts || {};
    clearTimeout(st.timer);
    if (st.running) { if (st.running.cancel) st.running.cancel(); st.running = null; }
    const old = layer;
    const prevPos = st.pos;
    st.pos = pos;
    st.ended = false;
    const idx = st.order[pos];
    const slide = L.pres.slides[idx];
    if (st.rehearse && old) recordTime(prevPos);
    st.history.push(pos);
    layer = h('div', { class: 'layer' });
    const el = renderSlideEl(idx);
    layer.appendChild(el);
    stage.appendChild(layer);
    st.steps = compile(slide, el, st.noAnim);
    st.step = 0;
    applyState(st.steps, opts.atEnd ? st.steps.length : 0);
    if (opts.atEnd) st.step = st.steps.length;
    st.slideStart = performance.now();
    redrawInk();
    const afterIn = () => {
      if (old && old.parentNode) old.remove();
      if (!st || !S.active) return;
      if (!opts.atEnd && st.steps[0] && st.steps[0].auto) runStep(0, autoChain);
      else scheduleAuto();
    };
    if (opts.transition && !st.noAnim && slide.trans && slide.trans.type && slide.trans.type !== 'none' && old) transition(old, layer, slide.trans, afterIn);
    else if (opts.transition && !st.noAnim && slide.trans && slide.trans.type && slide.trans.type !== 'none' && !old) {
      const black = h('div', { class: 'layer', style: `width:${L.pres.W}px;height:${L.pres.H}px;background:#000` });
      stage.insertBefore(black, layer);
      transition(black, layer, slide.trans, afterIn);
    } else afterIn();
    updateRehearse();
    L.bus.emit('show-slide', idx);
  }
  function autoChain() {
    if (!st) return;
    /* steps that start automatically after the previous ('after'/'with' at group start) chain on */
    const next = st.steps[st.step];
    if (next && next.auto) runStep(st.step, autoChain);
    else scheduleAuto();
  }
  function scheduleAuto() {
    if (!st) return;
    clearTimeout(st.timer);
    if (st.step < st.steps.length) return;
    const slide = L.pres.slides[st.order[st.pos]];
    const tr2 = slide.trans || {};
    if (st.rehearse || !st.useTimings || st.paused) return;
    if (tr2.after != null && tr2.after !== '') {
      const elapsed = performance.now() - st.slideStart;
      st.timer = setTimeout(() => S.next(true), Math.max(0, tr2.after - elapsed));
    }
  }

  function transition(oldL, newL, t, done, cx) {
    cx = cx || { stage, state: st };
    let tt = t;
    if (t.type === 'random') { const pool = TR.filter((x) => x.type !== 'none' && x.type !== 'random'); tt = Object.assign({}, pool[Math.floor(Math.random() * pool.length)], { spd: t.spd, dur: t.dur }); }
    const dur = tt.type === 'cut' && !tt.thruBlk ? 1 : A.transDur(tt);
    const W = L.pres.W, H = L.pres.H;
    const full = { x: 0, y: 0, w: W, h: H };
    const t0 = performance.now();
    let combClone = null;
    if (tt.type === 'comb') {
      combClone = newL.cloneNode(true);
      cx.stage.appendChild(combClone);
    }
    const vec = (d) => ({ l: [-1, 0], r: [1, 0], u: [0, -1], d: [0, 1], lu: [-1, -1], ru: [1, -1], ld: [-1, 1], rd: [1, 1] }[d] || [-1, 0]);
    const step = (now) => {
      const p = L.clamp((now - t0) / dur, 0, 1);
      clearFx(newL); clearFx(oldL);
      switch (tt.type) {
        case 'cut': if (tt.thruBlk) newL.style.opacity = p < 0.5 ? '0' : '1', oldL.style.opacity = p < 0.5 ? String(1 - p * 2) : '0'; break;
        case 'fade':
          if (tt.thruBlk) { oldL.style.opacity = String(Math.max(0, 1 - p * 2)); newL.style.opacity = String(Math.max(0, p * 2 - 1)); }
          else newL.style.opacity = String(p);
          break;
        case 'push': { const [vx, vy] = vec(tt.dir); const e = easeInOut(p); newL.style.translate = `${-vx * W * (1 - e)}px ${-vy * H * (1 - e)}px`; oldL.style.translate = `${vx * W * e}px ${vy * H * e}px`; break; }
        case 'cover': { const [vx, vy] = vec(tt.dir); const e = easeInOut(p); newL.style.translate = `${-vx * W * (1 - e)}px ${-vy * H * (1 - e)}px`; break; }
        case 'pull': { const [vx, vy] = vec(tt.dir); const e = easeInOut(p); cx.stage.appendChild(oldL); oldL.style.translate = `${vx * W * e}px ${vy * H * e}px`; break; }
        case 'wipe': reveal(newL, 'wipe', { d: 't', u: 'b', r: 'l', l: 'r' }[tt.dir] || 'l', p, full, W, H); break;
        case 'split': reveal(newL, 'split', (tt.orient === 'vert' ? 'vert' : 'horz') + (tt.dir === 'in' ? 'In' : 'Out'), p, full, W, H); break;
        case 'zoom': reveal(newL, 'box', tt.dir === 'in' ? 'in' : 'out', p, full, W, H); break;
        case 'blinds': reveal(newL, 'blinds', tt.dir === 'vert' ? 'vert' : 'horz', p, full, W, H); break;
        case 'checker': reveal(newL, 'checkerboard', tt.dir === 'vert' ? 'down' : 'across', p, full, W, H, 'tr'); break;
        case 'dissolve': reveal(newL, 'dissolve', null, p, full, W, H, 'tr'); break;
        case 'randomBar': reveal(newL, 'randomBars', tt.dir === 'vert' ? 'vert' : 'horz', p, full, W, H, 'tr'); break;
        case 'circle': reveal(newL, 'circle', 'out', p, full, W, H); break;
        case 'diamond': reveal(newL, 'diamond', 'out', p, full, W, H); break;
        case 'plus': reveal(newL, 'plus', 'out', p, full, W, H); break;
        case 'wedge': reveal(newL, 'wedge', null, p, full, W, H); break;
        case 'wheel': reveal(newL, 'wheel', String(tt.spokes || 4), p, full, W, H); break;
        case 'strips': reveal(newL, 'strips', tt.dir, p, full, W, H); break;
        case 'newsflash': newL.style.transformOrigin = `${W / 2}px ${H / 2}px`; newL.style.transform = `rotate(${(1 - p) * 720}deg) scale(${Math.max(0.001, p)})`; break;
        case 'comb': {
          const e = easeInOut(p);
          const bands = 8;
          const horiz = tt.dir !== 'vert';
          const g1 = horiz ? `repeating-linear-gradient(180deg, #000 0, #000 ${50 / bands}%, transparent ${50 / bands}%, transparent ${100 / bands}%)` : `repeating-linear-gradient(90deg, #000 0, #000 ${50 / bands}%, transparent ${50 / bands}%, transparent ${100 / bands}%)`;
          const g2 = horiz ? `repeating-linear-gradient(180deg, transparent 0, transparent ${50 / bands}%, #000 ${50 / bands}%, #000 ${100 / bands}%)` : `repeating-linear-gradient(90deg, transparent 0, transparent ${50 / bands}%, #000 ${50 / bands}%, #000 ${100 / bands}%)`;
          setMask(newL, g1); setMask(combClone, g2);
          if (horiz) { newL.style.translate = `${-W * (1 - e)}px 0`; combClone.style.translate = `${W * (1 - e)}px 0`; }
          else { newL.style.translate = `0 ${-H * (1 - e)}px`; combClone.style.translate = `0 ${H * (1 - e)}px`; }
          break;
        }
        default: newL.style.opacity = String(p);
      }
      if (p < 1) myRaf = requestAnimationFrame(step);
      else { clearFx(newL); if (combClone) combClone.remove(); cx.state.inTransition = false; done(); }
    };
    let myRaf = 0;
    cx.state.inTransition = { finish: () => { cancelAnimationFrame(myRaf); clearFx(newL); if (combClone) combClone.remove(); cx.state.inTransition = false; done(); } };
    myRaf = requestAnimationFrame(step);
  }

  /* ---- navigation ---- */
  S.next = function (auto) {
    if (!S.active) return;
    if (st.blank) { setBlank(null); return; }
    if (st.ended) { S.end(); return; }
    if (st.inTransition) { st.inTransition.finish(); return; }
    if (st.running) { st.running.finishNow(); return; }
    if (st.kiosk && !auto) return;
    if (st.step < st.steps.length) { runStep(st.step, autoChain); return; }
    let p = st.pos + 1;
    while (p < st.order.length && L.pres.slides[st.order[p]].hidden && !st.preview) p++;
    if (p >= st.order.length) {
      if (st.loop) { p = 0; while (p < st.order.length - 1 && L.pres.slides[st.order[p]].hidden) p++; showSlide(p, { transition: true }); return; }
      if (st.rehearse) recordTime(st.pos);
      endScreen();
      return;
    }
    showSlide(p, { transition: true });
  };
  S.prev = function () {
    if (!S.active) return;
    if (st.blank) { setBlank(null); return; }
    if (st.ended) { st.ended = false; const es = root.querySelector('.endscreen'); if (es) es.remove(); return; }
    if (st.running) { st.running.finishNow(); }
    if (st.step > 0) { st.step--; applyState(st.steps, st.step); return; }
    let p = st.pos - 1;
    while (p >= 0 && L.pres.slides[st.order[p]].hidden) p--;
    if (p < 0) return;
    showSlide(p, { atEnd: true });
  };
  S.goto = function (slideIndex) {
    const pos = st.order.indexOf(slideIndex);
    if (pos < 0) {
      if (slideIndex >= 0 && slideIndex < L.pres.slides.length) { st.order = st.order.concat([slideIndex]); showSlide(st.order.length - 1, { transition: true }); }
      return;
    }
    showSlide(pos, { transition: true });
  };
  function endScreen() {
    st.ended = true;
    clearTimeout(st.timer);
    const es = h('div', { class: 'endscreen', text: 'End of slide show, click to exit.' });
    root.appendChild(es);
  }
  function setBlank(color) {
    const b = root.querySelector('.blank');
    if (b) b.remove();
    st.blank = color;
    if (color) root.appendChild(h('div', { class: 'blank', style: `background:${color}` }));
  }

  S.end = async function () {
    if (!S.active) return;
    cancelAnimationFrame(raf);
    if (st.running && st.running.cancel) st.running.cancel();
    if (st.inTransition && st.inTransition.finish) { const tr0 = st.inTransition; st.inTransition = false; void tr0; }
    clearInterval(st.rTimer);
    clearTimeout(st.timer);
    if (st.rehearse && !st.ended) recordTime(st.pos);
    document.removeEventListener('keydown', onKey, true);
    window.removeEventListener('resize', layout);
    try { if (document.fullscreenElement) await document.exitFullscreen(); } catch (e) { /* ignore */ }
    try { if (st.wake) st.wake.release(); } catch (e) { /* ignore */ }
    const inkData = st.ink;
    const rehearse = st.rehearse ? st.times : null;
    const lastIdx = st.order[st.pos];
    root.remove();
    root = null; layer = null;
    S.active = false;
    const state = st;
    st = null;
    L.ui.closeMenus();
    const hasInk = Object.values(inkData).some((a) => a && a.length);
    if (hasInk) {
      const r = await L.ui.msg('Do you want to keep your ink annotations?', { icon: 'question', buttons: ['&Keep', '&Discard'] });
      if (r === 0) keepInk(inkData);
    }
    if (rehearse) {
      const total = Object.values(rehearse).reduce((a, b) => a + b, 0);
      const r = await L.ui.msg(`The total time for the slide show was ${fmtTime(total)}. Do you want to keep the new slide timings?`, { icon: 'question', buttons: ['&Yes', '&No'] });
      if (r === 0) {
        L.hist.push('Rehearse Timings');
        for (const k in rehearse) { const s = L.pres.slides[+k]; if (s) { s.trans = Object.assign({}, s.trans, { after: Math.round(rehearse[k]) }); } }
        L.pres.show.useTimings = true;
        L.bus.emit('slides-changed');
        if (L.app) L.app.setView('sorter');
      }
    }
    void state;
    if (L.ed && L.app && L.app.view !== 'sorter' && lastIdx != null) L.ed.goto(lastIdx);
    L.bus.emit('show-end');
    L.ed && L.ed.refocus();
  };
  function keepInk(data) {
    L.hist.push('Ink Annotations');
    const savedIdx = L.ed.idx, savedView = L.ed.view;
    for (const k in data) {
      const strokes = data[k];
      if (!strokes || !strokes.length) continue;
      const slide = L.pres.slides[+k];
      if (!slide) continue;
      L.ed.idx = +k;
      L.ed.view = 'normal';
      for (const s of strokes) {
        if (s.pts.length < 2) continue;
        const sh = L.ed.addFreeform(s.pts, false, true, { line: { c: s.color, w: s.w, dash: 'solid', a: s.alpha, cap: 'rnd', join: 'round' }, noHistory: true, noRender: true });
        if (sh) sh.name = 'Ink ' + slide.shapes.length;
      }
    }
    L.ed.idx = savedIdx; L.ed.view = savedView;
    L.ed.render();
    L.bus.emit('slides-changed');
  }

  /* ---- ink ---- */
  function redrawInk() {
    if (!inkCanvas || !st) return;
    const g = inkCanvas.getContext('2d');
    g.setTransform(1, 0, 0, 1, 0, 0);
    g.clearRect(0, 0, inkCanvas.width, inkCanvas.height);
    const k = (st.k || 1) * (window.devicePixelRatio || 1);
    g.setTransform(k, 0, 0, k, 0, 0);
    const strokes = st.ink[st.order[st.pos]] || [];
    for (const s of strokes) {
      g.globalAlpha = s.alpha == null ? 1 : s.alpha;
      g.strokeStyle = s.color; g.lineWidth = s.w; g.lineCap = 'round'; g.lineJoin = 'round';
      g.beginPath();
      s.pts.forEach((p, i) => (i ? g.lineTo(p[0], p[1]) : g.moveTo(p[0], p[1])));
      g.stroke();
    }
    g.globalAlpha = 1;
  }
  function toSlidePt(e) {
    const r = inkCanvas.getBoundingClientRect();
    return [(e.clientX - r.left) / st.k, (e.clientY - r.top) / st.k];
  }
  let drawing = null;
  function onPointerDown(e) {
    if (!st) return;
    if (e.target.closest('.popbar') || e.target.closest('.rehearse')) return;
    idleCursor();
    if (e.button === 2) return;
    if (st.pen && e.button === 0) {
      e.preventDefault();
      root.setPointerCapture(e.pointerId);
      const key = st.order[st.pos];
      if (st.pen === 'eraser') { drawing = { erase: true }; eraseAt(toSlidePt(e)); return; }
      const style = st.pen === 'highlighter' ? { w: 14, alpha: 0.4, color: '#FFFF00' } : st.pen === 'felt' ? { w: 5, alpha: 1, color: st.penColor } : { w: 2, alpha: 1, color: st.penColor };
      drawing = Object.assign({ pts: [toSlidePt(e)] }, style);
      (st.ink[key] = st.ink[key] || []).push(drawing);
      return;
    }
    const link = e.target.closest && e.target.closest('[data-link]');
    if (link && e.button === 0) { e.preventDefault(); follow(JSON.parse(link.dataset.link)); return; }
    if (e.button === 0) { st._down = true; }
  }
  function onPointerMove(e) {
    if (!st) return;
    idleCursor();
    if (drawing) {
      if (drawing.erase) { eraseAt(toSlidePt(e)); return; }
      drawing.pts.push(toSlidePt(e));
      redrawInk();
    }
  }
  function onPointerUp(e) {
    if (!st) return;
    if (drawing) { drawing = null; return; }
    if (st._down && e.button === 0 && !st.pen) { st._down = false; S.next(); }
    st._down = false;
  }
  function eraseAt(p) {
    const key = st.order[st.pos];
    const list = st.ink[key];
    if (!list) return;
    st.ink[key] = list.filter((s) => !s.pts.some((q) => Math.hypot(q[0] - p[0], q[1] - p[1]) < 10));
    redrawInk();
  }
  function onWheel(e) { if (!st || st.pen) return; if (e.deltaY > 0) S.next(); else if (e.deltaY < 0) S.prev(); }
  function follow(link) {
    if (link.url) { const a = h('a', { href: link.url, target: '_blank', rel: 'noopener noreferrer' }); document.body.appendChild(a); a.click(); a.remove(); return; }
    if (link.slide) { const i = L.pres.slides.findIndex((s) => s.id === link.slide); if (i >= 0) S.goto(i); return; }
    switch (link.action) {
      case 'next': S.next(); break;
      case 'prev': S.prev(); break;
      case 'first': showSlide(0, { transition: true }); break;
      case 'last': showSlide(st.order.length - 1, { transition: true }); break;
      case 'lastViewed': { const h2 = st.history; if (h2.length > 1) showSlide(h2[h2.length - 2], { transition: true }); break; }
      case 'end': S.end(); break;
      default: break;
    }
  }
  function setPen(kind) {
    st.pen = kind;
    root.classList.toggle('pen', !!kind);
    if (!kind) root.style.cursor = '';
  }
  function idleCursor() {
    if (!root) return;
    root.classList.remove('nocursor');
    clearTimeout(st && st._idle);
    if (st) st._idle = setTimeout(() => { if (root && st && !st.pen) root.classList.add('nocursor'); }, 3000);
  }

  /* ---- menus ---- */
  function pointerMenu() {
    return [
      { label: '&Arrow', run: () => setPen(null), checked: () => !st.pen },
      { label: '&Ballpoint Pen', run: () => setPen('pen'), checked: () => st.pen === 'pen' },
      { label: '&Felt Tip Pen', run: () => setPen('felt'), checked: () => st.pen === 'felt' },
      { label: '&Highlighter', run: () => setPen('highlighter'), checked: () => st.pen === 'highlighter' },
      { label: 'Ink &Color', sub: ['#FF0000', '#0000FF', '#00B050', '#000000', '#FFFFFF', '#FFC000', '#7030A0'].map((c) => ({ label: { '#FF0000': 'Red', '#0000FF': 'Blue', '#00B050': 'Green', '#000000': 'Black', '#FFFFFF': 'White', '#FFC000': 'Gold', '#7030A0': 'Purple' }[c], html: `<span style="display:inline-block;width:12px;height:12px;background:${c};border:1px solid #555"></span>`, run: () => { st.penColor = c; if (!st.pen) setPen('pen'); }, checked: () => st.penColor === c })) },
      '-',
      { label: '&Eraser', run: () => setPen('eraser'), checked: () => st.pen === 'eraser' },
      { label: 'E&rase All Ink on Slide', run: () => { st.ink[st.order[st.pos]] = []; redrawInk(); } },
    ];
  }
  function showMenu() {
    const slides = st.order.map((i, pos) => ({ label: `&${pos + 1} ${L.model.slideTitle(L.pres.slides[i]) || 'Slide ' + (i + 1)}`.slice(0, 60), run: () => showSlide(pos, { transition: true }), checked: () => st.pos === pos }));
    return [
      { label: '&Next', run: () => S.next() },
      { label: '&Previous', run: () => S.prev() },
      { label: 'Last &Viewed', run: () => follow({ action: 'lastViewed' }) },
      { label: '&Go to Slide', sub: slides.length ? slides : [{ label: '(none)', disabled: true }] },
      '-',
      { label: '&Screen', sub: [{ label: '&Black Screen', run: () => setBlank('#000') }, { label: '&White Screen', run: () => setBlank('#fff') }, { label: st.paused ? '&Resume' : '&Pause', run: () => togglePause() }] },
      { label: 'P&ointer Options', sub: pointerMenu() },
      { label: '&Help', run: showHelp },
      '-',
      { label: '&End Show', run: () => S.end() },
    ];
  }
  function onContext(e) {
    e.preventDefault();
    if (!st) return;
    if (st.pen) { setPen(null); }
    ui().openMenu(showMenu(), { x: e.clientX, y: e.clientY });
  }
  const ui = () => L.ui;
  function togglePause() {
    st.paused = !st.paused;
    if (st.paused) clearTimeout(st.timer); else scheduleAuto();
  }
  function showHelp() {
    L.ui.msg('Slide Show Help\n\nN, Enter, Page Down, Right Arrow, Down Arrow or Spacebar: next animation or slide\nP, Page Up, Left Arrow, Up Arrow or Backspace: previous animation or slide\nNumber + Enter: go to that slide\nB or Period: black / unblack screen\nW or Comma: white / unwhite screen\nS or Plus: stop or restart automatic show\nEsc, Ctrl+Break or Hyphen: end slide show\nE: erase drawing on screen\nCtrl+P: pen  ·  Ctrl+A: arrow  ·  Ctrl+E: eraser  ·  Ctrl+H: hide pointer\nHome: first slide  ·  End: last slide', { title: 'Slide Show Help', width: 500 });
  }

  function onKey(e) {
    if (!S.active) return;
    if (L.ui.dialogOpen()) return;
    if (L.ui.menuOpen()) { if (L.ui.menuKey(e)) { e.preventDefault(); e.stopPropagation(); } return; }
    const k = e.key;
    e.stopPropagation();
    if (/^[0-9]$/.test(k)) { st.gotoBuf += k; showGoto(); e.preventDefault(); return; }
    if (k === 'Enter' && st.gotoBuf) { const n = parseInt(st.gotoBuf, 10) - 1; st.gotoBuf = ''; hideGoto(); if (n >= 0) { const pos = st.order.indexOf(n); if (pos >= 0) showSlide(pos, { transition: true }); } e.preventDefault(); return; }
    if (st.gotoBuf && k !== 'Enter') { st.gotoBuf = ''; hideGoto(); }
    if (e.ctrlKey && (k === 'p' || k === 'P')) { e.preventDefault(); setPen('pen'); return; }
    if (e.ctrlKey && (k === 'a' || k === 'A')) { e.preventDefault(); setPen(null); return; }
    if (e.ctrlKey && (k === 'e' || k === 'E')) { e.preventDefault(); setPen('eraser'); return; }
    if (e.ctrlKey && (k === 'h' || k === 'H')) { e.preventDefault(); root.classList.add('nocursor'); return; }
    if (e.ctrlKey && k === 'Pause') { e.preventDefault(); S.end(); return; }
    switch (k) {
      case 'Escape': case '-': e.preventDefault(); if (st.pen) { setPen(null); return; } S.end(); return;
      case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'n': case 'N': case 'Enter':
        e.preventDefault(); S.next(); return;
      case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace': case 'p': case 'P':
        e.preventDefault(); S.prev(); return;
      case 'Home': e.preventDefault(); showSlide(0, { transition: true }); return;
      case 'End': e.preventDefault(); showSlide(st.order.length - 1, { transition: true }); return;
      case 'b': case 'B': case '.': e.preventDefault(); setBlank(st.blank === '#000' ? null : '#000'); return;
      case 'w': case 'W': case ',': e.preventDefault(); setBlank(st.blank === '#fff' ? null : '#fff'); return;
      case 'e': case 'E': e.preventDefault(); st.ink[st.order[st.pos]] = []; redrawInk(); return;
      case 's': case 'S': case '+': e.preventDefault(); togglePause(); return;
      case 'F1': e.preventDefault(); showHelp(); return;
      default: break;
    }
  }
  function showGoto() {
    let g = root.querySelector('.gotobox');
    if (!g) { g = h('div', { class: 'gotobox' }); root.appendChild(g); }
    g.textContent = 'Go to slide ' + st.gotoBuf;
  }
  function hideGoto() { const g = root && root.querySelector('.gotobox'); if (g) g.remove(); }

  /* ---- rehearsal ---- */
  function buildRehearse() {
    const bar = h('div', { class: 'rehearse', role: 'toolbar', 'aria-label': 'Rehearsal' },
      h('button', { type: 'button', class: 'tb-btn', 'aria-label': 'Next', html: L.icons.get('forward'), onclick: (e) => { e.stopPropagation(); S.next(); } }),
      h('button', { type: 'button', class: 'tb-btn', 'aria-label': 'Pause', text: '❚❚', onclick: (e) => { e.stopPropagation(); st.rPaused = !st.rPaused; if (st.rPaused) st.rPauseAt = performance.now(); else st.slideStart += performance.now() - st.rPauseAt; } }),
      h('span', { class: 'rt', text: '0:00:00' }),
      h('button', { type: 'button', class: 'tb-btn', 'aria-label': 'Repeat', html: L.icons.get('undo'), onclick: (e) => { e.stopPropagation(); st.slideStart = performance.now(); } }),
      h('span', { class: 'tot', text: '0:00:00' }));
    root.appendChild(bar);
    st.rTimer = setInterval(updateRehearse, 250);
  }
  function updateRehearse() {
    if (!st || !st.rehearse || !root) return;
    const bar = root.querySelector('.rehearse');
    if (!bar) return;
    const now = st.rPaused ? st.rPauseAt : performance.now();
    const cur = now - st.slideStart;
    const tot = Object.values(st.times).reduce((a, b) => a + b, 0) + cur;
    bar.querySelector('.rt').textContent = fmtTime(cur);
    bar.querySelector('.tot').textContent = fmtTime(tot);
  }
  function recordTime(pos) {
    if (!st || !st.rehearse) return;
    const idx = st.order[pos];
    st.times[idx] = performance.now() - st.slideStart;
  }
  function fmtTime(ms) { const s = Math.floor(ms / 1000); return `${Math.floor(s / 3600)}:${String(Math.floor(s / 60) % 60).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`; }

  /* preview helpers used by task panes (AutoPreview) */
  S.previewTransition = function (host, pres, slide, prevSlide) {
    const W = pres.W;
    const k = host.clientWidth / W;
    if (!k) return;
    L.clear(host);
    const stg = h('div', { style: `position:absolute;left:0;top:0;width:${pres.W}px;height:${pres.H}px;transform:scale(${k});transform-origin:0 0;overflow:hidden` });
    host.appendChild(stg);
    const a = h('div', { class: 'layer', style: 'position:absolute;left:0;top:0' }), b = h('div', { class: 'layer', style: 'position:absolute;left:0;top:0' });
    a.appendChild(prevSlide ? L.render.slide(pres, prevSlide, { mode: 'thumb' }) : h('div', { style: `width:${pres.W}px;height:${pres.H}px;background:#000` }));
    b.appendChild(L.render.slide(pres, slide, { mode: 'thumb' }));
    stg.append(a, b);
    transition(a, b, Object.assign({ type: 'fade' }, slide.trans && slide.trans.type !== 'none' ? slide.trans : { type: 'fade' }), () => { a.remove(); }, { stage: stg, state: {} });
  };
})();
