/* VibeOffice — collect visible DOM words and canvas text for the PDF text layer.
 * Original implementation using CSSOM View's Range geometry (font metrics, in client pixels):
 * https://www.w3.org/TR/cssom-view-1/#dom-range-getclientrects
 * DOM order supplies reading order; rectangles supply placement, never the page's text.
 */
(function () {
  'use strict';
  const L = window.L;
  const SKIP = 'script,style,canvas,defs,title,desc,annotation,annotation-xml,' +
    '[aria-hidden="true"],svg.warp,.format-marks,.sectmark,.hbrk,.crop,.balloons';
  const WARP_BODY = '.shp.warped > .tbx,.shp.warped > .tx';
  const SPACER = '.tb,.lbt,.lsp,img,br';
  const MATH = 'http://www.w3.org/1998/Math/MathML';
  const BLOCK = /^(block|flow-root|flex|grid|table-cell|list-item)$/;
  const CLIP = /^(hidden|clip|auto|scroll)$/;
  const segmenter = typeof Intl.Segmenter === 'function' ? new Intl.Segmenter(undefined, { granularity: 'grapheme' }) : null;
  const box = r => ({ left: r.left, top: r.top, right: r.right, bottom: r.bottom });
  const union = (a, b) => ({ left: Math.min(a.left, b.left), top: Math.min(a.top, b.top),
    right: Math.max(a.right, b.right), bottom: Math.max(a.bottom, b.bottom) });
  const intersection = (a, b) => ({ left: Math.max(a.left, b.left), top: Math.max(a.top, b.top),
    right: Math.min(a.right, b.right), bottom: Math.min(a.bottom, b.bottom) });
  const area = r => r.right > r.left && r.bottom > r.top;
  const inside = (r, clip) => r.left >= clip.left && r.right <= clip.right &&
    r.top >= clip.top && r.bottom <= clip.bottom;

  function sameLine(a, b, vertical) {
    const start = vertical ? 'left' : 'top', end = vertical ? 'right' : 'bottom';
    const overlap = Math.min(a[end], b[end]) - Math.max(a[start], b[start]);
    return overlap > Math.min(a[end] - a[start], b[end] - b[start]) * 0.35;
  }
  function boundaries(text) {
    const result = [0];
    if (segmenter) for (const part of segmenter.segment(text)) result.push(part.index + part.segment.length);
    else for (const ch of text) result.push(result[result.length - 1] + ch.length);
    return result;
  }

  const corners = r => [
    { x: r.left, y: r.top }, { x: r.right, y: r.top },
    { x: r.right, y: r.bottom }, { x: r.left, y: r.bottom },
  ];
  const bounds = ps => ({ left: Math.min(...ps.map(p => p.x)), right: Math.max(...ps.map(p => p.x)),
    top: Math.min(...ps.map(p => p.y)), bottom: Math.max(...ps.map(p => p.y)) });

  // Range rectangles lose their orientation at 45 degrees. Measure a temporary clone
  // without HTML transforms, then map its words through the original ancestor matrices.
  // CSS transforms do not change line layout. Quire's default path never makes this clone.
  // https://www.w3.org/TR/css-transforms-1/#transform-rendering
  function transformedPage(page) {
    const source = [page, ...page.querySelectorAll('*')], transforms = new Map();
    let rotated = false;
    for (let i = 0; i < source.length; i++) {
      const el = source[i];
      if (el.namespaceURI !== 'http://www.w3.org/1999/xhtml') continue;
      const css = getComputedStyle(el);
      if (css.transform === 'none') continue;
      const matrix = new DOMMatrix(css.transform);
      if (!matrix.is2D) return null;
      transforms.set(i, { matrix, origin: css.transformOrigin.split(' ').map(parseFloat) });
      rotated ||= Math.abs(matrix.b) > 1e-6 || Math.abs(matrix.c) > 1e-6 || matrix.a < 0 || matrix.d < 0;
    }
    if (!rotated) return null;
    const host = page.ownerDocument.createElement('div'), clone = page.cloneNode(true);
    const css = getComputedStyle(page);
    host.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;';
    host.style.font = css.font;
    host.style.direction = css.direction;
    host.inert = true;
    clone.style.position = 'relative';
    clone.style.left = clone.style.top = clone.style.margin = '0';
    const nodes = [clone, ...clone.querySelectorAll('*')];
    for (const index of transforms.keys()) nodes[index].style.setProperty('transform', 'none', 'important');
    host.append(clone);
    page.ownerDocument.body.append(host);
    try {
      const origin = clone.getBoundingClientRect(), matrices = new WeakMap(), identity = new DOMMatrix();
      for (let i = 0; i < nodes.length; i++) {
        const el = nodes[i], transform = transforms.get(i);
        let matrix = el === clone ? identity : matrices.get(el.parentElement);
        if (transform) {
          const r = el.getBoundingClientRect();
          const x = r.left - origin.left + transform.origin[0];
          const y = r.top - origin.top + transform.origin[1];
          matrix = matrix.multiply(new DOMMatrix().translate(x, y)
            .multiply(transform.matrix).translate(-x, -y));
        }
        matrices.set(el, matrix);
      }
      return { page: clone, matrices, identity, remove: () => host.remove() };
    } catch (e) {
      host.remove();
      throw e;
    }
  }

  function clipEdge(points, axis, edge, greater) {
    const result = [];
    const inside = p => greater ? p[axis] >= edge : p[axis] <= edge;
    for (let i = 0; i < points.length; i++) {
      const a = points[i], b = points[(i + 1) % points.length];
      const ai = inside(a), bi = inside(b);
      if (ai) result.push(a);
      if (ai !== bi) {
        const t = (edge - a[axis]) / (b[axis] - a[axis]);
        result.push({ x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t });
      }
    }
    return result;
  }

  /** scale is points per client pixel (for example 72 / 96 / the page's zoom).
   * Clipping stops at the page: scrolling the editor never removes words from a PDF.
   * The optional transform pass measures a clone; the original DOM is never changed. */
  function collect(page, scale, geometry) {
    const origin = page.getBoundingClientRect();
    if (!area(origin) || !(scale > 0)) return [];
    const states = new WeakMap(), pieces = new WeakMap(), lines = [];
    const range = page.ownerDocument.createRange();
    let previous, spaced = true;

    function state(el) {
      let value = states.get(el);
      if (value) return value;
      const parent = el === page ? { hidden: false, clip: box(origin), block: page,
        clips: geometry ? [{ matrix: geometry.identity, inverse: geometry.identity, x: true, y: true,
          rect: { left: 0, top: 0, right: origin.width, bottom: origin.height } }] : null } : state(el.parentElement);
      const css = getComputedStyle(el);
      const hidden = parent.hidden || el.matches(SKIP) || css.display === 'none' ||
        css.contentVisibility === 'hidden' || (+css.opacity === 0 && !el.matches(WARP_BODY));
      let clip = parent.clip, clips = parent.clips;
      const matrix = geometry?.matrices.get(el);
      if (!hidden && (CLIP.test(css.overflowX) || CLIP.test(css.overflowY))) {
        const r = el.getBoundingClientRect();
        clip = { ...clip };
        // overflow clips to the padding box. Client dimensions omit borders and scrollbars.
        const sx = el.offsetWidth ? r.width / el.offsetWidth : 1;
        const sy = el.offsetHeight ? r.height / el.offsetHeight : 1;
        if (geometry) {
          const left = r.left - origin.left + (el.clientLeft || 0) * sx;
          const top = r.top - origin.top + (el.clientTop || 0) * sy;
          clips = clips.concat({ matrix, inverse: matrix.inverse(),
            x: CLIP.test(css.overflowX), y: CLIP.test(css.overflowY),
            rect: { left, top, right: left + (el.clientWidth ?? r.width) * sx,
              bottom: top + (el.clientHeight ?? r.height) * sy } });
        } else if (CLIP.test(css.overflowX)) {
          const left = r.left + (el.clientLeft || 0) * sx;
          clip.left = Math.max(clip.left, left);
          clip.right = Math.min(clip.right, left + (el.clientWidth ?? r.width) * sx);
        }
        if (!geometry && CLIP.test(css.overflowY)) {
          const top = r.top + (el.clientTop || 0) * sy;
          clip.top = Math.max(clip.top, top);
          clip.bottom = Math.min(clip.bottom, top + (el.clientHeight ?? r.height) * sy);
        }
      }
      const block = BLOCK.test(css.display) || el.localName === 'text' ? el : parent.block;
      const group = el.namespaceURI === MATH && /^(mi|mn|mo|mtext|ms)$/.test(el.localName) ? el : parent.group;
      value = { hidden, visible: css.visibility === 'visible', clip, clips, matrix, block, group,
        vertical: /^(vertical|sideways)/.test(css.writingMode) };
      states.set(el, value);
      return value;
    }

    function rects(node, start, end) {
      range.setStart(node, start);
      range.setEnd(node, end);
      return Array.from(range.getClientRects()).filter(area).map(box);
    }
    const oneLine = (rs, info) => rs.length && rs.every(r => sameLine(rs[0], r, info.vertical));
    const local = r => ({ left: r.left - origin.left, right: r.right - origin.left,
      top: r.top - origin.top, bottom: r.bottom - origin.top });
    const points = (r, info) => {
      const b = local(r), m = info.matrix;
      if (!m) return { x: b.left * scale, y: b.top * scale,
        w: (b.right - b.left) * scale, h: (b.bottom - b.top) * scale };
      const sx = Math.hypot(m.a, m.b), sy = Math.hypot(m.c, m.d);
      if (m.a * m.d - m.b * m.c > 0 && Math.abs(m.a * m.c + m.b * m.d) < sx * sy * 1e-5) {
        const p = m.transformPoint({ x: b.left, y: b.top });
        return { x: p.x * scale, y: p.y * scale, w: (b.right - b.left) * sx * scale,
          h: (b.bottom - b.top) * sy * scale, angle: Math.atan2(m.b, m.a) * 180 / Math.PI };
      }
      // Skewed or reflected text still copies, with an axis-aligned selection box.
      const p = bounds(corners(b).map(p => m.transformPoint(p)));
      return { x: p.left * scale, y: p.top * scale, w: (p.right - p.left) * scale,
        h: (p.bottom - p.top) * scale };
    };

    function clipRect(r, info) {
      if (!geometry) return { rect: intersection(r, info.clip), whole: inside(r, info.clip) };
      let ps = corners(local(r)).map(p => info.matrix.transformPoint(p)), whole = true;
      for (const clip of info.clips) {
        ps = ps.map(p => clip.inverse.transformPoint(p));
        const b = clip.rect;
        if (clip.x) {
          whole &&= ps.every(p => p.x >= b.left - 1e-6 && p.x <= b.right + 1e-6);
          ps = clipEdge(clipEdge(ps, 'x', b.left, true), 'x', b.right, false);
        }
        if (clip.y) {
          whole &&= ps.every(p => p.y >= b.top - 1e-6 && p.y <= b.bottom + 1e-6);
          ps = clipEdge(clipEdge(ps, 'y', b.top, true), 'y', b.bottom, false);
        }
        if (!ps.length) return null;
        ps = ps.map(p => clip.matrix.transformPoint(p));
      }
      const inverse = info.matrix.inverse(), b = bounds(ps.map(p => inverse.transformPoint(p)));
      return { whole, rect: { left: b.left + origin.left, right: b.right + origin.left,
        top: b.top + origin.top, bottom: b.bottom + origin.top } };
    }

    function add(text, r, info, node, start, end) {
      if (!text || !area(r)) return;
      const continues = previous && previous.block === info.block && previous.matrix === info.matrix &&
        sameLine(previous.rect, r, info.vertical);
      if (!continues) lines.push([]);
      const line = lines[lines.length - 1];
      // A word may cross several styled spans. Only actual whitespace separates its pieces.
      if (continues && !spaced && previous.group === info.group) {
        const word = line[line.length - 1];
        word.text += text;
        previous.rect = union(previous.rect, r);
        Object.assign(word, points(previous.rect, info));
      } else {
        line.push({ text, ...points(r, info) });
        previous = { rect: r, block: info.block, group: info.group, matrix: info.matrix };
      }
      const word = line[line.length - 1];
      if (!pieces.has(word)) pieces.set(word, []);
      pieces.get(word).push({ node, start, end, vertical: info.vertical });
      spaced = false;
    }

    function clipped(node, start, end, rs, info) {
      const r = rs.reduce(union);
      const view = clipRect(r, info);
      if (view?.whole) {
        add(node.data.slice(start, end), r, info, node, start, end);
        return;
      }
      if (!view || !area(view.rect)) {
        spaced = true;
        return;
      }
      // Only a word touching a clip edge needs character ranges. Keep whole graphemes,
      // including surrogate pairs and combining sequences, and omit the hidden suffix.
      const cuts = boundaries(node.data.slice(start, end));
      let text = '', visible, from, to;
      const flush = () => {
        if (visible) add(text, visible, info, node, from, to);
        text = '';
        visible = null;
      };
      for (let i = 1; i < cuts.length; i++) {
        const chars = rects(node, start + cuts[i - 1], start + cuts[i]);
        const part = chars.length && clipRect(chars.reduce(union), info)?.rect;
        if (part && area(part)) {
          if (!visible) from = start + cuts[i - 1];
          to = start + cuts[i];
          text += node.data.slice(start + cuts[i - 1], start + cuts[i]);
          visible = visible ? union(visible, part) : part;
        } else {
          flush();
          spaced = true;
        }
      }
      flush();
    }

    function word(node, start, end, info) {
      let rs = rects(node, start, end);
      if (!rs.length) return;
      if (oneLine(rs, info)) {
        clipped(node, start, end, rs, info);
        return;
      }
      // CJK and long unbroken words may wrap. Find each line boundary by binary search
      // instead of measuring every character on every ordinary page.
      const cuts = boundaries(node.data.slice(start, end));
      let begin = 0;
      while (begin < cuts.length - 1) {
        let lo = begin + 1, hi = cuts.length - 1;
        while (lo < hi) {
          const mid = Math.ceil((lo + hi) / 2);
          if (oneLine(rects(node, start + cuts[begin], start + cuts[mid]), info)) lo = mid;
          else hi = mid - 1;
        }
        rs = rects(node, start + cuts[begin], start + cuts[lo]);
        if (rs.length) clipped(node, start + cuts[begin], start + cuts[lo], rs, info);
        begin = lo;
        spaced = true;
      }
    }

    const walker = page.ownerDocument.createTreeWalker(page, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
      acceptNode: node => node.nodeType === 1 && node.matches(SKIP) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT,
    });
    for (let node; (node = walker.nextNode());) {
      if (node.nodeType === 1) {
        if (node.matches(SPACER) && !state(node).hidden) {
          spaced = true;
          if (node.localName === 'br') previous = null;
        }
        continue;
      }
      if (!node.data) continue;
      const info = state(node.parentElement);
      if (info.hidden || !info.visible || !area(info.clip)) continue;
      for (const part of node.data.matchAll(/\s+|\S+/gu)) {
        if (/^\s/u.test(part[0])) spaced = true;
        else word(node, part.index, part.index + part[0].length, info);
      }
    }
    // PDF readers reconstruct bidi text from visual character order. For actual DOM
    // words, the browser already resolved punctuation, digits and nested directions:
    // retain that order rather than reversing individual letter runs around punctuation.
    // Only RTL words need these extra ranges; ordinary pages keep one range per word.
    for (const line of lines) for (const word of line) {
      if (!L.pdf.hasRTL(word.text)) continue;
      const glyphs = [];
      for (const part of pieces.get(word)) {
        const text = part.node.data.slice(part.start, part.end), cuts = boundaries(text);
        for (let i = 1; i < cuts.length; i++) {
          const rs = rects(part.node, part.start + cuts[i - 1], part.start + cuts[i]);
          if (!rs.length) continue;
          const r = rs.reduce(union);
          glyphs.push({ text: text.slice(cuts[i - 1], cuts[i]),
            at: part.vertical ? (r.top + r.bottom) / 2 : (r.left + r.right) / 2 });
        }
      }
      word.visual = glyphs.sort((a, b) => a.at - b.at).map(g => g.text).join('');
    }
    return lines;
  }

  L.pdf.collectDOM = function (page, scale = 0.75, options) {
    const geometry = options?.transforms ? transformedPage(page) : null;
    try { return collect(geometry?.page || page, scale, geometry); }
    finally { geometry?.remove(); }
  };

  const canvasRecords = new WeakMap();
  const FONT_STATE = ['font', 'fontKerning', 'fontStretch', 'fontVariantCaps', 'letterSpacing',
    'wordSpacing', 'textRendering', 'direction', 'lang'];
  const hasArea = r => r.right - r.left > 1e-6 && r.bottom - r.top > 1e-6;
  const canvasRect = (x, y, w, h) => ({ left: Math.min(x, x + w), top: Math.min(y, y + h),
    right: Math.max(x, x + w), bottom: Math.max(y, y + h) });
  function region(rect, matrix = new DOMMatrix()) {
    const device = !matrix.b && !matrix.c ? canvasRect(matrix.a * rect.left + matrix.e,
      matrix.d * rect.top + matrix.f, matrix.a * (rect.right - rect.left),
      matrix.d * (rect.bottom - rect.top)) : null;
    return { rect, matrix, inverse: matrix.inverse(), device };
  }

  function clipPolygon(ps, clip) {
    ps = ps.map(p => clip.inverse.transformPoint(p));
    const r = clip.rect;
    ps = clipEdge(clipEdge(ps, 'x', r.left, true), 'x', r.right, false);
    ps = clipEdge(clipEdge(ps, 'y', r.top, true), 'y', r.bottom, false);
    return ps.map(p => clip.matrix.transformPoint(p));
  }
  function canvasClip(ps, clips) {
    // Ordinary cells and page clips need only rectangle intersections. Keep the
    // polygon path for rotated text and rotated clips.
    const rect = bounds(ps);
    if (ps.length === 4 && ps.every(p => (p.x === rect.left || p.x === rect.right) &&
        (p.y === rect.top || p.y === rect.bottom)) && clips.every(rs => rs.length === 1 && rs[0].device)) {
      let clipped = rect;
      for (const rs of clips) clipped = intersection(clipped, rs[0].device);
      return hasArea(clipped) ? { whole: inside(rect, clipped), points: corners(clipped) } : null;
    }
    let polygons = [ps], whole = true;
    for (const regions of clips) {
      whole &&= regions.some(c => ps.every(p => {
        const q = c.inverse.transformPoint(p), r = c.rect;
        return q.x >= r.left - 1e-6 && q.x <= r.right + 1e-6 &&
          q.y >= r.top - 1e-6 && q.y <= r.bottom + 1e-6;
      }));
      polygons = polygons.flatMap(p => regions.map(c => clipPolygon(p, c)))
        .filter(p => p.length && hasArea(bounds(p)));
      if (!polygons.length) return null;
    }
    return { whole, points: polygons.flat() };
  }
  function canvasBox(r, m, scale) {
    const sx = Math.hypot(m.a, m.b), sy = Math.hypot(m.c, m.d);
    if (m.a * m.d - m.b * m.c > 0 && Math.abs(m.a * m.c + m.b * m.d) < sx * sy * 1e-5) {
      const p = m.transformPoint({ x: r.left, y: r.top });
      return { x: p.x * scale, y: p.y * scale, w: (r.right - r.left) * sx * scale,
        h: (r.bottom - r.top) * sy * scale, angle: Math.atan2(m.b, m.a) * 180 / Math.PI };
    }
    const p = bounds(corners(r).map(p => m.transformPoint(p)));
    return { x: p.left * scale, y: p.top * scale,
      w: (p.right - p.left) * scale, h: (p.bottom - p.top) * scale };
  }

  // The canvas exposes a whole string's metrics, not bidi subranges. Use the browser's
  // inline layout only for RTL strings, as collectDOM does, retaining its visual order.
  function canvasBidi(ctx, text, metrics) {
    const doc = ctx.canvas.ownerDocument || document, span = doc.createElement('span');
    span.style.cssText = 'all:initial;position:fixed;left:-100000px;top:0;white-space:pre;width:max-content';
    for (const key of FONT_STATE) if (key in span.style && ctx[key]) span.style[key] = ctx[key];
    span.style.direction = ctx.direction === 'inherit' && ctx.canvas instanceof Element
      ? getComputedStyle(ctx.canvas).direction : ctx.direction;
    span.textContent = text;
    doc.body.append(span);
    try {
      const origin = span.getBoundingClientRect(), range = doc.createRange(), out = [];
      const scale = origin.width ? metrics.width / origin.width : 1;
      for (const part of text.matchAll(/\S+/gu)) {
        const cuts = boundaries(part[0]), glyphs = [];
        for (let i = 1; i < cuts.length; i++) {
          range.setStart(span.firstChild, part.index + cuts[i - 1]);
          range.setEnd(span.firstChild, part.index + cuts[i]);
          const r = range.getBoundingClientRect();
          glyphs.push({ text: part[0].slice(cuts[i - 1], cuts[i]),
            left: (r.left - origin.left) * scale, right: (r.right - origin.left) * scale });
        }
        out.push({ text: part[0], left: Math.min(...glyphs.map(g => g.left)),
          right: Math.max(...glyphs.map(g => g.right)), glyphs });
      }
      return out;
    } finally { span.remove(); }
  }

  function canvasWords(ctx, text, x, y, maxWidth) {
    const metrics = ctx.measureText(text), width = metrics.width;
    if (!(width > 0) || !(maxWidth > 0)) return [];
    const squeeze = Math.min(1, maxWidth / width), matrix = ctx.getTransform();
    if (!matrix.is2D || Math.abs(matrix.a * matrix.d - matrix.b * matrix.c) < 1e-12) return [];
    const state = Object.fromEntries(FONT_STATE.filter(key => key in ctx).map(key => [key, ctx[key]]));
    const direction = ctx.direction === 'inherit' && ctx.canvas instanceof Element
      ? getComputedStyle(ctx.canvas).direction : ctx.direction;
    const align = ctx.textAlign;
    const right = align === 'right' || (align === 'start' && direction === 'rtl') ||
      (align === 'end' && direction !== 'rtl');
    const left = x - (right ? width * squeeze : align === 'center' ? width * squeeze / 2 : 0);
    const ascent = metrics.fontBoundingBoxAscent ?? metrics.actualBoundingBoxAscent;
    const descent = metrics.fontBoundingBoxDescent ?? metrics.actualBoundingBoxDescent;
    const top = y - ascent, bottom = y + descent;
    const cache = new Map([[0, 0], [text.length, width]]);
    const advance = end => {
      if (cache.has(end)) return cache.get(end);
      ctx.save();
      try {
        for (const key in state) ctx[key] = state[key];
        const w = ctx.measureText(text.slice(0, end)).width;
        cache.set(end, w);
        return w;
      } finally { ctx.restore(); }
    };
    const rtl = L.pdf.hasRTL(text);
    const parts = rtl ? canvasBidi(ctx, text, metrics) : Array.from(text.matchAll(/\S+/gu), part =>
      ({ text: part[0], index: part.index, left: advance(part.index), right: advance(part.index + part[0].length) }));
    return parts.map(part => {
      const rect = { left: left + part.left * squeeze, right: left + part.right * squeeze, top, bottom };
      const glyphs = (lo, hi) => {
        if (part.glyphs) return part.glyphs.map(g => ({ text: g.text,
          rect: { left: left + g.left * squeeze, right: left + g.right * squeeze, top, bottom } }));
        const cuts = boundaries(part.text), at = i => left + advance(part.index + cuts[i]) * squeeze;
        // A long overflowing cell needs only the characters near its visible interval.
        let a = 0, b = cuts.length - 1;
        while (a < b) { const m = (a + b) >> 1; if (at(m + 1) <= lo + 1e-6) a = m + 1; else b = m; }
        const first = a;
        b = cuts.length - 1;
        while (a < b) { const m = Math.ceil((a + b) / 2); if (at(m) < hi - 1e-6) a = m; else b = m - 1; }
        const out = [];
        for (let i = first; i <= a && i + 1 < cuts.length; i++) out.push({ text: part.text.slice(cuts[i], cuts[i + 1]),
          rect: { left: at(i), right: at(i + 1), top, bottom } });
        return out;
      };
      const visual = rtl ? [...part.glyphs].sort((a, b) => a.left - b.left).map(g => g.text).join('') : undefined;
      return { text: part.text, rect, matrix, glyphs, visual };
    });
  }

  /** A proxy for a fresh canvas paint. Calls still run on the real context. Rectangle
   * clips are retained in device coordinates, including the transform at rect(), and
   * the current path is deliberately separate from the save/restore state stack.
   * https://html.spec.whatwg.org/dev/canvas.html#drawing-text-to-the-bitmap */
  L.pdf.recordCanvas = function (ctx, options = {}) {
    const canvas = ctx.canvas, entries = [], stack = [], groups = new WeakMap();
    let clips = [[region(canvasRect(0, 0, canvas.width, canvas.height))]], path = [], otherPath = false;
    canvasRecords.set(canvas, entries);
    const methods = new Map();
    const hooks = {
      save() { stack.push(clips); },
      restore() { if (stack.length) clips = stack.pop(); },
      beginPath() {
        path = [];
        otherPath = false;
      },
      rect(x, y, w, h) { path.push(region(canvasRect(+x, +y, +w, +h), ctx.getTransform())); },
      clip(...args) {
        // The print renderer uses rectangular paths. Do not expose text behind an
        // unmeasured Path2D/curved clip if another caller supplies one.
        const external = args.length && typeof args[0] !== 'string';
        clips = clips.concat([external || otherPath ? [] : path.slice()]);
      },
      reset() {
        entries.length = stack.length = 0;
        clips = [[region(canvasRect(0, 0, canvas.width, canvas.height))]];
        path = [];
        otherPath = false;
      },
      fillText(value, x, y, maxWidth = Infinity) {
        if (![+x, +y].every(Number.isFinite) || !(maxWidth > 0) || !ctx.globalAlpha) return;
        const text = String(value).replace(/[\t\n\r\f]/g, ' ');
        const words = canvasWords(ctx, text, +x, +y, +maxWidth);
        if (!words.length) return;
        for (const word of words) word.clips = clips;
        let group, order;
        if (options.cells && clips.length > 1) {
          group = groups.get(clips);
          if (!group) {
            group = {};
            groups.set(clips, group);
          }
          const last = clips.at(-1)[0];
          const visible = last && canvasClip(corners(last.rect).map(p => last.matrix.transformPoint(p)), clips);
          if (visible) order = bounds(visible.points);
        }
        order ||= bounds(words.flatMap(w => corners(w.rect).map(p => w.matrix.transformPoint(p))));
        entries.push({ words, group, order, leading: /^\s/u.test(text), trailing: /\s$/u.test(text) });
      },
      drawImage(image, ...args) {
        const source = canvasRecords.get(image);
        if (!source || !ctx.globalAlpha) return;
        let sx = 0, sy = 0, sw = image.width, sh = image.height, dx, dy, dw = sw, dh = sh;
        if (args.length === 2) [dx, dy] = args;
        else if (args.length === 4) [dx, dy, dw, dh] = args;
        else if (args.length === 8) [sx, sy, sw, sh, dx, dy, dw, dh] = args;
        else return;
        if (![sx, sy, sw, sh, dx, dy, dw, dh].every(Number.isFinite) || !sw || !sh || !dw || !dh) return;
        if (sw < 0) {
          sx += sw;
          sw = -sw;
        }
        if (sh < 0) {
          sy += sh;
          sh = -sh;
        }
        if (dw < 0) {
          dx += dw;
          dw = -dw;
        }
        if (dh < 0) {
          dy += dh;
          dh = -dh;
        }
        const matrix = ctx.getTransform().translate(dx, dy).scale(dw / sw, dh / sh).translate(-sx, -sy);
        const crop = [region(canvasRect(sx, sy, sw, sh), matrix)], clipCache = new Map(), groupCache = new Map();
        const mappedClips = old => {
          if (!clipCache.has(old)) clipCache.set(old, old.map(rs => rs.map(r =>
            region(r.rect, matrix.multiply(r.matrix)))).concat([crop], clips));
          return clipCache.get(old);
        };
        for (const entry of source.slice()) {
          const words = entry.words.map(w => ({ ...w, matrix: matrix.multiply(w.matrix),
            clips: mappedClips(w.clips) }));
          let group;
          if (entry.group) {
            if (!groupCache.has(entry.group)) groupCache.set(entry.group, {});
            group = groupCache.get(entry.group);
          }
          const visible = canvasClip(corners(entry.order).map(p => matrix.transformPoint(p)), [crop, ...clips]);
          if (visible) entries.push({ ...entry, words, group, order: bounds(visible.points) });
        }
      },
    };
    const curves = ['moveTo', 'lineTo', 'arc', 'arcTo', 'ellipse', 'bezierCurveTo', 'quadraticCurveTo', 'roundRect'];
    for (const key of curves) {
      hooks[key] = () => { otherPath = true; };
    }
    return new Proxy(ctx, {
      get(target, key) {
        const value = Reflect.get(target, key, target);
        if (typeof value !== 'function') return value;
        if (!methods.has(key)) methods.set(key, hooks[key] ? (...args) => {
          const result = value.apply(target, args);
          hooks[key](...args);
          return result;
        } : value.bind(target));
        return methods.get(key);
      },
      set(target, key, value) { return Reflect.set(target, key, value, target); },
    });
  };

  function visibleCanvasWord(word, scale) {
    const { matrix, clips } = word, inverse = matrix.inverse();
    const visible = rect => {
      const view = canvasClip(corners(rect).map(p => matrix.transformPoint(p)), clips);
      if (!view) return null;
      return { whole: view.whole, rect: bounds(view.points.map(p => inverse.transformPoint(p))) };
    };
    const view = visible(word.rect);
    if (!view || !hasArea(view.rect)) return [];
    const heightOnly = !matrix.b && !matrix.c && clips.every(rs => rs.length === 1 && rs[0].device) &&
      Math.abs(view.rect.left - word.rect.left) < 1e-6 && Math.abs(view.rect.right - word.rect.right) < 1e-6;
    if (view.whole || heightOnly) return [{ text: word.text, ...canvasBox(view.rect, matrix, scale),
      ...(word.visual ? { visual: word.visual } : {}) }];
    const out = [];
    let text = '', rect, glyphs = [];
    const flush = () => {
      if (text && rect) out.push({ text, ...canvasBox(rect, matrix, scale), ...(word.visual
        ? { visual: glyphs.sort((a, b) => a.rect.left - b.rect.left).map(g => g.text).join('') } : {}) });
      text = '';
      rect = null;
      glyphs = [];
    };
    for (const glyph of word.glyphs(view.rect.left, view.rect.right)) {
      const r = visible(glyph.rect)?.rect;
      if (r && hasArea(r)) {
        text += glyph.text;
        rect = rect ? union(rect, r) : r;
        glyphs.push(glyph);
      } else flush();
    }
    flush();
    return out;
  }

  /** Associate measured SVG words with their raster image, before drawImage places it.
   * Subword positions use the supplied word box, as in the PDF font's fixed advances. */
  L.pdf.recordImage = function (image, lines, width, height) {
    const scale = new DOMMatrix().scale(image.width / width, image.height / height);
    const clips = [[region(canvasRect(0, 0, image.width, image.height))]];
    const entries = lines.map(line => {
      const words = line.map(w => {
        const matrix = scale.translate(w.x, w.y).rotate(w.angle || 0);
        const rect = canvasRect(0, 0, w.w, w.h), cuts = boundaries(w.text);
        const chars = cuts.slice(1).map((end, i) => w.text.slice(cuts[i], end));
        const order = w.visual ? boundaries(w.visual).slice(1).map((end, i, ends) =>
          w.visual.slice(i ? ends[i - 1] : 0, end)) : chars;
        const glyphs = () => {
          const remaining = order.slice();
          return chars.map(text => {
            const i = remaining.indexOf(text);
            remaining[i] = null;
            return { text, rect: canvasRect(i * w.w / chars.length, 0, w.w / chars.length, w.h) };
          });
        };
        return { text: w.text, rect, matrix, clips, glyphs, visual: w.visual };
      });
      return { words, order: bounds(words.flatMap(w => corners(w.rect).map(p => w.matrix.transformPoint(p)))) };
    });
    canvasRecords.set(image, entries);
  };

  /** Recorded words in canvas-pixel coordinates, or points when scale is supplied.
   * rows opts into cell order (recordCanvas's cells option), so separate repeated-title
   * canvases interleave without changing painting. consume releases the measurement
   * contexts when an export has finished collecting a page. */
  L.pdf.canvasLines = function (canvas, scale = 1, options = {}) {
    const entries = [...(canvasRecords.get(canvas) || [])], lines = [];
    if (options.rows) entries.sort((a, b) => Math.round(a.order.top * 100) - Math.round(b.order.top * 100) ||
      Math.round(a.order.left * 100) - Math.round(b.order.left * 100));
    let previous;
    for (const entry of entries) {
      const words = entry.words.flatMap(word => visibleCanvasWord(word, scale));
      if (!words.length) continue;
      const first = words[0], last = previous?.line.at(-1);
      const rad = (first.angle || 0) * Math.PI / 180;
      const cos = Math.cos(rad), sin = Math.sin(rad);
      const dx = last ? first.x - last.x : 0, dy = last ? first.y - last.y : 0;
      const same = last && previous.group === entry.group && Math.abs((last.angle || 0) - (first.angle || 0)) < .001 &&
        Math.abs(dy * cos - dx * sin) < Math.min(first.h, last.h) * .2;
      if (same) {
        const gap = dx * cos + dy * sin - last.w;
        if (!entry.leading && !previous.trailing && Math.abs(gap) < .1 * scale && !first.angle && !last.angle) {
          last.text += first.text;
          const right = first.x + first.w, bottom = Math.max(last.y + last.h, first.y + first.h);
          last.y = Math.min(last.y, first.y);
          last.h = bottom - last.y;
          last.w = right - last.x;
          if (last.visual || first.visual) last.visual = (last.visual || last.text.slice(0, -first.text.length)) +
            (first.visual || first.text);
          words.shift();
        }
        previous.line.push(...words);
      } else {
        lines.push(words);
        previous = { line: words, group: entry.group };
      }
      previous.trailing = entry.trailing;
    }
    if (options.consume) canvasRecords.delete(canvas);
    return lines;
  };
})();
