/* VibeOffice — collect the visible words of a DOM page for the PDF text layer.
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

  /** scale is points per client pixel (for example 72 / 96 / the page's zoom).
   * Clipping stops at the page: scrolling the editor never removes words from a PDF.
   * No DOM mutation or font measurement canvas is needed. */
  L.pdf.collectDOM = function (page, scale = 0.75) {
    const origin = page.getBoundingClientRect();
    if (!area(origin) || !(scale > 0)) return [];
    const states = new WeakMap(), pieces = new WeakMap(), lines = [];
    const range = page.ownerDocument.createRange();
    let previous, spaced = true;

    function state(el) {
      let value = states.get(el);
      if (value) return value;
      const parent = el === page ? { hidden: false, clip: box(origin), block: page } : state(el.parentElement);
      const css = getComputedStyle(el);
      const hidden = parent.hidden || el.matches(SKIP) || css.display === 'none' ||
        css.contentVisibility === 'hidden' || (+css.opacity === 0 && !el.matches(WARP_BODY));
      let clip = parent.clip;
      if (!hidden && (CLIP.test(css.overflowX) || CLIP.test(css.overflowY))) {
        const r = el.getBoundingClientRect();
        clip = { ...clip };
        // overflow clips to the padding box. Client dimensions omit borders and scrollbars.
        const sx = el.offsetWidth ? r.width / el.offsetWidth : 1;
        const sy = el.offsetHeight ? r.height / el.offsetHeight : 1;
        if (CLIP.test(css.overflowX)) {
          const left = r.left + (el.clientLeft || 0) * sx;
          clip.left = Math.max(clip.left, left);
          clip.right = Math.min(clip.right, left + (el.clientWidth ?? r.width) * sx);
        }
        if (CLIP.test(css.overflowY)) {
          const top = r.top + (el.clientTop || 0) * sy;
          clip.top = Math.max(clip.top, top);
          clip.bottom = Math.min(clip.bottom, top + (el.clientHeight ?? r.height) * sy);
        }
      }
      const block = BLOCK.test(css.display) || el.localName === 'text' ? el : parent.block;
      const group = el.namespaceURI === MATH && /^(mi|mn|mo|mtext|ms)$/.test(el.localName) ? el : parent.group;
      value = { hidden, visible: css.visibility === 'visible', clip, block, group,
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
    const points = r => ({ x: (r.left - origin.left) * scale, y: (r.top - origin.top) * scale,
      w: (r.right - r.left) * scale, h: (r.bottom - r.top) * scale });

    function add(text, r, info, node, start, end) {
      if (!text || !area(r)) return;
      const continues = previous && previous.block === info.block && sameLine(previous.rect, r, info.vertical);
      if (!continues) lines.push([]);
      const line = lines[lines.length - 1];
      // A word may cross several styled spans. Only actual whitespace separates its pieces.
      if (continues && !spaced && previous.group === info.group) {
        const word = line[line.length - 1];
        word.text += text;
        previous.rect = union(previous.rect, r);
        Object.assign(word, points(previous.rect));
      } else {
        line.push({ text, ...points(r) });
        previous = { rect: r, block: info.block, group: info.group };
      }
      const word = line[line.length - 1];
      if (!pieces.has(word)) pieces.set(word, []);
      pieces.get(word).push({ node, start, end, vertical: info.vertical });
      spaced = false;
    }

    function clipped(node, start, end, rs, info) {
      const r = rs.reduce(union);
      if (inside(r, info.clip)) {
        add(node.data.slice(start, end), r, info, node, start, end);
        return;
      }
      if (!area(intersection(r, info.clip))) {
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
        const part = chars.length && intersection(chars.reduce(union), info.clip);
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
  };
})();
