/* VibeOffice — PDF 1.7 page pictures with an invisible, searchable text layer.
 * Original implementation from ISO 32000-1 (7.3–7.9, 9.6–9.10, 12.3, 12.5):
 * https://opensource.adobe.com/dc-acrobat-sdk-docs/standards/pdfstandards/pdf/PDF32000_2008.pdf
 * The empty TrueType font is generated here, not read from a font file. Table definitions:
 * https://learn.microsoft.com/en-us/typography/opentype/spec/otff
 * https://learn.microsoft.com/en-us/typography/opentype/spec/ (head, hhea, maxp, hmtx, loca,
 * glyf, cmap, name, os2, post). No painted shapes, hinting, or third-party font data.
 */
(function (root) {
  'use strict';
  const L = root.L;
  const enc = new TextEncoder();
  const FONT_NAME = 'VibeOfficeText';
  const ASCENT = 800, DESCENT = -200, ADVANCE = 500, UNITS = 1000;
  const CID_LIMIT = 65535;
  const hex = (n, width = 4) => n.toString(16).toUpperCase().padStart(width, '0');
  const bytesHex = (bytes) => Array.from(bytes, (b) => hex(b, 2)).join('');
  const num = (n) => String(Math.round(n * 1e6) / 1e6);
  const ref = (id) => id + ' 0 R';
  const scalar = (ch) => {
    const cp = ch.codePointAt(0);
    return cp >= 0xd800 && cp <= 0xdfff ? '\ufffd' : ch;
  };
  const utf16 = (text) => {
    let result = '';
    for (let n = 0; n < text.length; n++) result += hex(text.charCodeAt(n));
    return result;
  };
  const str = (text) => '<FEFF' + utf16(String(text ?? '')) + '>';
  const literal = (text) => '(' + text.replace(/[\\()\r\n]/g, (ch) => {
    if (ch === '\r') return '\\r';
    if (ch === '\n') return '\\n';
    return '\\' + ch;
  }) + ')';

  function emptyFont() {
    const tables = new Map();
    const table = (tag, length) => {
      const bytes = new Uint8Array(length);
      tables.set(tag, bytes);
      return new DataView(bytes.buffer);
    };
    const head = table('head', 54);
    head.setUint32(0, 0x00010000); // version
    head.setUint32(4, 0x00010000); // font revision
    head.setUint32(12, 0x5f0f3cf5); // magic number
    head.setUint16(16, 3); // baseline at zero; left side bearings equal xMin
    head.setUint16(18, UNITS);
    // Fixed epoch in seconds since 1904; independent of export time.
    head.setUint32(24, 2082844800);
    head.setUint32(32, 2082844800);
    head.setInt16(38, DESCENT);
    head.setInt16(40, ADVANCE);
    head.setInt16(42, ASCENT);
    head.setUint16(46, 8); // smallest recommended pixel size
    head.setInt16(48, 2); // deprecated direction hint
    // Short loca offsets and glyphDataFormat are zero.
    const hhea = table('hhea', 36);
    hhea.setUint32(0, 0x00010000);
    hhea.setInt16(4, ASCENT);
    hhea.setInt16(6, DESCENT);
    hhea.setUint16(10, ADVANCE);
    hhea.setInt16(16, ADVANCE); // maximum extent
    hhea.setInt16(18, 1); // caret slope rise, upright
    hhea.setUint16(34, 2); // full horizontal metrics for both glyphs
    const maxp = table('maxp', 32);
    maxp.setUint32(0, 0x00010000);
    maxp.setUint16(4, 2); // .notdef and the empty glyph
    maxp.setUint16(6, 2); // two points on a retraced line, with no area to fill
    maxp.setUint16(8, 1); // one contour
    maxp.setUint16(14, 1); // one zone; no hinting instructions
    const hmtx = table('hmtx', 8);
    hmtx.setUint16(0, ADVANCE);
    hmtx.setUint16(4, ADVANCE);
    // A zero-contour glyph has no measured bounds in FreeType. PDFium then drops one-character
    // words and vertical text objects because their width is zero, even in invisible mode.
    // Glyph 1 instead goes from (0, -200) to (500, 800) and straight back: it still paints nothing,
    // but supplies the intended selection bounds. Glyph 0 remains an empty .notdef.
    const glyf = table('glyf', 36);
    glyf.setInt16(12, 1); // number of contours
    glyf.setInt16(16, DESCENT);
    glyf.setInt16(18, ADVANCE);
    glyf.setInt16(20, ASCENT);
    glyf.setUint16(22, 1); // last point of the contour
    glyf.setUint8(26, 1); // first point on the curve; full signed coordinate deltas
    glyf.setUint8(27, 1); // second point on the curve
    glyf.setInt16(30, ADVANCE); // X deltas: 0, 500
    glyf.setInt16(32, DESCENT); // Y deltas: -200, 1000
    glyf.setInt16(34, UNITS);
    const loca = table('loca', 6);
    loca.setUint16(2, 6); // short offsets are measured in two-byte units
    loca.setUint16(4, 18);

    // Windows Unicode BMP cmap: space -> empty glyph, then the required FFFF sentinel.
    // PDF uses CIDToGIDMap for every other character, including supplementary scalars.
    const cmap = table('cmap', 44);
    cmap.setUint16(2, 1);
    cmap.setUint16(4, 3);
    cmap.setUint16(6, 1);
    cmap.setUint32(8, 12);
    cmap.setUint16(12, 4);
    cmap.setUint16(14, 32);
    cmap.setUint16(18, 4);
    cmap.setUint16(20, 4);
    cmap.setUint16(22, 1);
    cmap.setUint16(26, 32);
    cmap.setUint16(28, 0xffff);
    cmap.setUint16(32, 32);
    cmap.setUint16(34, 0xffff);
    cmap.setInt16(36, 1 - 32);
    cmap.setUint16(38, 1);

    const names = [[1, FONT_NAME], [2, 'Regular'], [3, FONT_NAME + '-1'],
      [4, FONT_NAME], [5, 'Version 1.0'], [6, FONT_NAME]];
    const storage = 6 + names.length * 12;
    const name = table('name', storage + names.reduce((sum, entry) => sum + entry[1].length * 2, 0));
    name.setUint16(2, names.length);
    name.setUint16(4, storage);
    let offset = 0;
    names.forEach(([id, text], index) => {
      const record = 6 + index * 12;
      name.setUint16(record, 3); // Windows, Unicode BMP, English names
      name.setUint16(record + 2, 1);
      name.setUint16(record + 4, 0x0409);
      name.setUint16(record + 6, id);
      name.setUint16(record + 8, text.length * 2);
      name.setUint16(record + 10, offset);
      for (let n = 0; n < text.length; n++) name.setUint16(storage + offset + n * 2, text.charCodeAt(n));
      offset += text.length * 2;
    });
    const os2 = table('OS/2', 96);
    os2.setUint16(0, 4);
    os2.setInt16(2, ADVANCE);
    os2.setUint16(4, 400); // normal weight
    os2.setUint16(6, 5); // normal width; unrestricted embedding (fsType = 0)
    os2.setUint32(42, 1); // Basic Latin
    os2.setUint32(58, 0x56494245); // VIBE
    os2.setUint16(62, 0x00c0); // regular and USE_TYPO_METRICS
    os2.setUint16(64, 32);
    os2.setUint16(66, 32);
    os2.setInt16(68, ASCENT);
    os2.setInt16(70, DESCENT);
    os2.setUint16(74, ASCENT);
    os2.setUint16(76, -DESCENT);
    os2.setInt16(86, ASCENT);
    os2.setUint16(92, 32); // break character
    os2.setUint16(94, 1); // maximum context
    const post = table('post', 32);
    post.setUint32(0, 0x00030000); // no glyph names
    post.setUint32(12, 1); // fixed pitch

    const checksum = (bytes) => {
      let sum = 0;
      for (let n = 0; n < bytes.length; n += 4) {
        sum = (sum + ((bytes[n] << 24) | (bytes[n + 1] << 16) |
          (bytes[n + 2] << 8) | bytes[n + 3])) >>> 0;
      }
      return sum;
    };
    const entries = [...tables].sort(([a], [b]) => a < b ? -1 : 1);
    const aligned = (length) => (length + 3) & ~3;
    const font = new Uint8Array(12 + entries.length * 16 +
      entries.reduce((sum, [, data]) => sum + aligned(data.length), 0));
    const view = new DataView(font.buffer);
    view.setUint32(0, 0x00010000);
    view.setUint16(4, entries.length);
    const power = 2 ** Math.floor(Math.log2(entries.length));
    view.setUint16(6, power * 16);
    view.setUint16(8, Math.log2(power));
    view.setUint16(10, entries.length * 16 - power * 16);
    offset = 12 + entries.length * 16;
    let headOffset;
    entries.forEach(([tag, data], index) => {
      const record = 12 + index * 16;
      font.set(enc.encode(tag), record);
      view.setUint32(record + 4, checksum(data));
      view.setUint32(record + 8, offset);
      view.setUint32(record + 12, data.length);
      font.set(data, offset);
      if (tag === 'head') headOffset = offset;
      offset += aligned(data.length);
    });
    view.setUint32(headOffset + 8, (0xb1b0afba - checksum(font)) >>> 0);
    return font;
  }

  function unicodeMap(chars) {
    const ranges = [];
    for (let first = 0; first < chars.length;) {
      let last = first;
      const cp = chars[first].codePointAt(0);
      // bfrange increments the final UTF-16 byte: stop before its overflow, or a surrogate boundary.
      const low = chars[first].charCodeAt(chars[first].length - 1) & 255;
      while (last + 1 < chars.length && last - first < 255 - low &&
        chars[last + 1].length === chars[first].length &&
        chars[last + 1].codePointAt(0) === cp + last + 1 - first) last++;
      ranges.push(`<${hex(first + 1)}> <${hex(last + 1)}> <${utf16(chars[first])}>`);
      first = last + 1;
    }
    const lines = ['/CIDInit /ProcSet findresource begin', '12 dict begin', 'begincmap',
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
      '/CMapName /VibeOfficeUnicode def', '/CMapType 2 def',
      '1 begincodespacerange', '<0000> <FFFF>', 'endcodespacerange'];
    for (let n = 0; n < ranges.length; n += 100) {
      const batch = ranges.slice(n, n + 100);
      lines.push(batch.length + ' beginbfrange', ...batch, 'endbfrange');
    }
    lines.push('endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end');
    return enc.encode(lines.join('\n'));
  }

  const validWord = (word) => word && word.text && word.w > 0 && word.h > 0 &&
    [word.x, word.y, word.w, word.h, word.angle ?? 0].every(Number.isFinite);

  // Readers reverse RTL character runs when extracting PDF text. Supply those runs in visual
  // order, keeping numbers and Latin runs in the word's supplied order. Character positions within
  // a word are approximate (fixed advances); its complete selection box remains exact.
  const rtl = /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff\u{10840}-\u{10fff}\u{1e800}-\u{1eeff}]/u;
  function visualWord(text) {
    const letters = Array.from(String(text), scalar);
    for (let n = 0; n < letters.length; n++) {
      if (!rtl.test(letters[n]) || !/\p{Letter}/u.test(letters[n])) continue;
      const begin = n;
      while (n + 1 < letters.length && (/\p{Mark}/u.test(letters[n + 1]) ||
        (rtl.test(letters[n + 1]) && /\p{Letter}/u.test(letters[n + 1])))) n++;
      for (let a = begin, b = n; a < b; a++, b--) [letters[a], letters[b]] = [letters[b], letters[a]];
    }
    return letters;
  }

  function textLayer(page, chars) {
    const out = [];
    const word = (text, x, y, w, h, angle = 0, visual) => {
      const letters = visual ? Array.from(visual, scalar) : visualWord(text);
      if (!letters.length) return;
      const rad = angle * Math.PI / 180, cos = Math.cos(rad), sin = Math.sin(rad);
      const baseline = h * ASCENT / UNITS;
      out.push('BT 3 Tr', num(w * UNITS * 100 / (letters.length * ADVANCE * h)) + ' Tz',
        `${num(cos)} ${num(-sin)} ${num(sin)} ${num(cos)} ` +
        `${num(x - baseline * sin)} ${num(page.ph - y - baseline * cos)} Tm`);
      let font = -1, cids = '';
      for (const ch of letters) {
        const index = chars.get(ch), next = Math.floor(index / CID_LIMIT);
        if (next !== font) {
          if (cids) out.push('<' + cids + '> Tj');
          font = next;
          cids = '';
          out.push('/F' + (font + 1) + ' ' + num(h) + ' Tf');
        }
        cids += hex(index % CID_LIMIT + 1);
      }
      out.push('<' + cids + '> Tj', 'ET');
    };
    for (const line of page.lines || []) {
      let previous;
      for (const item of line) {
        if (!validWord(item)) continue;
        if (previous && !/\s$/.test(previous.text) && !/^\s/.test(item.text)) {
          const rad = (previous.angle || 0) * Math.PI / 180, cos = Math.cos(rad), sin = Math.sin(rad);
          const distance = (item.x - previous.x) * cos + (item.y - previous.y) * sin;
          const left = distance < 0 ? item : previous;
          const x = left.x + left.w * cos, y = left.y + left.w * sin;
          const gap = Math.max(0.01, Math.abs(distance) - left.w);
          word(' ', x, y, gap, previous.h, previous.angle);
        }
        word(item.text, item.x, item.y, item.w, item.h, item.angle, item.visual);
        previous = item;
      }
    }
    return out.join('\n');
  }

  /** Page and outline destinations use zero-based page indexes. Word angles are clockwise degrees
   * about their top-left corner; all rectangles and destinations are in points from the page's top-left.
   * A DOM collector may supply word.visual, the browser's visual character order for mixed bidi text.
   * info.date can fix the creation date for reproducible exports/tests. */
  L.pdf = L.pdf || {};
  L.pdf.hasRTL = text => rtl.test(text);
  /** Chromium may defer toBlob's encoder until an idle period, adding hundreds of milliseconds
   * per page after a preview/thumbnail. The same JPEG encoder runs promptly through toDataURL.
   * Decode one page at a time; the temporary base64 string never goes into the PDF or a cache. */
  L.pdf.jpeg = function (canvas) {
    const url = canvas.toDataURL('image/jpeg', 0.92);
    if (!url.startsWith('data:image/jpeg;base64,')) {
      throw new Error(L.t('This browser blocked the page image.'));
    }
    const binary = atob(url.slice(url.indexOf(',') + 1));
    const bytes = new Uint8Array(binary.length);
    for (let n = 0; n < bytes.length; n++) bytes[n] = binary.charCodeAt(n);
    return bytes;
  };

  /** Rasterize a few pages in one SVG so its font faces are decoded once. Each yielded canvas
   * belongs to this iteration only; release its backing store as soon as the caller encodes it.
   * Even the full SVG's scaled dimensions stay within the same cap as an individual page. */
  L.pdf.pictures = async function* (pages, prepare, fonts) {
    let next = 0, pending;
    while (next < pages.length || pending) {
      const batch = [];
      let width = 0, height = 0, scale;
      while (next < pages.length || pending) {
        const item = pending || { ...(await prepare(pages[next], next)), page: pages[next], index: next++ };
        // Nested SVGs share an ID namespace. Keep pages with local paint/filter/clip references
        // or anchors separate, so an identically named definition on another page cannot win.
        item.isolated = /\sid\s*=/i.test(item.svg) || /url\(\s*['"]?#|href\s*=\s*['"]#/i.test(item.svg);
        const w = Math.max(width, item.w), h = height + item.h;
        if (batch.length && (batch.length === 4 || !fonts || item.isolated || batch[0].isolated ||
          item.scale !== scale ||
          L.pdfScale(w, h, scale) !== scale)) {
          pending = item;
          break;
        }
        pending = null;
        item.y = height;
        batch.push(item);
        width = w;
        height = h;
        scale = item.scale;
      }
      const css = '<style>' + fonts.replace(/</g, '\\3c ') + '</style>';
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">` + css +
        batch.map((item) => item.svg.replace('<svg ', `<svg y="${item.y}" `)).join('') + '</svg>';
      const image = await L.loadImage('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg));
      for (const item of batch) {
        const canvas = document.createElement('canvas');
        canvas.width = Math.round(item.w * scale);
        canvas.height = Math.round(item.h * scale);
        try {
          const g = canvas.getContext('2d');
          g.fillStyle = '#fff';
          g.fillRect(0, 0, canvas.width, canvas.height);
          g.scale(scale, scale);
          g.drawImage(image, 0, item.y, item.w, item.h, 0, 0, item.w, item.h);
          yield { canvas, page: item.page, index: item.index };
        } finally {
          canvas.width = canvas.height = 1;
        }
      }
    }
  };

  L.pdf.write = async function (pages, info = {}) {
    const objects = [null];
    const reserve = () => objects.push(null) - 1;
    const object = (body) => {
      const id = reserve();
      objects[id] = { body };
      return id;
    };
    const stream = async (body, bytes, compress = true) => {
      if (compress && typeof CompressionStream !== 'undefined') {
        bytes = new Uint8Array(await new Response(new Blob([bytes]).stream()
          .pipeThrough(new CompressionStream('deflate'))).arrayBuffer());
        body += ' /Filter /FlateDecode';
      }
      const id = reserve();
      objects[id] = { body: '<< ' + body + ' /Length ' + bytes.length + ' >>', bytes };
      return id;
    };
    const catalog = reserve(), tree = reserve(), pageIds = pages.map(reserve);
    const chars = new Map();
    for (const page of pages) for (const line of page.lines || []) for (const word of line) {
      if (!validWord(word)) continue;
      if (!chars.size) chars.set(' ', 0);
      for (const ch of String(word.visual || word.text)) {
        const s = scalar(ch);
        if (!chars.has(s)) chars.set(s, chars.size);
      }
    }
    const fonts = [];
    if (chars.size) {
      const font = emptyFont();
      const file = await stream('/Length1 ' + font.length, font);
      const descriptor = object(`<< /Type /FontDescriptor /FontName /${FONT_NAME} /Flags 33 ` +
        `/FontBBox [0 ${DESCENT} ${ADVANCE} ${ASCENT}] /ItalicAngle 0 /Ascent ${ASCENT} ` +
        `/Descent ${DESCENT} /CapHeight ${ASCENT} /StemV 0 /FontFile2 ${ref(file)} >>`);
      const letters = [...chars.keys()];
      for (let n = 0; n < letters.length; n += CID_LIMIT) {
        const group = letters.slice(n, n + CID_LIMIT);
        const gids = new Uint8Array((group.length + 1) * 2);
        for (let cid = 1; cid <= group.length; cid++) gids[cid * 2 + 1] = 1;
        const mapping = await stream('', gids);
        const unicode = await stream('', unicodeMap(group));
        const descendant = object(`<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${FONT_NAME} ` +
          '/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> ' +
          `/FontDescriptor ${ref(descriptor)} /DW ${ADVANCE} /CIDToGIDMap ${ref(mapping)} >>`);
        fonts.push(object(`<< /Type /Font /Subtype /Type0 /BaseFont /${FONT_NAME} /Encoding /Identity-H ` +
          `/DescendantFonts [${ref(descendant)}] /ToUnicode ${ref(unicode)} >>`));
      }
    }
    const destination = (item) => {
      if (!Number.isInteger(item.page) || !pageIds[item.page]) return null;
      const top = Number.isFinite(item.top) ? item.top : 0;
      return `[${ref(pageIds[item.page])} /XYZ 0 ${num(pages[item.page].ph - top)} null]`;
    };
    for (let n = 0; n < pages.length; n++) {
      const page = pages[n];
      const image = await stream(`/Type /XObject /Subtype /Image /Width ${page.iw} /Height ${page.ih} ` +
        '/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode', page.jpeg, false);
      let content = `q ${num(page.pw)} 0 0 ${num(page.ph)} 0 0 cm /Im0 Do Q\n`;
      if (chars.size) content += textLayer(page, chars);
      const contents = await stream('', enc.encode(content));
      const links = [];
      for (const link of page.links || []) {
        if (![link.x, link.y, link.w, link.h].every(Number.isFinite) || link.w <= 0 || link.h <= 0) continue;
        let action;
        if (link.url) {
          const uri = encodeURI(String(link.url)).replace(/%25([\da-f]{2})/gi, '%$1');
          action = '/A << /S /URI /URI ' + literal(uri) + ' >>';
        }
        else {
          const dest = destination(link);
          if (!dest) continue;
          action = '/Dest ' + dest;
        }
        const rect = [link.x, page.ph - link.y - link.h, link.x + link.w, page.ph - link.y].map(num).join(' ');
        links.push(object(`<< /Type /Annot /Subtype /Link /Rect [${rect}] /Border [0 0 0] ${action} >>`));
      }
      const resources = `/XObject << /Im0 ${ref(image)} >>` + (fonts.length ? ' /Font << ' +
        fonts.map((id, i) => '/F' + (i + 1) + ' ' + ref(id)).join(' ') + ' >>' : '');
      objects[pageIds[n]] = { body: `<< /Type /Page /Parent ${ref(tree)} ` +
        `/MediaBox [0 0 ${num(page.pw)} ${num(page.ph)}] /Resources << ${resources} >> ` +
        `/Contents ${ref(contents)}` + (links.length ? ' /Annots [' + links.map(ref).join(' ') + ']' : '') + ' >>' };
    }
    let outlines = '';
    if (info.outline?.length) {
      const top = { id: reserve(), children: [], level: -Infinity }, stack = [top];
      for (const item of info.outline) {
        const dest = destination(item);
        if (!dest) continue;
        const level = Number.isFinite(item.level) ? item.level : 0;
        while (stack.length > 1 && stack[stack.length - 1].level >= level) stack.pop();
        const parent = stack[stack.length - 1];
        const node = { id: reserve(), children: [], parent, level, dest, title: item.title };
        parent.children.push(node);
        stack.push(node);
      }
      const write = (node) => {
        let count = 0;
        for (const child of node.children) count += 1 + write(child);
        let body;
        if (node === top) body = '/Type /Outlines';
        else {
          const siblings = node.parent.children, index = siblings.indexOf(node);
          body = `/Title ${str(node.title)} /Parent ${ref(node.parent.id)} /Dest ${node.dest}`;
          if (index) body += ' /Prev ' + ref(siblings[index - 1].id);
          if (index + 1 < siblings.length) body += ' /Next ' + ref(siblings[index + 1].id);
        }
        if (count) body += ' /First ' + ref(node.children[0].id) +
          ' /Last ' + ref(node.children[node.children.length - 1].id) + ' /Count ' + count;
        objects[node.id] = { body: '<< ' + body + ' >>' };
        return count;
      };
      if (write(top)) outlines = ' /Outlines ' + ref(top.id);
    }
    objects[catalog] = { body: `<< /Type /Catalog /Pages ${ref(tree)}${outlines}` +
      (info.lang ? ' /Lang ' + str(info.lang) : '') +
      (info.title ? ' /ViewerPreferences << /DisplayDocTitle true >>' : '') + ' >>' };
    objects[tree] = {
      body: '<< /Type /Pages /Kids [' + pageIds.map(ref).join(' ') + '] /Count ' + pages.length + ' >>',
    };
    const date = new Date(info.date ?? Date.now()).toISOString().replace(/[-:T]/g, '').slice(0, 14) + 'Z';
    const metadata = object(`<< /Title ${str(info.title)} /Author ${str(info.author)} ` +
      `/Producer ${str(L.APP)} /CreationDate (D:${date}) >>`);
    const chunks = [], offsets = [0], hashes = [];
    let length = 0;
    const push = (data, hash = true) => {
      const bytes = typeof data === 'string' ? enc.encode(data) : data;
      chunks.push(bytes);
      length += bytes.length;
      if (hash) hashes.push(L.sha.sha256(bytes));
    };
    push('%PDF-1.7\n');
    push(new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));
    for (let id = 1; id < objects.length; id++) {
      const obj = objects[id];
      offsets[id] = length;
      push(id + ' 0 obj\n' + obj.body);
      if (obj.bytes) {
        push('\nstream\n');
        push(obj.bytes);
        push('\nendstream');
      }
      push('\nendobj\n');
    }
    // Hash the ordered chunk digests: no second full-size copy of a large document just for its ID.
    const digestBytes = new Uint8Array(hashes.length * 32);
    hashes.forEach((hash, index) => digestBytes.set(hash, index * 32));
    const id = bytesHex(L.sha.sha256(digestBytes));
    const xref = length;
    push('xref\n0 ' + objects.length + '\n0000000000 65535 f \n', false);
    for (let n = 1; n < objects.length; n++) push(String(offsets[n]).padStart(10, '0') + ' 00000 n \n', false);
    push(`trailer\n<< /Size ${objects.length} /Root ${ref(catalog)} /Info ${ref(metadata)} ` +
      `/ID [<${id}> <${id}>] >>\nstartxref\n${xref}\n%%EOF\n`, false);
    return new Blob(chunks, { type: 'application/pdf' });
  };
})(typeof window !== 'undefined' ? window : globalThis);
