/* Derived from Apache POI hssf/record/{NoteRecord,TextObjectRecord,EscherAggregate}.java
 * (Apache-2.0). Derived from Apache OpenOffice main/sc/source/filter/excel/xiescher.cxx
 * (Apache-2.0): drawing/object/text association and BIFF5 NOTE continuations.
 * Comment parts/VML follow Office-authored comments.xlsx in the pinned POI corpus. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    X = B.xls,
    A = B.officeart,
    el = B.el,
    N = s => s,
    bad = s => B.fail('XLS_COMMENT', s);
  X.txo = function (records, index) {
    const r = records[index].r,
      n = r.u16(10),
      formatBytes = r.u16(12);
    let next = index + 1,
      pos = 0,
      chunk = null,
      value = '',
      remaining = n;
    const advance = () => {
      if (records[next]?.id !== 0x3c) bad('Missing comment continuation');
      chunk = records[next++].r;
      pos = 0;
    };
    while (remaining) {
      advance();
      const wide = chunk.u8(pos++) & 1,
        width = wide ? 2 : 1,
        take = Math.min(remaining, Math.floor((chunk.length - pos) / width));
      if (!take) bad('Empty comment text continuation');
      value += wide ? chunk.text(pos, take * 2, 1200) : B.latin16(chunk.sub(pos, take));
      pos += take * width;
      remaining -= take;
    }
    const runs = [];
    let formattingError = false;
    try {
      if (formatBytes % 8) bad('Invalid comment formatting runs');
      // Formatting starts in a new CONTINUE, including when the last text
      // chunk was short. Font-run bytes themselves have no encoding marker.
      let left = formatBytes;
      const bytes = new Uint8Array(left);
      let p = 0;
      while (left) {
        advance();
        const take = Math.min(left, chunk.length);
        if (!take) bad('Empty comment formatting continuation');
        bytes.set(chunk.sub(0, take), p);
        p += take;
        left -= take;
      }
      const f = new B.Reader(bytes);
      for (let p = 0; p < bytes.length; p += 8) {
        const at = f.u16(p);
        if (at <= n) runs.push({ at, font: f.u16(p + 2) });
      }
    } catch (e) {
      if (e.name !== 'LegacyError') throw e;
      formattingError = true;
    }
    return { value, runs, formattingError, next };
  };
  const vml = function (note, geo, styles, order, id) {
    let a = note.shape?.anchor;
    if (
      !a ||
      a.kind !== 'cells' ||
      a.col1 > 255 ||
      a.col2 > 255 ||
      a.row1 > 65535 ||
      a.row2 > 65535 ||
      a.col1 > a.col2 ||
      a.row1 > a.row2
    )
      a = {
        flags: 0,
        col1: Math.min(253, note.col + 1),
        col2: Math.min(255, note.col + 3),
        row1: Math.max(0, note.row - 1),
        row2: Math.min(65535, Math.max(0, note.row - 1) + 4),
        dx1: 240,
        dx2: 496,
        dy1: 30,
        dy2: 196
      };
    const { x, y, widths, heights } = geo,
      dx1 = Math.round((a.dx1 / 1024) * widths[a.col1]),
      dx2 = Math.round((a.dx2 / 1024) * widths[a.col2]),
      dy1 = Math.round((a.dy1 / 256) * heights[a.row1]),
      dy2 = Math.round((a.dy2 / 256) * heights[a.row2]);
    const left = x[a.col1] + dx1,
      top = y[a.row1] + dy1,
      width = Math.max(1, x[a.col2] + dx2 - left),
      height = Math.max(1, y[a.row2] + dy2 - top);
    const style =
      'position:absolute;margin-left:' +
      left * 0.75 +
      'pt;margin-top:' +
      top * 0.75 +
      'pt;width:' +
      width * 0.75 +
      'pt;height:' +
      height * 0.75 +
      'pt;z-index:' +
      order +
      ';visibility:' +
      (note.visible ? 'visible' : 'hidden');
    let fill = 'infoBackground [80]',
      line = '#000000';
    const shape = note.shape,
      palette = { ...styles.palette, 64: '000000', 65: 'FFFFFF', 80: 'FFFFE1', 81: '000000' };
    if (shape)
      for (const [prop, set] of [
        [385, c => (fill = c)],
        [448, c => (line = c)]
      ])
        if (shape.properties.has(prop)) {
          try {
            set(
              '#' +
                B.shapes.color(
                  shape.properties.get(prop).value,
                  {
                    palette,
                    scheme: palette,
                    system: { 5: 'FFFFFF', 8: '000000', 24: 'FFFFE1', 23: '000000' }
                  },
                  shape
                )
            );
          } catch (e) {
            if (e.name !== 'LegacyError') throw e;
          }
        }
    const anchor = [a.col1, dx1, a.row1, dy1, a.col2, dx2, a.row2, dy2].join(', '),
      text = el(
        'v:textbox',
        { style: 'mso-direction-alt:auto' },
        el('div', { style: 'text-align:left' }, '')
      );
    const client = el(
      'x:ClientData',
      { ObjectType: 'Note' },
      (a.flags & 1 ? '' : el('x:MoveWithCells')) +
        (a.flags & 2 ? '' : el('x:SizeWithCells')) +
        el('x:Anchor', {}, anchor) +
        el('x:AutoFill', {}, 'False') +
        el('x:Row', {}, note.row) +
        el('x:Column', {}, note.col) +
        (note.visible ? el('x:Visible') : '')
    );
    return el(
      'v:shape',
      {
        id: '_x0000_s' + id,
        type: '#_x0000_t202',
        style,
        fillcolor: fill,
        strokecolor: line,
        'o:insetmode': 'auto'
      },
      el('v:fill', { color2: fill }) +
        el('v:shadow', { on: 't', color: 'black', obscured: 't' }) +
        el('v:path', { 'o:connecttype': 'none' }) +
        text +
        client
    );
  };
  X.sheetComments = function (book, si, meta, styles, pkg, part, losses) {
    const records = meta.records,
      sheet = book.sheets[si],
      notes = records.filter(r => r.id === 0x1c);
    if (!notes.length) return { xml: '', stats: { commentSource: 0, comments: 0 } };
    const issue = (what, e, where = sheet.name, notify = true) =>
      X.featureLoss(losses, 'legacy-comments', what, where, e, notify);
    const drawing = X.readDrawingRecords(meta),
      objects = drawing.objects;
    for (const error of drawing.errors)
      issue('Some comment box positions could not be read', error, sheet.name, false);
    const comments = new Map();
    let source = 0;
    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      if (rec.id !== 0x1c) continue;
      const r = rec.r;
      let ref = sheet.name;
      source++;
      try {
        const row = r.u16(0),
          col = r.u16(2);
        if (col > 255) bad('Invalid comment cell');
        ref = sheet.name + '!' + X.address(row, col);
        let item;
        if (book.version === 'BIFF8') {
          const obj = objects.get(r.u16(6));
          if (!obj?.text) bad('The comment text is missing');
          item = {
            row,
            col,
            author: X.string(r, 8).value,
            visible: !!(r.u16(4) & 2),
            ...obj.text,
            shape: obj.shape
          };
          if (obj.text.formattingError)
            issue('Some comment text formatting could not be read', null, ref, false);
        } else {
          let n = r.u16(4),
            take = Math.min(n, r.length - 6),
            value = r.text(6, take, book.codepage);
          n -= take;
          while (n && records[i + 1]?.id === 0x1c) {
            const c = records[i + 1].r;
            if (c.u16(0) !== 65535) break;
            i++;
            const len = Math.min(n, c.u16(4));
            value += c.text(6, len, book.codepage);
            n -= len;
          }
          if (n) bad('The comment text is incomplete');
          item = { row, col, author: '', visible: false, value, runs: [] };
        }
        if (comments.has(row * 256 + col))
          issue(N('A duplicate comment on this cell was not kept'), null, ref);
        comments.set(row * 256 + col, item);
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        issue(N('A cell comment was not kept'), e, ref);
      }
    }
    if (!comments.size) return { xml: '', stats: { commentSource: source, comments: 0 } };
    const list = [...comments.values()].sort((a, b) => a.row - b.row || a.col - b.col),
      authors = [...new Set(list.map(c => c.author))],
      authorIds = new Map(authors.map((v, i) => [v, i]));
    const xml =
      B.xmlDeclaration +
      el(
        'comments',
        { xmlns: B.ns.s },
        el('authors', {}, authors.map(a => el('author', {}, B.xstring(a))).join('')) +
          el(
            'commentList',
            {},
            list
              .map(c =>
                el(
                  'comment',
                  { ref: X.address(c.row, c.col), authorId: authorIds.get(c.author) },
                  el('text', {}, X.richText(c, styles, losses, sheet.name + '!' + X.address(c.row, c.col)))
                )
              )
              .join('')
          )
      );
    const number = si + 1;
    pkg.add(
      'xl/comments' + number + '.xml',
      xml,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
    );
    pkg.rel(part, 'comments', '../comments' + number + '.xml');
    const geo = X.drawingGrid(meta),
      ids = new Set(),
      shapeIds = list.map((c, i) => {
        let id = c.shape?.id;
        if (!Number.isInteger(id) || id < 1025 || ids.has(id)) id = 1025 + i;
        while (ids.has(id)) id += 1024;
        ids.add(id);
        return id;
      });
    const layout = el(
      'o:shapelayout',
      { 'v:ext': 'edit' },
      el('o:idmap', {
        'v:ext': 'edit',
        data: [...new Set(shapeIds.map(id => Math.floor(id / 1024)))].join(',')
      })
    );
    const type = el(
      'v:shapetype',
      { id: '_x0000_t202', coordsize: '21600,21600', 'o:spt': 202, path: 'm,l,21600r21600,l21600,xe' },
      el('v:stroke', { joinstyle: 'miter' }) + el('v:path', { gradientshapeok: 't', 'o:connecttype': 'rect' })
    );
    pkg.add(
      'xl/drawings/vmlDrawing' + number + '.vml',
      el(
        'xml',
        {
          'xmlns:v': 'urn:schemas-microsoft-com:vml',
          'xmlns:o': 'urn:schemas-microsoft-com:office:office',
          'xmlns:x': 'urn:schemas-microsoft-com:office:excel'
        },
        layout + type + list.map((c, i) => vml(c, geo, styles, i + 1, shapeIds[i])).join('')
      ),
      'application/vnd.openxmlformats-officedocument.vmlDrawing'
    );
    return {
      xml: el('legacyDrawing', {
        'r:id': pkg.rel(part, 'vmlDrawing', '../drawings/vmlDrawing' + number + '.vml')
      }),
      stats: { commentSource: source, comments: list.length }
    };
  };
})(typeof window !== 'undefined' ? window : globalThis);
