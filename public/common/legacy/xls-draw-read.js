/* Binary Excel drawing/object association and cell anchors.
 * Derived from Apache OpenOffice main/sc/source/filter/excel/xiescher.cxx (Apache-2.0).
 * Derived from Apache POI hssf/record/EscherAggregate.java (Apache-2.0).
 * MS-XLS MSODRAWING/OBJ/TXO and MS-ODRAW ClientData/ClientTextbox boundaries. */
(function (root) {
  'use strict';
  const B = root.L.legacy, X = B.xls, A = B.officeart;
  const RECORD = Object.freeze({
    CONTINUE: 0x3c, OBJ: 0x5d, DRAWING_GROUP: 0xeb, DRAWING: 0xec, TXO: 0x1b6,
    DRAWING_ROOT: 0xf000, DRAWING_CONTAINER: 0xf002, GROUP: 0xf003, SHAPE: 0xf004, SOLVER: 0xf005,
    OPT: 0xf00b, TEXTBOX: 0xf00d, CLIENT: 0xf011,
    CONNECTOR: 0xf012, BLIP_STORE: 0xf001
  });
  const containers = new Set([RECORD.DRAWING_ROOT, RECORD.BLIP_STORE, RECORD.DRAWING_CONTAINER,
    RECORD.GROUP, RECORD.SHAPE, RECORD.SOLVER]);
  const OBJECT_PART = Object.freeze({ FLAGS: 8, PICTURE_LINK: 9, LIST: 19 });
  const N = s => s;
  const fail = message => B.fail('XLS_DRAWING', message);

  // Each bounded record is independent: a corrupt optional shape must not
  // prevent a later picture or comment from being read.
  function tree(bytes, errors, depth = 0, budget = { left: B.limits.records }) {
    if (depth >= B.limits.depth) fail('Drawing nesting exceeds the limit');
    const result = [];
    for (let at = 0; at < bytes.length;) {
      if (--budget.left < 0) fail('Drawing record count exceeds the limit');
      let record;
      try {
        record = A.record(bytes, at);
        // AOO dispatches these containers by type; some producers write version 0.
        if (record.version === 15 || containers.has(record.type)) {
          record.children = tree(record.r.bytes, errors, depth + 1, budget);
        }
        result.push(record);
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        errors.push(error);
        if (!record) break;
      }
      at = record.end;
    }
    return result;
  }

  function join(chunks, length) {
    const bytes = new Uint8Array(length);
    let at = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, at);
      at += chunk.length;
    }
    return bytes;
  }

  X.drawingGroup = function (book) {
    if (book.drawingGroup) return book.drawingGroup;
    const chunks = [], errors = [];
    let length = 0, continuing = false;
    for (const record of book.globals) {
      if (record.id === RECORD.DRAWING_GROUP || continuing && record.id === RECORD.CONTINUE) {
        length += record.r.length;
        if (length > B.limits.bytes) fail('The workbook drawing data is too large');
        chunks.push(record.r.bytes);
        continuing = true;
      } else continuing = false;
    }
    const records = tree(join(chunks, length), errors), defaults = new Map();
    for (const record of records) for (const child of record.children || []) {
      if (child.type !== RECORD.OPT) continue;
      try { A.mergeProperties(defaults, A.properties(child)); }
      catch (error) {
        if (error.name !== 'LegacyError') throw error;
        errors.push(error);
      }
    }
    const images = [];
    const stores = list => {
      for (const record of list) {
        if (record.type !== RECORD.BLIP_STORE) {
          stores(record.children || []);
          continue;
        }
        for (const entry of record.children || []) {
          try { images.push(...A.store([{ ...record, children: [entry] }])); }
          catch (error) {
            if (error.name !== 'LegacyError') throw error;
            images.push({ read: () => { throw error; } });
          }
        }
      }
    };
    stores(records);
    return book.drawingGroup = { images, defaults, errors };
  };

  function objectParts(object) {
    const r = object.record.r;
    for (let p = 0; p + 4 <= r.length;) {
      const id = r.u16(p), declared = r.u16(p + 2);
      p += 4;
      if (!id) break;
      // ftLbsData's length is a continuation hint, not the subrecord's size.
      const size = id === OBJECT_PART.LIST ? r.length - p : declared;
      const part = new B.Reader(r.sub(p, size));
      if (id === OBJECT_PART.FLAGS) object.control = !!(part.u16(0) & 16);
      else if (id === OBJECT_PART.PICTURE_LINK && part.u16(0) >= 7 && part.u16(2)) {
        const token = part.u8(8);
        object.embedded = token === 2;
        object.linked = token === 0x39;
        object.pictureLink = !object.embedded && !object.linked;
      } else if (id === OBJECT_PART.LIST && object.type === 20) {
        const formulaSize = part.u16(0);
        object.autoFilter = (part.u16(2 + formulaSize + 8) & 3) === 2;
      }
      p += size;
    }
  }

  X.readDrawingRecords = function (meta) {
    if (meta.drawings) return meta.drawings;
    const records = meta.records, objects = new Map(), markers = [], chunks = [], errors = [];
    let length = 0, inDrawing = false, lastObject = null;
    for (let i = 0; i < records.length; i++) {
      const record = records[i], r = record.r;
      if (record.id === RECORD.DRAWING || record.id === RECORD.CONTINUE && inDrawing) {
        length += r.length;
        if (length > B.limits.bytes) fail('The sheet drawing data is too large');
        chunks.push(r.bytes);
        inDrawing = true;
      } else if (record.id === RECORD.OBJ) {
        lastObject = null;
        try {
          // BIFF8 producers may retain a BIFF5 OBJ without an Escher drawing.
          const modern = r.u16(0) === 0x15 && r.u16(2) >= 6;
          const object = { type: r.u16(4), id: r.u16(6), record, modern };
          if (!modern) {
            object.anchor = { kind: 'cells', flags: 0, col1: r.u16(10), dx1: r.u16(12),
              row1: r.u16(14), dy1: r.u16(16), col2: r.u16(18), dx2: r.u16(20),
              row2: r.u16(22), dy2: r.u16(24) };
          }
          objects.set(object.id, object);
          lastObject = object;
          if (modern) markers.push({ kind: 'object', object, end: length });
          if (modern) {
            try { objectParts(object); }
            catch (error) {
              if (error.name !== 'LegacyError') throw error;
              object.partError = error;
            }
          }
        } catch (error) {
          if (error.name !== 'LegacyError') throw error;
          errors.push(error);
        }
      } else if (record.id === RECORD.TXO) {
        let text;
        try {
          text = { ...X.txo(records, i), options: r.u16(0), orientation: r.u16(2) };
          i = text.next - 1;
        } catch (error) {
          if (error.name !== 'LegacyError') throw error;
          errors.push(error);
          while (records[i + 1]?.id === RECORD.CONTINUE) i++;
          if (lastObject) lastObject.textError = error;
        }
        markers.push({ kind: 'text', text, end: length });
        if (lastObject && text) lastObject.text = text;
      } else if (record.id !== RECORD.CONTINUE) inDrawing = false;
    }
    const drawingTree = tree(join(chunks, length), errors), shapeByRecord = new Map();
    const readShape = record => {
      try {
        const shape = A.shapes([record], 'xls')[0];
        shapeByRecord.set(record, shape);
        return shape;
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        errors.push(error);
        return null;
      }
    };
    const shapes = list => list.flatMap(record => {
      if (record.type === RECORD.SHAPE) return [readShape(record)].filter(Boolean);
      if (record.type === RECORD.GROUP) {
        const children = record.children || [], header = children.find(c => c.type === RECORD.SHAPE);
        const group = header && readShape(header);
        const rest = shapes(children.filter(c => c !== header));
        if (!group) return rest;
        group.children = rest;
        return [group];
      }
      return shapes(record.children || []);
    });
    const result = shapes(drawingTree), byEnd = new Map(markers.map(m => [m.end, m]));
    const connectors = new Map();
    const walk = (list, base = 0, owner = null) => {
      for (const record of list) {
        const shape = shapeByRecord.get(record) || owner, marker = byEnd.get(base + record.end);
        if (record.type === RECORD.CLIENT && marker?.kind === 'object' && shape) {
          marker.object.shape = shape;
          shape.object = marker.object;
        } else if (record.type === RECORD.TEXTBOX && marker?.kind === 'text' && shape?.object) {
          shape.object.text = marker.text;
        } else if (record.type === RECORD.CONNECTOR) {
          try {
            const r = record.r;
            connectors.set(r.u32(12), { start: r.u32(4), end: r.u32(8),
              startSite: r.u32(16), endSite: r.u32(20) });
          } catch (error) {
            if (error.name !== 'LegacyError') throw error;
            errors.push(error);
          }
        }
        if (record.children) walk(record.children, base + record.at + 8, shape);
      }
    };
    walk(drawingTree);
    return meta.drawings = { objects, shapes: result, connectors, errors };
  };

  X.drawingGrid = function (meta) {
    const columns = 256, rows = 65536;
    const defaultWidth = meta.format.defaultColWidth ?? (meta.format.baseColWidth ?? 8) + 0.43;
    const widths = new Float64Array(columns), heights = new Float64Array(rows);
    widths.fill(Math.floor(defaultWidth * 7 + 5));
    heights.fill((meta.format.defaultRowHeight || 12.75) * 4 / 3);
    for (const col of meta.cols) for (let i = col.min - 1; i < col.max; i++) {
      widths[i] = col.hidden ? 0 : Math.floor(col.width * 7 + 5);
    }
    for (const [row, attrs] of meta.rows) {
      heights[row] = attrs.hidden ? 0 : attrs.ht > 0 && !attrs.defaultHeight ? attrs.ht * 4 / 3 : heights[row];
    }
    const x = new Float64Array(columns + 1), y = new Float64Array(rows + 1);
    for (let i = 0; i < columns; i++) x[i + 1] = x[i] + widths[i];
    for (let i = 0; i < rows; i++) y[i + 1] = y[i] + heights[i];
    return { x, y, widths, heights };
  };

  X.drawingAnchor = function (anchor, grid) {
    const a = anchor;
    if (!a || a.kind !== 'cells' || a.col1 > 255 || a.col2 > 255 || a.row1 > 65535 || a.row2 > 65535) {
      B.fail('XLS_DRAWING_ANCHOR', N('A drawing object has an invalid cell anchor'));
    }
    const from = { col: a.col1, colOff: Math.round(a.dx1 * grid.widths[a.col1] * 9525 / 1024),
      row: a.row1, rowOff: Math.round(a.dy1 * grid.heights[a.row1] * 9525 / 256) };
    const to = { col: a.col2, colOff: Math.round(a.dx2 * grid.widths[a.col2] * 9525 / 1024),
      row: a.row2, rowOff: Math.round(a.dy2 * grid.heights[a.row2] * 9525 / 256) };
    const x = Math.round(grid.x[a.col1] * 9525) + from.colOff;
    const y = Math.round(grid.y[a.row1] * 9525) + from.rowOff;
    const right = Math.round(grid.x[a.col2] * 9525) + to.colOff;
    const bottom = Math.round(grid.y[a.row2] * 9525) + to.rowOff;
    return { from, to, rect: { x, y, cx: Math.max(0, right - x), cy: Math.max(0, bottom - y) },
      editAs: a.flags & 1 ? 'absolute' : a.flags & 2 ? 'oneCell' : null };
  };
})(typeof window !== 'undefined' ? window : globalThis);
