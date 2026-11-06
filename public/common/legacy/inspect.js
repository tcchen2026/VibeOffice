/* Read-only classification and in-file references for legacy Office fixtures.
 * A converter must refuse unsupported inputs explicitly; inspection is not conversion. */
(function (root) {
  'use strict';
  const B = root.L.legacy;
  const isCFB = b => b.length >= 8 && [208, 207, 17, 224, 161, 177, 26, 225].every((v, i) => b[i] === v);
  function identify(bytes, cfb) {
    if (cfb) {
      const names = new Set(cfb.entries.filter(e => !e.path.includes('/')).map(e => e.name.toLowerCase()));
      if (names.has('worddocument')) {
        const f = B.doc.fib(cfb.stream('WordDocument'));
        return {
          format: 'doc',
          version: f.version,
          nFib: f.nFib,
          encrypted: f.encrypted,
          template: f.template
        };
      }
      if (names.has('workbook') || names.has('book')) {
        return { format: 'xls', ...B.xls.identify(cfb.stream(names.has('workbook') ? 'Workbook' : 'Book')) };
      }
      if (names.has('pp40')) return { format: 'ppt', version: 'PowerPoint95', encrypted: null };
      if (names.has('powerpoint document')) {
        const user = cfb.stream('Current User'),
          r = new B.Reader(user || new Uint8Array(0));
        if (r.length >= 4 && r.length < 28 && r.u32(0) + 4 === r.length)
          return { format: 'ppt', version: 'PowerPoint95', encrypted: false };
        return { format: 'ppt', version: 'PowerPoint97+', encrypted: r.u32(12) === 0xf3d1c4df };
      }
      if (names.has('encryptedpackage'))
        return { format: 'ooxml', version: 'encrypted package', encrypted: true };
      return { format: 'cfb', version: 'unknown', encrypted: null };
    }
    const r = new B.Reader(bytes);
    if (bytes.length >= 4 && [9, 0x209, 0x409, 0x809].includes(r.u16(0)))
      return { format: 'xls', ...B.xls.identify(bytes), raw: true };
    if (bytes.length >= 4 && [0xa5ec, 0xa5dc].includes(r.u16(0)))
      return { format: 'doc', version: B.doc.fib(bytes).version, encrypted: false, raw: true };
    if (bytes[0] === 80 && bytes[1] === 75) return { format: 'zip', version: 'ZIP/OOXML', encrypted: false };
    const head = B.decode(bytes.subarray(0, Math.min(bytes.length, 256)), 1252).trim();
    return {
      format: /^\{\\rtf/.test(head)
        ? 'rtf'
        : /^(?:<!doctype html|<html)/i.test(head)
          ? 'html'
          : /^<\?xml/.test(head)
            ? 'xml'
            : 'unknown',
      version: 'unknown',
      encrypted: null
    };
  }
  B.sniff = bytes => identify(bytes, isCFB(bytes) ? B.cfb(bytes) : null);
  for (const kind of ['xls', 'doc', 'ppt']) B[kind].sniff = B.sniff;
  B.inspect = function (bytes) {
    const cfb = isCFB(bytes) ? B.cfb(bytes) : null,
      result = { ...identify(bytes, cfb), size: bytes.length, features: {}, references: {} };
    if (cfb) {
      result.streams = cfb.entries.map(e => ({ path: e.path, type: e.type, size: e.size }));
      result.features.macros = cfb.entries.some(e =>
        /(?:^|\/)(VBA|_VBA_PROJECT_CUR|Macros)(?:\/|$)/i.test(e.path)
      );
      for (const [name, key] of [
        ['\x05SummaryInformation', 'summary'],
        ['\x05DocumentSummaryInformation', 'documentSummary']
      ]) {
        try {
          result.references[key] = B.properties(cfb.stream(name), { kind: key });
        } catch (e) {
          if (!['CFBError', 'LegacyError'].includes(e.name)) throw e;
          result.references[key + 'Error'] = { code: e.code, message: e.message };
        }
      }
    }
    if (result.encrypted) return result;
    if (result.format === 'doc' && result.version === 'Word97+' && cfb) {
      const doc = B.doc.read(cfb);
      Object.assign(result.features, doc.features);
      result.references.stories = doc.storyText;
      result.references.pieces = doc.pieces.map(({ text, ...piece }) => piece);
    } else if (result.format === 'xls') {
      const xls = B.xls.read(cfb ? cfb.stream('Workbook') || cfb.stream('Book') : bytes);
      result.features.recordCounts = xls.features;
      result.features.formulas = xls.features[6] || 0;
      result.features.charts = xls.features[0x1002] || 0;
      result.features.drawings = xls.features[0xec] || 0;
      result.features.comments = xls.features[0x1c] || 0;
      result.features.hyperlinks = xls.features[0x1b8] || 0;
      result.features.validation = xls.features[0x1be] || 0;
      result.features.conditionalFormats = (xls.features[0x1b1] || 0) + (xls.features[0x87a] || 0);
      result.features.filters = xls.features[0x9d] || 0;
      result.references.sheets = xls.sheets.map(s => ({
        name: s.name,
        cells: s.cells.map(({ formulaRecord, formula, ...c }) => ({
          ...c,
          formula: formula ? [...formula] : undefined
        }))
      }));
    } else if (result.format === 'ppt' && result.version === 'PowerPoint97+' && cfb) {
      const ppt = B.ppt.read(cfb);
      result.features.recordCounts = ppt.features;
      result.features.slides = ppt.slides.length;
      result.features.masters = ppt.masters.length;
      result.features.notes = ppt.notes.length;
      result.features.incrementalEdits = ppt.edits.length;
      result.features.tables = !!ppt.features.tables;
      result.features.drawings = ppt.features[1036] || 0;
      result.references.slides = ppt.slides.map(s => ({ id: s.id, persistRef: s.ref, text: s.text }));
    }
    return result;
  };
})(typeof window !== 'undefined' ? window : globalThis);
