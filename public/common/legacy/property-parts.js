/* Binary Office properties -> ECMA-376 core, extended and custom property parts.
 * Property IDs/semantics: MS-OSHARED §2.3.3.2. FILETIME editing duration is not a date. */
(function (root) {
  'use strict';
  const B = root.L.legacy, el = B.el;
  const ns = 'http://schemas.openxmlformats.org/officeDocument/2006/';
  const vt = ns + 'docPropsVTypes';
  const text = (name, value, attrs) => value == null ? '' : el(name, attrs, B.xml(value));
  const coreDate = value => typeof value === 'string' ? value.replace(/\.\d+Z$/, 'Z') : value;
  const names = { 2: 'title', 3: 'subject', 4: 'author', 5: 'keywords', 6: 'description', 7: 'template', 8: 'last author', 9: 'revision', 10: 'editing time', 11: 'last printed', 12: 'creation date', 13: 'modification date' };
  function variant(prop) {
    if (!prop || prop.unsupported) return null;
    const { type, value } = prop;
    if (type & 0x1000) {
      if (!prop.children) return null;
      const children = prop.children.map(variant); if (children.some(s => s == null)) return null;
      // A variant vector can represent every retained scalar without changing its type.
      return el('vt:vector', { size: children.length, baseType: 'variant' }, children.map(s => el('vt:variant', {}, s)).join(''));
    }
    const types = { 0: 'empty', 1: 'null', 2: 'i2', 3: 'i4', 4: 'r4', 5: 'r8', 6: 'cy', 7: 'date', 8: 'lpwstr', 10: 'error', 11: 'bool', 16: 'i1', 17: 'ui1', 18: 'ui2', 19: 'ui4', 20: 'i8', 21: 'ui8', 22: 'int', 23: 'uint', 30: 'lpwstr', 31: 'lpwstr', 64: 'filetime', 72: 'clsid' };
    const tag = types[type]; if (!tag || typeof value === 'number' && !Number.isFinite(value) || value == null && type > 1) return null;
    if (type <= 1) return el('vt:' + tag);
    return text('vt:' + tag, type === 11 ? value ? 'true' : 'false' : type === 72 ? '{' + value.toUpperCase() + '}' : type === 10 ? '0x' + value.toString(16).padStart(8, '0').toUpperCase() : value);
  }
  B.propertyParts = function (cfb, options = {}) {
    const losses = [], parts = [];
    const read = (name, kind) => {
      try { return B.properties(cfb.stream('\x05' + name), { kind }); }
      catch (e) { if (!['LegacyError', 'CFBError'].includes(e.name)) throw e; losses.push(B.loss('legacy-properties', 'Document properties could not be read', name, 'drop', 'Document properties')); return []; }
    };
    const summary = read('SummaryInformation', 'summary'), document = read('DocumentSummaryInformation', 'documentSummary');
    const sv = summary[0]?.values || {}, dv = document[0]?.values || {};
    const value = (obj, id) => obj[id]?.unsupported ? null : obj[id]?.value;
    const safe = (obj, id) => { const v = value(obj, id); return typeof v === 'number' ? Number.isFinite(v) ? v : null : v; };
    let core = '';
    for (const [id, tag] of [[2, 'dc:title'], [3, 'dc:subject'], [4, 'dc:creator'], [5, 'cp:keywords'], [6, 'dc:description'], [8, 'cp:lastModifiedBy'], [9, 'cp:revision'], [11, 'cp:lastPrinted']]) core += text(tag, id === 11 ? coreDate(safe(sv, id)) : safe(sv, id));
    for (const [id, tag] of [[12, 'dcterms:created'], [13, 'dcterms:modified']]) core += text(tag, coreDate(value(sv, id)), { 'xsi:type': 'dcterms:W3CDTF' });
    core += text('cp:category', value(dv, 2)) + text('cp:contentType', value(dv, 26)) + text('cp:contentStatus', value(dv, 27)) + text('dc:language', value(dv, 28)) + text('cp:version', value(dv, 29));
    parts.push({ name: 'docProps/core.xml', type: 'application/vnd.openxmlformats-package.core-properties+xml', relation: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', data: B.xmlDeclaration + el('cp:coreProperties', {
      'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', 'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
      'xmlns:dcterms': 'http://purl.org/dc/terms/', 'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/', 'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance'
    }, core) });
    const app = {};
    for (const [id, tag] of [[7, 'Template'], [14, 'Pages'], [15, 'Words'], [16, 'Characters'], [18, 'Application'], [19, 'DocSecurity']]) if (safe(sv, id) != null) app[tag] = safe(sv, id);
    if (sv[10]?.ticks != null) {
      const minutes = BigInt(sv[10].ticks) / 600000000n;
      if (minutes >= 0n && minutes <= 2147483647n) app.TotalTime = Number(minutes);
      else losses.push(B.loss('legacy-property-value', 'An invalid editing-time property was not kept', 'SummaryInformation', 'drop', 'Document properties'));
    }
    for (const [id, tag] of [[3, 'PresentationFormat'], [5, 'Lines'], [6, 'Paragraphs'], [7, 'Slides'], [8, 'Notes'], [9, 'HiddenSlides'], [10, 'MMClips'], [11, 'ScaleCrop'], [14, 'Manager'], [15, 'Company'], [16, 'LinksUpToDate'], [17, 'CharactersWithSpaces'], [19, 'SharedDoc'], [22, 'HyperlinksChanged']]) {
      let v = safe(dv, id); if (v != null) { if (id === 16) v = !v; app[tag] = v; }
    }
    // PIDDSI_VERSION's low word is a binary producer build, not an OOXML
    // minor version. Word 2016 refuses 11.6568/12.6568 with a converter
    // dialog; the same packages open with 11.0000/12.0000, as Office writes.
    if (typeof value(dv, 23) === 'number') { const ver = value(dv, 23) >>> 0; app.AppVersion = (ver >>> 16) + '.0000'; }
    Object.assign(app, options.extended || {});
    const vectors = {};
    const heads = dv[12]?.value, titles = dv[13]?.value;
    if (Array.isArray(heads) && heads.length % 2 === 0 && heads.every((v, i) => i % 2 ? Number.isInteger(v) && v >= 0 : typeof v === 'string')) {
      vectors.HeadingPairs = el('HeadingPairs', {}, el('vt:vector', { size: heads.length, baseType: 'variant' }, heads.map((v, i) => el('vt:variant', {}, text(i % 2 ? 'vt:i4' : 'vt:lpstr', v))).join('')));
    }
    if (Array.isArray(titles) && titles.every(v => typeof v === 'string')) vectors.TitlesOfParts = el('TitlesOfParts', {}, el('vt:vector', { size: titles.length, baseType: 'lpstr' }, titles.map(v => text('vt:lpstr', v)).join('')));
    const order = ['Template', 'TotalTime', 'Pages', 'Words', 'Characters', 'Application', 'DocSecurity', 'PresentationFormat', 'Lines', 'Paragraphs', 'Slides', 'Notes', 'HiddenSlides', 'MMClips', 'ScaleCrop', 'HeadingPairs', 'TitlesOfParts', 'Manager', 'Company', 'LinksUpToDate', 'CharactersWithSpaces', 'SharedDoc', 'HyperlinksChanged', 'AppVersion'];
    const extended = order.map(name => vectors[name] || text(name, typeof app[name] === 'boolean' ? app[name] ? 'true' : 'false' : app[name])).join('');
    parts.push({ name: 'docProps/app.xml', type: 'application/vnd.openxmlformats-officedocument.extended-properties+xml', relation: 'extended-properties', data: B.xmlDeclaration + el('Properties', { xmlns: ns + 'extended-properties', 'xmlns:vt': vt }, extended) });
    let custom = ''; const used = new Set(); let nextPid = 2;
    for (const set of document.slice(1)) {
      for (const [key, prop] of Object.entries(set.values)) {
        const id = +key; if (id < 2 || id >= 0x1000000) continue; // Linked-property strings are handled with their values below.
        const name = set.dictionary[id]; if (!name) continue;
        const body = variant(prop);
        if (body == null || used.has(name.toLowerCase())) { losses.push(B.loss('legacy-custom-property', 'A document property could not be kept', name, 'drop', 'Document properties')); continue; }
        used.add(name.toLowerCase()); const linked = value(set.values, id | 0x1000000);
        custom += el('property', { fmtid: '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}', pid: nextPid++, name, linkTarget: typeof linked === 'string' ? linked : null }, body);
      }
    }
    if (custom) parts.push({ name: 'docProps/custom.xml', type: 'application/vnd.openxmlformats-officedocument.custom-properties+xml', relation: 'custom-properties', data: B.xmlDeclaration + el('Properties', { xmlns: ns + 'custom-properties', 'xmlns:vt': vt }, custom) });
    for (const [sets, kind] of [[summary, 'SummaryInformation'], [document, 'DocumentSummaryInformation']]) for (const set of sets) for (const issue of set.issues) {
      losses.push({ ...B.loss('legacy-property-read', 'A document property could not be read', kind + (issue.id == null ? '' : ':' + issue.id), 'drop', 'Document properties'),
        detail: issue.code + (kind === 'SummaryInformation' && names[issue.id] ? ' (' + names[issue.id] + ')' : '') });
    }
    return { parts, losses, summary, document };
  };
  B.addProperties = (pkg, cfb, options) => {
    const result = B.propertyParts(cfb, options);
    for (const p of result.parts) { pkg.add(p.name, p.data, p.type); pkg.rel('', p.relation, p.name); }
    return result.losses;
  };
})(typeof window !== 'undefined' ? window : globalThis);
