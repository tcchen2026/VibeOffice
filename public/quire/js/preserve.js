/* Quire — package ownership and model-attached OOXML preservation. */
(function () {
  'use strict';
  const L = window.L, K = L.opc, N = K.NS, P = (L.preserve = {});
  const RT = name => N.rel + '/' + name;
  const canonical = K.relationshipType;
  const kids = el => Array.from(el?.children || []);
  const attr = (el, name) => Array.from(el?.attributes || []).find(a => a.localName === name && a.prefix !== 'xmlns')?.value;
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  const coreRel = N.pkg + '/metadata/core-properties';
  // Full relationship types, never just their last path segment. Opaque is the
  // default: reading a preview does not mean the target has been represented.
  P.parts = Object.freeze([
    ['word/styles.xml', RT('styles'), 'regenerated'],
    ['word/numbering.xml', RT('numbering'), 'regenerated'],
    ['word/settings.xml', RT('settings'), 'merged'],
    ['word/fontTable.xml', RT('fontTable'), 'merged'],
    ['word/webSettings.xml', RT('webSettings'), 'opaque'],
    ['word/theme/theme1.xml', RT('theme'), 'opaque'],
    ['word/footnotes.xml', RT('footnotes'), 'regenerated'],
    ['word/endnotes.xml', RT('endnotes'), 'regenerated'],
    ['word/comments.xml', RT('comments'), 'regenerated'],
    ['word/commentsExtended.xml', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'regenerated'],
    ['docProps/core.xml', coreRel, 'merged', true],
    ['docProps/app.xml', RT('extended-properties'), 'merged', true],
    ['docProps/custom.xml', RT('custom-properties'), 'merged', true],
  ]);
  P.settings = Object.freeze({ zoom: 'zoom', defTab: 'defaultTabStop', evenOdd: 'evenAndOddHeaders', track: 'trackRevisions', mirror: 'mirrorMargins', gutterTop: 'gutterAtTop', autoHyphen: 'autoHyphenation', showBg: 'displayBackgroundShape', updateFields: 'updateFields', protect: 'documentProtection', fnPr: 'footnotePr', enPr: 'endnotePr', compat: 'compat', compatXML: 'compat', docVars: 'docVars' });
  P.begin = function (doc, format) {
    const pkg = doc.pkg;
    const consumed = new Set([RT('officeDocument'), ...P.parts.filter(p => p[2] !== 'opaque').map(p => p[1]), ...['header', 'footer', 'image', 'chart', 'hyperlink'].map(RT)]);
    const output = K.output(pkg, { doc, format, consumes: (base, rel, type) => consumed.has(type),
      merge: (base, source, data, writer) => P.merge(doc, base, source, data, writer) });
    output.bind('word/document.xml', pkg?.main);
    for (const [base, type, mode, root] of P.parts) {
      const rel = pkg?.rels(root ? '' : pkg.main).find(r => canonical(r.type) === type && !r.external);
      output.bind(base, rel?.part, base === 'word/theme/theme1.xml' && !same(doc.theme, doc.keep?.values?.theme) ? 'regenerated' : mode);
    }
    return output;
  };
  P.merge = function (doc, base, source, generated, writer) {
    const original = writer.pkg.xml(source), fragment = K.fragment(original, { pkg: writer.pkg, part: source });
    const fresh = K.parse(generated), previous = doc.keep?.values || {};
    let xml = fragment.xml;
    if (base === 'word/settings.xml') {
      const replacements = {};
      for (const [key, tag] of Object.entries(P.settings)) if (!same(doc.settings[key], previous.settings?.[key])) {
        replacements['{' + N.w + '}' + tag] = kids(fresh).filter(e => e.localName === tag).map(K.raw);
      }
      xml = K.merge(xml, replacements, 'w:CT_Settings');
    } else if (base === 'word/fontTable.xml') {
      const tree = K.parse(xml), existing = new Set(kids(tree).map(e => attr(e, 'name')));
      const added = kids(fresh).filter(e => !existing.has(attr(e, 'name'))).map(K.raw);
      if (added.length) xml = K.mergeBag(xml, { __newFonts: added }, e => attr(e, 'name'));
    } else if (base === 'docProps/custom.xml') {
      xml = K.customProperties(xml, generated, doc.custom, previous.custom);
    } else {
      const owned = base === 'docProps/core.xml' ? { title: 'title', subject: 'subject', creator: 'creator', keywords: 'keywords', description: 'description', lastModifiedBy: 'lastModifiedBy', revision: 'revision', created: 'created', modified: 'modified', category: 'category', lastPrinted: 'lastPrinted', contentStatus: 'contentStatus', language: 'language', identifier: 'identifier', version: 'version' } : { manager: 'Manager', company: 'Company', template: 'Template' };
      xml = K.properties(xml, generated, doc.props, previous.props, owned);
    }
    return writer.emit(K.slice(fragment, xml), source);
  };
})();
