/* Quire — package ownership and model-attached OOXML preservation. */
(function () {
  'use strict';
  const L = window.L, D = L.D, K = L.opc, N = K.NS, P = (L.preserve = {});
  const RT = name => N.rel + '/' + name;
  const canonical = K.relationshipType;
  const kids = el => Array.from(el?.children || []);
  const attr = (el, name) => Array.from(el?.attributes || []).find(a => a.localName === name && a.prefix !== 'xmlns')?.value;
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  P.pictureBulletIntact = level => level.pictureKeep && ['picture', 'fmt', 'text', 'glyph'].every(key => same(level[key], level.pictureKeep.values[key]));
  const coreRel = N.pkg + '/metadata/core-properties';
  P.commentMetadata = doc => Object.entries(doc.comments).map(([id, c]) => [id, c.keep?.paraId, c.parent ?? null, !!c.done]);
  P.keepCommentExtension = doc => doc.keep?.commentExtension && same(P.commentMetadata(doc), doc.keep.commentExtension.values);
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
    const owned = new Set((doc.keep?.objects || []).flatMap(o => o.rels || []).filter(r => r.source === pkg?.id).map(r => r.owner + ':' + r.id));
    const output = K.output(pkg, { doc, format,
      contentType: (source, part) => { const mime = L.extToMime(part.split('.').pop()); return mime.startsWith('image/') ? mime : undefined; },
      consumes: (base, rel, type) => consumed.has(type) || owned.has((output.originals.get(base) || '') + ':' + rel.id),
      merge: (base, source, data, writer) => P.merge(doc, base, source, data, writer), audit: P.reportLosses });
    output.bind('word/document.xml', pkg?.main);
    for (const [base, type, mode, root] of P.parts) {
      const rel = pkg?.rels(root ? '' : pkg.main).find(r => canonical(r.type) === type && !r.external);
      output.bind(base, rel?.part, base === 'word/commentsExtended.xml' && P.keepCommentExtension(doc) ? 'opaque' :
        base === 'word/theme/theme1.xml' && !same(doc.theme, doc.keep?.values?.theme) ? 'regenerated' : mode);
    }
    return output;
  };
  P.reportLosses = writer => K.reportFeatures(writer, {
    accepts: el => wordNS(el.namespaceURI) && ['document', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments'].includes(el.localName),
    classify: el => {
      if (!wordNS(el.namespaceURI)) return null;
      const tag = el.localName;
      if (tag === 'body') return 'bodies';
      if (tag === 'sectPr' && wordNS(el.parentNode?.namespaceURI) && el.parentNode.localName === 'body') return 'finalSections';
      if (['smartTag', 'customXml'].includes(tag)) return 'xml';
      if (['dir', 'bdo'].includes(tag)) return 'direction';
      if (['moveFrom', 'moveTo'].includes(tag)) return 'moves';
      if (tag === 'fldSimple' || tag === 'fldChar' && attr(el, 'fldCharType') === 'begin') return 'fields';
      return ({ sdt: 'controls', bookmarkStart: 'bookmarks', permStart: 'permissions', comment: 'comments', footnote: 'footnotes', endnote: 'endnotes', ins: 'revisions', del: 'revisions' })[tag];
    },
    labels: { bodies: { text: 'Some text in this file is outside the main document and won\'t be saved', notify: true }, finalSections: 'Conflicting final section definitions could not all be retained', xml: 'Smart tags or inline XML wrappers were converted to ordinary content', direction: 'Inline direction wrappers were converted to ordinary text', moves: 'Move revisions were converted to ordinary revisions or text', fields: 'Some original fields were converted or removed', controls: 'Some original content-control wrappers were converted or removed', bookmarks: 'Some original bookmarks were removed', permissions: 'Some original permission ranges were removed', comments: 'Some original comments were removed', footnotes: 'Some original footnotes were removed', endnotes: 'Some original endnotes were removed', revisions: 'Some original revision markup was converted or removed' },
  });
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

  /* A property is regenerated only when the model fields owning it change.
     The original fragment keeps extension children and attributes on otherwise
     familiar properties (theme fonts/colours, language, spacing, and so on). */
  const wordNS = ns => ns === N.w || ns === 'http://purl.oclc.org/ooxml/wordprocessingml/main';
  const expanded = e => '{' + (wordNS(e.namespaceURI) ? N.w : e.namespaceURI) + '}' + e.localName;
  const originalProperties = new WeakMap();
  P.captureProperties = function (root) {
    // Capture before compatibility choices are normalised. A property fragment
    // belongs to this element, even when the element itself is inside a choice.
    const nodes = Array.from(root.getElementsByTagName('*'));
    for (const el of nodes) if (wordNS(el.namespaceURI) && ['rPr', 'pPr', 'sectPr', 'numbering', 'abstractNum', 'num', 'lvlOverride', 'lvl', 'numPicBullet', 'tbl', 'tblPr', 'drawing', 'pict', 'object', 'control', 'altChunk', 'contentPart'].includes(el.localName)) originalProperties.set(el, K.raw(el));
    // Leave an anchor when compatibility normalisation selects an empty branch.
    // captureAC still reads the immutable source text, without this sentinel.
    for (const el of nodes) if (el.namespaceURI === N.mc && ['Choice', 'Fallback'].includes(el.localName) && !el.children.length) el.appendChild(el.ownerDocument.createElementNS('urn:vibeoffice:keep', 'keep'));
    for (const el of nodes) if (el.namespaceURI === N.mc && el.localName === 'AlternateContent' && !kids(el).some(c => ['Choice', 'Fallback'].includes(c.localName))) {
      const fallback = el.ownerDocument.createElementNS(N.mc, 'mc:Fallback');
      fallback.appendChild(el.ownerDocument.createElementNS('urn:vibeoffice:keep', 'keep')); el.appendChild(fallback);
    }
  };
  P.pictureBullet = (el, context) => K.fragment(K.alternate(el) ? el : K.parse(originalProperties.get(el) || K.raw(el)), context);
  const alternatives = e => e.namespaceURI === N.mc && ['AlternateContent', 'Choice', 'Fallback'].includes(e.localName);
  const propertyChildren = el => kids(el).flatMap(c => alternatives(c) ? propertyChildren(c) : [c]);
  P.propertyFields = Object.freeze({
    rPr: { rStyle: ['style'], rFonts: ['font', 'fontEA', 'fontCS', 'hint'], b: ['b'], bCs: ['bCs'], i: ['i'], iCs: ['iCs'], caps: ['caps'], smallCaps: ['smallCaps'], strike: ['strike'], dstrike: ['dstrike'], outline: ['outline'], shadow: ['shadow'], emboss: ['emboss'], imprint: ['imprint'], noProof: ['noProof'], vanish: ['hidden'], webHidden: ['webHidden'], color: ['color'], spacing: ['spacing'], w: ['w'], kern: ['kern'], position: ['pos'], sz: ['sz'], szCs: ['szCs', 'sz'], highlight: ['hl'], u: ['u', 'uColor'], effect: ['effect'], bdr: ['border'], shd: ['shd'], vertAlign: ['vert'], rtl: ['rtl'], cs: ['cs'], em: ['em'], lang: ['lang', 'langEA'], specVanish: ['specVanish'], rPrChange: ['chg'] },
    pPr: { pStyle: ['style'], keepNext: ['keepNext'], keepLines: ['keepLines'], pageBreakBefore: ['pageBreakBefore'], framePr: ['frame', 'dropCap'], widowControl: ['widow'], numPr: ['num'], suppressLineNumbers: ['noLineNum'], pBdr: ['borders'], shd: ['shd'], tabs: ['tabs'], suppressAutoHyphens: ['noHyphen'], bidi: ['bidi'], snapToGrid: ['noSnap'], spacing: ['sp'], ind: ['ind'], contextualSpacing: ['contextual'], mirrorIndents: ['mirrorInd'], jc: ['jc'], textDirection: ['textDir'], textAlignment: ['textAlign'], outlineLvl: ['outline'], pPrChange: ['chg'] },
    tblPr: { tblStyle: ['style'], tblpPr: ['float'], bidiVisual: ['bidi'], tblStyleRowBandSize: ['rowBand'], tblStyleColBandSize: ['colBand'], tblW: ['w'], jc: ['jc'], tblCellSpacing: ['spacing'], tblInd: ['ind'], tblBorders: ['borders'], shd: ['shd'], tblLayout: ['layout'], tblCellMar: ['cellMar'], tblLook: ['look'], tblCaption: ['caption'], tblDescription: ['desc'] },
    sectPr: { headerReference: ['refs.hdr'], footerReference: ['refs.ftr'], footnotePr: ['fnPr'], endnotePr: ['enPr'], type: ['type'], pgSz: ['pgW', 'pgH', 'orient', 'paperCode'], pgMar: ['mt', 'mr', 'mb', 'ml', 'hdr', 'ftr', 'gutter'], paperSrc: ['paperSrc'], pgBorders: ['borders'], lnNumType: ['lnNum'], pgNumType: ['pgNum'], cols: ['cols'], formProt: ['formProt'], vAlign: ['vAlign'], titlePg: ['titlePg'], textDirection: ['textDir'], bidi: ['bidi'], docGrid: ['docGrid'] },
    lvl: { start: ['start'], numFmt: ['fmt', 'custFmt'], lvlRestart: ['restart'], pStyle: ['pStyle'], isLgl: ['isLgl'], suff: ['suff'], lvlText: ['text'], lvlPicBulletId: ['picture', 'fmt', 'text', 'glyph'], legacy: ['legacy'], lvlJc: ['jc'], pPr: ['pPr', 'ind', 'tabPos'], rPr: ['rPr'] },
    abstractNum: { multiLevelType: ['multi'], name: ['name'], styleLink: ['styleDef'], numStyleLink: ['styleLink'] },
    num: { abstractNumId: ['abs'] },
    lvlOverride: { startOverride: ['start'] },
    numbering: {},
  });
  const numberTypes = { numbering: 'w:CT_Numbering', abstractNum: 'w:CT_AbstractNum', num: 'w:CT_Num', lvlOverride: 'w:CT_NumLvl', lvl: 'w:CT_Lvl' };
  const ownedChildren = { pPr: ['rPr', 'sectPr'], numbering: ['numPicBullet', 'abstractNum', 'num'], abstractNum: ['lvl'], num: ['abstractNumId', 'lvlOverride'], lvlOverride: ['lvl'], lvl: ['pPr', 'rPr', 'lvlPicBulletId'] };
  const propertyKeys = (kind, el) => wordNS(el.namespaceURI) ? P.propertyFields[kind][el.localName] : kind === 'rPr' && /word\/2010\/wordml$/.test(el.namespaceURI || '') ? ({ shadow: ['shadow'], textFill: ['color'], textOutline: ['outline'] }[el.localName]) : undefined;
  const field = (model, path) => path.split('.').reduce((v, k) => v?.[k], model);
  const propertyValues = (model, keys) => JSON.stringify(keys.map(k => field(model, k) ?? null), (k, v) => k === 'x' && v?.fragment ? undefined : v);
  const sectionAttributes = {
    pgSz: { w: 'pgW', h: 'pgH', orient: 'orient', code: 'paperCode' },
    pgMar: { top: 'mt', right: 'mr', bottom: 'mb', left: 'ml', header: 'hdr', footer: 'ftr', gutter: 'gutter' },
    paperSrc: { first: 'paperSrc.first', other: 'paperSrc.other' },
    pgNumType: { fmt: 'pgNum.fmt', start: 'pgNum.start', chapStyle: 'pgNum.chapStyle', chapSep: 'pgNum.chapSep' },
    lnNumType: { countBy: 'lnNum.countBy', start: 'lnNum.start', distance: 'lnNum.distance', restart: 'lnNum.restart' },
    docGrid: { type: 'docGrid.type', linePitch: 'docGrid.linePitch', charSpace: 'docGrid.charSpace' },
  };
  function sectionChild(el, fresh, model, kept) {
    const fields = sectionAttributes[el.localName];
    if (!fields || fresh.length !== 1) return fresh.map(K.raw);
    const keys = P.propertyFields.sectPr[el.localName], values = JSON.parse(kept.values[expanded(el)]);
    const before = Object.fromEntries(keys.map((k, i) => [k, values[i]])), changes = {};
    for (const [name, path] of Object.entries(fields)) if (!same(field(model, path), field(before, path))) {
      // Margin start/end are the logical aliases used by Strict sources.
      const aliases = name === 'left' ? ['left', 'start'] : name === 'right' ? ['right', 'end'] : [name];
      const old = Array.from(el.attributes).filter(a => aliases.includes(a.localName) && a.prefix && wordNS(el.lookupNamespaceURI(a.prefix)));
      const value = attr(fresh[0], name);
      if (old.length) for (const a of old) changes[a.name] = value;
      else if (value != null) {
        let prefix = el.prefix || 'w', n = 0;
        while (el.lookupNamespaceURI(prefix) && el.lookupNamespaceURI(prefix) !== el.namespaceURI) prefix = 'w' + ++n;
        if (!el.lookupNamespaceURI(prefix)) changes['xmlns:' + prefix] = el.namespaceURI;
        changes[prefix + ':' + name] = value;
      }
    }
    return [K.attributes(K.raw(el), changes)];
  }
  const propertyRanks = new Map();
  P.properties = function (model, el, kind, context) {
    if (!el) return model;
    const fragment = K.fragment(K.parse(originalProperties.get(el) || K.raw(el)), context), tree = K.parse(fragment.xml);
    model.x = { key: D.nid(), fragment, owner: el.parentNode?.localName, values: Object.fromEntries(propertyChildren(tree).map(c => [expanded(c), propertyValues(model, propertyKeys(kind, c) || [])])) };
    // Absent section/list properties are meaningful. Do not introduce writer
    // defaults unless their owning model fields were edited.
    if (kind === 'sectPr' || numberTypes[kind]) for (const [name, keys] of Object.entries(P.propertyFields[kind])) model.x.values['{' + N.w + '}' + name] = propertyValues(model, keys);
    return model;
  };
  // Required attributes belong to the model. Preserve other attributes and the
  // source prefix instead of rebuilding the wrapper and discarding extensions.
  const propertyAttributes = (xml, values) => {
    if (!values || !Object.keys(values).length) return xml;
    const el = K.parse(xml), changes = {};
    for (const [name, value] of Object.entries(values)) {
      const old = Array.from(el.attributes).find(a => a.localName === name && a.prefix && wordNS(el.lookupNamespaceURI(a.prefix)));
      if (old) changes[old.name] = value;
      else {
        let prefix = el.prefix || 'w', n = 0;
        while (el.lookupNamespaceURI(prefix) && !wordNS(el.lookupNamespaceURI(prefix))) prefix = 'w' + ++n;
        if (!el.lookupNamespaceURI(prefix)) changes['xmlns:' + prefix] = N.w;
        changes[prefix + ':' + name] = value;
      }
    }
    return K.attributes(xml, changes);
  };
  P.propertyXML = function (kind, model, generated, ctx, attributes) {
    const kept = model.x, root = propertyAttributes('<w:' + kind + ' xmlns:w="' + N.w + '">' + generated + '</w:' + kind + '>', attributes);
    if (!kept) return generated ? root : '';
    let repaired = false;
    const type = numberTypes[kind] || { rPr: 'w:CT_RPr', pPr: 'w:CT_PPr', tblPr: 'w:CT_TblPr', sectPr: 'w:CT_SectPr' }[kind];
    if (!propertyRanks.has(type)) {
      const ranks = new Map();
      K.schema.types[type].forEach((slot, i) => slot.forEach(name => ranks.set('{' + K.schema.namespaces[name.split(':')[0]] + '}' + name.split(':')[1], i)));
      propertyRanks.set(type, ranks);
    }
    const ranks = propertyRanks.get(type);
    const original = K.parse(kept.fragment.xml), fresh = kids(K.parse(root)), present = new Set(propertyChildren(original).map(expanded));
    // Paragraph-mark revisions belong to CT_ParaRPr. They precede CT_RPr's
    // shared property sequence; pPr's nested mark/section are regenerated too.
    const mark = e => kind === 'rPr' && wordNS(e.namespaceURI) && ['ins', 'del'].includes(e.localName);
    const merge = (xml, outer = false) => {
      let tree = K.parse(xml);
      const patches = kids(tree).filter(alternatives).map(e => ({ ...location(e), value: merge(K.raw(e)) }));
      if (patches.length) { xml = K.patch(xml, patches); tree = K.parse(xml); }
      if (tree.namespaceURI === N.mc && tree.localName === 'AlternateContent') return xml;
      const replacements = {};
      for (const e of kids(tree)) {
        if (mark(e)) { replacements[expanded(e)] = []; continue; }
        const name = expanded(e), keys = propertyKeys(kind, e);
        const always = wordNS(e.namespaceURI) && ownedChildren[kind]?.includes(e.localName);
        if (always || keys && kept.values[name] !== propertyValues(model, keys)) {
          const next = fresh.filter(c => expanded(c) === name);
          replacements[name] = kind === 'sectPr' ? sectionChild(e, next, model, kept) : next.map(K.raw);
        }
      }
      if (outer) for (const c of fresh) if (!mark(c) && !present.has(expanded(c))) {
        const keys = propertyKeys(kind, c);
        const owned = wordNS(c.namespaceURI) && ownedChildren[kind]?.includes(c.localName);
        if (owned || kind !== 'sectPr' && !numberTypes[kind] || keys && kept.values[expanded(c)] !== propertyValues(model, keys)) replacements[expanded(c)] = fresh.filter(e => expanded(e) === expanded(c)).map(K.raw);
      }
      // Invalid source properties can be copied into several runs by text edits.
      // Normalize their spelling/order before that multiplies the source errors;
      // retain extension attributes and every unrelated property.
      const copiedSymbol = kind === 'rPr' && kept.owner === 'lvl' && kept.fragment.copy;
      const content = kind === 'pPr' && kept.owner === 'p' || kind === 'rPr' && ['r', 'pPr'].includes(kept.owner) || copiedSymbol;
      if (!content) return K.merge(xml, replacements, type);
      // The SDK's CT_RPrList excludes highlight (unlike ordinary CT_RPr).
      // Keep invalid originals in place, but do not multiply that error when
      // copying a list. No-highlight is redundant; colored highlights become
      // the equivalent supported background shading, with a conversion notice.
      if (copiedSymbol) {
        const name = '{' + N.w + '}highlight';
        const highlight = Object.hasOwn(replacements, name) ? replacements[name].map(K.parse) : kids(tree).filter(e => expanded(e) === name);
        if (highlight.length) {
          const value = attr(highlight[0], 'val'), color = L.R?.HIGHLIGHT[value];
          if (value === 'none' || color) {
            replacements[name] = [];
            if (color) replacements['{' + N.w + '}shd'] = ['<w:shd xmlns:w="' + N.w + '" w:val="clear" w:color="auto" w:fill="' + color.slice(1) + '"/>'];
            repaired = true;
          }
        }
      }
      const groups = new Map(); let previous = -1, unordered = false, unknown = false;
      for (const e of kids(tree)) {
        const name = expanded(e), rank = ranks.get(name);
        if (!wordNS(e.namespaceURI) || mark(e)) continue;
        if (rank == null) { unknown = true; continue; }
        unordered ||= rank < previous; previous = rank;
        if (!groups.has(name)) groups.set(name, []);
        groups.get(name).push(e);
      }
      // Only content properties are duplicated by splitting text. Leave styles,
      // table defaults and ambiguous source properties intact: choosing between
      // conflicting values or unknown Word children requires an explicit edit.
      const conflicting = [...groups.values()].some(group => group.some(e => K.raw(e) !== K.raw(group[0])));
      if (unknown || conflicting) return K.merge(xml, replacements, type);
      for (const [name, group] of groups) {
        if (Object.hasOwn(replacements, name)) continue;
        let value = K.raw(group[0]);
        const duplicate = group.length > 1 && group.every(e => K.raw(e) === value);
        if (wordNS(group[0].namespaceURI) && group[0].localName === 'color') {
          const color = attr(group[0], 'val'), key = (group[0].prefix ? group[0].prefix + ':' : '') + 'val';
          if (/^#[\da-f]{6}$/i.test(color || '')) value = K.attributes(value, { [key]: color.slice(1) });
          else if (color == null && attr(group[0], 'themeColor')) value = K.attributes(value, { [key]: 'auto' });
        }
        if (duplicate || value !== K.raw(group[0])) { replacements[name] = [value]; repaired = true; }
      }
      if (unordered) {
        // merge() keeps an untouched fragment byte-for-byte unless asked to
        // replace a child. Replacing an identical child also orders this repair.
        const first = [...groups].find(([name]) => !Object.hasOwn(replacements, name));
        if (first && !Object.keys(replacements).length) replacements[first[0]] = first[1].map(K.raw);
        repaired = true;
      }
      return K.merge(xml, replacements, type);
    };
    try {
      let xml = merge(kept.fragment.xml, true);
      const marks = fresh.filter(mark).map(K.raw).join('');
      if (marks) { const tree = K.parse(xml); xml = inner(xml, tree, marks + kids(tree).map(K.raw).join('')); }
      if (repaired) ctx.writer.loss({ id: 'property:repair:' + (kept.fragment.source || '') + ':' + kept.fragment.part, what: 'Invalid source formatting was normalized for compatibility.', where: kept.fragment.part, action: 'conversion' });
      return ctx.writer.emit(K.slice(kept.fragment, propertyAttributes(xml, attributes)), ctx.part);
    } catch (error) {
      ctx?.writer.loss({ id: 'property:' + kept.key, what: 'Some ' + (numberTypes[kind] ? 'list' : kind === 'rPr' ? 'text' : kind === 'sectPr' ? 'section' : 'paragraph') + ' formatting could not be retained: ' + error.message, where: kept.fragment.part, action: 'conversion' });
      return generated ? root : '';
    }
  };
  P.numberingLevels = function (abstract, present) {
    abstract.x.absentLevels = Object.fromEntries(abstract.levels.flatMap((level, i) => present.includes(i) ? [] : [[i, P.objectSignature(level, 'level')]]));
  };
  P.materializeNumbering = function (abstract) {
    if (!abstract.styleLink || !abstract.x?.linkedLevels || abstract.x.linkedLevels === P.objectSignature(abstract.levels, 'levels')) return abstract;
    // The dialog edited the effective levels of a numbering style. Make that
    // list independent; its former delegate and other users retain their values.
    const copy = K.duplicate(abstract);
    delete copy.styleLink;
    delete copy.x.absentLevels;
    delete copy.x.ownLevels;
    copy.x.materialized = true;
    return copy;
  };
  P.writeNumberingLevel = (abstract, i) => abstract.levels[i] && (!abstract.x?.absentLevels || !Object.hasOwn(abstract.x.absentLevels, i) || abstract.x.absentLevels[i] !== P.objectSignature(abstract.levels[i], 'level'));
  P.prepareNumbering = function (numbering, writer) {
    for (const [kind, models] of [['abstractNumId', numbering.abs], ['numId', numbering.nums]]) for (const [id, model] of Object.entries(models)) {
      writer.ids.reserve('numbering', kind, id);
      const f = model.x?.fragment, definition = f?.ids.find(i => i.kind === kind && i.definition);
      if (!definition) continue;
      if (f.copy) writer.ids.copy(f.copy, f.copyDefinitions || f.ids.filter(i => i.definition));
      writer.ids.bind(definition.source, 'numbering', kind, definition.id, id, 'numbering', f.copy);
    }
  };
  P.rangeMarker = function (el, context) {
    const name = el.localName, bookmark = name.startsWith('bookmark');
    return D.item(bookmark ? name.endsWith('Start') ? 'bs' : 'be' : 'perm', { id: attr(el, 'id'), name: bookmark ? attr(el, 'name') : undefined,
      end: name.endsWith('End'), keep: K.fragment(K.parse(K.raw(el)), context) });
  };
  P.prepareRanges = function (doc) {
    const valid = new Map(), visited = new Set();
    for (const story of D.stories(doc)) {
      const pairs = new Map();
      D.walk(story, p => {
        if (p.t !== 'p' || visited.has(p)) return;
        visited.add(p);
        for (const it of p.runs) if (['bs', 'be', 'perm'].includes(it.t)) {
          const key = (it.t === 'perm' ? 'permission:' : 'bookmark:') + it.id;
          if (!pairs.has(key)) pairs.set(key, { start: [], end: [] });
          pairs.get(key)[it.t === 'be' || it.end ? 'end' : 'start'].push(it);
        }
      });
      for (const pair of pairs.values()) for (const it of [...pair.start, ...pair.end]) valid.set(it, !!(pair.start.length && pair.end.length));
    }
    return valid;
  };
  P.markerXML = function (item, ctx) {
    if (ctx.ranges?.get(item) === false) {
      ctx.writer.loss({ id: 'range:' + item.t + ':' + (item.keep?.source || '') + ':' + ctx.part + ':' + item.id,
        what: (item.t === 'perm' ? 'Permission range' : 'Bookmark') + ': an incomplete boundary was removed.', where: ctx.part, action: 'drop' });
      return '';
    }
    if (!item.keep) return null;
    let fragment = item.keep;
    if (item.t === 'bs') {
      const tree = K.parse(fragment.xml), key = Array.from(tree.attributes).find(a => a.localName === 'name')?.name || (tree.prefix ? tree.prefix + ':' : '') + 'name';
      fragment = K.slice(fragment, K.attributes(K.serialize(tree), { [key]: item.name || '_bm' + item.id }));
    }
    return ctx.writer.emit(fragment, ctx.part);
  };

  /* Opaque content lives on its displayed model object. Compatibility choices
     can own several items or blocks; all of them must survive before their
     original wrapper can be emitted once. Bytes remain in the package graph. */
  const drawingTypes = new Set(['img', 'shape', 'group', 'chart']);
  const opaqueStore = (object, create = false) => object.t === 'p' ? object.pPr : object.t === 'tbl' ? object.tblPr : object.cells ? object.trPr : object.tcPr || object.keep || (create ? (object.keep = {}) : null);
  const geometry = it => ({ w: it.w, h: it.h, rot: it.rot || 0, flipH: !!it.flipH, flipV: !!it.flipV,
    x: it.float?.posH?.off, y: it.float?.posV?.off });
  const geometryKeys = new Set(['w', 'h', 'x', 'y', 'rot', 'flipH', 'flipV']);
  P.objectSignature = function (object, scope, scaled = true, properties = []) {
    if (scope === 'empty') return '';
    function value(v, key, parent, root = false) {
      if (v == null || typeof v !== 'object') return v;
      if (Array.isArray(v)) return v.map(c => value(c, '', v));
      const drawing = drawingTypes.has(v.t), out = {};
      for (const k of Object.keys(v).sort()) {
        if (root && properties.includes(k)) continue;
        if (['keep', 'ac', 'acBefore', 'acAfter', 'opaque', 'id', 'key', 'paras'].includes(k) || k.startsWith('_') || k === 'x' && v[k]?.fragment) continue;
        if (drawing && geometryKeys.has(k)) continue;
        if ((v.opaquePlaceholder || v.zero || root && v.t === 'p' && v.pPr?.opaque) && ['rPr', 'pPr'].includes(k)) continue;
        if (root && scope === 'run' && k === 'rPr') continue;
        if (drawing && k === 'float') {
          const f = L.clone(v.float);
          for (const side of ['posH', 'posV']) if (f[side]) for (const a of ['off', 'align', 'pct']) delete f[side][a];
          out.float = f; continue;
        }
        if (drawing && k === 'kids') {
          const w = scaled ? v.w || 1 : 1, h = scaled ? v.h || 1 : 1;
          out.kids = v.kids.map(c => ({ geometry: [(c.x || 0) / w, (c.y || 0) / h, c.w / w, c.h / h, c.rot || 0, !!c.flipH, !!c.flipV].map(n => typeof n === 'number' ? L.round(n, 5) : n), content: value(c, '', v) }));
        } else out[k] = value(v[k], k, v);
      }
      return out;
    }
    return JSON.stringify(value(object, '', null, true));
  };
  const objectBoxes = object => {
    const out = [];
    const visit = o => {
      if (!o || typeof o !== 'object') return;
      if (drawingTypes.has(o.t)) { out.push(geometry(o)); return; }
      if (Array.isArray(o)) { o.forEach(visit); return; }
      for (const [k, v] of Object.entries(o)) if (!['keep', 'ac', 'opaque', 'x', 'paras'].includes(k) && !k.startsWith('_')) visit(v);
    };
    visit(object); return out;
  };
  const objectKey = record => JSON.stringify([record.fragment.source, record.fragment.part, record.fragment.record || record.key, record.fragment.copy || '']);
  const typeKey = (fragment, id) => JSON.stringify([fragment.source, fragment.part, id.kind, id.id]);
  P.registerObject = function (fragment, doc, label, scope = 'run') {
    const key = fragment.record || D.nid();
    const record = { key, fragment, label, scope, count: 1, index: 0 };
    // Word's diagram data may name a drawing relationship on the story part
    // through dataModelExt@relId, outside the relationships namespace.
    for (const dep of fragment.deps) if (/\/diagramData$/.test(dep.type || '')) {
      const pkg = K.package(dep.source), tree = pkg?.xml(dep.part);
      for (const el of tree?.getElementsByTagName('*') || []) if (el.localName === 'dataModelExt') {
        const id = attr(el, 'relId'), rel = pkg.rels(fragment.part).find(r => r.id === id);
        if (rel) (record.related || (record.related = [])).push({ source: dep.source, owner: fragment.part, dataPart: dep.part, ...rel });
      }
    }
    const objects = doc.keep.objects || (doc.keep.objects = []);
    if (!objects.some(o => o.key === key)) objects.push({ key, part: fragment.part, label, rels: [...fragment.deps, ...(record.related || [])].map(r => ({ source: r.source, owner: r.owner, id: r.id })) });
    return record;
  };
  P.keepObject = function (object, el, ctx, doc, label) {
    if (!object) object = el.localName === 'pict' && kids(el).every(c => c.localName === 'shapetype') ? D.item('raw', { zero: true }) : D.item('raw', { label: label || 'Embedded object', opaquePlaceholder: true });
    let xml = originalProperties.get(el) || K.raw(el), tree = K.parse(xml);
    const addedTypes = new Set();
    if (ctx.vmlTypes && ['pict', 'object'].includes(el.localName)) {
      const nodes = [tree, ...tree.getElementsByTagName('*')], present = new Set(nodes.filter(e => e.localName === 'shapetype').map(e => attr(e, 'id'))), types = [];
      for (const e of nodes) {
        const id = (e.namespaceURI === K.KNOWN_NS.v && e.getAttribute('type') || '').replace(/^#/, '');
        if (id && !present.has(id) && ctx.vmlTypes.has(id)) { present.add(id); addedTypes.add(id); types.push(K.raw(ctx.vmlTypes.get(id))); }
      }
      if (types.length) { xml = inner(xml, tree, types.join('') + kids(tree).map(K.raw).join('')); tree = K.parse(xml); }
    }
    const fragment = K.fragment(tree, { pkg: doc.pkg, part: ctx.part });
    // A preview may refer to a type defined beside an earlier object. Make the
    // added definition private to this object, so both the original save and a
    // partial copy have one complete, unique type definition and matching refs.
    if (addedTypes.size) {
      const copy = 'vml-type:' + crypto.randomUUID();
      for (const id of fragment.ids) if (id.kind.startsWith('vmlType') && addedTypes.has((id.prefix || '').replace(/^#/, '') + id.id)) id.copy = copy;
    }
    opaqueStore(object, true).opaque = P.registerObject(fragment, doc, label || object.name || object.alt || 'Drawing');
    return object;
  };
  P.opaqueBlock = function (el, ctx, doc, label) {
    const p = D.para([D.item('raw', { label, opaquePlaceholder: true })]);
    const fragment = K.fragment(K.parse(originalProperties.get(el) || K.raw(el)), { pkg: doc.pkg, part: ctx.part });
    p.pPr.opaque = P.registerObject(fragment, doc, label, 'block');
    return p;
  };
  P.readAlternates = function (parent, output, ctx, doc, scope, rPr) {
    const groups = new Map(), parentRecord = K.alternate(parent);
    return {
      add(el, from) {
        const alternate = K.alternate(el);
        if (!alternate || alternate === parentRecord) return;
        if (!groups.has(alternate.id)) groups.set(alternate.id, { el, members: [] });
        const g = groups.get(alternate.id);
        if (output.length === from) output.push(scope === 'block' ? D.para([D.item('raw', { zero: true })]) : D.item('raw', { zero: true }, rPr));
        g.members.push(...output.slice(from));
      },
      finish() {
        for (const { el, members } of groups.values()) {
          const record = P.registerObject(K.fragment(el, { pkg: doc.pkg, part: ctx.part }), doc, 'Compatibility content', scope);
          if (scope === 'inline') record.wrapper = L.clone(rPr || {});
          members.forEach((o, index) => { opaqueStore(o, true).ac = { ...record, index, count: members.length }; });
        }
      },
    };
  };
  P.emptyAlternates = function (entries, models, ctx, doc) {
    for (const { el, index } of entries) {
      const next = models.slice(index).find(Boolean), previous = models.slice(0, index).findLast(Boolean);
      const owner = next || previous;
      if (!owner) continue; // An entirely empty table is retained as an opaque block.
      const store = opaqueStore(owner, true), key = next ? 'acBefore' : 'acAfter';
      (store[key] || (store[key] = [])).push(P.registerObject(K.fragment(el, { pkg: doc.pkg, part: ctx.part }), doc, 'Compatibility content', 'empty'));
    }
  };
  P.tableMemberXML = function (object, ctx, generate) {
    const store = opaqueStore(object), emit = name => (store?.[name] || []).map(record => P.objectXML(object, name, ctx, record) || '').join('');
    return emit('acBefore') + (P.objectXML(object, 'ac', ctx) ?? generate()) + emit('acAfter');
  };
  const objectRecords = doc => {
    const out = [], visited = new Set();
    const visit = (object, container, index) => {
      if (!object || typeof object !== 'object' || visited.has(object)) return;
      visited.add(object);
      if (object.t || object.cells || object.tcPr) for (const name of ['ac', 'opaque', 'acBefore', 'acAfter']) {
        const records = opaqueStore(object)?.[name];
        for (const record of [records].flat().filter(Boolean)) out.push({ object, name, record, container, index });
      }
      if (Array.isArray(object)) object.forEach((o, i) => visit(o, object, i));
      else for (const [k, v] of Object.entries(object)) if (!['keep', 'ac', 'acBefore', 'acAfter', 'opaque', 'x', 'paras'].includes(k) && !k.startsWith('_')) visit(v);
    };
    for (const story of D.stories(doc)) visit(story.blocks);
    return out;
  };
  P.finishObjects = function (doc) {
    const groups = new Map();
    for (const entry of objectRecords(doc)) {
      const { object, record } = entry, key = objectKey(record);
      record.drawing = record.count === 1 && P.drawingProperties?.(object, record.fragment);
      const properties = Object.keys(record.drawing || {});
      record.signature = P.objectSignature(object, record.scope, true, properties); record.absoluteSignature = P.objectSignature(object, record.scope, false, properties); record.boxes = record.scope === 'empty' ? [] : objectBoxes(object);
      const previews = new Map();
      const collect = o => {
        if (!o || typeof o !== 'object') return;
        if (o.watermark) record.watermark = true;
        if (o.media && doc.pkg && doc.keep.media?.[o.media]) previews.set(o.media, { id: o.media, source: doc.pkg.id, part: doc.keep.media[o.media] });
        for (const [k, v] of Object.entries(o)) if (!['keep', 'ac', 'opaque', 'x', 'paras'].includes(k) && !k.startsWith('_')) collect(v);
      };
      if (record.scope !== 'empty') collect(object);
      if (previews.size) record.previews = [...previews.values()];
      if (!groups.has(key)) groups.set(key, []); groups.get(key).push(record);
    }
    // Field normalisation may remove instruction markers while keeping results.
    for (const records of groups.values()) records.forEach((r, index) => { r.index = index; r.count = records.length; });
  };
  P.hydrateClipboard = async function (blocks) {
    const entries = objectRecords({ main: { blocks }, hf: {}, fn: {}, en: {}, comments: {} }), loaded = new Map();
    for (const { record } of entries) for (const ref of record.previews || []) {
      const key = ref.source + ':' + ref.part;
      if (loaded.has(key)) continue;
      const pkg = K.package(ref.source), bytes = pkg?.bytes(ref.part);
      if (!bytes) continue;
      const ext = ref.part.split('.').pop().toLowerCase(), mime = pkg.type(ref.part) || L.extToMime(ext);
      const view = /^(emf|wmf)$/.test(ext) ? await L.metafile.toPNG(bytes, ext, 1600) : null;
      loaded.set(key, L.media.add(new Blob([bytes], { type: mime }), ref.part.split('/').pop(), view));
    }
    for (const { object, record } of entries) {
      const ids = new Map((record.previews || []).map(r => [r.id, loaded.get(r.source + ':' + r.part)]).filter(([, id]) => id));
      const replace = o => {
        if (!o || typeof o !== 'object') return;
        if (ids.has(o.media)) o.media = ids.get(o.media);
        for (const [k, v] of Object.entries(o)) if (!['keep', 'ac', 'opaque', 'x', 'paras'].includes(k) && !k.startsWith('_')) replace(v);
      };
      replace(object);
      // Rebind preview handles without treating a pre-copy content edit as if
      // the original object were still intact.
      for (const key of ['signature', 'absoluteSignature']) if (record[key]) { const old = JSON.parse(record[key]); replace(old); record[key] = JSON.stringify(old); }
      for (const ref of record.previews || []) if (ids.has(ref.id)) ref.id = ids.get(ref.id);
    }
    return blocks;
  };
  P.prepareObjects = function (doc, writer) {
    const groups = new Map(), seen = new Set();
    for (const entry of objectRecords(doc)) {
      const key = objectKey(entry.record);
      if (!groups.has(key)) groups.set(key, { entries: [], emitted: false });
      groups.get(key).entries.push(entry); seen.add(entry.record.key);
    }
    for (const group of groups.values()) {
      const records = group.entries.map(e => e.record);
      group.intact = records.length === records[0].count && new Set(records.map(r => r.index)).size === records.length && group.entries.every(e => e.record.signature === P.objectSignature(e.object, e.record.scope, true, Object.keys(e.record.drawing || {})) || e.record.absoluteSignature === P.objectSignature(e.object, e.record.scope, false, Object.keys(e.record.drawing || {})));
      group.intact = group.intact && group.entries.every((e, i) => e.record.index === i && e.container === group.entries[0].container && e.index === group.entries[0].index + i);
      if (records.some(r => r.watermark) && !P.keepWatermark(doc)) group.intact = false;
      group.boxes = group.entries.flatMap(e => e.record.scope === 'empty' ? [] : objectBoxes(e.object));
      const previous = group.entries.flatMap(e => e.record.boxes || []);
      group.moved = !same(previous, group.boxes);
      if (group.moved && group.boxes.length !== 1) group.intact = false;
      if (group.moved && group.boxes.length === 1) group.boxChange = Object.fromEntries(Object.entries(group.boxes[0]).filter(([k, v]) => !same(v, previous[0]?.[k])));
    }
    for (const old of doc.keep?.objects || []) if (!seen.has(old.key)) writer.loss({ id: 'object:' + old.key, what: old.label + ' was removed.', where: old.part, action: 'drop' });
    groups.vmlTypes = new Set();
    for (const group of groups.values()) if (group.intact) for (const { record: { fragment } } of group.entries) {
      if (fragment.copy) continue;
      for (const id of fragment.ids) if (id.definition && id.kind.startsWith('vmlType') && !id.copy) groups.vmlTypes.add(typeKey(fragment, id));
    }
    return groups;
  };
  // Told when a kept object can only be saved as what Quire shows of it; compatibility wrappers stay silent
  const objectNouns = { SmartArt: 'SmartArt', 'Embedded object': 'embedded object', 'Form control': 'form control', 'Embedded content': 'embedded content', 'Embedded document': 'embedded document', Drawing: 'drawing' };
  const objectNotice = (label, why) => /^Compatibility /.test(label) ? { what: label + ' was converted' + why, notify: false } :
    { what: 'This ' + (objectNouns[label] || 'drawing') + ' will be saved as ordinary content' + why, place: objectNouns[label] ? undefined : label, notify: true };
  P.objectXML = function (object, name, ctx, record = opaqueStore(object)?.[name]) {
    if (!record) return null;
    if (object.watermark && !P.keepWatermark(ctx.doc)) {
      ctx.writer.loss({ id: 'object:' + record.key, what: 'The watermark was ' + (ctx.doc.watermark ? 'replaced.' : 'removed.'), where: ctx.part, action: ctx.doc.watermark ? 'conversion' : 'drop' });
      return '';
    }
    const group = ctx.objects.get(objectKey(record));
    if (!group?.intact) {
      ctx.writer.loss({ id: 'object:' + record.key + ':' + (record.fragment.copy || ''), ...objectNotice(record.label, ', because it was edited here.'), where: ctx.part, action: 'conversion' });
      return null;
    }
    if (group.emitted) return '';
    try {
      let fragment = record.fragment;
      if (!fragment.copy && fragment.source === ctx.writer.pkg?.id && fragment.part === ctx.part) {
        const shared = new Set(fragment.ids.filter(id => id.copy && id.kind.startsWith('vmlType') && ctx.objects.vmlTypes.has(typeKey(fragment, id))).map(id => typeKey(fragment, id)));
        if (shared.size) {
          const tree = K.parse(fragment.xml), patches = [];
          for (const el of [tree, ...tree.getElementsByTagName('*')]) if (el.localName === 'shapetype') {
            const index = /^\u0001id:(\d+)\u0001$/.exec(attr(el, 'id') || '')?.[1];
            if (index != null && shared.has(typeKey(fragment, fragment.ids[+index]))) patches.push({ ...location(el), value: '' });
          }
          fragment = K.slice(fragment, K.patch(fragment.xml, patches));
          fragment.ids = fragment.ids.map(id => shared.has(typeKey(fragment, id)) ? { ...id, copy: undefined } : id);
        }
      }
      const diagrams = new Map();
      for (const rel of record.related || []) {
        const id = ctx.writer.keepRel(ctx.part, rel);
        if (id === rel.id) continue;
        const key = rel.source + ':' + rel.dataPart;
        if (!diagrams.has(key)) diagrams.set(key, { source: rel.source, part: rel.dataPart, ids: new Map() });
        diagrams.get(key).ids.set(rel.id, id);
      }
      for (const { source, part, ids } of diagrams.values()) {
        const pkg = K.package(source), tree = pkg?.xml(part);
        if (!tree) throw new Error('The copied diagram has no data part.');
        const cache = ctx.objects.diagrams || (ctx.objects.diagrams = new Map());
        const key = JSON.stringify([ctx.part, source, part, [...ids]]);
        let dest = cache.get(key);
        if (!dest) {
          let xml = K.raw(tree);
          const patches = Array.from(K.parse(xml).getElementsByTagName('*')).filter(e => e.localName === 'dataModelExt' && ids.has(attr(e, 'relId'))).map(e => ({ ...location(e), value: K.attributes(K.raw(e), { relId: ids.get(attr(e, 'relId')) }) }));
          xml = K.patch(xml, patches);
          const slash = part.lastIndexOf('/');
          dest = ctx.writer.name(part.slice(0, slash), 'data', 'xml');
          ctx.writer.copyPart(pkg, part, dest);
          ctx.writer.put(dest, xml, pkg.type(part)); cache.set(key, dest);
        }
        // The implicit dataModelExt link belongs to the destination story. A
        // private data part keeps another pasted instance/story independent.
        fragment = K.slice(fragment, fragment.xml.replace(/\u0001rel:(\d+)\u0001/g, (token, index) => {
          const dep = fragment.deps[+index];
          return dep.source === source && dep.part === part ? ctx.writer.rels(ctx.part).add(dep.type, K.relative(ctx.part, dest), false) : token;
        }));
      }
      let xml = ctx.writer.emit(fragment, ctx.part, group.moved ? { box: group.boxChange } : {});
      if (record.drawing) xml = P.mergeDrawing(xml, object, record.drawing, ctx);
      group.emitted = true; return xml;
    } catch (error) {
      const notice = objectNotice(record.label, ', because the original couldn\'t be kept.');
      ctx.writer.loss({ id: 'object:' + record.key, ...notice, notify: notice.notify && !K.sourceMissing(error), detail: error.message, where: ctx.part, action: 'conversion' });
      group.intact = false; return null;
    }
  };
  P.keepWatermark = doc => !!doc.keep?.values && Object.hasOwn(doc.keep.values, 'watermark') && same(doc.watermark, doc.keep.values.watermark);

  /* Content controls are boundaries, not text. The shell owns its properties and
     dependencies; the editor continues to own the paragraphs/runs inside it. */
  const SLOT = '\u0002content\u0002';
  const child = (el, name) => kids(el).find(e => e.localName === name);
  const location = el => L.xmlTree.source.get(el);
  const inner = (xml, el, value) => {
    const p = location(el);
    return /\/>$/.test(xml.slice(p.start, p.openEnd))
      ? K.patch(xml, [{ start: p.openEnd - 2, end: p.end, value: '>' + value + '</' + el.nodeName + '>' }])
      : K.patch(xml, [{ start: p.openEnd, end: xml.lastIndexOf('</', p.end - 1), value }]);
  };
  const setAttrs = (xml, el, values) => {
    const p = location(el);
    // This property is being edited. Serialise its attributes with double quotes
    // before assigning values, which may themselves contain apostrophes.
    return K.patch(xml, [{ start: p.start, end: p.end, value: K.attributes(K.serialize(el), values) }]);
  };
  P.control = function (el, ctx, doc) {
    let xml = K.raw(el), tree = K.parse(xml), content = child(tree, 'sdtContent');
    if (!content) return null;
    xml = inner(xml, content, SLOT);
    const pr = child(el, 'sdtPr'), binding = child(pr, 'dataBinding');
    const type = ['text', 'date', 'dropDownList', 'comboBox', 'checkbox', 'picture'].find(t => child(pr, t)) || 'richText';
    const control = { key: D.nid(), shell: K.fragment(K.parse(xml), { pkg: doc.pkg, part: ctx.part }), type,
      lock: attr(child(pr, 'lock'), 'val'), name: attr(child(pr, 'alias'), 'val') || attr(child(pr, 'tag'), 'val') || 'Content control' };
    if (child(pr, 'placeholder')) control.related = (doc.pkg?.rels(doc.pkg.main) || []).filter(r => canonical(r.type) === RT('glossaryDocument')).map(r => ({ source: doc.pkg.id, owner: doc.pkg.main, ...r }));
    if (binding) {
      const storeID = (attr(binding, 'storeItemID') || '').toUpperCase();
      control.binding = { source: doc.pkg?.id, part: doc.keep.stores?.[storeID], storeID,
        xpath: attr(binding, 'xpath'), mappings: attr(binding, 'prefixMappings') || '' };
      const store = ctx.stores?.get(storeID), ns = {};
      try {
        for (const m of control.binding.mappings.matchAll(/xmlns:([\w.-]+)\s*=\s*['"]([^'"]*)['"]/g)) ns[m[1]] = m[2];
        const sd = store?.ownerDocument || store;
        const node = sd?.evaluate(control.binding.xpath, sd, p => ns[p] || null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        control.binding.complex = !node || (node.nodeType === 1 && node.children.length > 0) || (node.nodeType === 3 && node.parentNode.childNodes.length !== 1) || /^\s*<(\?xml|pkg:package|html|w:)/.test(node.textContent || '');
      } catch (error) { control.binding.complex = true; }
    }
    const tp = child(pr, type);
    if (type === 'dropDownList' || type === 'comboBox') control.choices = kids(tp).filter(e => e.localName === 'listItem').map(e => ({ value: attr(e, 'value') || '', text: attr(e, 'displayText') || attr(e, 'value') || '' }));
    if (type === 'date') control.dateFormat = attr(child(tp, 'dateFormat'), 'val');
    if (type === 'checkbox') control.checked = /^(1|true|on)$/.test(attr(child(tp, 'checked'), 'val') || '0');
    doc.keep.controls = true;
    return control;
  };
  P.markControl = function (objects, property, control) {
    if (!control || !objects.length) return;
    const first = objects[0][property], last = objects[objects.length - 1][property];
    (first.sdts || (first.sdts = [])).unshift(control);
    (last.sdte || (last.sdte = [])).push(control.key);
  };
  const properties = b => b.t === 'p' ? b.pPr : b.tblPr;
  const propertyStores = new Set(['{6C3C8BC8-F283-45AE-878A-BAB7291924A1}', '{6668398D-A668-4E3E-A5EB-62B293D839F1}']);
  const plainPr = pr => { const o = { ...pr }; delete o.sdts; delete o.sdte; return o; };
  const typed = c => ['date', 'dropDownList', 'comboBox', 'checkbox'].includes(c.type);
  const controlKinds = { date: 'date picker', dropDownList: 'drop-down list', comboBox: 'combo box', checkbox: 'check box' };
  // Silent unless notify: removals the user made and binding bookkeeping are recorded only
  const controlLoss = (doc, c, what, { writer, notify = false, detail } = {}) => {
    const entry = { id: 'control:' + c.key, what: notify ? 'This ' + (controlKinds[c.type] || 'content control') + ' ' + what : c.name + ': ' + what, detail, where: c.shell.part,
      place: notify && c.name !== 'Content control' ? c.name : undefined, action: 'conversion', notify };
    if (writer) writer.loss(entry);
    else { D.touchKey(doc, 'losses'); K.loss(doc, entry); }
  };
  // A scan contains model references only for this operation. Nothing here is
  // stored in history, and there are no back-links from a fragment to its owner.
  P.controls = function (doc, track = true) {
    const records = new Map(), markers = [], stack = [];
    const feed = (value, text = '') => { if (!stack.length) return; const v = JSON.stringify(L.clone(value)); for (const r of stack) { r.parts.push(v); r.text += text; } };
    const start = (control, ref) => {
      feed(['control', control.key]);
      const r = { control, state: L.clone(control), start: ref, text: '', parts: [], complete: false };
      records.set(control.key, r); stack.push(r); markers.push({ ...ref, key: control.key });
    };
    const end = (key, ref) => {
      markers.push({ ...ref, key });
      const r = records.get(key), i = stack.findIndex(r => r.control.key === key);
      if (r && i >= 0) { r.end = ref; r.complete = i === stack.length - 1 && r.start.scope === ref.scope; stack.splice(i, 1); }
    };
    const marks = (pr, owner, touch, scope, side, pos) => {
      const list = pr[side] || [];
      for (const v of list) {
        const ref = { owner, pr, scope, pos, touch, remove() { touch(); pr[side] = pr[side].filter(x => x !== v); if (!pr[side].length) delete pr[side]; } };
        if (side === 'sdts') start(v, ref); else end(v, ref);
      }
    };
    function sequence(objects, prop, touch, scope) {
      if (!track) touch = () => {};
      for (const o of objects) {
        const ps = o.t === 'p' ? [o] : o.cells ? o.cells.flatMap(c => [D.firstPara(c.blocks), D.lastPara(c.blocks)]) : o.blocks ? [D.firstPara(o.blocks), D.lastPara(o.blocks)] : [D.firstPara([o]), D.lastPara([o])];
        const first = ps.find(Boolean), last = ps.filter(Boolean).at(-1), pr = prop(o);
        marks(pr, o, () => touch(o), scope, 'sdts', first && D.pos(first, 0));
        if (o.t === 'p') {
          feed(['p', plainPr(pr), o.rPr]);
          let offset = 0;
          for (const it of o.runs) {
            const ref = { owner: o, scope: o, pos: D.pos(o, offset), touch: () => { if (track) D.touch(o); }, remove() { if (track) D.touch(o); o.runs = o.runs.filter(x => x !== it); } };
            if (it.t === 'sdts') start(it.control, ref);
            else if (it.t === 'sdte') end(it.key, ref);
            else {
              feed(it, it.rPr?.del ? '' : it.t === 'text' ? it.text : it.t === 'sym' ? it.char || '' : it.t === 'tab' ? '\t' : it.t === 'br' ? '\n' : D.ilen(it) ? '\ufffc' : '');
            }
            offset += D.ilen(it);
          }
          feed(['/p'], '\n');
        } else if (o.t === 'tbl') {
          feed(['tbl', plainPr(pr), o.grid]); sequence(o.rows, r => r.trPr, () => D.touchTbl(o), o.rows); feed(['/tbl']);
        } else if (o.cells) {
          feed(['row', plainPr(pr)]); sequence(o.cells, c => c.tcPr, () => touch(o), o.cells); feed(['/row']);
        } else {
          feed(['cell', plainPr(pr)]); sequence(o.blocks, properties, b => b.t === 'p' ? D.touch(b) : D.touchTbl(b), o.blocks); feed(['/cell']);
        }
        marks(pr, o, () => touch(o), scope, 'sdte', last && D.pos(last, D.plen(last)));
      }
    }
    for (const story of D.stories(doc)) {
      sequence(story.blocks || [], properties, b => b.t === 'p' ? D.touch(b) : D.touchTbl(b), story.blocks);
      // Text boxes form their own control scopes, even inside a body control.
      const boxes = [];
      D.walk(story, b => { if (b.t === 'p') for (const it of b.runs) if (it.t === 'group') boxes.push(...D.groupTextboxes(it)); else if (it.tb) boxes.push(it.tb); });
      stack.length = 0;
      for (const box of new Set(boxes)) { sequence(box.blocks, properties, b => b.t === 'p' ? D.touch(b) : D.touchTbl(b), box.blocks); stack.length = 0; }
    }
    for (const r of records.values()) { r.signature = r.parts.join(''); delete r.parts; if (r.start.scope !== r.start.owner) r.text = r.text.replace(/\n$/, ''); }
    return { records, markers };
  };
  P.removeControl = function (scan, key) { for (const m of scan.markers) if (m.key === key) m.remove(); };
  P.beforeEdit = doc => doc?.keep?.controls ? P.controls(doc) : null;
  P.afterEdit = function (doc, before) {
    if (!before && !doc.keep?.controls) return;
    const after = P.controls(doc);
    for (const [key, old] of before?.records || []) {
      const next = after.records.get(key), c = old.state;
      const deleted = !next?.complete;
      if ((deleted && (c.lock === 'sdtLocked' || c.lock === 'sdtContentLocked')) || ((!next || old.signature !== next.signature || !same(c.value, next.control.value)) && (c.lock === 'contentLocked' || c.lock === 'sdtContentLocked'))) {
        const error = new Error('This modification is not allowed because the content control is locked.'); error.code = 'locked-control'; throw error;
      }
      if (deleted) { P.removeControl(after, key); controlLoss(doc, c, 'the control was removed with its boundary.'); continue; }
      if (old.text === next.text && same(c.value, next.control.value)) continue;
      next.start.touch();
      next.control.edited = true;
      if (typed(c) && same(c.value, next.control.value)) {
        next.control.converted = true; delete next.control.binding;
        controlLoss(doc, c, 'will be saved as ordinary text, because text was typed into it.', { notify: true });
      } else if (next.control.binding) {
        const binding = next.control.binding;
        if (binding.part && binding.xpath && !binding.complex) {
          D.touchKey(doc, 'keep');
          binding.value = typed(c) ? next.control.value : next.text;
          (doc.keep.boundUpdates || (doc.keep.boundUpdates = {}))[binding.storeID + ':' + binding.xpath] = { ...binding };
        } else {
          delete next.control.binding; next.control.unbound = true;
          controlLoss(doc, c, 'has a new value that can\'t be stored with the document\'s data, so Word may show its old value.', { notify: true });
        }
      }
    }
    for (const [key, r] of after.records) if (!r.complete) { P.removeControl(after, key); controlLoss(doc, r.control, 'an incomplete control boundary was removed.'); }
  };
  P.deleteControls = function (doc, a, b) {
    if (!doc.keep?.controls) return;
    const scan = P.controls(doc);
    for (const [key, r] of scan.records) {
      if (!r.complete || !r.start.pos || !r.end.pos || D.storyOf(doc, r.start.pos.p) !== D.storyOf(doc, a.p)) continue;
      const startInside = D.cmp(doc, a, r.start.pos) <= 0 && D.cmp(doc, r.start.pos, b) < 0;
      const endInside = D.cmp(doc, a, r.end.pos) < 0 && D.cmp(doc, r.end.pos, b) <= 0;
      if (startInside || endInside) P.removeControl(scan, key);
    }
  };
  P.duplicateControls = function (blocks) {
    const copy = K.duplicate(blocks), scan = P.controls({ main: { blocks: copy }, hf: {}, fn: {}, en: {}, comments: {} }, false), ids = new Map(), stores = new Map();
    for (const [key, r] of scan.records) {
      if (!r.complete) { P.removeControl(scan, key); continue; }
      ids.set(key, D.nid()); r.control.key = ids.get(key);
      const b = r.control.binding;
      if (b && b.source !== D.doc?.pkg?.id) {
        if (!b.part || propertyStores.has(b.storeID)) {
          delete r.control.binding; r.control.unbound = true;
          if (D.doc) controlLoss(D.doc, r.control, !b.part ? 'the copied XML binding had no available store.' : 'copying to another document removed its document-property binding.');
        } else {
          const storeKey = b.source + ':' + b.part;
          if (!stores.has(storeKey)) {
            let id = D.doc?.keep?.storeCopies?.[storeKey];
            if (!id) {
              id = '{' + crypto.randomUUID().toUpperCase() + '}';
              if (D.doc) { D.touchKey(D.doc, 'keep'); D.doc.keep = D.doc.keep || {}; (D.doc.keep.storeCopies || (D.doc.keep.storeCopies = {}))[storeKey] = id; }
            }
            stores.set(storeKey, id);
          }
          b.originalStoreID = b.storeID; b.storeID = stores.get(storeKey);
        }
      }
    }
    for (const m of scan.markers) if (!ids.has(m.key)) m.remove();
    // Incomplete copied selections keep the text, never a dangling wrapper.
    const visit = value => {
      if (!value || typeof value !== 'object') return;
      if (value.t === 'sdte') value.key = ids.get(value.key) || value.key;
      if (Array.isArray(value.sdte)) value.sdte = value.sdte.filter(k => ids.has(k)).map(k => ids.get(k));
      for (const v of Object.values(value)) visit(v);
    };
    visit(copy);
    const bookmarks = new Map(), paragraphs = [];
    D.walk({ blocks: copy }, b => {
      if (b.t !== 'p') return;
      paragraphs.push(b);
      for (const it of b.runs) if (['bs', 'be', 'perm'].includes(it.t)) {
        const key = (it.t === 'perm' ? 'perm:' : 'bm:') + it.id; if (!bookmarks.has(key)) bookmarks.set(key, {});
        bookmarks.get(key)[it.t === 'be' || it.end ? 'be' : 'bs'] = it;
      }
    });
    const names = new Map();
    for (const pair of bookmarks.values()) if (pair.bs && pair.be) {
      pair.id = 'copy' + D.nid();
      if (pair.bs.t === 'bs') {
        const name = (pair.bs.name || 'Bookmark').slice(0, 24) + '_' + pair.id;
        names.set(pair.bs.name, name); pair.bs.name = name;
      }
    }
    for (const p of paragraphs) p.runs = p.runs.filter(it => {
      if (['bs', 'be', 'perm'].includes(it.t)) {
        const pair = bookmarks.get((it.t === 'perm' ? 'perm:' : 'bm:') + it.id); if (!pair.id) return false; it.id = pair.id;
      }
      if (it.rPr?.link?.anchor && names.has(it.rPr.link.anchor)) it.rPr.link.anchor = names.get(it.rPr.link.anchor);
      return true;
    });
    return copy;
  };
  P.copyParagraphProperties = function (p, from, to) {
    const pr = L.clone(p.pPr);
    if (from > 0) delete pr.sdts;
    if (to < D.plen(p)) delete pr.sdte;
    return pr;
  };
  P.splitControls = function (doc, p, next) {
    delete next.pPr.sdts;
    if (p.pPr.sdte) { next.pPr.sdte = p.pPr.sdte; delete p.pPr.sdte; }
    const open = [], offsets = new Map(); let offset = 0;
    for (const it of p.runs) {
      if (it.t === 'sdts') { open.push(it.control); offsets.set(it.control.key, offset); }
      else if (it.t === 'sdte') { const i = open.findIndex(c => c.key === it.key); if (i >= 0) open.splice(i, 1); }
      offset += D.ilen(it);
    }
    if (!open.length) return;
    const ends = new Map(); offset = 0;
    for (const it of next.runs) { if (it.t === 'sdte') ends.set(it.key, offset); offset += D.ilen(it); }
    if (open.every(c => offsets.get(c.key) === 0 && ends.get(c.key) === offset)) {
      // A control covering the whole paragraph becomes a block control when
      // Enter gives it a second paragraph. Its identity and binding stay intact.
      const keys = new Set(open.map(c => c.key));
      p.runs = p.runs.filter(it => it.t !== 'sdts' || !keys.has(it.control.key));
      next.runs = next.runs.filter(it => it.t !== 'sdte' || !keys.has(it.key));
      p.pPr.sdts = [...(p.pPr.sdts || []), ...open];
      next.pPr.sdte = [...open.map(c => c.key).reverse(), ...(next.pPr.sdte || [])];
    } else {
      // Word cannot put half of two paragraphs inside one run-level sdt. Keep
      // both editable halves as controls and report the necessary unbinding.
      const starts = [];
      for (const c of open) {
        const copy = K.duplicate(c); copy.key = D.nid();
        for (const control of [c, copy]) { delete control.binding; control.unbound = true; if (typed(control)) control.converted = true; }
        for (const it of next.runs) if (it.t === 'sdte' && it.key === c.key) it.key = copy.key;
        starts.push(D.item('sdts', { control: copy }));
        controlLoss(doc, c, 'splitting its paragraph created separate unbound controls.');
      }
      p.runs.push(...open.map(c => D.item('sdte', { key: c.key })).reverse());
      next.runs.unshift(...starts);
    }
  };
  P.inlineControls = function (p) {
    if (!p.pPr.sdts?.length || !p.pPr.sdte?.length) return;
    const closed = new Set(p.pPr.sdte), controls = p.pPr.sdts.filter(c => closed.has(c.key));
    if (!controls.length) return;
    const keys = new Set(controls.map(c => c.key));
    p.runs.unshift(...controls.map(control => D.item('sdts', { control })));
    p.runs.push(...p.pPr.sdte.filter(key => keys.has(key)).map(key => D.item('sdte', { key })));
    p.pPr.sdts = p.pPr.sdts.filter(c => !keys.has(c.key)); p.pPr.sdte = p.pPr.sdte.filter(k => !keys.has(k));
    if (!p.pPr.sdts.length) delete p.pPr.sdts; if (!p.pPr.sdte.length) delete p.pPr.sdte;
  };
  P.joinControls = function (doc, p, q) {
    if (!doc.keep?.controls) return;
    P.inlineControls(p); P.inlineControls(q);
    const scan = P.controls(doc);
    for (const [key, r] of scan.records) if (r.start.pr && ((r.end?.owner === p) || (r.start.owner === q))) {
      P.removeControl(scan, key); controlLoss(doc, r.control, 'will be saved as ordinary text, because paragraphs were joined across it.', { notify: true });
    }
    if (q.pPr.sdte) p.pPr.sdte = [...(p.pPr.sdte || []), ...q.pPr.sdte];
    delete q.pPr.sdte;
  };
  P.controlAt = function (pos) {
    if (!pos || !D.doc?.keep?.controls) return null;
    let found = null;
    for (const r of P.controls(D.doc).records.values()) if (r.complete && r.start.pos && r.end.pos && D.storyOf(D.doc, r.start.pos.p) === D.storyOf(D.doc, pos.p) && D.cmp(D.doc, r.start.pos, pos) <= 0 && D.cmp(D.doc, pos, r.end.pos) <= 0) found = r;
    return found;
  };
  P.setControlValue = function (key, value) {
    const r = P.controls(D.doc).records.get(key);
    if (!r?.complete || !typed(r.control) || r.control.converted || r.start.pos?.p !== r.end.pos?.p) return false;
    const c = r.control, p = r.start.pos.p;
    r.start.touch(); D.touch(p);
    let display = String(value);
    if (c.type === 'checkbox') {
      value = value === true || value === 'true' ? 'true' : 'false'; c.checked = value === 'true';
      const tp = child(child(K.parse(c.shell.xml), 'sdtPr'), 'checkbox'), state = child(tp, c.checked ? 'checkedState' : 'uncheckedState');
      display = state ? String.fromCodePoint(parseInt(attr(state, 'val'), 16)) : c.checked ? '\u2612' : '\u2610';
    } else if (c.choices) display = c.choices.find(v => v.value === value)?.text ?? String(value);
    else if (c.type === 'date') {
      const date = new Date(String(value).slice(0, 10) + 'T12:00:00');
      if (!Number.isFinite(date.getTime())) return false;
      value = String(value).slice(0, 10) + 'T00:00:00Z';
      if (c.dateFormat && L.fields?.formatDate) display = L.fields.formatDate(date, c.dateFormat);
    }
    c.value = String(value);
    const from = r.start.pos.o, to = r.end.pos.o;
    D.splitAt(p, from); D.splitAt(p, to);
    let start, end;
    if (r.start.pr) { start = 0; end = p.runs.length; }
    else { start = p.runs.findIndex(it => it.t === 'sdts' && it.control.key === key) + 1; end = p.runs.findIndex(it => it.t === 'sdte' && it.key === key); }
    const format = p.runs.slice(start, end).find(it => D.ilen(it))?.rPr || {};
    p.runs.splice(start, end - start, D.text(display, L.clone(format)));
    return D.pos(p, from + display.length);
  };
  P.controlDialog = function (record) {
    const c = record?.control;
    if (!c || !typed(c) || c.converted) return;
    const apply = value => L.ed.edit('Content Control Value', () => P.setControlValue(c.key, value));
    if (c.type === 'checkbox') { apply(!c.checked); return; }
    const input = c.type === 'dropDownList' ? L.h('select', {}, ...(c.choices || []).map(v => L.h('option', { value: v.value, selected: c.value === v.value || record.text === v.text }, v.text)))
      : L.h('input', { type: c.type === 'date' ? 'date' : 'text', value: c.type === 'date' ? (c.value || attr(child(child(K.parse(c.shell.xml), 'sdtPr'), 'date'), 'fullDate') || '').slice(0, 10) : c.value ?? record.text });
    L.ui.dialog({ title: c.name, width: 340, body: L.h('label', { class: 'col' }, 'Value:', input), buttons: [
      { label: 'OK', primary: true, onClick: () => { if (c.type === 'date' && !input.value) { input.focus(); return false; } apply(input.value); } }, { label: 'Cancel' } ] });
  };
  P.clipboard = function (blocks) {
    const refs = [];
    const visit = value => {
      if (!value || typeof value !== 'object') return;
      if (value.source && value.part && value.storeID) refs.push({ source: value.source, part: value.part });
      if (Array.isArray(value.related)) for (const r of value.related) if (!r.external && r.part) refs.push({ source: r.source, part: r.part });
      if (Array.isArray(value.previews)) for (const r of value.previews) refs.push({ source: r.source, part: r.part });
      if (Array.isArray(value.deps)) for (const d of value.deps) if (!d.external && d.part) refs.push({ source: d.source, part: d.part });
      for (const v of Object.values(value)) visit(v);
    };
    visit(blocks);
    try { return K.export(refs); } catch (error) { return null; }
  };
  P.controlXML = function (c, content, ctx) {
    let xml = c.shell.xml;
    const tree = K.parse(xml), pr = child(tree, 'sdtPr'), changes = [];
    for (const e of kids(pr)) {
      const remove = (c.converted && ['date', 'dropDownList', 'comboBox', 'checkbox', 'text', 'dataBinding', 'showingPlcHdr'].includes(e.localName)) || (c.unbound && e.localName === 'dataBinding') || (c.edited && e.localName === 'showingPlcHdr');
      if (remove) { const p = location(e); changes.push({ start: p.start, end: p.end, value: '' }); }
    }
    xml = K.patch(xml, changes);
    if (c.binding?.originalStoreID) {
      const binding = child(child(K.parse(xml), 'sdtPr'), 'dataBinding');
      if (binding) xml = setAttrs(xml, binding, { [binding.prefix + ':storeItemID']: c.binding.storeID });
    }
    if (c.value !== undefined && !c.converted) {
      const tp = child(child(K.parse(xml), 'sdtPr'), c.type);
      if (tp && c.type === 'date') xml = setAttrs(xml, tp, { [tp.prefix + ':fullDate']: c.value });
      if (tp && (c.type === 'dropDownList' || c.type === 'comboBox')) xml = setAttrs(xml, tp, { [tp.prefix + ':lastValue']: c.value });
      if (tp && c.type === 'checkbox') {
        const checked = child(tp, 'checked');
        if (checked) xml = setAttrs(xml, checked, { [checked.prefix + ':val']: c.value === 'true' ? '1' : '0' });
      }
    }
    try {
      for (const dependency of c.related || []) if (dependency.source !== ctx.writer.pkg?.id) ctx.writer.keepRel(ctx.mainPart, dependency);
      // Data bindings refer to a document-wide store ID. Headers and footers
      // cannot own CustomXmlPart relationships; attach the store to the main part.
      if (c.binding?.part && !propertyStores.has(c.binding.storeID)) ctx.writer.keepRel(ctx.mainPart, { source: c.binding.source, owner: c.shell.part, type: RT('customXml'), part: c.binding.part, target: K.relative(c.shell.part, c.binding.part) });
      return ctx.writer.emit(K.slice(c.shell, xml), ctx.part).replace(SLOT, () => content);
    }
    catch (error) { controlLoss(ctx.doc, c, 'couldn\'t be kept and will be saved as ordinary text.', { writer: ctx.writer, notify: !K.sourceMissing(error), detail: error.message }); return content; }
  };
  P.controlSequence = function (objects, prop, render, ctx) {
    const stack = [{ xml: '' }];
    for (const o of objects) {
      const pr = prop(o);
      for (const c of pr?.sdts || []) stack.push({ xml: '', control: c });
      stack.at(-1).xml += render(o);
      for (const key of pr?.sdte || []) {
        const current = stack.at(-1);
        if (current.control?.key !== key) continue;
        stack.pop(); stack.at(-1).xml += P.controlXML(current.control, current.xml, ctx);
      }
    }
    while (stack.length > 1) { const r = stack.pop(); controlLoss(ctx.doc, r.control, 'an incomplete control boundary was removed.', { writer: ctx.writer }); stack.at(-1).xml += r.xml; }
    return stack[0].xml;
  };
  P.writeBindings = function (doc, writer) {
    const updates = {};
    if (doc.keep?.controls) for (const r of P.controls(doc).records.values()) {
      const b = r.control.binding;
      if (b?.value !== undefined) updates[b.storeID + ':' + b.xpath] = b;
      if (b?.originalStoreID) {
        const source = K.package(b.source), props = source?.rels(b.part).find(r => canonical(r.type) === RT('customXmlProps'))?.part;
        if (props) {
          const tree = source.xml(props), id = Array.from(tree.attributes).find(a => a.localName === 'itemID');
          if (id) {
            const dest = writer.claim(props, 'merged', writer.target(source, props), source);
            writer.put(dest, K.partXML(source.text(props), K.attributes(K.raw(tree), { [id.name]: b.storeID })), source.type(props));
            writer.carryRels(source, props, dest);
          }
        }
      }
    }
    Object.assign(updates, doc.keep?.boundUpdates);
    for (const update of Object.values(updates)) {
      const source = K.package(update.source) || doc.pkg, part = update.part;
      try {
        if (!source?.has(part)) throw new Error('The XML store is missing');
        const dest = writer.claim(part, 'merged', writer.target(source, part), source);
        const raw = writer.parts.get(dest) || source.text(part);
        const xml = typeof raw === 'string' ? raw : L.xmlTree.decode(raw);
        const dom = new DOMParser().parseFromString(xml, 'application/xml'), ns = {};
        for (const m of update.mappings.matchAll(/xmlns:([\w.-]+)\s*=\s*['"]([^'"]*)['"]/g)) ns[m[1]] = m[2];
        let node = dom.evaluate(update.xpath, dom, p => ns[p] || null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (!node) throw new Error('The binding XPath has no target');
        if (node.nodeType === 3 && node.parentNode.childNodes.length === 1) node = node.parentNode;
        const element = node.nodeType === 2 ? node.ownerElement : node, path = [];
        if (element.nodeType !== 1) throw new Error('The binding target is not a scalar value');
        for (let e = element; e.parentElement; e = e.parentElement) path.unshift(Array.from(e.parentElement.children).indexOf(e));
        let target = K.parse(xml); for (const i of path) target = kids(target)[i];
        let output;
        if (node.nodeType === 2) {
          const a = location(target).attrs.find(a => a.name === node.name);
          output = K.patch(xml, [{ start: a.start, end: a.end, value: K.esc(String(update.value)).replace(/'/g, '&apos;') }]);
        } else output = inner(xml, target, K.esc(String(update.value)));
        writer.put(dest, output, source.type(part)); writer.carryRels(source, part, dest);
      } catch (error) {
        writer.loss({ id: 'binding:' + update.storeID + ':' + update.xpath, what: 'A content control has a new value that can\'t be stored with the document\'s data, so Word may show its old value.', detail: error.message, where: part || update.storeID, action: 'conversion', notify: true });
      }
    }
  };
})();
