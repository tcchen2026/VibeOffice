/* VibeOffice — lossless Open Packaging Conventions carry.
 * Immutable source packages stay outside models/history. Models hold JSON fragments
 * and source identities; writers own reservations, graph traversal and loss reports.
 * Requires common/xml.js; works in Node without a browser DOM.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {}), XML = L.xmlTree;
  const K = (L.opc = {});
  const NS = K.NS = Object.freeze({
    rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    strictRel: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
    pkg: 'http://schemas.openxmlformats.org/package/2006/relationships',
    ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    xml: 'http://www.w3.org/XML/1998/namespace', xmlns: 'http://www.w3.org/2000/xmlns/',
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  });
  K.KNOWN_NS = Object.freeze({
    a: NS.a, r: NS.rel, w: NS.w, p: NS.p, mc: NS.mc,
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    a14: 'http://schemas.microsoft.com/office/drawing/2010/main', a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
    x14: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', xm: 'http://schemas.microsoft.com/office/excel/2006/main',
    x14ac: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac', c14: 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
    c16: 'http://schemas.microsoft.com/office/drawing/2014/chart', v: 'urn:schemas-microsoft-com:vml',
    o: 'urn:schemas-microsoft-com:office:office', x: 'urn:schemas-microsoft-com:office:excel',
  });
  // Main-part types distinguish templates and macro/slideshow variants even when
  // a file was renamed. These are file-format facts, independent of an app UI.
  K.formats = Object.freeze(Object.fromEntries([
    ['docx', 'wordprocessingml.document', 'ms-word.document', false, false],
    ['docm', 'wordprocessingml.document', 'ms-word.document', true, false],
    ['dotx', 'wordprocessingml.template', 'ms-word.template', false, true],
    ['dotm', 'wordprocessingml.template', 'ms-word.template', true, true],
    ['xlsx', 'spreadsheetml.sheet', 'ms-excel.sheet', false, false],
    ['xlsm', 'spreadsheetml.sheet', 'ms-excel.sheet', true, false],
    ['xltx', 'spreadsheetml.template', 'ms-excel.template', false, true],
    ['xltm', 'spreadsheetml.template', 'ms-excel.template', true, true],
    ['pptx', 'presentationml.presentation', 'ms-powerpoint.presentation', false, false],
    ['pptm', 'presentationml.presentation', 'ms-powerpoint.presentation', true, false],
    ['ppsx', 'presentationml.slideshow', 'ms-powerpoint.slideshow', false, false],
    ['ppsm', 'presentationml.slideshow', 'ms-powerpoint.slideshow', true, false],
    ['potx', 'presentationml.template', 'ms-powerpoint.template', false, true],
    ['potm', 'presentationml.template', 'ms-powerpoint.template', true, true],
  ].map(([ext, standard, enabled, macro, template]) => {
    const stem = macro ? 'application/vnd.' + enabled : 'application/vnd.openxmlformats-officedocument.' + standard;
    return [ext, Object.freeze({ ext, macro, template, family: standard.split('.')[0],
      contentType: stem + (macro ? ext === 'dotm' ? '.macroEnabledTemplate.main+xml' : '.macroEnabled.main+xml' : '.main+xml'),
      mime: stem + (macro ? '.macroEnabled.12' : '') })];
  })));
  K.variant = (pkg, fallback) => Object.values(K.formats).find(f => f.contentType === pkg?.type(pkg.main))?.ext || fallback;
  K.relationshipType = type => {
    if (type === NS.strictRel + '/extendedProperties') return NS.rel + '/extended-properties';
    if (type === NS.strictRel + '/customProperties') return NS.rel + '/custom-properties';
    return type?.replace(NS.strictRel + '/', NS.rel + '/');
  };
  const strictURI = uri => uri.replace(/^http:\/\/schemas\.openxmlformats\.org\/(officeDocument|wordprocessingml|drawingml|spreadsheetml|presentationml)\/2006\/(.*)$/, (_, family, rest) =>
    'http://purl.oclc.org/ooxml/' + family + '/' + rest.replace(/(^|\/)(extended-properties|custom-properties)$/, (m, slash, name) => slash + (name === 'extended-properties' ? 'extendedProperties' : 'customProperties')));
  K.strictXML = xml => xml.replace(/(\b(?:xmlns(?::[\w.-]+)?|uri|Type)\s*=\s*)(["'])(.*?)\2/g, (_, key, quote, uri) => key + quote + strictURI(uri) + quote);
  K.format = (doc, requested, fallback) => {
    const type = requested || doc.ooxmlFormat || K.variant(doc.pkg, fallback);
    if (!K.formats[type]) throw new Error('Unsupported Office file variant: ' + type);
    return K.formats[type];
  };
  const esc = K.esc = XML.esc;
  const uuid = () => root.crypto?.randomUUID ? root.crypto.randomUUID() : Date.now().toString(36) + '-' + Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
  const elements = el => Array.from(el.children || []);
  const walk = (el, fn) => { if (el.nodeType === 1) fn(el); for (const c of elements(el)) walk(c, fn); };
  const attrs = el => Array.from(el.attributes || []);
  const ans = (el, a) => a.namespaceURI || (a.prefix ? a.prefix === 'xmlns' ? NS.xmlns : el.lookupNamespaceURI(a.prefix) : a.name === 'xmlns' ? NS.xmlns : '');
  const relAttr = (el, a) => [NS.rel, NS.strictRel].includes(ans(el, a)) || ans(el, a) === K.KNOWN_NS.o && a.localName === 'relid';
  K.parse = text => XML.parse(text, { keepWhitespace: true, source: true, strict: true });
  K.patch = function (text, changes) {
    let end = text.length;
    for (const c of changes.slice().sort((a, b) => b.start - a.start || b.end - a.end)) {
      if (c.start < 0 || c.end < c.start || c.end > end) throw new Error('Overlapping XML edits');
      text = text.slice(0, c.start) + c.value + text.slice(c.end); end = c.start;
    }
    return text;
  };
  K.partXML = function (original, root) {
    const before = XML.source.get(K.parse(original)), after = XML.source.get(K.parse(root));
    return original.slice(0, before.start) + root.slice(after.start, after.end) + original.slice(before.end);
  };
  // Carry all in-scope bindings, including prefixes that appear only inside values
  // (Requires, xsi:type, vendor attributes). Inner redeclarations keep their scope.
  function inherited(el) {
    const chain = []; for (let e = el; e && e.nodeType === 1; e = e.parentNode) chain.unshift(e);
    const bindings = {}, contextual = {}, lists = new Map();
    for (const e of chain) for (const a of attrs(e)) {
      if (a.name === 'xmlns' || a.prefix === 'xmlns') bindings[a.name] = a.value;
      else if (ans(e, a) === NS.xml) contextual[a.name] = a.value;
      else if (ans(e, a) === NS.mc) {
        const key = a.name;
        if (!lists.has(key)) lists.set(key, new Set());
        for (const v of a.value.split(/\s+/).filter(Boolean)) lists.get(key).add(v);
      }
    }
    for (const [key, values] of lists) contextual[key] = Array.from(values).join(' ');
    // Detached lightweight fragments from older models may lack their ancestor.
    walk(el, e => {
      const prefixes = [e.prefix];
      for (const a of attrs(e)) {
        if (a.prefix && a.prefix !== 'xmlns') prefixes.push(a.prefix);
        if (ans(e, a) === NS.mc || a.localName === 'Requires' && e.namespaceURI === NS.mc || a.name === 'xsi:type') {
          for (const value of a.value.split(/\s+/)) prefixes.push(value.split(':')[0]);
        }
      }
      for (const prefix of prefixes) if (prefix && prefix !== 'xml' && !e.lookupNamespaceURI(prefix) && K.KNOWN_NS[prefix]) bindings['xmlns:' + prefix] = K.KNOWN_NS[prefix];
    });
    return { ...bindings, ...contextual };
  }
  K.serialize = function (el) {
    if (!el) return '';
    const context = inherited(el);
    function write(e, first) {
      if (e.nodeType === 3) return (e.nodeValue ?? e.data).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      if (e.nodeType === 4) return '<![CDATA[' + e.data + ']]>';
      if (e.nodeType === 8) return '<!--' + e.data + '-->';
      if (e.nodeType === 7) return '<?' + e.target + ' ' + e.data + '?>';
      if (e.nodeType !== 1) return Array.from(e.childNodes || []).map(n => write(n, true)).join('');
      const values = Object.fromEntries(attrs(e).map(a => [a.name, a.value]));
      if (first) for (const [key, value] of Object.entries(context)) {
        if (!(key in values)) values[key] = value;
        else if (listsAttribute(e, key)) values[key] = Array.from(new Set((value + ' ' + values[key]).split(/\s+/))).join(' ');
      }
      const tag = '<' + e.nodeName + Object.entries(values).map(([k, v]) => ' ' + k + '="' + esc(v) + '"').join('');
      return e.childNodes.length ? tag + '>' + Array.from(e.childNodes).map(n => write(n, false)).join('') + '</' + e.nodeName + '>' : tag + '/>';
    }
    return write(el, true);
  };
  function listsAttribute(el, name) { const prefix = name.split(':')[0]; return name.includes(':') && el.lookupNamespaceURI(prefix) === NS.mc; }
  K.raw = function (el) {
    const source = XML.source.get(el);
    if (!source || source.end == null) return K.serialize(el);
    const changes = [], added = [];
    for (const [name, value] of Object.entries(inherited(el))) {
      const attr = source.attrs.find(a => a.name === name);
      if (!attr) added.push(' ' + name + '="' + esc(value) + '"');
      else if (listsAttribute(el, name) && el.getAttribute(name) !== value) changes.push({ start: attr.start - source.start, end: attr.end - source.start, value: esc(value) });
    }
    let pos = source.openEnd - source.start - 1;
    if (source.text[source.start + pos - 1] === '/') pos--;
    if (added.length) changes.push({ start: pos, end: pos, value: added.join('') });
    return K.patch(source.text.slice(source.start, source.end), changes);
  };

  K.relsPath = part => { const i = part.lastIndexOf('/'); return part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels'; };
  K.owner = name => name === '_rels/.rels' ? '' : name.replace(/(^|\/)\_rels\/([^/]+)\.rels$/, '$1$2');
  K.resolve = function (owner, uri) {
    if (/^[a-z][a-z\d+.-]*:|^\/\/|\\/i.test(uri)) throw new Error('Invalid internal part URI: ' + uri);
    const hash = uri.indexOf('#'), fragment = hash < 0 ? '' : uri.slice(hash);
    const target = decodeURIComponent(hash < 0 ? uri : uri.slice(0, hash));
    const path = target ? target.startsWith('/') ? [] : owner.split('/').slice(0, -1) : owner.split('/');
    for (const segment of target.split('/')) {
      if (!segment || segment === '.') continue;
      if (segment === '..') { if (!path.length) throw new Error('Part URI escapes package'); path.pop(); }
      else path.push(segment);
    }
    return { part: path.join('/'), fragment };
  };
  K.relative = function (owner, target) {
    const from = owner.split('/').slice(0, -1), to = target.split('/');
    while (from.length && to.length && from[0] === to[0]) { from.shift(); to.shift(); }
    return '../'.repeat(from.length) + to.map(s => encodeURIComponent(s).replace(/%3A/gi, ':')).join('/');
  };
  const escapedURI = value => value.replace(/[\u0000-\u0020\u007f-\uffff]+/g, s => encodeURI(s));
  const packages = new Map(), privateData = new WeakMap();
  K.package = id => packages.get(id);
  K.open = async function (zip, options = {}) {
    if (!(zip instanceof Map)) zip = await L.zip.read(zip);
    const bytes = new Map();
    for (const entry of new Set(zip.values())) if (!entry.name.endsWith('/')) {
      const name = decodeURIComponent(entry.name);
      if (bytes.has(name)) throw new Error('Duplicate package part URI: ' + name);
      bytes.set(name, (await entry.bytes()).slice());
    }
    return makePackage(bytes, options);
  };
  function makePackage(bytes, options = {}) {
    const rels = new Map(), types = new Map(), parsed = new Map(), problems = [];
    const parse = name => {
      if (!parsed.has(name)) {
        try { parsed.set(name, bytes.has(name) ? K.parse(XML.decode(bytes.get(name))) : null); }
        catch (e) { problems.push(Object.freeze({ part: name, why: e.message })); parsed.set(name, null); }
      }
      return parsed.get(name);
    };
    const defaults = new Map(), overrides = new Map();
    for (const el of elements(parse('[Content_Types].xml') || {})) {
      if (el.localName === 'Default') defaults.set((el.getAttribute('Extension') || '').toLowerCase(), el.getAttribute('ContentType'));
      if (el.localName === 'Override') overrides.set(decodeURIComponent((el.getAttribute('PartName') || '').replace(/^\//, '')), el.getAttribute('ContentType'));
    }
    for (const name of bytes.keys()) {
      types.set(name, overrides.get(name) || defaults.get(name.split('.').pop().toLowerCase()) || null);
      if (!name.endsWith('.rels')) continue;
      const owner = K.owner(name), list = [];
      for (const el of elements(parse(name) || {})) {
        if (el.localName !== 'Relationship' || el.namespaceURI !== NS.pkg) continue;
        const external = el.getAttribute('TargetMode') === 'External', target = el.getAttribute('Target') || '';
        try {
          const dest = external ? {} : K.resolve(owner, target);
          list.push(Object.freeze({ id: el.getAttribute('Id'), type: el.getAttribute('Type'), target, external, ...dest }));
        } catch (e) { problems.push(Object.freeze({ part: name, why: e.message })); }
      }
      rels.set(owner, Object.freeze(list));
    }
    const rootRel = (rels.get('') || []).find(r => r.type === NS.rel + '/officeDocument' || r.type === NS.strictRel + '/officeDocument');
    const pkg = Object.freeze({
      id: options.id || uuid(), names: Object.freeze(Array.from(bytes.keys())), main: rootRel?.part || null,
      has: name => bytes.has(name), type: name => types.get(name) || null,
      rels: owner => rels.get(owner) || Object.freeze([]),
      bytes: name => bytes.get(name)?.slice(), text: name => bytes.has(name) ? XML.decode(bytes.get(name)) : null,
      problems: () => problems.slice(),
      // Give readers independent trees. A reader may normalise its own DOM freely.
      xml: name => bytes.has(name) ? K.parse(XML.decode(bytes.get(name))) : null,
    });
    if (packages.has(pkg.id)) throw new Error('Duplicate source package identity');
    privateData.set(pkg, { bytes, rels, types, parse }); packages.set(pkg.id, pkg);
    return pkg;
  }
  K.attach = function (doc, pkg) { Object.defineProperty(doc, 'pkg', { value: pkg, configurable: true, writable: true, enumerable: false }); return doc; };
  K.subtree = function (pkg, start) {
    const names = new Set(), pending = [start];
    while (pending.length) {
      const part = pending.pop(); if (names.has(part)) continue;
      if (!pkg.has(part)) throw new Error('Missing dependency: ' + part);
      names.add(part);
      for (const r of pkg.rels(part)) if (!r.external) pending.push(r.part);
    }
    return { source: pkg.id, part: start, parts: Array.from(names) };
  };

  class Identities {
    constructor() { this.spaces = new Map(); this.mapping = new Map(); this.copies = new Map(); }
    canonical(id) { return /^[+-]?\d+$/.test(String(id)) && Number.isSafeInteger(+id) ? String(+id) : String(id); }
    space(scope, kind) { const key = JSON.stringify([scope, kind]); if (!this.spaces.has(key)) this.spaces.set(key, { used: new Set(), high: 0 }); return this.spaces.get(key); }
    reserve(scope, kind, id) { const s = this.space(scope, kind), value = this.canonical(id); s.used.add(value); if (/^\d+$/.test(value)) s.high = Math.max(s.high, +value); return String(id); }
    fresh(scope, kind) {
      const s = this.space(scope, kind), limits = K.schema?.limits[kind] || { min: 0, max: Number.MAX_SAFE_INTEGER };
      let candidate = Math.max(limits.min, s.high + 1);
      // Bounded OOXML integer spaces can end at the original maximum. In that
      // case use a free legal value, never overflow and never rebind an old ID.
      if (candidate > limits.max) candidate = Math.max(1, limits.min);
      while (s.used.has(String(candidate))) candidate++;
      if (candidate > limits.max) throw new Error('No available ' + kind + ' identity');
      return this.reserve(scope, kind, String(candidate));
    }
    key(source, scope, kind, id) { return JSON.stringify([source, scope, kind, this.canonical(id)]); }
    copy(instance, definitions) { this.copies.set(instance, new Set(definitions.map(d => this.key(d.source, d.scope, d.kind, d.id)))); }
    bind(source, scope, kind, id, target, destination = scope) {
      this.mapping.set(JSON.stringify([this.key(source, scope, kind, id), destination, '']), this.reserve(destination, kind, target));
    }
    resolve(source, scope, kind, id, options = {}) {
      const original = this.key(source, scope, kind, id), dest = options.scope || scope;
      const copied = options.copy && this.copies.get(options.copy)?.has(original);
      const key = JSON.stringify([original, dest, copied ? options.copy : '']);
      if (this.mapping.has(key)) return this.mapping.get(key);
      if (options.primary === source && dest === scope && !copied) return this.reserve(dest, kind, id);
      if (!this.mapping.has(key)) this.mapping.set(key, this.fresh(dest, kind));
      return this.mapping.get(key);
    }
  }
  K.Identities = Identities;
  function identity(el, a, part) {
    const tag = el.localName, name = a.localName, ns = el.namespaceURI || '';
    const isW = ns === NS.w || ns === 'http://purl.oclc.org/ooxml/wordprocessingml/main';
    const isP = ns === NS.p || ns === 'http://purl.oclc.org/ooxml/presentationml/main';
    const isS = ns === NS.s || ns === 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    let kind, scope = part, definition = false;
    if (tag === 'docPr' && name === 'id' && ns.includes('wordprocessingDrawing')) { kind = 'docPr'; scope = 'document'; definition = true; }
    else if (tag === 'cNvPr' && name === 'id') { kind = 'shape'; definition = true; }
    else if (isP && tag === 'cTn' && name === 'id') { kind = 'timing'; definition = true; }
    else if (isP && tag === 'tn' && name === 'val') kind = 'timing';
    else if ((isP && name === 'spid') || ['stCxn', 'endCxn'].includes(tag) && name === 'id') kind = 'shape';
    else if (isP && name === 'grpId' && ['cTn', 'bldP', 'bldDgm', 'bldOleChart', 'bldGraphic'].includes(tag)) { kind = 'build'; definition = tag !== 'cTn'; }
    else if (isW && name === 'id' && /^(bookmark|perm)(Start|End)$/.test(tag)) { kind = tag.startsWith('bookmark') ? 'bookmark' : 'permission'; definition = tag.endsWith('Start'); }
    else if (isP && name === 'id' && ['sldId', 'sldMasterId', 'sldLayoutId'].includes(tag)) { kind = tag; scope = 'presentation'; definition = true; }
    else if (isS && name === 'sheetId') { kind = 'sheet'; scope = 'workbook'; definition = tag === 'sheet'; }
    else if (isS && name === 'cacheId') { kind = 'cache'; scope = 'workbook'; definition = tag === 'pivotCache'; }
    else if (isW && name === 'val' && ['numId', 'abstractNumId'].includes(tag)) { kind = tag; scope = 'numbering'; }
    else if (isW && ((tag === 'num' && name === 'numId') || (tag === 'abstractNum' && name === 'abstractNumId'))) { kind = name; scope = 'numbering'; definition = true; }
    else if (isW && tag === 'id' && name === 'val' && el.parentNode?.localName === 'sdtPr') { kind = 'control'; scope = 'document'; definition = true; }
    else if (isS && name === 'dxfId') { kind = 'dxf'; scope = 'styles'; }
    return kind ? { kind, scope, id: a.value, definition } : null;
  }
  K.identityAttribute = identity;
  const alternate = new WeakMap(), records = new WeakMap();
  K.alternate = el => alternate.get(el);
  K.alternates = el => records.get(el) || [];
  K.captureAC = function (dom, text) {
    if (records.has(dom)) return records.get(dom);
    const original = dom.nodeName === '#document' ? elements(dom)[0] : dom.documentElement || dom, found = [];
    if (!original) return found;
    let hasAC = false; walk(original, e => { if (e.namespaceURI === NS.mc && e.localName === 'AlternateContent') hasAC = true; });
    if (!hasAC) { records.set(dom, Object.freeze(found)); return found; }
    const raw = text ? K.parse(text) : null;
    function visit(node, source, owner) {
      if (node.namespaceURI === NS.mc && node.localName === 'AlternateContent' && !owner) {
        owner = Object.freeze({ id: uuid(), xml: K.raw(source || node) }); found.push(owner);
      }
      if (owner) alternate.set(node, owner);
      const children = elements(node), rawChildren = source ? elements(source) : [];
      children.forEach((child, i) => visit(child, rawChildren[i], owner));
    }
    visit(original, raw, null); Object.freeze(found); records.set(dom, found); records.set(original, found); return found;
  };
  K.fragment = function (el, context = {}) {
    const ac = alternate.get(el), xml = ac ? ac.xml : K.raw(el);
    const source = context.pkg?.id || context.source, part = context.part || '';
    const parsed = K.parse(xml), deps = [], ids = [], changes = [];
    walk(parsed, e => {
      const pos = XML.source.get(e);
      for (const a of attrs(e)) {
        let token;
        if (relAttr(e, a) && a.value) {
          const rel = context.pkg?.rels(part).find(r => r.id === a.value);
          token = '\u0001rel:' + deps.length + '\u0001'; deps.push({ source, owner: part, id: a.value, ...rel });
        } else {
          const ref = identity(e, a, part);
          if (ref) { token = '\u0001id:' + ids.length + '\u0001'; ids.push({ source, ...ref }); }
        }
        if (token) { const p = pos.attrs.find(p => p.name === a.name); changes.push({ start: p.start, end: p.end, value: token }); }
      }
    });
    return { xml: K.patch(xml, changes), deps, ids, source, part, record: ac?.id || null };
  };
  // Derive a property/subtree from an existing tokenised fragment. Discard metadata
  // for tokens outside the selection, so copied definitions have the right scope.
  K.slice = function (fragment, xml) {
    const deps = [], ids = [], seen = new Map();
    xml = xml.replace(/\u0001(rel|id):(\d+)\u0001/g, (token, kind, n) => {
      if (!seen.has(token)) {
        const input = kind === 'rel' ? fragment.deps : fragment.ids, output = kind === 'rel' ? deps : ids;
        if (!input[+n]) throw new Error('Invalid fragment token');
        seen.set(token, '\u0001' + kind + ':' + output.length + '\u0001'); output.push(input[+n]);
      }
      return seen.get(token);
    });
    return { ...fragment, xml, deps, ids, record: null };
  };
  K.duplicate = function (object) {
    const copy = JSON.parse(JSON.stringify(object)), fragments = [];
    const visit = v => {
      if (!v || typeof v !== 'object') return;
      if (typeof v.xml === 'string' && Array.isArray(v.deps) && Array.isArray(v.ids)) { fragments.push(v); return; }
      for (const child of Object.values(v)) visit(child);
    };
    visit(copy);
    const cohort = uuid(), definitions = fragments.flatMap(f => f.ids.filter(i => i.definition));
    const recordIds = new Map();
    for (const f of fragments) {
      f.copy = cohort; f.copyDefinitions = definitions;
      if (f.record) { if (!recordIds.has(f.record)) recordIds.set(f.record, uuid()); f.record = recordIds.get(f.record); }
    }
    return copy;
  };
  K.loss = function (doc, entry) {
    if (!entry.id || !entry.what || !['conversion', 'drop'].includes(entry.action)) throw new Error('Loss needs identity, description and action');
    const list = doc.losses || (doc.losses = []), previous = list.findIndex(e => e.id === entry.id);
    if (previous < 0) list.push({ ...entry }); else list[previous] = { ...entry };
    return entry;
  };
  const lossKey = e => JSON.stringify([e.id, e.what, e.where || '', e.action]);
  K.pendingLosses = (doc, entries = doc.losses || []) => entries.filter(e => !(doc.acknowledgedLosses || []).includes(lossKey(e)));
  K.acknowledge = (doc, entries) => { doc.acknowledgedLosses = Array.from(new Set([...(doc.acknowledgedLosses || []), ...entries.map(lossKey)])); };
  K.lossState = doc => ({ entries: doc.losses || [], acknowledged: doc.acknowledgedLosses || [] });
  K.recoverLosses = function (doc, state) {
    if (!state) return;
    // A draft may already contain a conversion. Its original save-time loss can
    // no longer be rediscovered from that package, so keep it after recovery.
    for (const entry of state.entries || []) K.loss(doc, { ...entry, phase: 'recovery' });
    doc.acknowledgedLosses = Array.from(new Set([...(doc.acknowledgedLosses || []), ...(state.acknowledged || [])]));
  };

  /** Replace owned children only; preserve every other child and its relative anchor.
   * replacements: expanded element name -> XML string(s), or [] to remove that property.
   * type: generated schema type, e.g. 'a:CT_ShapeProperties'.
   */
  K.merge = function (xml, replacements, type) {
    if (!Object.keys(replacements).length) return xml;
    const tree = K.parse(xml), pos = XML.source.get(tree), slots = K.schema?.types[type];
    if (!slots) throw new Error('No schema order for ' + type);
    const aliases = Object.fromEntries(['w', 'p', 'a', 's', 'rel'].map(k => [NS[k], 'http://purl.oclc.org/ooxml/' + ({ w: 'wordprocessingml/main', p: 'presentationml/main', a: 'drawingml/main', s: 'spreadsheetml/main', rel: 'officeDocument/relationships' })[k]]));
    const canonical = name => { for (const [transitional, strict] of Object.entries(aliases)) name = name.replace('{' + strict + '}', '{' + transitional + '}'); return name; };
    const expanded = name => canonical(name.startsWith('{') ? name : '{' + K.schema.namespaces[name.split(':')[0]] + '}' + name.split(':')[1]);
    const strict = Object.values(aliases).includes(tree.namespaceURI);
    const replacement = xml => strict ? xml.replace(/(xmlns(?::[\w.-]+)?\s*=\s*)(["'])(.*?)\2/g, (m, key, quote, uri) => key + quote + (aliases[uri] || uri) + quote) : xml;
    replacements = Object.fromEntries(Object.entries(replacements).map(([name, values]) => [canonical(name), [values].flat().map(replacement)]));
    const rank = new Map(); slots.forEach((slot, i) => slot.forEach(name => { if (!rank.has(expanded(name))) rank.set(expanded(name), i); }));
    const entries = [], seen = new Set(); let previous = -1, cursor = pos.openEnd, serial = 0;
    const add = (text, order) => { if (text) entries.push({ text, order, serial: serial++ }); };
    for (const child of elements(tree)) {
      const cp = XML.source.get(child), name = canonical('{' + child.namespaceURI + '}' + child.localName);
      add(xml.slice(cursor, cp.start), previous);
      const order = rank.get(name) ?? previous;
      if (rank.has(name)) previous = order;
      if (Object.prototype.hasOwnProperty.call(replacements, name)) {
        if (!seen.has(name)) for (const value of [replacements[name]].flat()) add(value, order);
        seen.add(name);
      } else add(xml.slice(cp.start, cp.end), order);
      cursor = cp.end;
    }
    const self = /\/>$/.test(xml.slice(pos.start, pos.openEnd));
    const closeAt = self ? pos.openEnd : xml.lastIndexOf('</', pos.end - 1);
    add(xml.slice(cursor, closeAt), previous);
    for (const [name, values] of Object.entries(replacements)) if (!seen.has(name)) {
      if (!rank.has(name)) throw new Error('Owned child is absent from the schema table: ' + name);
      for (const value of [values].flat()) add(value, rank.get(name));
    }
    const inner = entries.sort((a, b) => a.order - b.order || a.serial - b.serial).map(e => e.text).join('');
    return self ? xml.slice(0, pos.openEnd - 2) + '>' + inner + '</' + tree.nodeName + '>' + xml.slice(pos.end) : xml.slice(0, pos.openEnd) + inner + xml.slice(closeAt);
  };

  // Unordered settings/property bags only. Content and ordered schema elements
  // use model generation or merge() with its schema table instead.
  K.mergeBag = function (xml, replacements, key = el => '{' + el.namespaceURI + '}' + el.localName) {
    const tree = K.parse(xml), pos = XML.source.get(tree), edits = [], seen = new Set();
    for (const el of elements(tree)) {
      const id = key(el);
      if (!Object.prototype.hasOwnProperty.call(replacements, id)) continue;
      const p = XML.source.get(el), value = seen.has(id) ? '' : [replacements[id]].flat().join('');
      seen.add(id); edits.push({ start: p.start, end: p.end, value });
    }
    const extra = Object.entries(replacements).filter(([id]) => !seen.has(id)).flatMap(([, value]) => [value].flat()).join('');
    if (extra) {
      const self = xml[pos.openEnd - 2] === '/';
      const at = self ? pos.openEnd - 2 : pos.end - tree.nodeName.length - 3;
      edits.push({ start: at, end: self ? pos.openEnd : at, value: self ? '>' + extra + '</' + tree.nodeName + '>' : extra });
    }
    return K.patch(xml, edits);
  };

  // Settings attributes are a property bag too. Keep unknown attributes and
  // children lexically, including the original prefixes and quote style.
  K.attributes = function (xml, changes) {
    const tree = K.parse(xml), pos = XML.source.get(tree), edits = [], added = [];
    for (const [name, value] of Object.entries(changes)) {
      const a = pos.attrs.find(a => a.name === name);
      if (a) {
        if (value == null) {
          const start = xml.lastIndexOf(name, a.start - 2);
          edits.push({ start, end: a.end + 1, value: '' });
        } else edits.push({ start: a.start, end: a.end, value: esc(value) });
      } else if (value != null) added.push(' ' + name + '="' + esc(value) + '"');
    }
    if (added.length) {
      const at = pos.openEnd - (xml[pos.openEnd - 2] === '/' ? 2 : 1);
      edits.push({ start: at, end: at, value: added.join('') });
    }
    return K.patch(xml, edits);
  };
  const equal = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  K.properties = function (xml, generated, current, previous, owned) {
    const replacements = {}, old = elements(K.parse(xml)), fresh = elements(K.parse(generated));
    for (const [key, tag] of Object.entries(owned)) if (!equal(current?.[key], previous?.[key])) {
      const a = old.find(e => e.localName === tag), b = fresh.find(e => e.localName === tag), node = a || b;
      if (node) replacements['{' + node.namespaceURI + '}' + tag] = b ? K.raw(b) : [];
    }
    return K.mergeBag(xml, replacements);
  };
  K.customProperties = function (xml, generated, current = {}, previous = {}) {
    const tree = K.parse(xml), fresh = elements(K.parse(generated)), replacements = {}, names = new Set();
    let pid = Math.max(1, ...elements(tree).map(e => +e.getAttribute('pid') || 0));
    for (const el of elements(tree)) {
      const name = el.getAttribute('name'); names.add(name);
      if (!Object.prototype.hasOwnProperty.call(current, name)) replacements[name] = [];
      else if (!equal(current[name], previous[name])) {
        const node = fresh.find(e => e.getAttribute('name') === name);
        if (node) { node.setAttribute('pid', el.getAttribute('pid')); replacements[name] = K.serialize(node); }
      }
    }
    for (const el of fresh) if (!names.has(el.getAttribute('name'))) {
      el.setAttribute('pid', String(++pid)); replacements[el.getAttribute('name')] = K.serialize(el);
    }
    return K.mergeBag(xml, replacements, e => e.getAttribute('name'));
  };

  const direct = (el, name) => elements(el).find(e => e.localName === name);
  function outerFrames(tree) {
    const names = new Set(['sp', 'pic', 'graphicFrame', 'grpSp', 'wsp', 'wgp', 'spTree', 'shape', 'rect', 'oval', 'group']);
    const out = [];
    function scan(e, inside) {
      if (e !== tree && inside && names.has(e.localName)) return;
      if (e.localName === 'xfrm') out.push(e);
      for (const c of elements(e)) scan(c, inside || names.has(e.localName));
    }
    scan(tree, false); return out;
  }
  function anchorNodes(tree) {
    const all = []; walk(tree, e => all.push(e));
    return {
      frames: outerFrames(tree),
      word: all.filter(e => ['inline', 'anchor'].includes(e.localName) && (e.namespaceURI || '').includes('wordprocessingDrawing')),
      sheet: all.filter(e => ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(e.localName) && (e.namespaceURI || '').includes('spreadsheetDrawing')),
      vml: all.filter(e => e.namespaceURI === K.KNOWN_NS.v && e.hasAttribute('style') && !all.some(p => p !== e && p.namespaceURI === K.KNOWN_NS.v && p.localName === 'group' && isAncestor(p, e))),
    };
  }
  function isAncestor(parent, child) { for (let p = child.parentNode; p; p = p.parentNode) if (p === parent) return true; return false; }
  K.getBox = function (xml) {
    const root = K.parse(xml), nodes = anchorNodes(root), pt = v => +v / 12700;
    const frame = nodes.frames[0], off = frame && direct(frame, 'off'), ext = frame && direct(frame, 'ext');
    const box = off && ext ? { x: pt(off.getAttribute('x')), y: pt(off.getAttribute('y')), w: pt(ext.getAttribute('cx')), h: pt(ext.getAttribute('cy')), rot: +(frame.getAttribute('rot') || 0) / 60000, flipH: frame.getAttribute('flipH') === '1', flipV: frame.getAttribute('flipV') === '1' } : {};
    if (nodes.word.length) {
      const w = nodes.word[0], size = direct(w, 'extent'); box.kind = w.localName;
      if (size) { box.w = pt(size.getAttribute('cx')); box.h = pt(size.getAttribute('cy')); }
      for (const [key, tag] of [['x', 'positionH'], ['y', 'positionV']]) {
        const offset = direct(direct(w, tag) || {}, 'posOffset'); if (offset) box[key] = pt(offset.textContent);
      }
      const simple = direct(w, 'simplePos');
      if (simple && ['1', 'true'].includes(w.getAttribute('simplePos'))) { box.x = pt(simple.getAttribute('x')); box.y = pt(simple.getAttribute('y')); }
    } else if (nodes.sheet.length) {
      const s = nodes.sheet[0]; box.kind = s.localName;
      for (const side of ['from', 'to']) {
        const anchor = direct(s, side); if (!anchor) continue;
        box[side] = { c: +direct(anchor, 'col').textContent, r: +direct(anchor, 'row').textContent, cOff: pt(direct(anchor, 'colOff').textContent), rOff: pt(direct(anchor, 'rowOff').textContent) };
      }
      const position = direct(s, 'pos'), size = direct(s, 'ext');
      if (position) { box.x = pt(position.getAttribute('x')); box.y = pt(position.getAttribute('y')); }
      if (size) { box.w = pt(size.getAttribute('cx')); box.h = pt(size.getAttribute('cy')); }
    } else if (nodes.vml.length) {
      box.kind = 'vml';
      const style = Object.fromEntries(nodes.vml[0].getAttribute('style').split(';').filter(s => s.includes(':')).map(s => { const i = s.indexOf(':'); return [s.slice(0, i).trim(), s.slice(i + 1).trim()]; }));
      const points = v => { const n = parseFloat(v); return /in$/i.test(v) ? n * 72 : /cm$/i.test(v) ? n * 72 / 2.54 : /mm$/i.test(v) ? n * 72 / 25.4 : /px$/i.test(v) ? n * 0.75 : n; };
      for (const [key, name] of [['x', 'left'], ['y', 'top'], ['w', 'width'], ['h', 'height']]) if (style[name] || style['margin-' + name]) box[key] = points(style['margin-' + name] || style[name]);
      walk(nodes.vml[0], el => {
        if (el.localName !== 'Anchor' || el.namespaceURI !== K.KNOWN_NS.x) return;
        const values = el.textContent.split(',').map(Number);
        if (values.length === 8) for (const [side, base] of [['from', 0], ['to', 4]]) box[side] = { c: values[base], r: values[base + 2], cFraction: values[base + 1] / 1024, rFraction: values[base + 3] / 256 };
      });
    } else box.kind = 'drawing';
    return box;
  };
  K.setBox = function (xml, box) {
    for (const key of ['x', 'y', 'w', 'h', 'rot']) if (box[key] != null && !Number.isFinite(box[key])) throw new Error('Invalid box ' + key);
    const root = K.parse(xml), nodes = anchorNodes(root), edits = [], insertions = new Map();
    const emu = n => Math.round(n * 12700);
    function attr(el, name, value) {
      if (!el || value == null) return;
      const previous = el.getAttribute(name);
      if (previous === String(value) || previous != null && typeof value === 'number' && +previous === value) return;
      const p = XML.source.get(el), a = p.attrs.find(a => a.name === name), v = esc(value).replace(/'/g, '&apos;');
      if (a) edits.push({ start: a.start, end: a.end, value: v });
      else { const pos = p.openEnd - (xml[p.openEnd - 2] === '/' ? 2 : 1); insertions.set(pos, (insertions.get(pos) || '') + ' ' + name + '="' + v + '"'); }
    }
    function text(el, value) {
      if (!el || value == null) return;
      const p = XML.source.get(el), self = xml[p.openEnd - 2] === '/';
      edits.push(self ? { start: p.openEnd - 2, end: p.openEnd, value: '>' + esc(value) + '</' + el.nodeName + '>' } : { start: p.openEnd, end: xml.lastIndexOf('</', p.end - 1), value: esc(value) });
    }
    for (const frame of nodes.frames) {
      const off = direct(frame, 'off'), ext = direct(frame, 'ext');
      if (!nodes.word.length && !nodes.sheet.length) { if (box.x != null) attr(off, 'x', emu(box.x)); if (box.y != null) attr(off, 'y', emu(box.y)); }
      if (box.w != null) attr(ext, 'cx', Math.max(0, emu(box.w))); if (box.h != null) attr(ext, 'cy', Math.max(0, emu(box.h)));
      if (box.rot != null) attr(frame, 'rot', Math.round(box.rot * 60000));
      for (const key of ['flipH', 'flipV']) if (box[key] != null) attr(frame, key, box[key] ? '1' : '0');
    }
    for (const word of nodes.word) {
      const ext = direct(word, 'extent'); if (box.w != null) attr(ext, 'cx', Math.max(0, emu(box.w))); if (box.h != null) attr(ext, 'cy', Math.max(0, emu(box.h)));
      for (const [key, tag] of [['x', 'positionH'], ['y', 'positionV']]) if (box[key] != null) {
        const position = direct(word, tag), offset = position && direct(position, 'posOffset');
        if (offset) text(offset, emu(box[key]));
        else if (position) {
          const align = direct(position, 'align');
          if (align) { const p = XML.source.get(align), prefix = position.prefix ? position.prefix + ':' : ''; edits.push({ start: p.start, end: p.end, value: '<' + prefix + 'posOffset>' + emu(box[key]) + '</' + prefix + 'posOffset>' }); }
        }
      }
      const simple = direct(word, 'simplePos'); if (['1', 'true'].includes(word.getAttribute('simplePos'))) { if (box.x != null) attr(simple, 'x', emu(box.x)); if (box.y != null) attr(simple, 'y', emu(box.y)); }
      if (box.effectExtent) for (const key of ['l', 't', 'r', 'b']) if (box.effectExtent[key] != null) attr(direct(word, 'effectExtent'), key, emu(box.effectExtent[key]));
    }
    for (const sheet of nodes.sheet) {
      for (const side of ['from', 'to']) if (box[side]) {
        const anchor = direct(sheet, side); if (!anchor) continue;
        for (const [key, name] of [['c', 'col'], ['r', 'row'], ['cOff', 'colOff'], ['rOff', 'rowOff']]) if (box[side][key] != null) text(direct(anchor, name), key.endsWith('Off') ? emu(box[side][key]) : box[side][key]);
      }
      const off = direct(sheet, 'pos'), ext = direct(sheet, 'ext');
      if (box.x != null) attr(off, 'x', emu(box.x)); if (box.y != null) attr(off, 'y', emu(box.y));
      if (box.w != null) attr(ext, 'cx', Math.max(0, emu(box.w))); if (box.h != null) attr(ext, 'cy', Math.max(0, emu(box.h)));
    }
    for (const vml of nodes.vml) {
      let style = vml.getAttribute('style');
      for (const [key, property] of [['x', 'left'], ['y', 'top'], ['w', 'width'], ['h', 'height']]) if (box[key] != null) {
        const name = new RegExp('(?:^|;)\\s*margin-' + property + '\\s*:').test(style) ? 'margin-' + property : property;
        const re = new RegExp('(^|;)(\\s*' + name + '\\s*:)[^;]*');
        style = re.test(style) ? style.replace(re, (_, before, label) => before + label + box[key] + 'pt') : style + (style.endsWith(';') ? '' : ';') + name + ':' + box[key] + 'pt';
      }
      attr(vml, 'style', style);
      walk(vml, el => {
        if (el.localName !== 'Anchor' || el.namespaceURI !== K.KNOWN_NS.x) return;
        const values = el.textContent.split(',').map(Number);
        if (values.length !== 8) throw new Error('Invalid VML cell anchor');
        for (const [side, base] of [['from', 0], ['to', 4]]) if (box[side]) {
          const a = box[side]; if (a.c != null) values[base] = a.c; if (a.r != null) values[base + 2] = a.r;
          if (a.cFraction != null) values[base + 1] = Math.round(a.cFraction * 1024); if (a.rFraction != null) values[base + 3] = Math.round(a.rFraction * 256);
        }
        if (box.from || box.to) text(el, values.join(', '));
      });
    }
    for (const [pos, value] of insertions) edits.push({ start: pos, end: pos, value });
    return K.patch(xml, edits);
  };

  class Rels {
    constructor(original = []) {
      this.original = original; this.list = []; this.reserved = new Set(original.map(r => r.id)); this.high = 0;
      for (const id of this.reserved) { const m = /^rId(\d+)$/.exec(id); if (m) this.high = Math.max(this.high, +m[1]); }
    }
    add(type, target, external = false, preferred) {
      const same = r => r.type === type && r.target === target && !!r.external === !!external;
      if (preferred) {
        const hit = this.list.find(r => r.id === preferred);
        if (hit) { if (!same(hit)) throw new Error('Relationship identity rebound: ' + preferred); return hit.id; }
      } else {
        const hit = this.list.find(same); if (hit) return hit.id;
        preferred = this.original.find(r => same(r) && !this.list.some(x => x.id === r.id))?.id;
      }
      let id = preferred;
      if (!id) { do { id = 'rId' + ++this.high; } while (this.reserved.has(id)); }
      this.reserved.add(id); this.list.push({ id, type, target, external: !!external }); return id;
    }
    xml() {
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + NS.pkg + '">' + this.list.map(r => '<Relationship Id="' + esc(r.id) + '" Type="' + esc(r.type) + '" Target="' + esc(r.external ? r.target : escapedURI(r.target)) + '"' + (r.external ? ' TargetMode="External"' : '') + '/>').join('') + '</Relationships>';
    }
  }
  K.Rels = Rels;
  const signatureRel = type => /^http:\/\/schemas\.openxmlformats\.org\/package\/2006\/relationships\/digital-signature\//.test(type || '');
  class Writer {
    constructor(pkg, options = {}) {
      this.pkg = pkg; this.doc = options.doc || {}; this.parts = new Map(); this.types = new Map();
      this.contentType = options.contentType;
      // Save-time losses are recomputed. Cancelling a macro-free Save As must
      // not warn again when the next save keeps the original macro variant.
      this.doc.losses = (this.doc.losses || []).filter(e => e.phase !== 'save');
      this.names = new Set((pkg?.names || []).map(n => n.toLowerCase())); this.mapping = new Map();
      this.classes = new Map(); this.omitted = new Set(); this.carried = new Set(); this.relationships = new Map(); this.emitted = new Set();
      this.ids = new Identities();
      if (pkg) {
        for (const name of pkg.names) {
          this.mapping.set(this.key(pkg, name), name);
          const text = /\.(xml|vml)$/i.test(name) ? pkg.text(name) : null;
          if (!text || !/docPr|cNvPr|cTn|bookmark|permStart|permEnd|sldId|sldMasterId|sldLayoutId|sheetId|cacheId|numId|abstractNumId|sdtPr|dxf/.test(text)) continue;
          let tree; try { tree = K.parse(text); } catch (e) { continue; } // carried opaque if the original XML is malformed
          walk(tree, el => {
            if (el.localName === 'dxfs' && el.namespaceURI === NS.s) elements(el).forEach((e, i) => this.ids.reserve('styles', 'dxf', i));
            for (const a of attrs(el)) { const ref = identity(el, a, name); if (ref) this.ids.reserve(ref.scope, ref.kind, ref.id); }
          });
        }
        for (const [name, mode] of Object.entries(options.classes || {})) this.claim(name, mode);
      }
    }
    loss(entry) { return K.loss(this.doc, { ...entry, phase: 'save' }); }
    key(pkg, part) { return JSON.stringify([pkg.id, part]); }
    name(dir, base, ext) {
      dir = dir ? dir.replace(/\/$/, '') + '/' : ''; ext = ext ? '.' + ext.replace(/^\./, '') : '';
      let n = 1, name;
      do { name = dir + base + n++ + ext; } while (this.names.has(name.toLowerCase()));
      this.names.add(name.toLowerCase()); return name;
    }
    target(pkg, part) {
      const key = this.key(pkg, part);
      if (this.omitted.has(key)) return null;
      if (!this.mapping.has(key)) {
        let name = part;
        if (this.names.has(name.toLowerCase())) {
          const slash = part.lastIndexOf('/'), dot = part.lastIndexOf('.');
          const hasExt = dot > slash;
          const base = part.slice(slash + 1, hasExt ? dot : undefined).replace(/\d+$/, '') || 'part';
          name = this.name(part.slice(0, slash < 0 ? 0 : slash), base, hasExt ? part.slice(dot + 1) : '');
        } else this.names.add(name.toLowerCase());
        this.mapping.set(key, name);
      }
      return this.mapping.get(key);
    }
    claim(part, mode = 'regenerated', target = part, pkg = this.pkg) {
      if (!['opaque', 'merged', 'regenerated'].includes(mode)) throw new Error('Unknown part ownership: ' + mode);
      const key = this.key(pkg, part);
      if (Array.from(this.mapping).some(([other, name]) => other !== key && name.toLowerCase() === target.toLowerCase())) throw new Error('Part name collision: ' + target);
      this.classes.set(key, mode); this.mapping.set(key, target); this.names.add(target.toLowerCase());
      return target;
    }
    mode(pkg, part) { return this.classes.get(this.key(pkg, part)) || 'opaque'; }
    omit(part, why, pkg = this.pkg) {
      const key = this.key(pkg, part); this.omitted.add(key);
      const target = this.mapping.get(key); if (target) { this.parts.delete(target); this.relationships.delete(target); }
      this.loss({ id: 'part:' + key, what: why, where: part, action: 'drop' });
    }
    put(name, data, type) {
      if (data == null) throw new Error('No bytes for ' + name);
      this.names.add(name.toLowerCase()); this.parts.set(name, typeof data === 'string' ? data : new Uint8Array(data).slice());
      if (type) this.types.set(name, type);
      return name;
    }
    rels(owner) {
      if (!this.relationships.has(owner)) {
        let original = [];
        if (this.pkg) {
          const origin = ['', ...this.pkg.names].find(n => (n ? this.mapping.get(this.key(this.pkg, n)) : '') === owner);
          if (origin !== undefined) original = this.pkg.rels(origin);
        }
        this.relationships.set(owner, new Rels(original));
      }
      return this.relationships.get(owner);
    }
    keepRel(owner, dep) {
      const pkg = packages.get(dep.source) || this.pkg;
      if (!pkg || !dep.type) throw new Error('Unresolved relationship ' + (dep.id || '') + ' in ' + (dep.owner || owner));
      if (signatureRel(dep.type)) {
        this.loss({ id: 'signature:' + pkg.id, what: 'The digital signature becomes invalid when this file is saved.', where: dep.owner || '/', action: 'drop' });
        return null;
      }
      let target = dep.target;
      if (!dep.external) {
        const part = dep.part || K.resolve(dep.owner || '', dep.target).part;
        const dest = this.carry(pkg, part);
        if (!dest) return null;
        target = pkg === this.pkg && dest === part && owner === dep.owner ? dep.target : K.relative(owner, dest) + (dep.fragment || '');
      }
      const preferred = pkg === this.pkg && owner === (dep.owner ? this.target(pkg, dep.owner) : '') ? dep.id : undefined;
      return this.rels(owner).add(dep.type, target, dep.external, preferred);
    }
    carryRels(pkg, from, owner = this.target(pkg, from), predicate = () => true) {
      for (const r of pkg.rels(from)) if (predicate(r)) this.keepRel(owner, { source: pkg.id, owner: from, ...r });
    }
    carry(pkg, part) {
      // Check the whole opaque graph before mutating the output. A malformed
      // source may need a model conversion; a failed carry must not leave half
      // a graph behind and poison that otherwise valid fallback.
      if (this.omitted.has(this.key(pkg, part))) return null;
      const seen = new Set(), pending = [part];
      while (pending.length) {
        const name = pending.pop(), key = this.key(pkg, name);
        if (seen.has(key)) continue;
        seen.add(key);
        if (this.omitted.has(key)) throw new Error('Cannot preserve ' + part + ': deleted dependency ' + name);
        if (!pkg.has(name)) throw new Error('Missing dependency: ' + name);
        if (this.mode(pkg, name) !== 'opaque') continue;
        for (const r of pkg.rels(name)) if (!r.external && !signatureRel(r.type)) pending.push(r.part);
      }
      return this._carry(pkg, part);
    }
    _carry(pkg, part) {
      const key = this.key(pkg, part), dest = this.target(pkg, part);
      if (!dest || this.carried.has(key)) return dest;
      if (!pkg.has(part)) throw new Error('Missing dependency: ' + part);
      this.carried.add(key); // mark before following edges: layouts can point back at masters
      if (this.mode(pkg, part) !== 'opaque') return dest;
      this.put(dest, pkg.bytes(part), pkg.type(part) || this.contentType?.(pkg, part));
      const relationships = pkg.rels(part), rewritten = [];
      for (const r of relationships) {
        if (signatureRel(r.type)) {
          this.keepRel(dest, { source: pkg.id, owner: part, ...r }); continue;
        }
        const target = r.external ? r.target : this._carry(pkg, r.part);
        if (target == null) {
          this.loss({ id: 'dependency:' + key + ':' + r.id, what: 'A preserved part refers to a deleted part.', where: part, action: 'drop' });
          throw new Error('Cannot preserve ' + part + ': deleted dependency ' + r.part);
        }
        rewritten.push({ ...r, target: r.external ? r.target : escapedURI(dest === part && target === r.part ? r.target : K.relative(dest, target) + (r.fragment || '')) });
      }
      if (relationships.length || pkg.has(K.relsPath(part))) {
        const unchanged = relationships.length === rewritten.length && relationships.every((r, i) => r.target === rewritten[i].target);
        if (unchanged) this.put(K.relsPath(dest), pkg.bytes(K.relsPath(part)), 'application/vnd.openxmlformats-package.relationships+xml');
        else {
          const rels = new Rels(); for (const r of rewritten) rels.add(r.type, r.target, r.external, r.id);
          this.put(K.relsPath(dest), rels.xml(), 'application/vnd.openxmlformats-package.relationships+xml');
        }
      }
      return dest;
    }
    copyPart(pkg, part, dest, references = {}) {
      if (!pkg.has(part)) throw new Error('Missing dependency: ' + part);
      const rels = new Rels();
      for (const r of pkg.rels(part)) {
        if (signatureRel(r.type)) { this.keepRel(dest, { source: pkg.id, owner: part, ...r }); continue; }
        const target = r.external ? r.target : references[r.part] || this.carry(pkg, r.part);
        if (!target) throw new Error('Cannot copy ' + part + ': deleted dependency ' + r.part);
        rels.add(r.type, r.external ? target : K.relative(dest, target) + (r.fragment || ''), r.external, r.id);
      }
      this.put(dest, pkg.bytes(part), pkg.type(part));
      if (rels.list.length) this.put(K.relsPath(dest), rels.xml(), 'application/vnd.openxmlformats-package.relationships+xml');
      return dest;
    }
    emit(fragment, owner, options = {}) {
      const record = fragment.record && JSON.stringify([owner, fragment.record]);
      if (record && this.emitted.has(record)) return '';
      const copy = options.copy || fragment.copy;
      const definitions = fragment.copyDefinitions || fragment.ids.filter(i => i.definition);
      if (copy) this.ids.copy(copy, definitions);
      if (fragment.source !== this.pkg?.id || owner !== fragment.part) {
        const available = new Set(definitions.concat(options.references || []).map(i => this.ids.key(i.source, i.scope, i.kind, i.id)));
        const unresolved = fragment.ids.find(i => !i.definition && ['shape', 'timing'].includes(i.kind) && i.scope === fragment.part && !available.has(this.ids.key(i.source, i.scope, i.kind, i.id)));
        if (unresolved) {
          this.loss({ id: 'copy-reference:' + fragment.source + ':' + fragment.part + ':' + unresolved.kind + ':' + unresolved.id, what: 'The copied object refers to an object outside the copied selection.', where: owner, action: 'conversion' });
          const error = new Error('Unresolved copied ' + unresolved.kind + ' target: ' + unresolved.id);
          error.code = 'OOXML_COPY_REFERENCE'; throw error;
        }
      }
      const dependencies = fragment.deps.map(d => {
        const id = this.keepRel(owner, d); if (!id) throw new Error('Fragment dependency was removed'); return id;
      });
      const ids = fragment.ids.map(i => this.ids.resolve(i.source, i.scope, i.kind, i.id, {
        primary: this.pkg?.id, copy, scope: i.scope === fragment.part ? owner : i.scope,
      }));
      let xml = fragment.xml.replace(/\u0001(rel|id):(\d+)\u0001/g, (_, kind, n) => {
        const value = (kind === 'rel' ? dependencies : ids)[+n];
        if (value == null) throw new Error('Invalid fragment token'); return esc(value);
      });
      if (options.box) xml = K.setBox(xml, options.box);
      if (record) this.emitted.add(record);
      return xml;
    }
    finish() {
      for (const [owner, rels] of this.relationships) if (rels.list.length) this.put(K.relsPath(owner), rels.xml(), 'application/vnd.openxmlformats-package.relationships+xml');
      for (const [name, data] of this.parts) if (name.endsWith('.rels')) {
        const owner = K.owner(name), tree = K.parse(typeof data === 'string' ? data : XML.decode(data));
        for (const rel of elements(tree)) if (rel.getAttribute('TargetMode') !== 'External') {
          const target = K.resolve(owner, rel.getAttribute('Target')).part;
          if (!this.parts.has(target)) throw new Error('Unwritten relationship target: ' + target + ' from ' + owner);
        }
      }
      if (this.pkg) for (const name of this.pkg.names) {
        if (name === '[Content_Types].xml' || name.endsWith('.rels') || this.mode(this.pkg, name) !== 'opaque') continue;
        if (!this.parts.has(this.mapping.get(this.key(this.pkg, name))) && !this.omitted.has(this.key(this.pkg, name))) {
          this.loss({ id: 'unreferenced:' + this.pkg.id + ':' + name, what: 'This preserved part has no remaining reference in the saved file.', where: name, action: 'drop' });
        }
      }
      const overrides = [];
      for (const [name] of this.parts) {
        if (name === '[Content_Types].xml') continue;
        const type = this.types.get(name) || this.pkg?.type(name);
        if (!type) throw new Error('Missing content type for ' + name);
        overrides.push('<Override PartName="/' + esc(K.relative('', name)) + '" ContentType="' + esc(type) + '"/>');
      }
      this.put('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?>\n<Types xmlns="' + NS.ct + '">' + overrides.join('') + '</Types>');
      return { files: Array.from(this.parts, ([name, data]) => ({ name: name === '[Content_Types].xml' ? name : K.relative('', name), data })), dropped: (this.doc.losses || []).slice() };
    }
  }
  K.Writer = Writer;

  // Bridge canonical generator paths to original OPC names. Apps provide their
  // ownership table and settings merge; content remains model-generated.
  K.output = function (pkg, options = {}) {
    const writer = new Writer(pkg, options), names = new Map(), originals = new Map(), modes = new Map(), views = new Map();
    const strict = pkg?.main && pkg.xml(pkg.main)?.namespaceURI?.startsWith('http://purl.oclc.org/ooxml/');
    const bind = (base, original, mode = 'regenerated') => {
      if (original && pkg?.has(original)) {
        let target = writer.target(pkg, original);
        if (target === original && pkg.names.some(n => n !== original && n.toLowerCase() === original.toLowerCase())) {
          const slash = original.lastIndexOf('/'), dot = original.lastIndexOf('.');
          target = writer.name(original.slice(0, Math.max(0, slash)), original.slice(slash + 1, dot).replace(/\d+$/, ''), original.slice(dot + 1));
        }
        names.set(base, target); originals.set(base, original); modes.set(base, mode); writer.claim(original, mode, target);
      }
    };
    const part = base => {
      if (!base) return '';
      if (base.endsWith('.rels')) return K.relsPath(part(K.owner(base)));
      if (!names.has(base)) {
        const slash = base.lastIndexOf('/'), dot = base.lastIndexOf('.');
        const name = writer.names.has(base.toLowerCase()) ? writer.name(base.slice(0, Math.max(0, slash)), base.slice(slash + 1, dot < slash ? undefined : dot).replace(/\d+$/, ''), dot > slash ? base.slice(dot + 1) : '') : base;
        names.set(base, name); writer.names.add(name.toLowerCase());
      }
      return names.get(base);
    };
    const rels = base => {
      const owner = part(base);
      if (!views.has(base)) {
        const r = writer.rels(owner);
        const add = (type, target, external, preferred) => {
          const mapped = value => {
            if (external) return value;
            const resolved = K.resolve(base, value);
            return K.relative(owner, part(resolved.part)) + (resolved.fragment || '');
          };
          if (typeof target === 'function') return r.add(strict ? strictURI(type) : type, () => mapped(target()), external, preferred);
          let destination = mapped(target);
          const source = base ? originals.get(base) : '';
          const original = source !== undefined && pkg?.rels(source).find(rel =>
            K.relationshipType(rel.type) === K.relationshipType(type) && !!rel.external === !!external &&
            (external ? rel.target === destination : writer.target(pkg, rel.part) === K.resolve(owner, destination).part && (rel.fragment || '') === K.resolve(owner, destination).fragment));
          if (original && preferred == null) {
            preferred = original.id;
            if (owner === source && (external || writer.target(pkg, original.part) === original.part)) destination = original.target;
          }
          return r.add(strict ? strictURI(type) : type, destination, external, preferred);
        };
        views.set(base, { base, owner, list: r.list, add, xml: () => r.xml() });
      }
      return views.get(base);
    };
    const put = (base, data, type) => {
      if (base === '[Content_Types].xml') return;
      if (base.endsWith('.rels')) {
        const ownerBase = K.owner(base), owner = part(ownerBase);
        if (modes.get(ownerBase) === 'opaque') return;
        if (writer.relationships.has(owner)) return;
        const rels = writer.rels(owner), tree = K.parse(typeof data === 'string' ? data : XML.decode(data));
        for (const r of elements(tree)) {
          const external = r.getAttribute('TargetMode') === 'External', target = r.getAttribute('Target');
          rels.add(r.getAttribute('Type'), external ? target : K.relative(owner, part(K.resolve(ownerBase, target).part)), external, r.getAttribute('Id'));
        }
        return;
      }
      const original = originals.get(base), mode = modes.get(base), target = part(base);
      if (original && mode === 'opaque') {
        try { writer.carry(pkg, original); writer.relationships.delete(target); return; }
        catch (error) {
          const converted = options.convert?.(base, original, data, error);
          if (converted == null) throw error;
          data = converted; modes.set(base, 'regenerated'); writer.claim(original, 'regenerated', target);
          writer.loss({ id: 'converted-part:' + original, what: 'The part was converted because its original dependencies are incomplete: ' + error.message, where: original, action: 'conversion' });
        }
      }
      if (original && mode === 'merged') {
        data = options.merge(base, original, data, writer);
        // A part's declaration, internal entity definitions, and processing
        // instructions are outside its root fragment but still belong to it.
        data = K.partXML(pkg.text(original), data);
      }
      writer.put(target, strict && typeof data === 'string' ? K.strictXML(data) : data, type || pkg?.type(original));
    };
    const finish = () => {
      if (pkg) {
        for (const [base, source] of [['', ''], ...originals]) {
          const owner = base ? part(base) : '';
          if (base && !writer.parts.has(owner)) continue;
          if (modes.get(base) === 'opaque') continue;
          for (const rel of pkg.rels(source)) {
            const type = K.relationshipType(rel.type);
            if (options.format && !options.format.macro && /\/vbaProject$/.test(type)) { writer.omit(rel.part, 'Macros cannot be saved in the selected file type.'); continue; }
            if (options.consumes?.(base, rel, type)) continue;
            try { writer.keepRel(owner, { source: pkg.id, owner: source, ...rel }); }
            catch (error) { writer.loss({ id: 'relationship:' + source + ':' + rel.id, what: 'The original relationship could not be retained: ' + error.message, where: source || '/', action: 'drop' }); }
          }
        }
      }
      return writer.finish();
    };
    const use = name => { names.set(name, name); writer.names.add(name.toLowerCase()); return name; };
    return { writer, bind, use, part, rels, put, finish, originals };
  };

  K.emit = (fragment, writer, owner, options) => writer.emit(fragment, owner, options);

  // Clipboard transport is explicit: bytes never enter model/history JSON.
  const base64 = bytes => { let s = ''; for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode(...bytes.subarray(i, i + 32768)); return btoa(s); };
  K.export = function (references) {
    const sources = new Map();
    for (const ref of references) {
      const pkg = packages.get(ref.source); if (!pkg) throw new Error('Source package is no longer available');
      if (!sources.has(pkg.id)) sources.set(pkg.id, new Set());
      const selected = sources.get(pkg.id);
      for (const name of K.subtree(pkg, ref.part).parts) { selected.add(name); if (pkg.has(K.relsPath(name))) selected.add(K.relsPath(name)); }
    }
    return Array.from(sources, ([source, names]) => {
      const pkg = packages.get(source);
      return { source, parts: Array.from(names, name => ({ name, type: pkg.type(name), data: base64(pkg.bytes(name)) })) };
    });
  };
  K.import = function (bundle) {
    const mapping = {};
    for (const source of bundle) {
      const bytes = new Map(source.parts.map(p => [p.name, Uint8Array.from(atob(p.data), c => c.charCodeAt(0))]));
      const ct = '<Types xmlns="' + NS.ct + '">' + source.parts.map(p => '<Override PartName="/' + esc(K.relative('', p.name)) + '" ContentType="' + esc(p.type) + '"/>').join('') + '</Types>';
      bytes.set('[Content_Types].xml', new TextEncoder().encode(ct));
      mapping[source.source] = makePackage(bytes);
    }
    return mapping;
  };
  K.remapSources = function (object, mapping) {
    const copy = JSON.parse(JSON.stringify(object));
    const visit = v => {
      if (!v || typeof v !== 'object') return;
      if (v.source && mapping[v.source]) v.source = mapping[v.source].id;
      for (const child of Object.values(v)) visit(child);
    };
    visit(copy); return copy;
  };
})(typeof window !== 'undefined' ? window : globalThis);
