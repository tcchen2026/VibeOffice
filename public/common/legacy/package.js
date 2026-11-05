/* Minimal OOXML package assembly, using the suite's ZIP implementation.
 * ECMA-376 Part 2: part names, content types and package relationships. No DOM. */
(function (root) {
  'use strict';
  const L = root.L, B = L.legacy;
  B.ns = Object.freeze({
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
    officeRel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/',
    content: 'http://schemas.openxmlformats.org/package/2006/content-types',
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main'
  });
  B.xmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  B.attr = value => B.xml(value).replace(/\r/g, '&#13;').replace(/\n/g, '&#10;').replace(/\t/g, '&#9;');
  const nameOK = name => /^[A-Za-z_][\w.:-]*$/.test(name);
  B.el = (name, attrs = {}, body) => {
    if (!nameOK(name)) B.fail('XML_NAME', 'Invalid XML name');
    let out = '<' + name;
    for (const [key, value] of Object.entries(attrs)) {
      if (!nameOK(key)) B.fail('XML_NAME', 'Invalid XML attribute');
      if (value != null) out += ' ' + key + '="' + B.attr(value) + '"';
    }
    return body == null ? out + '/>' : out + '>' + body + '</' + name + '>';
  };
  B.XML = class {
    constructor() { this.parts = []; this.stack = []; this.length = 0; }
    append(s) { this.length += s.length; if (this.length > B.limits.bytes) B.fail('XML_LIMIT', 'The generated XML is too large'); this.parts.push(s); return this; }
    start(name, attrs) { if (this.stack.length >= B.limits.depth) B.fail('XML_LIMIT', 'The generated XML is too deeply nested'); this.append(B.el(name, attrs).slice(0, -2) + '>'); this.stack.push(name); return this; }
    text(value) { return this.append(B.xml(value)); }
    empty(name, attrs) { return this.append(B.el(name, attrs)); }
    end() { const name = this.stack.pop(); if (!name) B.fail('XML_STRUCTURE', 'Unbalanced XML writer'); return this.append('</' + name + '>'); }
    finish() { if (this.stack.length) B.fail('XML_STRUCTURE', 'Unclosed XML element'); return B.xmlDeclaration + this.parts.join(''); }
  };
  const partName = name => {
    if (typeof name !== 'string' || !name || name.length > 1024 || name.startsWith('/') || /[\\\x00-\x20?#]/.test(name) || name.split('/').some(p => !p || p === '.' || p === '..')) B.fail('PART_NAME', 'Invalid package part name');
    return name;
  };
  B.Package = class {
    constructor() { this.parts = new Map(); this.types = new Map(); this.rels = new Map(); this.bytes = 0; }
    add(name, data, type) {
      partName(name);
      if (this.parts.has(name)) B.fail('PART_DUPLICATE', 'Duplicate package part: ' + name);
      if (!(typeof data === 'string' || data instanceof Uint8Array)) B.fail('PART_DATA', 'Invalid package data');
      const size = typeof data === 'string' ? L.zip.encode(data).length : data.length;
      this.bytes += size; if (this.bytes > B.limits.bytes || this.parts.size >= 60000) B.fail('PACKAGE_LIMIT', 'The converted package is too large');
      this.parts.set(name, data); if (type) this.types.set(name, type); return name;
    }
    rel(from, type, target, options = {}) {
      if (from) partName(from);
      const list = this.rels.get(from) || []; this.rels.set(from, list);
      const id = options.id || 'rId' + (list.length + 1);
      if (!nameOK(id) || list.some(r => r.Id === id)) B.fail('REL_ID', 'Invalid or duplicate relationship ID');
      if (typeof target !== 'string' || /[\x00-\x1f]/.test(target)) B.fail('REL_TARGET', 'Invalid relationship target');
      if (options.external && /^(?:javascript|vbscript|data):/i.test(target.trim())) B.fail('REL_TARGET', 'Unsupported executable link target');
      list.push({ Id: id, Type: type.includes('://') ? type : B.ns.officeRel + type, Target: target, ...(options.external ? { TargetMode: 'External' } : {}) }); return id;
    }
    async finish() {
      for (const [from, list] of this.rels) {
        const slash = from.lastIndexOf('/'), name = from ? from.slice(0, slash + 1) + '_rels/' + from.slice(slash + 1) + '.rels' : '_rels/.rels';
        this.add(name, B.xmlDeclaration + B.el('Relationships', { xmlns: B.ns.rel }, list.map(r => B.el('Relationship', r)).join('')));
      }
      const defaults = B.el('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }) +
        B.el('Default', { Extension: 'xml', ContentType: 'application/xml' });
      this.add('[Content_Types].xml', B.xmlDeclaration + B.el('Types', { xmlns: B.ns.content }, defaults +
        [...this.types].map(([name, type]) => B.el('Override', { PartName: '/' + name, ContentType: type })).join('')));
      const blob = await L.zip.write([...this.parts].map(([name, data]) => ({ name, data })));
      if (blob.size > B.limits.bytes) B.fail('PACKAGE_LIMIT', 'The converted package is too large');
      return new Uint8Array(await blob.arrayBuffer());
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);
