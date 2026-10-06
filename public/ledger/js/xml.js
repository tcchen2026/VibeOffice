/* Ledger — small, fast XML parser.
 * L.xml.parse(text) → a lightweight element tree that offers the subset of the DOM API the readers use
 *   (localName, prefix, children, childNodes, getAttribute, getAttributeNS, attributes, textContent,
 *    getElementsByTagName, parentNode, replaceWith, remove, namespaceURI).
 * L.xml.scan(text, from, to, fn) → a regex-driven tag scanner for very large parts (worksheet cell data).
 * Works the same in browsers and in Node (no DOMParser needed).
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const XML = (L.xml = {});

  const ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
  function unescape(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e) => {
      if (e[0] === '#') {
        const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        if (!(cp >= 0 && cp <= 0x10ffff)) return m;
        try { return String.fromCodePoint(cp); } catch (err) { return m; }
      }
      return ENT[e] != null ? ENT[e] : m;
    });
  }
  XML.unescape = unescape;
  XML.esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  /** text node content: also protect characters XML 1.0 forbids (Excel's _xHHHH_ escape) */
  XML.escText = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, (c) => '_x' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') + '_');
  /** decode Excel's _xHHHH_ escapes (used in shared strings and attribute text) */
  XML.unx = (s) => (s.indexOf('_x') < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16))));

  class Txt {
    constructor(data, parent) { this.data = data; this.parentNode = parent; }
    get nodeType() { return 3; }
    get textContent() { return this.data; }
    get nodeValue() { return this.data; }
    get localName() { return null; }
    remove() { const p = this.parentNode; if (!p) return; const i = p.childNodes.indexOf(this); if (i >= 0) p.childNodes.splice(i, 1); this.parentNode = null; }
  }
  class El {
    constructor(name, parent) {
      this.nodeName = name;
      const i = name.indexOf(':');
      this.prefix = i < 0 ? null : name.slice(0, i);
      this.localName = i < 0 ? name : name.slice(i + 1);
      this.attrs = null;
      this.children = [];
      this.childNodes = [];
      this.parentNode = parent;
    }
    get nodeType() { return 1; }
    get tagName() { return this.nodeName; }
    getAttribute(q) { const a = this.attrs; if (!a) return null; const v = a[q]; return v === undefined ? null : v; }
    hasAttribute(q) { return this.getAttribute(q) != null; }
    setAttribute(q, v) { if (!this.attrs) this.attrs = Object.create(null); this.attrs[q] = String(v); }
    removeAttribute(q) { if (this.attrs) delete this.attrs[q]; }
    get attributes() {
      const out = [];
      if (this.attrs) for (const k in this.attrs) { const i = k.indexOf(':'); out.push({ name: k, prefix: i < 0 ? null : k.slice(0, i), localName: i < 0 ? k : k.slice(i + 1), value: this.attrs[k] }); }
      return out;
    }
    lookupNamespaceURI(prefix) {
      const key = prefix ? 'xmlns:' + prefix : 'xmlns';
      for (let e = this; e && e.nodeType === 1; e = e.parentNode) { const a = e.attrs; if (a && a[key] !== undefined) return a[key]; }
      if (prefix === 'xml') return 'http://www.w3.org/XML/1998/namespace';
      return null;
    }
    get namespaceURI() { return this.lookupNamespaceURI(this.prefix); }
    getAttributeNS(ns, local) {
      const a = this.attrs;
      if (!a) return null;
      if (!ns) { const v = a[local]; return v === undefined ? null : v; }
      for (const k in a) {
        const i = k.indexOf(':');
        if (i < 0 || k.slice(i + 1) !== local) continue;
        const p = k.slice(0, i);
        if (p === 'xmlns') continue;
        if (this.lookupNamespaceURI(p) === ns) return a[k];
      }
      return null;
    }
    get textContent() {
      if (this.childNodes.length === 1 && this.childNodes[0].nodeType === 3) return this.childNodes[0].data;
      let s = '';
      const walk = (e) => { for (const n of e.childNodes) s += n.nodeType === 3 ? n.data : (walk(n), ''); };
      walk(this);
      return s;
    }
    set textContent(v) { this.children = []; this.childNodes = [new Txt(String(v), this)]; }
    get firstElementChild() { return this.children[0] || null; }
    get firstChild() { return this.childNodes[0] || null; }
    get parentElement() { return this.parentNode && this.parentNode.nodeType === 1 ? this.parentNode : null; }
    get documentElement() { return this; }
    getElementsByTagName(name) {
      const out = [];
      const any = name === '*';
      const walk = (e) => { for (const c of e.children) { if (any || c.nodeName === name) out.push(c); if (c.children.length) walk(c); } };
      walk(this);
      return out;
    }
    getElementsByTagNameNS(ns, local) {
      const out = [];
      const walk = (e) => { for (const c of e.children) { if ((local === '*' || c.localName === local) && (ns === '*' || c.namespaceURI === ns)) out.push(c); if (c.children.length) walk(c); } };
      walk(this);
      return out;
    }
    querySelector() { return null; }
    appendChild(n) { if (n.parentNode) n.remove(); n.parentNode = this; this.childNodes.push(n); if (n.nodeType === 1) this.children.push(n); return n; }
    remove() {
      const p = this.parentNode;
      if (!p) return;
      let i = p.childNodes.indexOf(this); if (i >= 0) p.childNodes.splice(i, 1);
      i = p.children.indexOf(this); if (i >= 0) p.children.splice(i, 1);
      this.parentNode = null;
    }
    replaceWith(...nodes) {
      const p = this.parentNode;
      if (!p) return;
      nodes = nodes.slice();
      for (const n of nodes) if (n.parentNode) n.remove();
      let i = p.childNodes.indexOf(this);
      p.childNodes.splice(i, 1, ...nodes);
      for (const n of nodes) n.parentNode = p;
      p.children = p.childNodes.filter((n) => n.nodeType === 1);
      this.parentNode = null;
    }
    cloneNode(deep) {
      const e = new El(this.nodeName, null);
      if (this.attrs) { e.attrs = Object.create(null); Object.assign(e.attrs, this.attrs); }
      if (deep) for (const n of this.childNodes) e.appendChild(n.nodeType === 3 ? new Txt(n.data, e) : n.cloneNode(true));
      return e;
    }
  }
  XML.El = El; XML.Txt = Txt;

  const reAttr = /([^\s=\/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  function parseAttrs(s, el) {
    reAttr.lastIndex = 0;
    let m;
    while ((m = reAttr.exec(s))) {
      if (!el.attrs) el.attrs = Object.create(null);
      const v = m[2] !== undefined ? m[2] : m[3];
      el.attrs[m[1]] = v.indexOf('&') < 0 ? v : unescape(v);
    }
  }
  XML.attrs = function (s) {
    const o = Object.create(null);
    reAttr.lastIndex = 0;
    let m;
    while ((m = reAttr.exec(s))) { const v = m[2] !== undefined ? m[2] : m[3]; o[m[1]] = v.indexOf('&') < 0 ? v : unescape(v); }
    return o;
  };

  /** parse a document; returns the document element (with a .doc wrapper parent) */
  XML.parse = function (text, opts) {
    if (text && text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    const keepWs = !!(opts && opts.keepWhitespace);
    const doc = new El('#document', null);
    let cur = doc;
    let i = 0;
    const n = text.length;
    while (i < n) {
      const lt = text.indexOf('<', i);
      if (lt < 0) { addText(text.slice(i)); break; }
      if (lt > i) addText(text.slice(i, lt));
      const c = text.charCodeAt(lt + 1);
      if (c === 47 /* / */) {
        const gt = text.indexOf('>', lt);
        if (gt < 0) throw new Error('XML parse error: unterminated end tag');
        const name = text.slice(lt + 2, gt).trim();
        /* tolerate mismatches: pop to the matching open element */
        let e = cur;
        while (e && e !== doc && e.nodeName !== name) e = e.parentNode;
        if (e && e !== doc) cur = e.parentNode;
        i = gt + 1;
      } else if (c === 33 /* ! */) {
        if (text.startsWith('<!--', lt)) { const e = text.indexOf('-->', lt + 4); i = e < 0 ? n : e + 3; }
        else if (text.startsWith('<![CDATA[', lt)) {
          const e = text.indexOf(']]>', lt + 9);
          const data = text.slice(lt + 9, e < 0 ? n : e);
          if (cur !== doc) cur.childNodes.push(new Txt(data, cur));
          i = e < 0 ? n : e + 3;
        } else {
          /* DOCTYPE (with a possible internal subset) */
          let depth = 0, j = lt + 2;
          for (; j < n; j++) { const ch = text[j]; if (ch === '[') depth++; else if (ch === ']') depth--; else if (ch === '>' && depth <= 0) break; }
          i = j + 1;
        }
      } else if (c === 63 /* ? */) {
        const e = text.indexOf('?>', lt + 2);
        i = e < 0 ? n : e + 2;
      } else {
        /* start tag: find the closing '>' outside quotes */
        let j = lt + 1, q = 0;
        for (; j < n; j++) {
          const ch = text.charCodeAt(j);
          if (q) { if (ch === q) q = 0; }
          else if (ch === 34 || ch === 39) q = ch;
          else if (ch === 62) break;
        }
        if (j >= n) throw new Error('XML parse error: unterminated start tag');
        let k = lt + 1;
        while (k < j) { const ch = text.charCodeAt(k); if (ch === 32 || ch === 9 || ch === 10 || ch === 13 || ch === 47 || ch === 62) break; k++; }
        const name = text.slice(lt + 1, k);
        const self = text.charCodeAt(j - 1) === 47;
        const el = new El(name, cur);
        const attrEnd = self ? j - 1 : j;
        if (k < attrEnd) parseAttrs(text.slice(k, attrEnd), el);
        cur.childNodes.push(el);
        cur.children.push(el);
        if (!self) cur = el;
        i = j + 1;
      }
    }
    function addText(s) {
      if (cur === doc) return;
      if (!keepWs && /^\s*$/.test(s) && !preserve(cur)) return;
      cur.childNodes.push(new Txt(unescape(s), cur));
    }
    function preserve(e) {
      /* whitespace-only text matters inside text-bearing elements (t, v, f, a:t ...) */
      const ln = e.localName;
      return ln === 't' || ln === 'v' || ln === 'f' || ln === 'formula' || ln === 'formula1' || ln === 'formula2' || ln === 'oddHeader' || ln === 'oddFooter' || ln === 'evenHeader' || ln === 'evenFooter' || ln === 'firstHeader' || ln === 'firstFooter' || ln === 'text' || ln === 'Data' || ln === 'Font' || ln === 'B' || ln === 'I' || ln === 'U' || ln === 'S' || ln === 'Sup' || ln === 'Sub';
    }
    const de = doc.children[0];
    if (!de) throw new Error('XML parse error: no root element');
    return de;
  };

  /**
   * Scan tags between [from, to): fn(local, kind, attrText, textBefore) where kind is 1 open, 2 close, 3 self-closing.
   * textBefore is the raw (escaped) text since the previous tag.
   */
  XML.scan = function (text, from, to, fn) {
    const re = /<(\/?)(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)([^>]*?)(\/?)>/g;
    re.lastIndex = from;
    let last = from, m;
    while ((m = re.exec(text)) && m.index < to) {
      const txt = m.index > last ? text.slice(last, m.index) : '';
      fn(m[2], m[1] ? 2 : m[4] ? 3 : 1, m[3], txt);
      last = re.lastIndex;
    }
  };

  /** bytes → string honouring a UTF-16 byte-order mark */
  XML.decode = function (u8) {
    if (u8.length >= 2 && u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8.subarray(2));
    if (u8.length >= 2 && u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8.subarray(2));
    return new TextDecoder('utf-8').decode(u8);
  };
  /* ------------------------------------------------------------ serialisation (kept fragments) */
  /** well-known prefixes, the safety net when a fragment was cut from the element that declared them */
  const KNOWN_NS = {
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main', r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006', a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
    a16: 'http://schemas.microsoft.com/office/drawing/2014/main', x14: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    xm: 'http://schemas.microsoft.com/office/excel/2006/main', x14ac: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
    c14: 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart', c16: 'http://schemas.microsoft.com/office/drawing/2014/chart',
    v: 'urn:schemas-microsoft-com:vml', o: 'urn:schemas-microsoft-com:office:office', x: 'urn:schemas-microsoft-com:office:excel',
  };
  XML.serialize = function (el) {
    const out = [];
    const w = (e) => {
      if (e.nodeType === 3) { out.push(e.data.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')); return; }
      out.push('<' + e.nodeName);
      if (e.attrs) for (const k in e.attrs) out.push(' ' + k + '="' + XML.esc(e.attrs[k]) + '"');
      /* carry namespace declarations from ancestors that the fragment uses */
      if (e === el) {
        const used = new Set();
        const collect = (x) => {
          if (x.nodeType !== 1) return;
          if (x.prefix) used.add(x.prefix);
          if (x.attrs) for (const k in x.attrs) {
            const i = k.indexOf(':');
            if (i > 0 && k.slice(0, i) !== 'xmlns') used.add(k.slice(0, i));
            /* mc:Ignorable / mc:Requires name prefixes in their values */
            if (/^(?:\w+:)?(Ignorable|Requires|ProcessContent|MustUnderstand)$/.test(k)) for (const p of x.attrs[k].split(/\s+/)) if (p && !p.includes(':')) used.add(p);
          }
          x.children.forEach(collect);
        };
        collect(e);
        for (const p of used) {
          if (p === 'xml') continue;
          if (e.attrs && e.attrs['xmlns:' + p] !== undefined) continue;
          const ns = e.lookupNamespaceURI(p) || KNOWN_NS[p];
          if (ns) out.push(' xmlns:' + p + '="' + XML.esc(ns) + '"');
        }
        if (!e.prefix && !(e.attrs && e.attrs.xmlns !== undefined)) { const ns = e.lookupNamespaceURI(null); if (ns) out.push(' xmlns="' + XML.esc(ns) + '"'); }
      }
      if (!e.childNodes.length) { out.push('/>'); return; }
      out.push('>');
      e.childNodes.forEach(w);
      out.push('</' + e.nodeName + '>');
    };
    w(el);
    return out.join('');
  };
})(typeof window !== 'undefined' ? window : globalThis);
