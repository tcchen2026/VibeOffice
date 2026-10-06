/* Quire — Office Math (OMML) to MathML, so equations from Word render as real two-dimensional math.
 * Covers the structures Word's equation tools produce: fractions, scripts, radicals, delimiters,
 * n-ary operators, functions, accents, bars, group characters, limits, matrices, equation arrays,
 * boxes, phantoms and pre-scripts. Rendering relies on the browser's MathML Core support. */
(function () {
  'use strict';
  const L = window.L;
  const MNS = 'http://www.w3.org/1998/Math/MathML';
  const M = (L.omml = {});

  const lname = (n) => n.localName || n.nodeName.replace(/^.*:/, '');
  const kids = (n, name) => Array.from(n.childNodes).filter((c) => c.nodeType === 1 && (!name || lname(c) === name));
  const kid = (n, name) => (n ? kids(n, name)[0] || null : null);
  const val = (n) => {
    if (!n) return null;
    for (const a of Array.from(n.attributes || [])) if (a.localName === 'val') return a.value;
    return null;
  };
  const prop = (n, pr, name) => val(kid(kid(n, pr), name));
  const onOff = (v) => v != null && v !== '0' && v !== 'false' && v !== 'off';

  function mk(tag, text, attrs) {
    const e = document.createElementNS(MNS, tag);
    if (text != null) e.textContent = text;
    if (attrs) for (const k in attrs) if (attrs[k] != null) e.setAttribute(k, attrs[k]);
    return e;
  }
  function row(nodes) {
    const list = nodes.filter(Boolean);
    if (list.length === 1) return list[0];
    const r = mk('mrow');
    for (const n of list) r.appendChild(n);
    return r;
  }

  /* ---------- styled alphabets (m:scr) ---------- */
  const HOLES = {
    'double-struck': { C: 0x2102, H: 0x210D, N: 0x2115, P: 0x2119, Q: 0x211A, R: 0x211D, Z: 0x2124 },
    script: { B: 0x212C, E: 0x2130, F: 0x2131, H: 0x210B, I: 0x2110, L: 0x2112, M: 0x2133, R: 0x211B, e: 0x212F, g: 0x210A, o: 0x2134 },
    fraktur: { C: 0x212D, H: 0x210C, I: 0x2111, R: 0x211C, Z: 0x2128 },
  };
  const BASE = { 'double-struck': 0x1D538, script: 0x1D49C, fraktur: 0x1D504, 'sans-serif': 0x1D5A0, monospace: 0x1D670 };
  function styled(ch, scr) {
    const b = BASE[scr];
    if (!b) return ch;
    const h = HOLES[scr] && HOLES[scr][ch];
    if (h) return String.fromCodePoint(h);
    const c = ch.charCodeAt(0);
    if (c >= 65 && c <= 90) return String.fromCodePoint(b + c - 65);
    if (c >= 97 && c <= 122) return String.fromCodePoint(b + 26 + c - 97);
    if (c >= 48 && c <= 57 && (scr === 'double-struck' || scr === 'sans-serif' || scr === 'monospace')) {
      return String.fromCodePoint({ 'double-struck': 0x1D7D8, 'sans-serif': 0x1D7E2, monospace: 0x1D7F6 }[scr] + c - 48);
    }
    return ch;
  }

  /* ---------- runs ---------- */
  const LETTER = /\p{L}|\p{M}/u;
  const BRACKETS = /[()[\]{}⟨⟩|‖⌈⌉⌊⌋〈〉⟦⟧]/;
  function runNodes(r) {
    const rp = kid(r, 'rPr');
    const sty = prop(r, 'rPr', 'sty');
    const nor = rp && kid(rp, 'nor') && onOff(val(kid(rp, 'nor')) == null ? '1' : val(kid(rp, 'nor')));
    const scr = prop(r, 'rPr', 'scr');
    let text = '';
    for (const c of kids(r)) {
      const n = lname(c);
      if (n === 't') text += c.textContent;
      else if (n === 'br') text += ' ';
    }
    if (!text) return [];
    const bold = sty === 'b' || sty === 'bi';
    const upright = sty === 'p' || sty === 'b';
    if (nor) {
      const t = mk('mtext', text);
      if (bold) t.setAttribute('style', 'font-weight:bold');
      return [t];
    }
    const out = [];
    const chars = Array.from(text);
    let i = 0;
    while (i < chars.length) {
      const ch = chars[i];
      if (/\s/.test(ch)) { out.push(mk('mspace', null, { width: '0.25em' })); i++; continue; }
      if (/[0-9]/.test(ch)) {
        let s = ch; i++;
        while (i < chars.length && (/[0-9]/.test(chars[i]) || (chars[i] === '.' && /[0-9]/.test(chars[i + 1] || '')))) s += chars[i++];
        out.push(mk('mn', scr ? Array.from(s).map((c) => styled(c, scr)).join('') : s));
        continue;
      }
      if (LETTER.test(ch) || ch === '∞' || ch === '∂' || ch === '∅' || ch === 'ℏ') {
        let s = ch; i++;
        while (i < chars.length && (LETTER.test(chars[i]))) s += chars[i++];
        if (scr) s = Array.from(s).map((c) => styled(c, scr)).join('');
        if (upright || scr) out.push(mk('mi', s, { mathvariant: 'normal' }));
        else for (const c of Array.from(s)) out.push(mk('mi', c));
        continue;
      }
      const mo = mk('mo', ch === '-' && !upright ? '\u2212' : ch === '*' ? '\u2217' : ch);
      /* brackets typed as plain characters take no operator spacing (their form is not inferable here) */
      if (BRACKETS.test(ch)) { mo.setAttribute('lspace', '0'); mo.setAttribute('rspace', '0'); mo.setAttribute('stretchy', 'false'); }
      else if (ch === ',' || ch === ';') { mo.setAttribute('lspace', '0'); mo.setAttribute('rspace', '0.1667em'); }
      out.push(mo);
      i++;
    }
    if (bold) for (const n of out) n.setAttribute('style', 'font-weight:bold');
    return out;
  }

  /* ---------- structures ---------- */
  const ACCENT = { '̀': '`', '́': '´', '̂': 'ˆ', '̃': '˜', '̄': '¯', '̅': '¯', '̆': '˘', '̇': '˙', '̈': '¨', '̌': 'ˇ', '⃗': '→', '⃖': '←', '⃡': '↔', '⃐': '↼', '⃑': '⇀', '̑': '⌢', '̿': '‗' };
  const INTEGRALS = /[∫∬∭∮∯∰∱∲∳⨌]/;

  /** convert the children of an argument element (m:e, m:num, …) */
  function arg(n) { return n ? row(seq(n)) : mk('mrow'); }
  function seq(n) {
    const out = [];
    for (const c of kids(n)) {
      const x = conv(c);
      if (Array.isArray(x)) out.push(...x); else if (x) out.push(x);
    }
    return out;
  }
  function conv(c) {
    switch (lname(c)) {
      case 'r': return runNodes(c);
      case 'f': {
        const type = prop(c, 'fPr', 'type') || 'bar';
        const num = arg(kid(c, 'num')), den = arg(kid(c, 'den'));
        if (type === 'lin' || type === 'skw') return row([num, mk('mo', '/'), den]);
        const f = mk('mfrac');
        if (type === 'noBar') f.setAttribute('linethickness', '0');
        f.append(num, den);
        return f;
      }
      case 'sSup': { const e = mk('msup'); e.append(arg(kid(c, 'e')), arg(kid(c, 'sup'))); return e; }
      case 'sSub': { const e = mk('msub'); e.append(arg(kid(c, 'e')), arg(kid(c, 'sub'))); return e; }
      case 'sSubSup': { const e = mk('msubsup'); e.append(arg(kid(c, 'e')), arg(kid(c, 'sub')), arg(kid(c, 'sup'))); return e; }
      case 'sPre': {
        const e = mk('mmultiscripts');
        e.append(arg(kid(c, 'e')), mk('mprescripts'), arg(kid(c, 'sub')), arg(kid(c, 'sup')));
        return e;
      }
      case 'rad': {
        const hide = onOff(prop(c, 'radPr', 'degHide'));
        const deg = kid(c, 'deg');
        const e = arg(kid(c, 'e'));
        if (hide || !deg || !deg.textContent.trim()) { const s = mk('msqrt'); s.appendChild(e); return s; }
        const r = mk('mroot'); r.append(e, arg(deg)); return r;
      }
      case 'd': {
        const pr = kid(c, 'dPr');
        const g = (name, dflt) => { const v = val(kid(pr, name)); return v == null ? dflt : v; };
        const beg = g('begChr', '('), end = g('endChr', ')'), sep = g('sepChr', '|');
        const parts = [];
        const grow = val(kid(pr, 'grow'));
        const args = kids(c, 'e').map(arg);
        /* grow only around tall content; a stretched variant of a plain bracket is too wide */
        const tall = args.some((a) => a.querySelector && (/^(mfrac|munderover|munder|mover|mtable)$/.test(a.localName) || a.querySelector('mfrac, munderover, munder, mtable')));
        const st = (grow != null && !onOff(grow)) || !tall ? 'false' : 'true';
        if (beg) parts.push(mk('mo', beg, { fence: 'true', form: 'prefix', stretchy: st, lspace: '0', rspace: '0', symmetric: 'true' }));
        args.forEach((a, i) => { if (i && sep) parts.push(mk('mo', sep, { separator: 'true', stretchy: st, lspace: '0', rspace: '0.1667em' })); parts.push(a); });
        if (end) parts.push(mk('mo', end, { fence: 'true', form: 'postfix', stretchy: st, lspace: '0', rspace: '0', symmetric: 'true' }));
        const r = mk('mrow'); for (const p of parts) r.appendChild(p); return r;
      }
      case 'nary': {
        const pr = kid(c, 'naryPr');
        const chr = val(kid(pr, 'chr')) || '∫';
        const subHide = onOff(val(kid(pr, 'subHide'))), supHide = onOff(val(kid(pr, 'supHide')));
        const loc = val(kid(pr, 'limLoc')) || (INTEGRALS.test(chr) ? 'subSup' : 'undOvr');
        const op = mk('mo', chr, { largeop: 'true', movablelimits: 'false' });
        const sub = subHide ? null : kid(c, 'sub'), sup = supHide ? null : kid(c, 'sup');
        const filled = (n) => !!n && kids(n).some((k) => lname(k) !== 'ctrlPr' && lname(k) !== 'argPr');
        const hasSub = filled(sub), hasSup = filled(sup);
        let base = op;
        const under = loc === 'undOvr';
        if (hasSub && hasSup) { base = mk(under ? 'munderover' : 'msubsup'); base.append(op, arg(sub), arg(sup)); }
        else if (hasSub) { base = mk(under ? 'munder' : 'msub'); base.append(op, arg(sub)); }
        else if (hasSup) { base = mk(under ? 'mover' : 'msup'); base.append(op, arg(sup)); }
        return row([base, arg(kid(c, 'e'))]);
      }
      case 'func': {
        const name = kid(c, 'fName');
        const fn = name ? row(seq(name).map((n) => { if (n.localName === 'mi' && n.textContent.length === 1) n.setAttribute('mathvariant', 'normal'); return n; })) : null;
        /* function names are upright: merge consecutive single letters into one identifier */
        if (fn && fn.localName === 'mrow') {
          const txt = Array.from(fn.childNodes).every((n) => n.localName === 'mi') ? fn.textContent : null;
          if (txt) return row([mk('mi', txt), mk('mo', '⁡'), arg(kid(c, 'e'))]);
        }
        return row([fn, mk('mo', '⁡'), arg(kid(c, 'e'))]);
      }
      case 'acc': {
        const ch = prop(c, 'accPr', 'chr') || '̂';
        const e = mk('mover', null, { accent: 'true' });
        e.append(arg(kid(c, 'e')), mk('mo', ACCENT[ch] || ch, { stretchy: /[→←↔¯]/.test(ACCENT[ch] || ch) ? 'true' : 'false' }));
        return e;
      }
      case 'bar': {
        const top = prop(c, 'barPr', 'pos') === 'top';
        const e = mk(top ? 'mover' : 'munder', null, top ? { accent: 'true' } : { accentunder: 'true' });
        e.append(arg(kid(c, 'e')), mk('mo', top ? '¯' : '_', { stretchy: 'true' }));
        return e;
      }
      case 'groupChr': {
        const pr = kid(c, 'groupChrPr');
        const chr = val(kid(pr, 'chr')) || '⏟';
        const top = val(kid(pr, 'pos')) === 'top';
        const e = mk(top ? 'mover' : 'munder');
        e.append(arg(kid(c, 'e')), mk('mo', chr, { stretchy: 'true' }));
        return e;
      }
      case 'limLow': { const e = mk('munder'); e.append(arg(kid(c, 'e')), arg(kid(c, 'lim'))); return e; }
      case 'limUpp': { const e = mk('mover'); e.append(arg(kid(c, 'e')), arg(kid(c, 'lim'))); return e; }
      case 'm': {
        const t = mk('mtable');
        for (const mr of kids(c, 'mr')) {
          const tr = mk('mtr');
          for (const e of kids(mr, 'e')) { const td = mk('mtd'); td.appendChild(arg(e)); tr.appendChild(td); }
          t.appendChild(tr);
        }
        return t;
      }
      case 'eqArr': {
        const t = mk('mtable', null, { columnalign: 'left' });
        for (const e of kids(c, 'e')) { const tr = mk('mtr'); const td = mk('mtd'); td.appendChild(arg(e)); tr.appendChild(td); t.appendChild(tr); }
        return t;
      }
      case 'box': return arg(kid(c, 'e'));
      case 'borderBox': { const r = mk('mrow', null, { style: 'border:1px solid currentColor;padding:1px 2px' }); r.appendChild(arg(kid(c, 'e'))); return r; }
      case 'phant': {
        const show = prop(c, 'phantPr', 'show');
        if (show != null && !onOff(show)) { const p = mk('mphantom'); p.appendChild(arg(kid(c, 'e'))); return p; }
        return arg(kid(c, 'e'));
      }
      case 'oMath': return seq(c);
      /* wrappers that carry runs: revisions, content controls, smart tags */
      case 'ins': case 'sdt': case 'sdtContent': case 'customXml': case 'smartTag': case 'moveTo': return seq(c);
      case 'del': case 'moveFrom': return null;
      default: return null;
    }
  }

  /** first explicit font size (pt) in the equation, from w:rPr/w:sz inside its runs */
  M.fontSize = function (node) {
    const all = node.getElementsByTagNameNS ? node.getElementsByTagNameNS('*', 'sz') : [];
    for (const s of Array.from(all)) {
      const p = s.parentNode;
      if (p && lname(p) === 'rPr' && p.parentNode && lname(p.parentNode) === 'r') { const v = +val(s); if (v) return v / 2; }
    }
    return null;
  };

  /**
   * Build a <math> element from an m:oMath / m:oMathPara element or its serialized XML.
   * Returns {math, display: bool, jc} or null when the input cannot be parsed.
   */
  M.toMathML = function (src) {
    let node = src;
    if (typeof src === 'string') {
      const xml = /xmlns:m=/.test(src) ? src : src.replace(/^<m:(oMath\w*)/, '<m:$1 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      if (doc.getElementsByTagName('parsererror').length) return null;
      node = doc.documentElement;
    }
    if (!node) return null;
    const isPara = lname(node) === 'oMathPara';
    const math = mk('math', null, isPara ? { display: 'block' } : null);
    if (isPara) {
      const maths = kids(node, 'oMath');
      if (maths.length > 1) {
        /* several equations in one paragraph: one per line */
        const t = mk('mtable');
        for (const m of maths) { const tr = mk('mtr'); const td = mk('mtd'); td.appendChild(row(seq(m))); tr.appendChild(td); t.appendChild(tr); }
        math.appendChild(t);
      } else if (maths[0]) for (const n of seq(maths[0])) math.appendChild(n);
    } else for (const n of seq(node)) math.appendChild(n);
    const jc = isPara ? prop(node, 'oMathParaPr', 'jc') || 'centerGroup' : null;
    return { math, display: isPara, jc };
  };
})();
